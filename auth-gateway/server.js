"use strict";

const express = require("express");
const { OAuth2Client } = require("google-auth-library");
const { issueSession, verifySession, parseCookies } = require("./lib/session.cjs");
const { authorizeAction, sanitizeClientPayload } = require("./lib/policy.cjs");

function envRequired(name) {
  const value = String(process.env[name] || "").trim();
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

function envSecret(name, minBytes = 32) {
  const value = envRequired(name);
  if (Buffer.byteLength(value) < minBytes) throw new Error(`${name} must be at least ${minBytes} bytes`);
  if (/(?:replace|example|change[-_ ]?me|your[-_ ]?(?:secret|token|key))/i.test(value)) {
    throw new Error(`${name} still contains a placeholder value`);
  }
  return value;
}

function envBoolean(name, fallback) {
  if (process.env[name] === undefined) return fallback;
  return String(process.env[name]).trim().toLowerCase() === "true";
}

function loadConfig() {
  const cookieSecure = envBoolean("COOKIE_SECURE", process.env.NODE_ENV === "production");
  const cookieSameSite = String(process.env.COOKIE_SAME_SITE || "lax").trim().toLowerCase();
  if (!["lax", "strict", "none"].includes(cookieSameSite)) throw new Error("COOKIE_SAME_SITE must be lax, strict, or none");
  if (cookieSameSite === "none" && !cookieSecure) throw new Error("SameSite=None requires COOKIE_SECURE=true");
  if (process.env.NODE_ENV === "production" && !cookieSecure) throw new Error("COOKIE_SECURE must be true in production");
  const sessionSecret = envSecret("SESSION_SECRET");
  const ttlSeconds = Number.parseInt(process.env.SESSION_TTL_SECONDS || "28800", 10);
  if (!Number.isFinite(ttlSeconds) || ttlSeconds < 300 || ttlSeconds > 86400) throw new Error("SESSION_TTL_SECONDS must be between 300 and 86400");
  const origins = envRequired("ALLOWED_ORIGINS").split(",").map((value) => value.trim().replace(/\/$/, "")).filter(Boolean);
  if (!origins.length) throw new Error("ALLOWED_ORIGINS must contain at least one origin");
  if (origins.some((origin) => origin === "*")) throw new Error("ALLOWED_ORIGINS cannot contain * when credential cookies are enabled");
  origins.forEach((origin) => {
    let parsed;
    try { parsed = new URL(origin); }
    catch (error) { throw new Error(`Invalid ALLOWED_ORIGINS entry: ${origin}`); }
    if (!["http:", "https:"].includes(parsed.protocol) || parsed.origin !== origin) throw new Error(`ALLOWED_ORIGINS entries must be exact HTTP(S) origins: ${origin}`);
    if (process.env.NODE_ENV === "production" && parsed.protocol !== "https:") throw new Error("ALLOWED_ORIGINS must use HTTPS in production");
  });
  const googleClientId = envRequired("GOOGLE_CLIENT_ID");
  if (!/^[A-Za-z0-9._-]+\.apps\.googleusercontent\.com$/.test(googleClientId) || /example/i.test(googleClientId)) {
    throw new Error("GOOGLE_CLIENT_ID must be a real Web OAuth client ID");
  }
  const allowedDomain = String(process.env.ALLOWED_GOOGLE_DOMAIN || "spu.ac.th").trim().toLowerCase();
  if (!/^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/.test(allowedDomain)) throw new Error("ALLOWED_GOOGLE_DOMAIN is invalid");
  const appsScriptUrl = envRequired("APPS_SCRIPT_URL");
  let parsedAppsScriptUrl;
  try { parsedAppsScriptUrl = new URL(appsScriptUrl); }
  catch (error) { throw new Error("APPS_SCRIPT_URL is invalid"); }
  if (parsedAppsScriptUrl.protocol !== "https:" || parsedAppsScriptUrl.hostname !== "script.google.com" || !/\/exec\/?$/.test(parsedAppsScriptUrl.pathname) || /replace/i.test(appsScriptUrl)) {
    throw new Error("APPS_SCRIPT_URL must be an HTTPS script.google.com Web App /exec URL");
  }
  const connectToken = envSecret("APPS_SCRIPT_CONNECT_TOKEN");
  const gatewayToken = envSecret("APPS_SCRIPT_GATEWAY_TOKEN");
  if (new Set([sessionSecret, connectToken, gatewayToken]).size !== 3) throw new Error("SESSION_SECRET and Apps Script tokens must all be different");
  const port = Number.parseInt(process.env.PORT || "8080", 10);
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error("PORT must be between 1 and 65535");
  return {
    port,
    googleClientId,
    allowedDomain,
    allowedOrigins: new Set(origins),
    sessionSecret,
    sessionTtlSeconds: ttlSeconds,
    cookieSecure,
    cookieSameSite,
    cookieName: cookieSecure ? "__Host-spu_session" : "spu_session",
    requirePersonnel: envBoolean("REQUIRE_PERSONNEL", false),
    appsScriptUrl,
    connectToken,
    gatewayToken
  };
}

function publicUser(user) {
  return {
    sub: String(user.sub || ""),
    email: String(user.email || "").trim().toLowerCase(),
    name: String(user.name || ""),
    picture: String(user.picture || ""),
    role: String(user.role || "REQUESTER").toUpperCase()
  };
}

function createRateLimiter({ windowMs, max, key }) {
  const buckets = new Map();
  return function rateLimit(req, res, next) {
    const now = Date.now();
    if (buckets.size > 5000) {
      for (const [storedKey, value] of buckets) {
        if (value.resetAt <= now || buckets.size > 5000) buckets.delete(storedKey);
      }
    }
    const bucketKey = key(req);
    const current = buckets.get(bucketKey);
    if (!current || current.resetAt <= now) {
      buckets.set(bucketKey, { count: 1, resetAt: now + windowMs });
      return next();
    }
    current.count += 1;
    if (current.count > max) {
      res.set("Retry-After", String(Math.max(1, Math.ceil((current.resetAt - now) / 1000))));
      return res.status(429).json({ ok: false, error: "too_many_attempts" });
    }
    next();
  };
}

function createAppsScriptClient(config) {
  return async function callAppsScript(action, input = {}, actor = null) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);
    const payload = {
      ...input,
      action,
      token: config.connectToken,
      gatewayToken: config.gatewayToken
    };
    if (actor) {
      payload.actorEmail = String(actor.email || "").trim().toLowerCase();
      payload.actorRole = String(actor.role || "").toUpperCase();
      payload.actorSub = String(actor.sub || "");
      // Temporary aliases keep the current Apps Script functions compatible;
      // Code.gs overwrites/trusts them only after validating gatewayToken.
      payload.email = payload.actorEmail;
      payload.senderEmail = payload.actorEmail;
      payload.readerEmail = payload.actorEmail;
      payload.role = payload.actorRole;
      payload.senderRole = payload.actorRole;
      payload.readerRole = payload.actorRole;
      payload.by = payload.actorEmail;
    }
    try {
      const response = await fetch(config.appsScriptUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8", Accept: "application/json" },
        body: JSON.stringify(payload),
        signal: controller.signal,
        redirect: "follow"
      });
      const text = await response.text();
      let data;
      try { data = JSON.parse(text); }
      catch (error) { throw new Error("Apps Script returned a non-JSON response"); }
      if (!response.ok) throw new Error(`Apps Script HTTP ${response.status}`);
      return data;
    } finally {
      clearTimeout(timeout);
    }
  };
}

function createApp(config, dependencies = {}) {
  const app = express();
  const googleClient = dependencies.googleClient || new OAuth2Client(config.googleClientId);
  const callAppsScript = dependencies.callAppsScript || createAppsScriptClient(config);
  app.disable("x-powered-by");
  app.set("trust proxy", 1);

  app.use((req, res, next) => {
    console.log(`[${new Date().toLocaleTimeString()}] ${req.method} ${req.url} (Origin: ${req.get("origin") || "none"})`);
    res.set("Cache-Control", "no-store");
    res.set("X-Content-Type-Options", "nosniff");
    res.set("Referrer-Policy", "no-referrer");
    const origin = String(req.get("origin") || "").replace(/\/$/, "");
    if (origin && config.allowedOrigins.has(origin)) {
      res.set("Access-Control-Allow-Origin", origin);
      res.set("Access-Control-Allow-Credentials", "true");
      res.set("Access-Control-Allow-Headers", "Content-Type, Authorization");
      res.vary("Origin");
    }
    if (req.method === "OPTIONS") {
      if (!origin || !config.allowedOrigins.has(origin)) {
        console.warn(`[CORS OPTIONS BLOCKED] Origin: "${origin}". Allowed:`, Array.from(config.allowedOrigins));
        return res.status(403).json({ ok: false, error: "origin_forbidden", receivedOrigin: origin });
      }
      res.set("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
      res.set("Access-Control-Allow-Headers", "Content-Type, Authorization");
      return res.status(204).end();
    }
    next();
  });
  app.use(express.json({ limit: "30mb", strict: true }));

  function requireAllowedOrigin(req, res, next) {
    const origin = String(req.get("origin") || "").replace(/\/$/, "");
    if (!origin || !config.allowedOrigins.has(origin)) {
      console.warn(`[ORIGIN BLOCKED] Origin: "${origin}". Allowed:`, Array.from(config.allowedOrigins));
      return res.status(403).json({ ok: false, error: "origin_forbidden", receivedOrigin: origin });
    }
    next();
  }

  function readSession(req) {
    const authHeader = req.get("authorization");
    if (authHeader && authHeader.toLowerCase().startsWith("bearer ")) {
      const bearerToken = authHeader.slice(7).trim();
      const verified = verifySession(bearerToken, config.sessionSecret);
      if (verified) return verified;
    }
    const cookies = parseCookies(req.headers.cookie);
    return verifySession(cookies[config.cookieName], config.sessionSecret);
  }

  function requireSession(req, res, next) {
    const session = readSession(req);
    if (!session) return res.status(401).json({ ok: false, error: "auth_required" });
    req.authUser = session;
    next();
  }

  function writeSession(res, user) {
    const token = issueSession(user, config.sessionSecret, config.sessionTtlSeconds);
    res.cookie(config.cookieName, token, {
      httpOnly: true,
      secure: config.cookieSecure,
      sameSite: config.cookieSameSite,
      path: "/",
      maxAge: config.sessionTtlSeconds * 1000
    });
    return token;
  }

  function clearSession(res) {
    res.clearCookie(config.cookieName, {
      httpOnly: true,
      secure: config.cookieSecure,
      sameSite: config.cookieSameSite,
      path: "/"
    });
  }

  async function readPersonnel(user) {
    try {
      const personnel = await callAppsScript("lookupPersonnel", {}, user);
      if (personnel && personnel.ok) return personnel;
      if (config.requirePersonnel) throw new Error("Personnel lookup failed");
    } catch (error) {
      if (config.requirePersonnel) throw error;
    }
    return { ok: true, found: false, warning: "personnel_lookup_unavailable" };
  }

  app.get("/health", (req, res) => res.json({ ok: true }));
  app.get("/auth/config", (req, res) => res.json({
    ok: true,
    googleClientId: config.googleClientId,
    allowedDomain: config.allowedDomain
  }));

  app.get("/auth/session", async (req, res, next) => {
    const user = readSession(req);
    if (!user) return res.json({ ok: true, authenticated: false });
    try {
      let personnel = null;
      if (String(user.role || "").toUpperCase() === "REQUESTER") {
        personnel = await readPersonnel(user);
        if (config.requirePersonnel && !personnel.found) {
          clearSession(res);
          return res.status(403).json({ ok: false, error: "personnel_not_authorized" });
        }
      }
      return res.json({ ok: true, authenticated: true, user: publicUser(user), personnel, token: issueSession(user, config.sessionSecret, config.sessionTtlSeconds) });
    } catch (error) {
      next(error);
    }
  });

  const googleLoginLimit = createRateLimiter({ windowMs: 15 * 60 * 1000, max: 30, key: (req) => `google:${req.ip}` });
  const adminIpLoginLimit = createRateLimiter({ windowMs: 15 * 60 * 1000, max: 30, key: (req) => `admin-ip:${req.ip}` });
  const adminLoginLimit = createRateLimiter({
    windowMs: 15 * 60 * 1000,
    max: 10,
    key: (req) => `admin:${req.ip}:${String((req.body && req.body.email) || "").trim().toLowerCase()}`
  });
  const apiLimit = createRateLimiter({
    windowMs: 5 * 60 * 1000,
    max: 300,
    key: (req) => `api:${req.ip}:${String((req.authUser && (req.authUser.sub || req.authUser.email)) || "anonymous")}`
  });
  const costlyApiLimit = createRateLimiter({
    windowMs: 15 * 60 * 1000,
    max: 30,
    key: (req) => `costly:${req.ip}:${String((req.authUser && (req.authUser.sub || req.authUser.email)) || "anonymous")}`
  });
  const costlyActions = new Set(["submitJob", "uploadFile", "sendReportEmail"]);
  function rateLimitCostlyActions(req, res, next) {
    if (!costlyActions.has(String((req.body && req.body.action) || ""))) return next();
    return costlyApiLimit(req, res, next);
  }

  app.post("/auth/google", requireAllowedOrigin, googleLoginLimit, async (req, res, next) => {
    try {
      const credential = String((req.body && req.body.credential) || "");
      if (!credential || credential.length > 20000) return res.status(400).json({ ok: false, error: "invalid_google_credential" });
      const ticket = await googleClient.verifyIdToken({ idToken: credential, audience: config.googleClientId });
      const claims = ticket.getPayload() || {};
      const email = String(claims.email || "").trim().toLowerCase();
      const hostedDomain = String(claims.hd || "").trim().toLowerCase();
      if (claims.email_verified !== true || !claims.sub || hostedDomain !== config.allowedDomain || !email.endsWith(`@${config.allowedDomain}`)) {
        return res.status(403).json({ ok: false, error: "workspace_account_required", message: `Use a verified @${config.allowedDomain} Google Workspace account.` });
      }
      const user = publicUser({ sub: claims.sub, email, name: claims.name, picture: claims.picture, role: "REQUESTER" });
      const personnel = await readPersonnel(user);
      if (config.requirePersonnel && !personnel.found) return res.status(403).json({ ok: false, error: "personnel_not_authorized" });
      const token = writeSession(res, user);
      return res.json({ ok: true, user, personnel, token });
    } catch (error) {
      if (/token|audience|issuer|signature|expired/i.test(String(error && error.message))) {
        return res.status(401).json({ ok: false, error: "invalid_google_credential" });
      }
      next(error);
    }
  });

  app.post("/auth/admin", requireAllowedOrigin, adminIpLoginLimit, adminLoginLimit, async (req, res, next) => {
    try {
      const email = String((req.body && req.body.email) || "").trim().toLowerCase();
      const password = String((req.body && req.body.password) || "");
      if (!email || email.length > 320 || !email.endsWith(`@${config.allowedDomain}`) || !password || password.length > 1024) {
        return res.status(401).json({ ok: false, error: "invalid_credentials" });
      }
      const result = await callAppsScript("checkAdminUser", { email, password });
      console.log(`[Admin Login] email: "${email}", Apps Script result:`, result);
      if (!result || !result.ok || !result.isAdmin || !["ADMIN", "EXECUTIVE"].includes(String(result.role || "").toUpperCase())) {
        return res.status(401).json({ ok: false, error: "invalid_credentials" });
      }
      const user = publicUser({
        sub: `staff:${String(result.email || email).toLowerCase()}`,
        email: result.email || email,
        name: result.name || email,
        role: result.role
      });
      const token = writeSession(res, user);
      return res.json({ ok: true, user, token });
    } catch (error) {
      next(error);
    }
  });

  app.post("/auth/logout", requireAllowedOrigin, (req, res) => {
    clearSession(res);
    res.json({ ok: true });
  });

  app.post("/api", requireAllowedOrigin, requireSession, apiLimit, rateLimitCostlyActions, async (req, res, next) => {
    try {
      const action = String((req.body && req.body.action) || "");
      const policy = authorizeAction(req.authUser.role, action, req.body || {});
      if (!policy.ok) return res.status(403).json(policy);
      const clean = sanitizeClientPayload(req.body || {});
      delete clean.action;
      const result = await callAppsScript(action, clean, req.authUser);
      if (result && result.error === "staff_session_revoked") {
        clearSession(res);
        return res.status(401).json({ ok: false, error: "session_revoked" });
      }
      if (result && ["gateway_unauthorized", "gateway_token_not_set", "actor_required", "actor_sub_required"].includes(result.error)) {
        return res.status(502).json({ ok: false, error: "backend_auth_misconfigured" });
      }
      return res.json(result);
    } catch (error) {
      next(error);
    }
  });

  app.use((req, res) => res.status(404).json({ ok: false, error: "not_found" }));

  app.use((error, req, res, next) => {
    if (res.headersSent) return next(error);
    // Do not expose upstream URLs, tokens, credentials, or stack traces.
    const isBodyTooLarge = error && (error.type === "entity.too.large" || error.status === 413);
    const isInvalidJson = error && error instanceof SyntaxError && error.status === 400 && Object.prototype.hasOwnProperty.call(error, "body");
    const status = isBodyTooLarge ? 413 : (isInvalidJson ? 400 : 502);
    const code = isBodyTooLarge ? "request_too_large" : (isInvalidJson ? "invalid_json" : "gateway_error");
    res.status(status).json({ ok: false, error: code });
  });

  return app;
}

if (require.main === module) {
  const config = loadConfig();
  createApp(config).listen(config.port, () => {
    // Avoid printing secrets or upstream identifiers in logs.
    console.log(`SPU auth gateway listening on port ${config.port}`);
  });
}

module.exports = { loadConfig, createApp, createAppsScriptClient, publicUser, createRateLimiter };
