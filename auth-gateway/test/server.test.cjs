"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { createApp } = require("../server.js");

function config(overrides = {}) {
  return {
    googleClientId: "client.apps.googleusercontent.com",
    allowedDomain: "spu.ac.th",
    allowedOrigins: new Set(["http://frontend.test"]),
    sessionSecret: "0123456789abcdef0123456789abcdef",
    sessionTtlSeconds: 3600,
    cookieSecure: false,
    cookieSameSite: "lax",
    cookieName: "spu_session",
    requirePersonnel: false,
    appsScriptUrl: "https://example.invalid/exec",
    connectToken: "connect",
    gatewayToken: "gateway",
    ...overrides
  };
}

async function withServer(callback, configOverrides = {}) {
  const calls = [];
  const googleClient = {
    async verifyIdToken({ idToken, audience }) {
      assert.equal(audience, "client.apps.googleusercontent.com");
      return {
        getPayload: () => idToken === "wrong-domain"
          ? { sub: "evil", email: "user@evil.test", email_verified: true, hd: "evil.test" }
          : { sub: "google-123", email: idToken === "missing-personnel" ? "profileless@spu.ac.th" : "User@SPU.AC.TH", email_verified: true, hd: "spu.ac.th", name: "SPU User" }
      };
    }
  };
  const callAppsScript = async (action, input, actor) => {
    calls.push({ action, input, actor });
    if (action === "lookupPersonnel") return actor.email === "profileless@spu.ac.th" ? { ok: false, error: "personnel_sheet_not_found" } : { ok: true, found: true, name: "SPU User", unit: "Unit" };
    if (action === "checkAdminUser") return input.password === "wrong" ? { ok: false, error: "user_not_found", message: "detail" } : { ok: true, isAdmin: true, email: input.email, name: "Admin", role: "ADMIN" };
    if (input && input.simulateRevoked) return { ok: false, error: "staff_session_revoked" };
    if (action === "listJobs") return { ok: true, jobs: [], items: [] };
    return { ok: true };
  };
  const server = createApp(config(configOverrides), { googleClient, callAppsScript }).listen(0);
  await new Promise((resolve) => server.once("listening", resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  try { await callback({ base, calls }); }
  finally { await new Promise((resolve) => server.close(resolve)); }
}

async function jsonRequest(url, options = {}) {
  const response = await fetch(url, options);
  return { response, body: await response.json() };
}

test("verified Workspace login creates an HttpOnly session and API ignores spoofed identity", async () => {
  await withServer(async ({ base, calls }) => {
    const login = await jsonRequest(`${base}/auth/google`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: JSON.stringify({ credential: "valid" })
    });
    assert.equal(login.response.status, 200);
    assert.equal(login.body.user.email, "user@spu.ac.th");
    const setCookie = login.response.headers.get("set-cookie");
    assert.match(setCookie, /spu_session=/);
    assert.match(setCookie, /HttpOnly/i);
    const cookie = setCookie.split(";")[0];

    const restored = await jsonRequest(`${base}/auth/session`, { headers: { Cookie: cookie } });
    assert.equal(restored.response.status, 200);
    assert.equal(restored.body.authenticated, true);
    assert.equal(restored.body.personnel.unit, "Unit");

    const api = await jsonRequest(`${base}/api`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json", Cookie: cookie },
      body: JSON.stringify({
        action: "listJobs", email: "victim@spu.ac.th", role: "ADMIN", actorSub: "spoof",
        token: "spoof", gatewayToken: "spoof", spreadsheetId: "spoof", folderId: "spoof"
      })
    });
    assert.equal(api.response.status, 200);
    const proxied = calls.find((call) => call.action === "listJobs");
    assert.deepEqual(proxied.input, {});
    assert.equal(proxied.actor.email, "user@spu.ac.th");
    assert.equal(proxied.actor.role, "REQUESTER");
    assert.equal(proxied.actor.sub, "google-123");

    const forbidden = await jsonRequest(`${base}/api`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json", Cookie: cookie },
      body: JSON.stringify({ action: "togglePriceRule", ruleId: "R01" })
    });
    assert.equal(forbidden.response.status, 403);
    assert.equal(forbidden.body.error, "action_forbidden");
  });
});

test("wrong Workspace domain and untrusted browser origin are rejected", async () => {
  await withServer(async ({ base }) => {
    const domain = await jsonRequest(`${base}/auth/google`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: JSON.stringify({ credential: "wrong-domain" })
    });
    assert.equal(domain.response.status, 403);
    assert.equal(domain.body.error, "workspace_account_required");

    const origin = await jsonRequest(`${base}/auth/google`, {
      method: "POST",
      headers: { Origin: "https://attacker.test", "Content-Type": "application/json" },
      body: JSON.stringify({ credential: "valid" })
    });
    assert.equal(origin.response.status, 403);
    assert.equal(origin.body.error, "origin_forbidden");

    const missingOrigin = await jsonRequest(`${base}/auth/google`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ credential: "valid" })
    });
    assert.equal(missingOrigin.response.status, 403);
    assert.equal(missingOrigin.body.error, "origin_forbidden");
  });
});

test("optional personnel lookup does not become an authentication dependency", async () => {
  await withServer(async ({ base }) => {
    const optional = await jsonRequest(`${base}/auth/google`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: JSON.stringify({ credential: "missing-personnel" })
    });
    assert.equal(optional.response.status, 200);
    assert.equal(optional.body.user.email, "profileless@spu.ac.th");
    assert.equal(optional.body.personnel.found, false);
    assert.equal(optional.body.personnel.warning, "personnel_lookup_unavailable");
  });
});

test("admin authentication is domain-bound and does not enumerate accounts", async () => {
  await withServer(async ({ base, calls }) => {
    const external = await jsonRequest(`${base}/auth/admin`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: JSON.stringify({ email: "admin@outside.test", password: "secret" })
    });
    assert.equal(external.response.status, 401);
    assert.deepEqual(external.body, { ok: false, error: "invalid_credentials" });
    assert.equal(calls.some((call) => call.action === "checkAdminUser"), false);

    const unknown = await jsonRequest(`${base}/auth/admin`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: JSON.stringify({ email: "unknown@spu.ac.th", password: "wrong" })
    });
    assert.equal(unknown.response.status, 401);
    assert.deepEqual(unknown.body, { ok: false, error: "invalid_credentials" });
  });
});

test("revoked staff sessions are cleared and malformed requests fail safely", async () => {
  await withServer(async ({ base }) => {
    const login = await jsonRequest(`${base}/auth/admin`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: JSON.stringify({ email: "admin@spu.ac.th", password: "secret" })
    });
    const cookie = login.response.headers.get("set-cookie").split(";")[0];

    const revoked = await jsonRequest(`${base}/api`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json", Cookie: cookie },
      body: JSON.stringify({ action: "listJobs", simulateRevoked: true })
    });
    assert.equal(revoked.response.status, 401);
    assert.equal(revoked.body.error, "session_revoked");
    assert.match(revoked.response.headers.get("set-cookie"), /spu_session=;/);

    const invalidJson = await jsonRequest(`${base}/auth/google`, {
      method: "POST",
      headers: { Origin: "http://frontend.test", "Content-Type": "application/json" },
      body: "{"
    });
    assert.equal(invalidJson.response.status, 400);
    assert.equal(invalidJson.body.error, "invalid_json");

    const missing = await jsonRequest(`${base}/missing`);
    assert.equal(missing.response.status, 404);
    assert.equal(missing.body.error, "not_found");
  });
});
