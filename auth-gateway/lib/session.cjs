"use strict";

const crypto = require("node:crypto");

function base64url(input) {
  return Buffer.from(input).toString("base64url");
}

function sign(value, secret) {
  return crypto.createHmac("sha256", secret).update(value).digest("base64url");
}

function safeEqual(left, right) {
  const a = Buffer.from(String(left));
  const b = Buffer.from(String(right));
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function issueSession(user, secret, ttlSeconds, nowSeconds = Math.floor(Date.now() / 1000)) {
  if (!secret || Buffer.byteLength(secret) < 32) throw new Error("SESSION_SECRET must be at least 32 bytes");
  if (!user || !user.email || !user.role) throw new Error("Session user is incomplete");
  const payload = {
    v: 1,
    sub: String(user.sub || ""),
    email: String(user.email).trim().toLowerCase(),
    name: String(user.name || ""),
    picture: String(user.picture || ""),
    role: String(user.role).toUpperCase(),
    iat: nowSeconds,
    exp: nowSeconds + ttlSeconds
  };
  const encoded = base64url(JSON.stringify(payload));
  return `${encoded}.${sign(encoded, secret)}`;
}

function verifySession(token, secret, nowSeconds = Math.floor(Date.now() / 1000)) {
  if (!token || !secret) return null;
  const parts = String(token).split(".");
  if (parts.length !== 2 || !safeEqual(sign(parts[0], secret), parts[1])) return null;
  try {
    const payload = JSON.parse(Buffer.from(parts[0], "base64url").toString("utf8"));
    if (payload.v !== 1 || !payload.email || !payload.role || !Number.isFinite(payload.exp) || payload.exp <= nowSeconds) return null;
    return payload;
  } catch (error) {
    return null;
  }
}

function parseCookies(header) {
  const out = {};
  String(header || "").split(";").forEach((part) => {
    const index = part.indexOf("=");
    if (index < 1) return;
    const key = part.slice(0, index).trim();
    try { out[key] = decodeURIComponent(part.slice(index + 1).trim()); }
    catch (error) { out[key] = ""; }
  });
  return out;
}

module.exports = { issueSession, verifySession, parseCookies };
