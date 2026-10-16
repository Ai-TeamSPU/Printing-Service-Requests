"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { loadConfig } = require("../server.js");

const ENV_KEYS = [
  "NODE_ENV", "GOOGLE_CLIENT_ID", "ALLOWED_GOOGLE_DOMAIN", "ALLOWED_ORIGINS",
  "SESSION_SECRET", "SESSION_TTL_SECONDS", "COOKIE_SECURE", "COOKIE_SAME_SITE",
  "APPS_SCRIPT_URL", "APPS_SCRIPT_CONNECT_TOKEN", "APPS_SCRIPT_GATEWAY_TOKEN",
  "REQUIRE_PERSONNEL", "PORT"
];

function withEnv(values, callback) {
  const previous = {};
  ENV_KEYS.forEach((key) => {
    previous[key] = process.env[key];
    delete process.env[key];
  });
  Object.entries(values).forEach(([key, value]) => { process.env[key] = value; });
  try { return callback(); }
  finally {
    ENV_KEYS.forEach((key) => {
      if (previous[key] === undefined) delete process.env[key];
      else process.env[key] = previous[key];
    });
  }
}

function validEnv() {
  return {
    NODE_ENV: "production",
    GOOGLE_CLIENT_ID: "1234567890-realclient.apps.googleusercontent.com",
    ALLOWED_GOOGLE_DOMAIN: "spu.ac.th",
    ALLOWED_ORIGINS: "https://printing.spu.ac.th",
    SESSION_SECRET: "session-Qq7EueRCU4q9CwG2hPfm6ZAK8wYp",
    SESSION_TTL_SECONDS: "3600",
    COOKIE_SECURE: "true",
    COOKIE_SAME_SITE: "lax",
    APPS_SCRIPT_URL: "https://script.google.com/macros/s/abc123/exec",
    APPS_SCRIPT_CONNECT_TOKEN: "connect-WtJ5Mj6QxZ8oV4bK2nR7cFp9HsLd",
    APPS_SCRIPT_GATEWAY_TOKEN: "gateway-Rm8YvT3sN6pL2wQ5kC9xFa7JdUeB",
    REQUIRE_PERSONNEL: "false",
    PORT: "8080"
  };
}

test("production config accepts real-shaped values and rejects sample secrets", () => {
  withEnv(validEnv(), () => {
    const config = loadConfig();
    assert.equal(config.cookieSecure, true);
    assert.equal(config.allowedOrigins.has("https://printing.spu.ac.th"), true);
  });

  withEnv({ ...validEnv(), SESSION_SECRET: "replace-with-a-long-random-secret-at-least-32-characters" }, () => {
    assert.throws(() => loadConfig(), /placeholder/i);
  });

  const duplicate = validEnv();
  duplicate.APPS_SCRIPT_GATEWAY_TOKEN = duplicate.APPS_SCRIPT_CONNECT_TOKEN;
  withEnv(duplicate, () => assert.throws(() => loadConfig(), /must all be different/i));

  withEnv({ ...validEnv(), COOKIE_SECURE: "false" }, () => {
    assert.throws(() => loadConfig(), /COOKIE_SECURE must be true/i);
  });
});
