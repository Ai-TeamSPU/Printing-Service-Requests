"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { issueSession, verifySession, parseCookies } = require("../lib/session.cjs");

const secret = "0123456789abcdef0123456789abcdef";

test("signed session round-trips and normalizes identity", () => {
  const token = issueSession({ sub: "google-sub", email: "USER@SPU.AC.TH", name: "User", role: "requester" }, secret, 600, 1000);
  assert.deepEqual(verifySession(token, secret, 1100), {
    v: 1, sub: "google-sub", email: "user@spu.ac.th", name: "User", picture: "",
    role: "REQUESTER", iat: 1000, exp: 1600
  });
});

test("tampered and expired sessions fail closed", () => {
  const token = issueSession({ email: "user@spu.ac.th", role: "REQUESTER" }, secret, 600, 1000);
  assert.equal(verifySession(`${token}x`, secret, 1100), null);
  assert.equal(verifySession(token, secret, 1600), null);
  assert.equal(verifySession(token, `${secret}x`, 1100), null);
});

test("cookie parser does not throw on malformed encodings", () => {
  assert.deepEqual(parseCookies("a=1; session=hello%20world; bad=%ZZ"), { a: "1", session: "hello world", bad: "" });
});
