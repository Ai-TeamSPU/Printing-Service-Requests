"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { authorizeAction, sanitizeClientPayload } = require("../lib/policy.cjs");

test("requester may read/submit own flow but not call admin actions", () => {
  assert.equal(authorizeAction("REQUESTER", "listJobs").ok, true);
  assert.equal(authorizeAction("REQUESTER", "submitJob").ok, true);
  assert.equal(authorizeAction("REQUESTER", "togglePriceRule").ok, false);
  assert.equal(authorizeAction("REQUESTER", "listUsers").ok, false);
});

test("requester status and upload categories are narrowed", () => {
  assert.equal(authorizeAction("REQUESTER", "updateStatus", { status: "SERVICE_DONE" }).ok, true);
  assert.equal(authorizeAction("REQUESTER", "updateStatus", { status: "IN_PRODUCTION" }).ok, false);
  assert.equal(authorizeAction("REQUESTER", "uploadFile", { category: "requester-files" }).ok, true);
  assert.equal(authorizeAction("REQUESTER", "uploadFile", { category: "admin-files" }).ok, false);
});

test("executive is read/report only and admin can operate", () => {
  assert.equal(authorizeAction("EXECUTIVE", "listJobs").ok, true);
  assert.equal(authorizeAction("EXECUTIVE", "sendReportEmail").ok, true);
  assert.equal(authorizeAction("EXECUTIVE", "updateStatus", { status: "PROD_DONE" }).ok, false);
  assert.equal(authorizeAction("ADMIN", "updateStatus", { status: "PROD_DONE" }).ok, true);
  assert.equal(authorizeAction("ADMIN", "setConfirmedAmount").ok, true);
});

test("client-controlled identity and backend coordinates are stripped", () => {
  const clean = sanitizeClientPayload({
    jobNo: "PRN-1", email: "victim@spu.ac.th", role: "ADMIN", by: "victim",
    token: "leak", gatewayToken: "leak2", spreadsheetId: "other-sheet", folderId: "other-folder"
  });
  assert.deepEqual(clean, { jobNo: "PRN-1" });
});
