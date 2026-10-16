"use strict";

const COMMON_CHAT = ["listChatThreads", "listMessages", "sendMessage", "markChatRead", "getUnreadCounts"];
const REQUESTER_ACTIONS = new Set([
  "lookupPersonnel", "submitJob", "uploadFile", "listJobs", "listFilesForJob", "listPriceRules", "updateStatus",
  ...COMMON_CHAT
]);
const ADMIN_ACTIONS = new Set([
  ...REQUESTER_ACTIONS,
  "testSheet", "testDrive", "testAccess", "togglePriceRule", "setConfirmedAmount", "listUsers", "sendReportEmail"
]);
const EXECUTIVE_ACTIONS = new Set([
  "listJobs", "listFilesForJob", "listPriceRules", "sendReportEmail", ...COMMON_CHAT
]);

const ROLE_ACTIONS = {
  REQUESTER: REQUESTER_ACTIONS,
  ADMIN: ADMIN_ACTIONS,
  EXECUTIVE: EXECUTIVE_ACTIONS
};

function authorizeAction(role, action, body = {}) {
  const normalizedRole = String(role || "").toUpperCase();
  const normalizedAction = String(action || "");
  const allowed = ROLE_ACTIONS[normalizedRole];
  if (!allowed || !allowed.has(normalizedAction)) return { ok: false, error: "action_forbidden" };

  if (normalizedRole === "REQUESTER" && normalizedAction === "updateStatus" && String(body.status || "") !== "SERVICE_DONE") {
    return { ok: false, error: "status_forbidden" };
  }
  if (normalizedRole === "REQUESTER" && normalizedAction === "uploadFile") {
    const category = String(body.category || "requester-files");
    if (!["requester-files", "payment-slip", "chat-attachments"].includes(category)) {
      return { ok: false, error: "file_category_forbidden" };
    }
  }
  return { ok: true };
}

function sanitizeClientPayload(input) {
  const clean = { ...(input || {}) };
  [
    "token", "gatewayToken", "adminToken", "spreadsheetId", "folderId",
    "actorEmail", "actorRole", "actorSub", "email", "senderEmail", "readerEmail", "senderRole", "readerRole", "role", "by"
  ].forEach((key) => delete clean[key]);
  return clean;
}

module.exports = { ROLE_ACTIONS, authorizeAction, sanitizeClientPayload };
