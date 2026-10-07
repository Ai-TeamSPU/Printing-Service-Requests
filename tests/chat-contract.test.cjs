const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const backend = fs.readFileSync(path.join(root, "google-apps-script", "Code.gs"), "utf8");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const dataModule = fs.readFileSync(path.join(root, "spu-data.js"), "utf8");
const runtime = fs.readFileSync(path.join(root, "support.js"), "utf8");

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function balancedBlock(source, braceStart, label) {
  assert.notEqual(braceStart, -1, `${label}: opening brace not found`);
  let depth = 0;
  let quote = null;
  let escaped = false;
  let lineComment = false;
  let blockComment = false;

  for (let i = braceStart; i < source.length; i += 1) {
    const ch = source[i];
    const next = source[i + 1];

    if (lineComment) {
      if (ch === "\n") lineComment = false;
      continue;
    }
    if (blockComment) {
      if (ch === "*" && next === "/") {
        blockComment = false;
        i += 1;
      }
      continue;
    }
    if (quote) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === quote) quote = null;
      continue;
    }
    if (ch === "/" && next === "/") {
      lineComment = true;
      i += 1;
      continue;
    }
    if (ch === "/" && next === "*") {
      blockComment = true;
      i += 1;
      continue;
    }
    if (ch === "\"" || ch === "'" || ch === "`") {
      quote = ch;
      continue;
    }
    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(braceStart, i + 1);
    }
  }
  assert.fail(`${label}: closing brace not found`);
}

function functionSource(source, name) {
  const match = new RegExp(`function\\s+${escapeRegex(name)}\\s*\\(`).exec(source);
  assert.ok(match, `function ${name} is missing`);
  return balancedBlock(source, source.indexOf("{", match.index), `function ${name}`);
}

function methodSource(source, name) {
  const match = new RegExp(`^\\s{2}(?:async\\s+)?${escapeRegex(name)}\\s*\\([^\\n]*\\)\\s*\\{`, "m").exec(source);
  assert.ok(match, `frontend method ${name} is missing`);
  return balancedBlock(source, source.indexOf("{", match.index), `method ${name}`);
}

function indexOfOrFail(source, needle, label) {
  const index = source.indexOf(needle);
  assert.notEqual(index, -1, `${label}: expected ${JSON.stringify(needle)}`);
  return index;
}

function evaluateBackend() {
  const context = vm.createContext({});
  new vm.Script(backend, { filename: "google-apps-script/Code.gs" }).runInContext(context);
  return context;
}

function evaluateDataModule() {
  const context = vm.createContext({});
  const script = dataModule.replace(/^export\s+/gm, "") + "\n;globalThis.__sheetTabs = SHEET_TABS;";
  new vm.Script(script, { filename: "spu-data.js" }).runInContext(context);
  return plain(context.__sheetTabs);
}

function extractFrontendLogic() {
  const match = html.match(/<script\b[^>]*type=["']text\/x-dc["'][^>]*data-dc-script[^>]*>([\s\S]*?)<\/script>/i);
  assert.ok(match, "inline data-dc-script block is missing");
  return match[1];
}

test("chat JavaScript sources parse locally", () => {
  assert.doesNotThrow(() => new vm.Script(backend, { filename: "google-apps-script/Code.gs" }));
  assert.doesNotThrow(() => new vm.Script(extractFrontendLogic(), { filename: "index.html#data-dc-script" }));
  assert.doesNotThrow(() => new vm.Script(dataModule.replace(/^export\s+/gm, ""), { filename: "spu-data.js" }));
  assert.doesNotThrow(() => new vm.Script(runtime, { filename: "support.js" }));
});

test("templated slip image stays inert until the runtime resolves its loop item", () => {
  assert.match(html, /<img\s+sc-camel-src=["']\{\{\s*f\.imgSrc\s*\}\}["']\s+sc-camel-on-error=["']\{\{\s*f\.onImgError\s*\}\}["']/);
  assert.doesNotMatch(html, /<img\s+src=["']\{\{\s*f\.imgSrc\s*\}\}["']/);
  assert.doesNotMatch(html, /onError=["']\{\{\s*f\.onImgError\s*\}\}["']/);
});

test("backend exposes the required chat schemas, routes, and functions", () => {
  const context = evaluateBackend();
  const expectedSchemas = {
    JOB_MESSAGES: [
      "message_id", "job_no", "sequence", "sender_email", "sender_role",
      "message_type", "message_text", "attachment_file_ids", "created_at", "client_message_id",
    ],
    CHAT_READS: [
      "job_no", "reader_email", "reader_role", "last_read_at", "last_read_message_id", "updated_at",
    ],
  };

  for (const [name, columns] of Object.entries(expectedSchemas)) {
    assert.deepEqual(plain(context.SHEET_SCHEMA[name]), columns, `${name} schema drifted`);
  }
  assert.ok(plain(context.SHEET_SCHEMA.FILES).includes("client_upload_id"), "FILES must retain its upload idempotency key");

  const routes = {
    listChatThreads: "listChatThreads_",
    listMessages: "listMessages_",
    sendMessage: "sendMessage_",
    markChatRead: "markChatRead_",
    getUnreadCounts: "getUnreadCounts_",
  };
  for (const [action, fn] of Object.entries(routes)) {
    assert.match(backend, new RegExp(`case\\s+["']${action}["']\\s*:\\s*return\\s+json_\\(${fn}\\(p\\)\\)`));
    assert.equal(typeof context[fn], "function", `${fn} must remain callable by Apps Script`);
  }

  const testSheet = functionSource(backend, "testSheet_");
  assert.match(
    testSheet,
    /\[\s*["']FILES["']\s*,\s*["']JOB_MESSAGES["']\s*,\s*["']CHAT_READS["']\s*\]\.forEach/,
    "testSheet_ must validate FILES and both chat tab schemas before reporting a healthy connection",
  );

  const displayTabs = new Map(evaluateDataModule().map((tab) => [tab.name, tab.cols.map((col) => col[1])]));
  assert.deepEqual(displayTabs.get("JOB_MESSAGES"), expectedSchemas.JOB_MESSAGES, "spu-data JOB_MESSAGES must match Code.gs");
  assert.deepEqual(displayTabs.get("CHAT_READS"), expectedSchemas.CHAT_READS, "spu-data CHAT_READS must match Code.gs");
  assert.ok(displayTabs.get("FILES")?.includes("client_upload_id"), "spu-data FILES must show client_upload_id");
});

test("job allocation and chat writes retain locking and idempotency", () => {
  const submit = functionSource(backend, "submitJob_");
  const submitLock = indexOfOrFail(submit, "LockService.getScriptLock()", "submitJob lock");
  const allocate = indexOfOrFail(submit, "nextJobNo_(jobsSh)", "job number allocation");
  const appendJob = indexOfOrFail(submit, "jobsSh.appendRow", "JOBS append");
  assert.ok(submitLock < allocate && allocate < appendJob, "job number allocation and JOBS append must happen under the same lock");
  assert.match(submit, /finally\s*\{[\s\S]*lock\.releaseLock\(\)/);

  const send = functionSource(backend, "sendMessage_");
  assert.match(send, /missing_client_message_id/);
  assert.match(send, /invalid_client_message_id/);
  assert.match(send, /LockService\.getScriptLock\(\)/);
  assert.match(send, /finally\s*\{[\s\S]*lock\.releaseLock\(\)/);
  const duplicateCheck = indexOfOrFail(send, "existing.clientMessageId === clientMessageId", "message deduplication");
  const appendMessage = indexOfOrFail(send, "appendChatMessage_(table, message)", "message append");
  assert.ok(duplicateCheck < appendMessage, "deduplication must run before appending the message");
  assert.match(send, /existing\.jobNo\s*===\s*jobNo/);
  assert.match(send, /existing\.senderEmail\s*===\s*auth\.actor\.email/);
  assert.match(send, /duplicate:\s*true/);

  const markRead = functionSource(backend, "markChatRead_");
  assert.match(markRead, /LockService\.getScriptLock\(\)/);
  assert.match(markRead, /finally\s*\{[\s\S]*lock\.releaseLock\(\)/);

  const upload = functionSource(backend, "uploadChatAttachment_");
  assert.match(upload, /clientUploadId/);
  assert.match(upload, /duplicate:\s*true/);
  assert.match(upload, /LockService\.getScriptLock\(\)/);
  assert.ok(
    indexOfOrFail(upload, "authorizeChat_(ss, p, jobNo)", "attachment authorization") <
      indexOfOrFail(upload, "catFolder.createFile(blob)", "Drive file creation"),
    "chat attachment authorization must happen before writing to Drive",
  );
  assert.match(upload, /file\.setTrashed\(true\)/, "failed FILES writes should clean up the Drive file");
});

test("server-side chat authorization, validation, and sheet-text guards remain in place", () => {
  const authorize = functionSource(backend, "authorizeChat_");
  assert.match(authorize, /job_not_found/);
  assert.match(authorize, /duplicate_job_no/);
  assert.match(authorize, /job\.requesterEmail\s*!==\s*email/);
  assert.match(authorize, /findActiveChatStaff_\(ss, email\)/);
  assert.match(authorize, /staff_forbidden/);

  for (const name of ["listChatThreads_", "listMessages_", "sendMessage_", "markChatRead_"]) {
    assert.match(functionSource(backend, name), /authorizeChat_\(ss, p,/i, `${name} must authorize against Sheets`);
  }

  const send = functionSource(backend, "sendMessage_");
  assert.match(send, /message_too_long/);
  assert.match(send, /empty_message/);
  assert.match(send, /too_many_attachments/);
  assert.match(send, /attachment_not_found/);
  assert.match(send, /fileMap\[jobNo\s*\+\s*"\\n"\s*\+\s*attachmentIds\[f\]\]/);

  const append = functionSource(backend, "appendChatMessage_");
  assert.match(append, /safeSheetText_\(message\.text\)/, "message text must not become a Sheet formula");
  assert.match(functionSource(backend, "safeSheetText_"), /\^\[=\+\\-@\]/);

  const upload = functionSource(backend, "uploadChatAttachment_");
  for (const guard of ["unsupported_file_type", "file_type_mismatch", "empty_file", "file_too_large"]) {
    assert.match(upload, new RegExp(guard));
  }
});

test("frontend keeps a stable message retry id until a confirmed success", () => {
  const logic = extractFrontendLogic();
  assert.match(logic, /chatPendingClientId\s*:\s*""/);
  const handleFiles = methodSource(logic, "_handleChatFiles");
  assert.match(handleFiles, /accepted\.push\(\{\s*file\s*,\s*id\s*:\s*this\._chatClientId\(\)\s*\}\)/);
  assert.match(handleFiles, /id\s*:\s*entry\.id/);
  const send = methodSource(logic, "_sendChatMessage");
  assert.match(send, /const\s+clientMessageId\s*=\s*S\.chatPendingClientId\s*\|\|\s*this\._chatClientId\(\)/);
  assert.match(send, /chatPendingClientId\s*:\s*clientMessageId/);
  assert.match(send, /clientMessageId\s*,\s*attachmentFileIds/);
  assert.match(send, /if\s*\(!r\s*\|\|\s*!r\.ok\)\s*throw/);
  assert.match(
    send,
    /clientUploadId\s*:\s*file\.id/,
    "each chat upload must reuse the stable per-file idempotency key retained in draft state",
  );

  const catchIndex = indexOfOrFail(send, "catch(err)", "send failure handler");
  const catchBlock = send.slice(catchIndex);
  assert.doesNotMatch(catchBlock, /chatPendingClientId\s*:\s*""/, "a failed/uncertain send must preserve its idempotency key");

  const successClear = send.match(/this\.setState\(\{[^}]*chatDraft\s*:\s*""[^}]*chatPendingClientId\s*:\s*""[^}]*\}\)/s);
  assert.ok(successClear, "confirmed send success must clear the pending retry id with the draft");
});

test("frontend polling stops cleanly and ignores stale message responses", () => {
  const logic = extractFrontendLogic();
  const start = methodSource(logic, "_startChatPolling");
  const stop = methodSource(logic, "_stopChatPolling");
  const load = methodSource(logic, "_loadChatMessages");
  const unmount = methodSource(logic, "componentWillUnmount");
  const go = methodSource(logic, "go");

  assert.match(start, /this\._stopChatPolling\(\)/);
  assert.match(start, /this\.state\.page\s*!==\s*"chat"/);
  assert.match(start, /document\.visibilityState\s*===\s*"hidden"/);
  assert.match(start, /this\._chatPollBusy/);
  assert.match(stop, /clearInterval\(this\._chatPollTimer\)/);
  assert.match(unmount, /this\._stopChatPolling\(\)/);
  assert.match(go, /this\._stopChatPolling\(\)/, "SPA page changes must stop polling because the root component stays mounted");

  assert.match(load, /const\s+requestId\s*=\s*\(this\._chatMessageRequestId\s*\|\|\s*0\)\s*\+\s*1/);
  assert.match(load, /this\._chatMessageRequestId\s*=\s*requestId/);
  assert.match(load, /requestId\s*!==\s*this\._chatMessageRequestId/);
  assert.match(load, /target\s*!==\s*this\.state\.chatSelectedJobNo/);
  assert.match(load, /row\.sequence\s*\|\|\s*row\.sequence_no\s*\|\|\s*row\.sequenceNo/);
  assert.match(load, /const\s+payload\s*=\s*\{[^}]*limit\s*:\s*100[^}]*\}/s);
  assert.match(load, /if\s*\(incremental[^)]*\)\s*payload\.afterSequence\s*=/);
  assert.doesNotMatch(
    load,
    /const\s+payload\s*=\s*\{[^}]*afterSequence/s,
    "initial history loads must omit afterSequence so the backend returns the latest page",
  );
});

test("frontend applies backend-confirmed unread state only after markChatRead succeeds", () => {
  const load = methodSource(extractFrontendLogic(), "_loadChatMessages");
  const call = indexOfOrFail(load, 'this._gasCall("markChatRead"', "markChatRead call");
  const successGuard = indexOfOrFail(load, "if(marked&&marked.ok", "markChatRead success guard");
  const backendUnread = indexOfOrFail(load, "Number(marked.unreadCount)", "backend unread result");
  const updateUnread = indexOfOrFail(load, "unread_count:remaining", "local unread update");
  assert.ok(
    call < successGuard && successGuard < backendUnread && backendUnread < updateUnread,
    "unread state may update only from a successful backend mark-read result",
  );
  assert.match(load, /unreadBeforeRead\s*>\s*0\s*\|\|\s*hasNewMessageFromOther/);
  assert.match(load, /document\.visibilityState\s*!==\s*"hidden"/);
});

test("frontend and backend agree on chat limits", () => {
  const context = evaluateBackend();
  const textarea = html.match(/<textarea\b[^>]*value=["']\{\{\s*chatDraft\s*\}\}["'][^>]*>/i);
  assert.ok(textarea, "chat composer textarea is missing");
  const maxlength = textarea[0].match(/maxlength=["'](\d+)["']/i);
  assert.ok(maxlength, "chat composer maxlength is missing");
  assert.equal(Number(maxlength[1]), context.CHAT_MAX_MESSAGE_LENGTH_);
  assert.equal(context.CHAT_MAX_ATTACHMENTS_, 5, "UI currently exposes at most five chat attachments");
  assert.equal(context.CHAT_MAX_PAGE_SIZE_, 100, "frontend incremental loads use a 100-message page");
});
