"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const backend = fs.readFileSync(path.join(root, "google-apps-script", "Code.gs"), "utf8");
const gateway = fs.readFileSync(path.join(root, "auth-gateway", "server.js"), "utf8");
const policy = fs.readFileSync(path.join(root, "auth-gateway", "lib", "policy.cjs"), "utf8");

function balancedBlock(source, braceStart, label) {
  let depth = 0, quote = null, escaped = false, lineComment = false, blockComment = false;
  for (let i = braceStart; i < source.length; i += 1) {
    const ch = source[i], next = source[i + 1];
    if (lineComment) { if (ch === "\n") lineComment = false; continue; }
    if (blockComment) { if (ch === "*" && next === "/") { blockComment = false; i += 1; } continue; }
    if (quote) { if (escaped) escaped = false; else if (ch === "\\") escaped = true; else if (ch === quote) quote = null; continue; }
    if (ch === "/" && next === "/") { lineComment = true; i += 1; continue; }
    if (ch === "/" && next === "*") { blockComment = true; i += 1; continue; }
    if (["\"", "'", "`"].includes(ch)) { quote = ch; continue; }
    if (ch === "{") depth += 1;
    if (ch === "}" && --depth === 0) return source.slice(braceStart, i + 1);
  }
  assert.fail(`${label}: closing brace not found`);
}

function functionSource(source, name) {
  const match = new RegExp(`function\\s+${name}\\s*\\(`).exec(source);
  assert.ok(match, `function ${name} is missing`);
  return balancedBlock(source, source.indexOf("{", match.index), name);
}

function frontendLogic() {
  const match = html.match(/<script\b[^>]*type=["']text\/x-dc["'][^>]*data-dc-script[^>]*>([\s\S]*?)<\/script>/i);
  assert.ok(match);
  return match[1];
}

function methodSource(name) {
  const source = frontendLogic();
  const match = new RegExp(`^\\s{2}(?:async\\s+)?${name}\\s*\\([^\\n]*\\)\\s*\\{`, "m").exec(source);
  assert.ok(match, `frontend method ${name} is missing`);
  return balancedBlock(source, source.indexOf("{", match.index), name);
}

test("browser contains no Apps Script shared secret or direct backend coordinates", () => {
  assert.doesNotMatch(html, /script\.google\.com\/macros\/s\//);
  assert.doesNotMatch(html, /6552b83f-9a7a-4dd1-8169-05193c6744e3/);
  assert.doesNotMatch(html, /body\s*:\s*JSON\.stringify\(\{[^}]*token\s*:/s);
  assert.match(html, /spu-google-signin/);
  assert.doesNotMatch(html, /<input[^>]+value=["']\{\{\s*gateEmail\s*\}\}/i);
});

test("frontend restores a server session and never loads jobs directly on mount", () => {
  const mounted = methodSource("componentDidMount");
  assert.match(mounted, /this\._bootstrapAuth\(\)/);
  assert.doesNotMatch(mounted, /this\._loadLiveJobs\(\)/);
  assert.match(methodSource("_authFetch"), /credentials\s*:\s*["']include["']/);
  assert.match(methodSource("_gasCall"), /["']\/api["']/);
  assert.doesNotMatch(frontendLogic(), /localStorage\.(?:getItem|setItem)\([^)]*(?:token|auth|session)/i);
});

test("frontend clears requester data when the authenticated account changes", () => {
  const clear = methodSource("_clearAuthenticatedState");
  assert.match(clear, /form\s*:\s*\{\s*name\s*:\s*["']["']/);
  assert.match(clear, /liveJobs\s*:\s*\[\]/);
  assert.match(clear, /uploadedFiles\s*:\s*\[\]/);
  assert.match(methodSource("componentWillUnmount"), /clearTimeout\(this\._googleButtonTimer\)/);
});

test("Apps Script requires a second server-only gateway token for every browser action", () => {
  const handle = functionSource(backend, "handle_");
  assert.match(handle, /GATEWAY_REQUIRED_ACTIONS_\[p\.action\]/);
  assert.match(handle, /gatewayAuthorized_\(p\)/);
  const gatewayAuth = functionSource(backend, "gatewayAuthorized_");
  assert.match(gatewayAuth, /getGatewayToken_\(\)/);
  assert.match(gatewayAuth, /p\.gatewayToken/);
  assert.match(gatewayAuth, /ACTION_ROLES_\[p\.action\]/);
  assert.match(functionSource(backend, "openSheet_"), /getProperty\(["']SPREADSHEET_ID["']\)/);
  assert.match(functionSource(backend, "openFolder_"), /getProperty\(["']FOLDER_ID["']\)/);
});

test("requester data is scoped and privileged mutations are server-authorized", () => {
  const listJobs = functionSource(backend, "listJobs_");
  assert.match(listJobs, /actorOwnsJob_\(actor, job\)/);
  assert.match(listJobs, /allowedJobNos/);
  const updateStatus = functionSource(backend, "updateStatus_");
  assert.match(updateStatus, /old\s*!==\s*["']PROD_DONE["']/);
  assert.match(updateStatus, /next\s*!==\s*["']SERVICE_DONE["']/);
  assert.match(updateStatus, /actorOwnsJob_/);
  const upload = functionSource(backend, "uploadFile_");
  assert.ok(upload.indexOf("authorizeJobAccess_") < upload.indexOf("createFile(blob)"));
  assert.match(functionSource(backend, "listFilesForJob_"), /authorizeJobAccess_/);
  assert.match(functionSource(backend, "setConfirmedAmount_"), /actor\.role\s*!==\s*["']ADMIN["']/);
});

test("stable Google subject is stored and legacy email remains only as fallback", () => {
  assert.match(backend, /["']requester_google_sub["']/);
  assert.match(functionSource(backend, "actorOwnsJob_"), /ownerSub\s*&&\s*actor\.sub/);
  assert.match(functionSource(backend, "submitJob_"), /actor\.sub/);
});

test("gateway verifies Google claims and strips client identity", () => {
  assert.match(gateway, /verifyIdToken\(\{\s*idToken:\s*credential,\s*audience:\s*config\.googleClientId\s*\}\)/);
  assert.match(gateway, /claims\.email_verified\s*!==\s*true/);
  assert.match(gateway, /hostedDomain\s*!==\s*config\.allowedDomain/);
  assert.match(gateway, /httpOnly:\s*true/);
  assert.match(gateway, /sameSite:\s*config\.cookieSameSite/);
  assert.match(policy, /"actorEmail"[^\n]*"actorRole"[^\n]*"actorSub"/);
  assert.match(policy, /"spreadsheetId"[^\n]*"folderId"/);
  assert.match(gateway, /callAppsScript\(["']lookupPersonnel["'],\s*\{\},\s*user\)/);
  assert.match(gateway, /!origin\s*\|\|\s*!config\.allowedOrigins\.has\(origin\)/);
});

test("listUsers never returns password cells", () => {
  assert.match(functionSource(backend, "listUsers_"), /h\s*!==\s*["']password["']/);
});
