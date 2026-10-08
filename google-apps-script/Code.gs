/**
 * SPU Printing Service — Google Apps Script backend
 * ---------------------------------------------------
 * ผูกกับ Google Sheet ที่เก็บข้อมูลระบบ (JOBS, JOB_ITEMS, PRICE_RULES, ...)
 * และอ่าน/เขียนไฟล์ผ่าน Google Drive
 *
 * วิธีติดตั้ง: ดู README.md ในโฟลเดอร์นี้
 */

// ---------- โครงสร้างแท็บที่ระบบต้องใช้ (ต้องตรงกับ spu-data.js -> SHEET_TABS) ----------
var SHEET_SCHEMA = {
  JOBS: ["job_no","submitted_at","requester_email","requester_name","position_type","phone","unit_code","required_date","purpose","status","estimated_amount","confirmed_amount","completed_at","drive_folder_id","budget_source","budget_acct"],
  // variant_snapshot (คอลัมน์ D): เลิกเขียนค่าใหม่แล้ว (ทางเลือก A — เก็บคอลัมน์ไว้ ไม่ลบ ข้อมูลเก่าไม่หาย) ใช้คอลัมน์แยกท้ายแถวแทน
  JOB_ITEMS: ["item_id","job_no","service_code","variant_snapshot","quantity","unit","unit_price","estimated_amount","confirmed_amount","price_rule_id","price_rule_snapshot","override_reason","color_mode","paper_size","paper_type","sides","staple","own_paper","qty_original","qty_sets","exam_type","exam_subject","service_group","note"],
  PRICE_RULES: ["rule_id","service_code","condition","price_type","price","min_price","max_price","effective_from","effective_to","active","note"],
  UNITS: ["unit_code","unit_name","parent_group","display_order","active"],
  USERS: ["email","full_name","role","unit_code","phone","active","password"],
  FILES: ["file_id","job_no","file_name","mime_type","file_size","file_category","uploaded_by","uploaded_at","web_view_link","client_upload_id"],
  STATUS_LOG: ["log_id","job_no","old_status","new_status","changed_by","changed_at","note","channel"],
  JOB_MESSAGES: ["message_id","job_no","sequence","sender_email","sender_role","message_type","message_text","attachment_file_ids","created_at","client_message_id"],
  CHAT_READS: ["job_no","reader_email","reader_role","last_read_at","last_read_message_id","updated_at"],
  NOTIFY_LOG: ["notify_id","job_no","template_code","recipient","status","sent_at","error"],
  SETTINGS: ["key","value","updated_by","updated_at","note"]
};

// Chat is intentionally keyed by job_no: an existing job is the conversation.
// These limits are enforced on the server even if the browser also validates them.
var CHAT_MAX_MESSAGE_LENGTH_ = 5000;
var CHAT_MAX_ATTACHMENTS_ = 5;
var CHAT_MAX_PAGE_SIZE_ = 100;

/**
 * รันฟังก์ชันนี้ "ครั้งเดียว" จากตัวแก้ไข Apps Script (เลือก setupSheets แล้วกด Run)
 * เพื่อสร้างแท็บทั้งหมดพร้อมหัวตารางในสเปรดชีตที่สคริปต์นี้ผูกอยู่
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var headers = SHEET_SCHEMA[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  // ลบแท็บเริ่มต้น "Sheet1" ถ้ายังว่างอยู่และไม่ได้ใช้
  var def = ss.getSheetByName("Sheet1");
  if (def && def.getLastRow() === 0 && ss.getSheets().length > Object.keys(SHEET_SCHEMA).length) {
    ss.deleteSheet(def);
  }
  Logger.log("สร้างแท็บครบแล้ว: " + Object.keys(SHEET_SCHEMA).join(", "));
}

// ---------- แท็บ Personnel_Information (ทะเบียนบุคลากร ใช้เติมฟอร์มอัตโนมัติ) ----------
// แยกออกจาก setupSheets() ข้างบนโดยตั้งใจ: setupSheets() เขียนทับหัวตารางทุกครั้งที่รัน (เพื่อรับ schema ใหม่ของแท็บระบบ)
// แต่แท็บนี้อาจเป็นทะเบียนบุคลากรจริงที่คุณสร้าง/กรอกไว้เองล่วงหน้าด้วยชื่อคอลัมน์ของคุณเอง จึง "สร้างให้ก็ต่อเมื่อยังไม่มีแท็บนี้อยู่เลย" เท่านั้น ไม่แตะของเดิม
var PERSONNEL_SHEET_NAME = "Personnel_Information";

/**
 * รันฟังก์ชันนี้ "ครั้งเดียว" ถ้ายังไม่เคยมีแท็บ Personnel_Information ในชีต — จะสร้างแท็บเปล่าพร้อมหัวตารางให้
 * ถ้ามีแท็บนี้อยู่แล้ว (ไม่ว่าจะหัวตารางแบบไหน) จะไม่แตะต้องอะไรเลย ปลอดภัยกับข้อมูลที่กรอกไว้แล้ว
 */
function setupPersonnelSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PERSONNEL_SHEET_NAME);
  if (sh) {
    Logger.log("มีแท็บ " + PERSONNEL_SHEET_NAME + " อยู่แล้ว ไม่แตะต้องหัวตาราง/ข้อมูลเดิม");
    return;
  }
  sh = ss.insertSheet(PERSONNEL_SHEET_NAME);
  var headers = ["requester_email", "full_name", "unit_name"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  Logger.log("สร้างแท็บ " + PERSONNEL_SHEET_NAME + " ใหม่ พร้อมหัวตาราง: " + headers.join(", "));
}

// ---------- แท็บ PRICE_RULES (ตารางราคากลาง) — เพิ่งเชื่อมต่อให้หน้า "ตารางราคา" อ่าน/เขียนได้จริง ----------
/**
 * รันฟังก์ชันนี้ "ครั้งเดียว" ถ้าแท็บ PRICE_RULES ยังว่างอยู่ (มีแค่หัวตาราง ไม่มีข้อมูล) — จะใส่ราคาตั้งต้นให้ตรงกับ
 * ที่เคยแสดงเป็นตัวอย่างในหน้าเว็บ ถ้ามีข้อมูลอยู่แล้ว (ไม่ว่าจะแก้ไขไปแค่ไหน) จะไม่แตะต้องอะไรเลย ปลอดภัย
 */
function seedPriceRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("PRICE_RULES");
  if (!sh) { Logger.log("ไม่พบแท็บ PRICE_RULES — กรุณารัน setupSheets ก่อน"); return; }
  if (sh.getLastRow() > 1) { Logger.log("แท็บ PRICE_RULES มีข้อมูลอยู่แล้ว ไม่แตะต้องของเดิม"); return; }
  // [rule_id, service_code, condition, price_type, price, min_price, max_price, effective_from, effective_to, active, note]
  var rows = [
    ["R01","copy","A4 ปอนด์ 70 แกรม ขาวดำ|A4 bond 70 gsm B/W","PER_SHEET","0.50","","","01/10/2567","","TRUE",""],
    ["R02","copy","A4 ปอนด์สี 80 แกรม|A4 coloured bond 80 gsm","PER_SHEET","2.00","","","01/10/2567","","TRUE",""],
    ["R03","copy","A3 ปอนด์ 80 แกรม|A3 bond 80 gsm","PER_SHEET","5.00","","","01/10/2567","","TRUE",""],
    ["R04","dup","A4 ปอนด์ 70 แกรม|A4 bond 70 gsm","PER_SHEET","2.00","","","01/10/2567","","TRUE",""],
    ["R05","dup","การ์ดสี 180 แกรม ปกชุด|Coloured card 180 gsm, cover","PER_SHEET","5.00","","","01/10/2567","","TRUE",""],
    ["R06","exam","ข้อสอบ A4 ปอนด์ 70 แกรม|Exam paper, A4 bond 70 gsm","PER_SHEET","2.00","","","01/10/2567","","TRUE",""],
    ["R07","card","1 ด้าน ซิมโบแมทพลัส 250 แกรม|1-side Simbo Matt Plus 250 gsm","PER_ITEM","2.50","","","01/10/2567","","TRUE",""],
    ["R08","card","2 ด้าน แบบมหาวิทยาลัย|2-side university template","PER_ITEM","3.00","","","01/10/2567","","TRUE",""],
    ["R09","cert","โลโก้ SPU + อักษรสีดำ|SPU logo + black text","PER_SHEET","10.00","","","01/10/2567","","TRUE",""],
    ["R10","cert","โลโก้ SPU + โลโก้อื่น + อักษรสี|SPU + other logos, colour text","PER_SHEET","15.00","","","01/10/2567","","TRUE",""],
    ["R11","cert","เจียนขอบ 4 ด้าน เฉพาะวุฒิบัตร|Trim 4 edges, certificates only","ADD_ON","+3.00","","","01/10/2567","","TRUE",""],
    ["R12","a4","A4 ปอนด์ 70 แกรม|A4 bond 70 gsm","PER_SHEET","15.00","","","01/10/2567","","TRUE",""],
    ["R13","a4","A4 ซิมโบแมทพลัส 250 แกรม|A4 Simbo Matt Plus 250 gsm","PER_SHEET","20.00","","","01/10/2567","","TRUE",""],
    ["R14","a4","MOU สครีม 180 แกรม หน้าโลโก้ 4 สี|MOU, Scream 180 gsm, 4-colour logo page","PER_PAGE","20.00","","","01/10/2567","","TRUE",""],
    ["R15","a4","MOU สครีม 180 แกรม หน้าอักษรสีดำ|MOU, Scream 180 gsm, black text page","PER_PAGE","10.00","","","01/10/2567","","TRUE",""],
    ["R16","a4","F4 นำกระดาษมาเอง|F4, own paper","PER_SHEET","20.00","","","01/10/2567","","TRUE",""],
    ["R17","a3","A3 ปอนด์ 80 แกรม|A3 bond 80 gsm","PRICE_RANGE","","30","50","01/10/2567","","TRUE",""],
    ["R18","a3","A3 ซิมโบแมทพลัส 250 แกรม|A3 Simbo Matt Plus 250 gsm","PER_SHEET","60.00","","","01/10/2567","","TRUE",""],
    ["R19","layout","จัดรูปเล่มขาวดำ|B/W layout","PER_PAGE","50.00","","","01/10/2567","","TRUE",""],
    ["R20","layout","จัดรูปเล่ม 4 สี|4-colour layout","PRICE_RANGE","","100","200","01/10/2567","","TRUE",""],
    ["R21","a5","พริ้นสี A5|Colour print A5","MANUAL_QUOTE","","","","01/10/2567","","TRUE",""],
    ["R22","bind","เข้าเล่มทั่วไป|General binding","MANUAL_QUOTE","","","","01/10/2567","","TRUE",""]
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log("ใส่ราคาตั้งต้น " + rows.length + " รายการให้แท็บ PRICE_RULES แล้ว");
}

function listPriceRules_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName("PRICE_RULES");
  if (!sh) return { ok: false, error: "price_rules_sheet_not_found" };
  var values = sh.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, rules: [] };
  var headers = values.shift();
  var rules = values.map(function (row) {
    var o = {};
    headers.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
  return { ok: true, rules: rules };
}

// เปิด/ปิดใช้งานกฎราคาหนึ่งแถว (คอลัมน์ active) — ใช้โดยสวิตช์เปิด/ปิดในหน้า "ตารางราคา"
function togglePriceRule_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName("PRICE_RULES");
  if (!sh) return { ok: false, error: "price_rules_sheet_not_found" };
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var idIdx = headers.indexOf("rule_id");
  var activeIdx = headers.indexOf("active");
  if (idIdx === -1 || activeIdx === -1) return { ok: false, error: "columns_not_found" };
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(p.ruleId)) {
      var newVal = p.active === true || String(p.active).toUpperCase() === "TRUE";
      sh.getRange(i + 1, activeIdx + 1).setValue(newVal);
      return { ok: true, ruleId: p.ruleId, active: newVal };
    }
  }
  return { ok: false, error: "rule_not_found" };
}

// ---------- จุดเข้า HTTP ----------
function doGet(e) {
  return handle_((e && e.parameter) || {});
}
function doPost(e) {
  var params = {};
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    params = (e && e.parameter) || {};
  }
  return handle_(params);
}

function handle_(p) {
  try {
    var token = getToken_();
    if (!token || p.token !== token) {
      return json_({ ok: false, error: "unauthorized" });
    }
    if (BULK_ACTIONS_[p.action]) {
      var ba = bulkAuthorized_(p);
      if (!ba.ok) return json_(ba);
    }
    switch (p.action) {
      case "testSheet": return json_(testSheet_(p));
      case "testDrive": return json_(testDrive_(p));
      case "testAccess": return json_(testAccess_(p));
      case "submitJob": return json_(submitJob_(p));
      case "updateStatus": return json_(updateStatus_(p));
      case "importJobs": return json_(importJobs_(p));
      case "patchJobDates": return json_(patchJobDates_(p));
      case "uploadFile": return json_(uploadFile_(p));
      case "listJobs": return json_(listJobs_(p));
      case "checkAdminUser": return json_(checkAdminUser_(p));
      case "listUsers": return json_(listUsers_(p));
      case "sendReportEmail": return json_(sendReportEmail_(p));
      case "lookupPersonnel": return json_(lookupPersonnel_(p));
      case "listPriceRules": return json_(listPriceRules_(p));
      case "togglePriceRule": return json_(togglePriceRule_(p));
      case "listFilesForJob": return json_(listFilesForJob_(p));
      case "setConfirmedAmount": return json_(setConfirmedAmount_(p));
      case "listChatThreads": return json_(listChatThreads_(p));
      case "listMessages": return json_(listMessages_(p));
      case "sendMessage": return json_(sendMessage_(p));
      case "markChatRead": return json_(markChatRead_(p));
      case "getUnreadCounts": return json_(getUnreadCounts_(p));
      default: return json_({ ok: false, error: "unknown_action" });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- ตั้งค่า / ยืนยันสิทธิ์ ----------
// รหัสลับที่หน้าเว็บต้องส่งมาด้วยทุกคำขอ ตั้งค่าใน Project Settings > Script properties คีย์ CONNECT_TOKEN
function getToken_() {
  return PropertiesService.getScriptProperties().getProperty("CONNECT_TOKEN") || "";
}

// โทเคนตัวที่ 2 สำหรับคำสั่งที่แก้ข้อมูลทีละหลายแถว (importJobs, patchJobDates)
// ตั้งใน Project Settings > Script properties คีย์ ADMIN_TOKEN และ "ห้าม" ใส่ไว้ในหน้าเว็บเด็ดขาด
// เพราะ CONNECT_TOKEN ฝังอยู่ใน index.html ใครเปิด View Source ก็เห็น จึงกันคำสั่งอันตรายไม่ได้
function getAdminToken_() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_TOKEN") || "";
}

// คำสั่งที่ลบ/เขียนทับข้อมูลเป็นชุด ต้องผ่านโทเคนตัวที่ 2 เสมอ
// ถ้ายังไม่ได้ตั้ง ADMIN_TOKEN ให้ปฏิเสธไว้ก่อน (fail closed) ไม่ใช่ปล่อยผ่าน
var BULK_ACTIONS_ = { importJobs: true, patchJobDates: true };
function bulkAuthorized_(p) {
  var need = getAdminToken_();
  if (!need) return { ok: false, error: "admin_token_not_set",
    message: "ยังไม่ได้ตั้ง ADMIN_TOKEN ใน Script properties จึงใช้คำสั่งนี้ไม่ได้" };
  if (String(p.adminToken || "") !== need) return { ok: false, error: "admin_unauthorized",
    message: "adminToken ไม่ถูกต้อง" };
  return { ok: true };
}

// เผื่อมีคนส่ง URL เต็มหรือ "ID/edit?gid=..." มาแทนตัว ID ล้วน ๆ (เช่น เรียก API ตรง ๆ ไม่ผ่านหน้าเว็บ) ตัดส่วนเกินออกให้
function cleanId_(raw, isFolder) {
  var s = String(raw || "").trim();
  var re = isFolder ? /\/folders\/([a-zA-Z0-9_-]+)/ : /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;
  var m = s.match(re);
  if (m) return m[1];
  s = s.split("?")[0].split("#")[0];
  var slash = s.indexOf("/");
  return slash === -1 ? s : s.slice(0, slash);
}

function openSheet_(p) {
  if (!p.spreadsheetId) throw new Error("missing_spreadsheetId");
  return SpreadsheetApp.openById(cleanId_(p.spreadsheetId, false));
}
function openFolder_(p) {
  if (!p.folderId) throw new Error("missing_folderId");
  return DriveApp.getFolderById(cleanId_(p.folderId, true));
}

// ---------- ทดสอบการเชื่อมต่อ (ใช้โดยหน้า "ตั้งค่าเชื่อมต่อ") ----------
function testSheet_(p) {
  var ss = openSheet_(p);
  var names = ss.getSheets().map(function (s) { return s.getName(); });
  var missing = Object.keys(SHEET_SCHEMA).filter(function (n) { return names.indexOf(n) === -1; });
  var schemaErrors = [];
  ["FILES", "JOB_MESSAGES", "CHAT_READS"].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var expected = SHEET_SCHEMA[name];
    var actual = sh.getRange(1, 1, 1, expected.length).getValues()[0].map(function (x) { return String(x || ""); });
    var missingHeaders = expected.filter(function (header) { return actual.indexOf(header) === -1; });
    var wrongOrder = missingHeaders.length === 0 && expected.some(function (header, i) { return actual[i] !== header; });
    if (missingHeaders.length || wrongOrder) {
      schemaErrors.push({ tab: name, missingHeaders: missingHeaders, wrongOrder: wrongOrder });
    }
  });
  return { ok: missing.length === 0 && schemaErrors.length === 0, name: ss.getName(), tabs: names, missing: missing, schemaErrors: schemaErrors };
}

function testDrive_(p) {
  var folder = openFolder_(p);
  // เขียนไฟล์ทดสอบเล็ก ๆ แล้วลบทิ้งทันที เพื่อยืนยันว่าเขียนได้จริง ไม่ได้แค่อ่านสิทธิ์
  var probe = folder.createFile("__connection_test__.txt", "ok " + new Date().toISOString());
  var name = folder.getName();
  probe.setTrashed(true);
  return { ok: true, name: name };
}

function testAccess_(p) {
  var email = Session.getEffectiveUser().getEmail();
  return { ok: true, connectedAs: email || "(ไม่ทราบอีเมล — ตรวจสอบสิทธิ์การ Deploy)" };
}

// ---------- งานจริง: ส่งคำขอ / เปลี่ยนสถานะ / อัปโหลดไฟล์ / อ่านรายการ ----------
function nextJobNo_(sheet) {
  var ymPrefix = "PRN-" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMM");
  var lastRow = sheet.getLastRow();
  var max = 0;
  if (lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    data.forEach(function (row) {
      var v = row[0];
      if (typeof v === "string" && v.indexOf(ymPrefix) === 0) {
        var n = parseInt(v.split("-").pop(), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  return ymPrefix + "-" + ("0000" + (max + 1)).slice(-4);
}

// จัดรูปแบบวันเวลาให้อ่านง่ายแบบไทย (พ.ศ. เขตเวลา Asia/Bangkok) ใช้กับคอลัมน์ note
function stampThai_(d) {
  var tz = "Asia/Bangkok";
  var be = Number(Utilities.formatDate(d, tz, "yyyy")) + 543;
  return Utilities.formatDate(d, tz, "d/M/") + be + Utilities.formatDate(d, tz, " HH:mm") + " น.";
}

function logStatus_(ss, jobNo, oldStatus, newStatus, by, channel, note) {
  ss.getSheetByName("STATUS_LOG").appendRow([
    Utilities.getUuid(), jobNo, oldStatus || "", newStatus, by || "", new Date(), note || "", channel || "web"
  ]);
}

function submitJob_(p) {
  var ss = openSheet_(p);
  var jobsSh = ss.getSheetByName("JOBS");
  var itemsSh = ss.getSheetByName("JOB_ITEMS");
  var jobNo = "", now = null, jobRow = -1;

  // job_no is also the chat conversation key. Allocate it and commit the JOBS row under one lock
  // so two simultaneous submissions cannot accidentally share a conversation.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    jobNo = nextJobNo_(jobsSh);
    now = new Date();
    jobRow = jobsSh.getLastRow() + 1;
    jobsSh.appendRow([
      jobNo, now, p.email || "", p.name || "", p.position || "", p.phone || "",
      p.unit || "", p.needBy || "", p.purpose || "", "RECEIVED",
      p.estimatedAmount || "", "", "", "", p.budgetSource || "", p.budgetAcct || ""
    ]);
  } finally {
    lock.releaseLock();
  }

  // สร้างโฟลเดอร์งานบน Drive แบบ "ทำได้ก็ทำ" — ถ้ายังไม่ได้ตั้งค่า Drive (ไม่มี folderId) หรือเขียนไม่ได้
  // ให้ข้ามส่วนนี้ไป ไม่ทำให้การบันทึกแถวลง Sheet ล้มเหลวตามไปด้วย
  var jobFolderId = "";
  if (p.folderId) {
    try {
      var folder = openFolder_(p);
      var jobFolder = findOrCreateFolder_(folder, jobNo);
      // สร้างแค่ "requester-files" ล่วงหน้า (โฟลเดอร์เดียวที่ใช้จริงตอนส่งคำขอ) — admin-files/proof/final/payment-slip
      // ยังไม่สร้างตอนนี้ จะสร้างเองอัตโนมัติทีหลังตอนมีการอัปโหลดไฟล์เข้าหมวดนั้นจริง ๆ ผ่าน findOrCreateFolder_ กันโฟลเดอร์เปล่าคาอยู่ใน Drive
      findOrCreateFolder_(jobFolder, "requester-files");
      var createdFolderId = jobFolder.getId();
      jobsSh.getRange(jobRow, 14).setValue(createdFolderId); // column N = drive_folder_id
      jobFolderId = createdFolderId;
    } catch (err) {
      // เก็บงานลง Sheet ต่อไปได้ แม้ Drive จะยังเชื่อมต่อไม่ได้
    }
  }

  (p.items || []).forEach(function (it, i) {
    itemsSh.appendRow([
      jobNo + "-" + (i + 1), jobNo, it.serviceCode || "", "", // variant_snapshot: เลิกเขียนแล้ว (ทางเลือก A) ใช้คอลัมน์แยกด้านล่างแทน
      it.qty || "", it.unit || "", it.unitPrice || "", it.estimatedAmount || "",
      "", it.priceRuleId || "", it.priceRuleSnapshot || "", "",
      it.colorMode || "", it.paperSize || "", it.paperType || "", it.sides || "",
      it.staple || "", it.ownPaper || "", it.qtyOriginal || "", it.qtySets || "",
      it.examType || "", it.examSubject || "", it.serviceGroup || "",
      // note: ประทับเวลาที่กดส่งคำขอ เฉพาะรายการที่ฝั่งเว็บสั่งมา (ตอนนี้คือข้อสอบแบบ "สอบนอกตาราง")
      // สร้างเวลาที่ฝั่งเซิร์ฟเวอร์ ไม่ใช่เอาจากเบราว์เซอร์ จะได้ตรงกับ submitted_at เสมอ
      it.stampNote ? stampThai_(now) : (it.note || "")
    ]);
  });
  logStatus_(ss, jobNo, "", "RECEIVED", p.email || "system", "web", "ส่งคำขอใหม่ผ่านเว็บ");
  return { ok: true, jobNo: jobNo, folderId: jobFolderId };
}

function updateStatus_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSheet_(p);
    var sh = ss.getSheetByName("JOBS");
    var values = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === p.jobNo) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { ok: false, error: "job_not_found" };
    var old = values[rowIndex][9]; // column J = status
    sh.getRange(rowIndex + 1, 10).setValue(p.status);
    if (p.status === "SERVICE_DONE") sh.getRange(rowIndex + 1, 13).setValue(new Date()); // column M = completed_at
    logStatus_(ss, p.jobNo, old, p.status, p.by || "", p.channel || "web", p.note || "");
    var chatMessageAppended = false, chatWarning = "";
    if (String(old) !== String(p.status)) {
      try { chatMessageAppended = appendStatusSystemMessage_(ss, p.jobNo, old, p.status); }
      catch (chatErr) { chatWarning = "system_message_write_failed"; }
    }
    return { ok: true, chatMessageAppended: chatMessageAppended, chatWarning: chatWarning };
  } finally {
    lock.releaseLock();
  }
}

// นำเข้างานย้อนหลังเป็นชุด (เช่น จากไฟล์ Excel/Google Form เดิม) — เขียนลง JOBS, JOB_ITEMS และ STATUS_LOG พร้อมกัน
// ข้ามรายการที่มีอยู่แล้ว (อีเมล + เวลาส่ง + วัตถุประสงค์ตรงกัน) จึงรันซ้ำได้ปลอดภัย ส่ง dryRun:true เพื่อดูผลก่อนเขียนจริง
// ส่ง replaceAll:true เพื่อล้างข้อมูลเดิมทั้งหมด (เก็บหัวตารางไว้) แล้วนำเข้าชุดใหม่แทน — เลขที่งานจะเริ่มนับใหม่ตั้งแต่ 0001
function importJobs_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSheet_(p);
    var jobsSh = ss.getSheetByName("JOBS");
    var itemsSh = ss.getSheetByName("JOB_ITEMS");
    var logSh = ss.getSheetByName("STATUS_LOG");
    var incoming = (p.jobs || []).slice().sort(function (a, b) { return new Date(a.submittedAt) - new Date(b.submittedAt); });

    var clearedCount = 0;
    if (p.replaceAll) {
      clearedCount = Math.max(0, jobsSh.getLastRow() - 1);
      if (!p.dryRun) {
        // ใช้ clearContent ไม่ใช่ deleteRows เพราะ Google Sheets ห้ามลบแถวที่ไม่ได้ตรึงไว้จนหมดทั้งแท็บ
        // (แท็บที่จำนวนแถวพอดีกับข้อมูลจะลบไม่ผ่าน แล้วทำให้แท็บก่อนหน้าถูกล้างไปแล้วแต่เขียนกลับไม่ได้)
        [jobsSh, itemsSh, logSh].forEach(function (sh) {
          var n = sh.getLastRow() - 1;
          if (n > 0) sh.getRange(2, 1, n, sh.getMaxColumns()).clearContent();
        });
      }
    }

    var seen = {};
    var maxByPrefix = {};
    if (!p.replaceAll) {
      var jv = jobsSh.getDataRange().getValues();
      for (var i = 1; i < jv.length; i++) {
        var sub = jv[i][1] instanceof Date ? jv[i][1].getTime() : new Date(jv[i][1]).getTime();
        seen[String(jv[i][2]).toLowerCase() + "|" + sub + "|" + jv[i][8]] = true;
        var no = String(jv[i][0] || "");
        var m = no.match(/^(PRN-\d{6})-(\d+)$/);
        if (m) maxByPrefix[m[1]] = Math.max(maxByPrefix[m[1]] || 0, parseInt(m[2], 10));
      }
    }

    var jobRows = [], itemRows = [], logRows = [], skipped = 0, created = [];
    incoming.forEach(function (j) {
      var when = new Date(j.submittedAt);
      var key = String(j.email || "").toLowerCase() + "|" + when.getTime() + "|" + (j.purpose || "");
      if (seen[key]) { skipped++; return; }
      seen[key] = true;
      var prefix = "PRN-" + Utilities.formatDate(when, "Asia/Bangkok", "yyyyMM");
      maxByPrefix[prefix] = (maxByPrefix[prefix] || 0) + 1;
      var jobNo = prefix + "-" + ("0000" + maxByPrefix[prefix]).slice(-4);
      var done = j.completedAt ? new Date(j.completedAt) : when;
      var status = j.status || "SERVICE_DONE";
      var it = j.item || {};
      jobRows.push([jobNo, when, j.email || "", j.name || "", j.position || "", j.phone || "", j.unit || "",
        j.needBy || "", j.purpose || "", status,
        j.estimatedAmount === undefined ? "" : j.estimatedAmount,
        j.confirmedAmount === undefined ? "" : j.confirmedAmount,
        status === "SERVICE_DONE" ? done : "", "", j.budgetSource || "", j.budgetAcct || ""]);
      itemRows.push([jobNo + "-1", jobNo, it.serviceCode || "", "", it.qty === undefined ? "" : it.qty, it.unit || "",
        it.unitPrice === undefined ? "" : it.unitPrice,
        it.estimatedAmount === undefined ? "" : it.estimatedAmount,
        it.confirmedAmount === undefined ? "" : it.confirmedAmount, "", "", "",
        it.colorMode || "", it.paperSize || "", it.paperType || "", it.sides || "", it.staple || "", it.ownPaper || "",
        it.qtyOriginal === undefined ? "" : it.qtyOriginal, it.qtySets === undefined ? "" : it.qtySets,
        it.examType || "", it.examSubject || "", it.serviceGroup || "", it.note || ""]);
      var by = j.importBy || "นำเข้าข้อมูลย้อนหลัง";
      logRows.push([Utilities.getUuid(), jobNo, "", "RECEIVED", j.email || by, when, by, "import"]);
      if (status !== "RECEIVED") logRows.push([Utilities.getUuid(), jobNo, "RECEIVED", status, by, done, by, "import"]);
      created.push(jobNo);
    });

    if (!p.dryRun && jobRows.length) {
      var r = jobsSh.getLastRow() + 1;
      jobsSh.getRange(r, 8, jobRows.length, 1).setNumberFormat("@"); // required_date เก็บเป็นข้อความ กันเลื่อนวัน
      jobsSh.getRange(r, 1, jobRows.length, jobRows[0].length).setValues(jobRows);
      var r2 = itemsSh.getLastRow() + 1;
      itemsSh.getRange(r2, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
      var r3 = logSh.getLastRow() + 1;
      logSh.getRange(r3, 1, logRows.length, logRows[0].length).setValues(logRows);
    }
    return { ok: true, dryRun: !!p.dryRun, replaceAll: !!p.replaceAll, cleared: clearedCount,
      received: incoming.length, imported: jobRows.length, skipped: skipped, jobNos: created };
  } finally {
    lock.releaseLock();
  }
}

// ซ่อมเฉพาะคอลัมน์วันที่ของงานที่มีอยู่แล้ว (submitted_at / completed_at) โดยอ้างจาก job_no
// ใช้ตอนที่ข้อมูลนำเข้าย้อนหลังมีวันที่หาย แต่ยอดเงิน/จำนวน/สถานะยังถูกต้อง จึงไม่อยากล้างแล้วนำเข้าใหม่
// อ่าน-เขียนทีละคอลัมน์ (ไม่ใช่ทีละเซลล์) เพราะ 100+ แถวถ้ายิง setValue รายเซลล์จะช้าจนหมดเวลา
function patchJobDates_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSheet_(p);
    var sh = ss.getSheetByName("JOBS");
    var values = sh.getDataRange().getValues();
    var headers = values[0];
    var cNo = headers.indexOf("job_no"), cSub = headers.indexOf("submitted_at"), cDone = headers.indexOf("completed_at");
    if (cNo < 0 || cSub < 0 || cDone < 0) return { ok: false, error: "missing_columns" };
    var n = values.length - 1;
    if (n < 1) return { ok: false, error: "no_rows" };

    var rowOf = {};
    for (var i = 1; i < values.length; i++) rowOf[String(values[i][cNo])] = i - 1;   // index ในอาร์เรย์คอลัมน์

    var subCol = sh.getRange(2, cSub + 1, n, 1).getValues();
    var doneCol = sh.getRange(2, cDone + 1, n, 1).getValues();

    var toDate_ = function (v) {
      if (!v) return null;
      var dt = new Date(v);
      return isNaN(dt.getTime()) ? null : dt;
    };
    var updated = 0, untouched = 0, missing = [], changes = [];
    (p.dates || []).forEach(function (x) {
      var idx = rowOf[String(x.jobNo)];
      if (idx === undefined) { missing.push(x.jobNo); return; }
      var sub = toDate_(x.submittedAt), done = toDate_(x.completedAt);
      var hit = false;
      if (sub && String(subCol[idx][0]) !== String(sub)) { subCol[idx][0] = sub; hit = true; }
      if (done && String(doneCol[idx][0]) !== String(done)) { doneCol[idx][0] = done; hit = true; }
      if (hit) { updated++; if (changes.length < 5) changes.push(x.jobNo); }
      else untouched++;
    });

    if (!p.dryRun && updated) {
      sh.getRange(2, cSub + 1, n, 1).setValues(subCol);
      sh.getRange(2, cDone + 1, n, 1).setValues(doneCol);
    }
    return { ok: true, dryRun: !!p.dryRun, received: (p.dates || []).length,
      updated: updated, untouched: untouched, missing: missing, sample: changes };
  } finally {
    lock.releaseLock();
  }
}

// ---------- รหัสผ่าน ----------
// เก็บในชีตเป็น "sha256:<salt>:<hash>" แทนรหัสจริง คนเปิดชีตจะเห็นเป็นตัวอักษรสุ่ม อ่านย้อนกลับไม่ได้
// ยังรองรับรหัสแบบข้อความธรรมดาไว้ชั่วคราว เพื่อไม่ให้แอดมินที่ยังไม่ได้แปลงล็อกอินไม่ได้
// *** เมื่อแปลงครบทุกคนแล้ว ควรลบ 2 บรรทัดที่ทำ fallback ออก ไม่งั้นแถวที่ยังเป็นข้อความธรรมดาก็ยังใช้ได้อยู่ ***
function hashPassword_(plain, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + plain, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}

function passwordMatches_(input, stored) {
  var st = String(stored || "");
  if (st.indexOf("sha256:") === 0) {
    var parts = st.split(":");
    if (parts.length !== 3) return false;
    return hashPassword_(String(input || ""), parts[1]) === parts[2];
  }
  return String(input || "") === st;   // fallback: รหัสแบบข้อความธรรมดาที่ยังไม่ได้แปลง
}

/**
 * เครื่องมือสำหรับแอดมิน — รันในหน้าต่าง Apps Script Editor เท่านั้น ไม่ได้เปิดเป็น action ทางเว็บ
 * ตั้งใจให้รหัสผ่านจริงไม่ต้องวิ่งผ่านอินเทอร์เน็ตเลย
 *
 * วิธีใช้
 *   1. แก้บรรทัด PLAIN ข้างล่างเป็นรหัสผ่านที่ต้องการ
 *   2. กด Run แล้วเปิด Execution log
 *   3. คัดลอกข้อความที่ขึ้นต้นด้วย sha256: ไปวางทับคอลัมน์ password ของคนนั้นในแท็บ USERS
 *   4. ลบรหัสผ่านที่พิมพ์ไว้ในบรรทัด PLAIN ออก แล้วบันทึก
 */
function makePasswordHash() {
  var PLAIN = "ใส่รหัสผ่านที่ต้องการตรงนี้";
  var salt = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  var out = "sha256:" + salt + ":" + hashPassword_(PLAIN, salt);
  Logger.log("คัดลอกบรรทัดล่างนี้ไปวางในคอลัมน์ password ของแท็บ USERS");
  Logger.log(out);
  return out;
}

// ---------- ราคายืนยันโดยแอดมิน ----------
// เขียน confirmed_amount ลงทั้ง JOB_ITEMS (รายบรรทัด) และ JOBS (ยอดรวมของใบ) พร้อมเหตุผลใน override_reason
// เงื่อนไข: ผู้บันทึกต้องเป็น ADMIN ที่ active อยู่ในแท็บ USERS และงานต้องถึงสถานะ "ผลิตเสร็จ" แล้ว
//
// ข้อจำกัดที่ต้องรู้: ตรวจได้แค่ว่า "อีเมลนี้เป็น ADMIN จริงหรือไม่" แต่ยืนยันไม่ได้ว่าคนส่งคำสั่งคือเจ้าของอีเมล
// เพราะระบบยังไม่มีการยืนยันตัวตน (ข้อ 3 ในแผนพัฒนา ซึ่งตัดสินใจว่ายังไม่ทำ)
function setConfirmedAmount_(p) {
  var ALLOWED_STATUS = { PROD_DONE: true, SERVICE_DONE: true };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSheet_(p);
    var by = String(p.by || "").trim().toLowerCase();
    if (!isActiveAdmin_(ss, by)) return { ok: false, error: "not_admin", message: "ต้องเป็นผู้ดูแลระบบ (ADMIN) เท่านั้นจึงจะยืนยันราคาได้" };

    // ต้องมีรายการที่ใช้ได้อย่างน้อย 1 รายการ ก่อนจะแตะข้อมูลใด ๆ
    // ถ้าไม่กันตรงนี้ คำสั่งที่ส่งรายการว่างมาจะเขียนยอดรวมเป็น 0 ทับของเดิมทันที
    var incoming = p.items || [];
    var valid = [];
    for (var v = 0; v < incoming.length; v++) {
      var amt = Number(incoming[v].amount);
      if (incoming[v].itemId && isFinite(amt) && amt >= 0) valid.push(incoming[v]);
    }
    if (!valid.length) return { ok: false, error: "no_items",
      message: "ต้องส่งราคายืนยันอย่างน้อย 1 รายการ (ไม่มีการแก้ไขข้อมูลใด ๆ)" };

    var jobsSh = ss.getSheetByName("JOBS");
    var jv = jobsSh.getDataRange().getValues();
    var jh = jv[0];
    var cNo = jh.indexOf("job_no"), cStatus = jh.indexOf("status"), cConf = jh.indexOf("confirmed_amount");
    if (cNo < 0 || cStatus < 0 || cConf < 0) return { ok: false, error: "missing_columns" };

    var row = -1;
    for (var i = 1; i < jv.length; i++) if (String(jv[i][cNo]) === String(p.jobNo)) { row = i + 1; break; }
    if (row < 0) return { ok: false, error: "job_not_found", message: "ไม่พบเลขที่งาน " + p.jobNo };

    var status = String(jv[row - 1][cStatus] || "");
    if (!ALLOWED_STATUS[status]) return { ok: false, error: "status_not_allowed",
      message: "ยืนยันราคาได้เมื่องานถึงสถานะ \"งานเสร็จแล้ว\" เป็นต้นไปเท่านั้น (สถานะปัจจุบัน: " + status + ")" };

    // รายบรรทัดใน JOB_ITEMS — p.items = [{ itemId, amount }]
    var itemsSh = ss.getSheetByName("JOB_ITEMS");
    var iv = itemsSh.getDataRange().getValues();
    var ih = iv[0];
    var cItemId = ih.indexOf("item_id"), cJobNo = ih.indexOf("job_no");
    var cIConf = ih.indexOf("confirmed_amount"), cReason = ih.indexOf("override_reason");
    var wanted = {};
    for (var k = 0; k < valid.length; k++) wanted[String(valid[k].itemId)] = valid[k].amount;

    var updated = 0, sum = 0, touched = [];
    for (var r = 1; r < iv.length; r++) {
      if (String(iv[r][cJobNo]) !== String(p.jobNo)) continue;
      var id = String(iv[r][cItemId]);
      var val = wanted.hasOwnProperty(id) ? Number(wanted[id]) : Number(iv[r][cIConf]);
      if (wanted.hasOwnProperty(id) && isFinite(Number(wanted[id]))) {
        itemsSh.getRange(r + 1, cIConf + 1).setValue(Number(wanted[id]));
        if (cReason >= 0 && p.reason) itemsSh.getRange(r + 1, cReason + 1).setValue(String(p.reason));
        updated++; touched.push(id);
      }
      if (isFinite(val)) sum += val;
    }
    sum = Math.round(sum * 100) / 100;
    jobsSh.getRange(row, cConf + 1).setValue(sum);
    logStatus_(ss, p.jobNo, status, status, by, "web",
      "ยืนยันราคา " + sum.toFixed(2) + " บาท" + (p.reason ? (" — " + p.reason) : ""));
    return { ok: true, jobNo: p.jobNo, itemsUpdated: updated, jobTotal: sum, itemIds: touched };
  } finally {
    lock.releaseLock();
  }
}

// ตรวจว่าอีเมลนี้เป็น ADMIN ที่ยังใช้งานอยู่หรือไม่ (ไม่ตรวจรหัสผ่าน ดูหมายเหตุใน setConfirmedAmount_)
function isActiveAdmin_(ss, email) {
  var sh = ss.getSheetByName("USERS");
  if (!sh || !email) return false;
  var v = sh.getDataRange().getValues(), h = v[0];
  var cEm = h.indexOf("email"), cRole = h.indexOf("role"), cAct = h.indexOf("active");
  if (cEm < 0 || cRole < 0) return false;
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][cEm] || "").trim().toLowerCase() !== email) continue;
    var act = cAct < 0 ? true : (v[i][cAct] === true || String(v[i][cAct]).toUpperCase() === "TRUE" || v[i][cAct] === 1 || v[i][cAct] === "1");
    return act && String(v[i][cRole] || "").trim().toUpperCase() === "ADMIN";
  }
  return false;
}

function uploadFile_(p) {
  if (String(p.category || "") === "chat-attachments") return uploadChatAttachment_(p);
  var folder = openFolder_(p);
  var jobFolder = findOrCreateFolder_(folder, p.jobNo);
  var catFolder = findOrCreateFolder_(jobFolder, p.category || "requester-files");
  var bytes = Utilities.base64Decode(p.base64);
  var blob = Utilities.newBlob(bytes, p.mimeType || "application/octet-stream", p.fileName || "file");
  var file = catFolder.createFile(blob);
  var ss = openSheet_(p);
  ss.getSheetByName("FILES").appendRow([
    file.getId(), p.jobNo, p.fileName || "", p.mimeType || "", bytes.length,
    p.category || "requester-files", p.email || "", new Date(), file.getUrl()
  ]);
  return { ok: true, fileId: file.getId(), link: file.getUrl() };
}

function chatViewerEmails_(ss, jobNo, actorEmail) {
  var emails = {}, jobs = readChatJobs_(ss);
  if (isValidChatEmail_(normalizeChatEmail_(actorEmail))) emails[normalizeChatEmail_(actorEmail)] = true;
  if (jobs.ok && jobs.byNo[jobNo] && isValidChatEmail_(jobs.byNo[jobNo].requesterEmail)) {
    emails[jobs.byNo[jobNo].requesterEmail] = true;
  }
  var users = ss.getSheetByName("USERS");
  if (users) {
    var values = users.getDataRange().getValues(), h = values[0] || [];
    var cEmail = h.indexOf("email"), cRole = h.indexOf("role"), cActive = h.indexOf("active");
    if (cEmail >= 0 && cRole >= 0) {
      for (var i = 1; i < values.length; i++) {
        var role = normalizeChatRole_(values[i][cRole]), email = normalizeChatEmail_(values[i][cEmail]);
        if ((role === "ADMIN" || role === "EXECUTIVE") && activeCell_(cActive < 0 ? "" : values[i][cActive], cActive >= 0) && isValidChatEmail_(email)) {
          emails[email] = true;
        }
      }
    }
  }
  return Object.keys(emails);
}

function grantChatFileViewers_(ss, file, jobNo, actorEmail) {
  var failed = 0, emails = chatViewerEmails_(ss, jobNo, actorEmail);
  for (var i = 0; i < emails.length; i++) {
    try { file.addViewer(emails[i]); }
    catch (err) { failed++; }
  }
  return { complete: failed === 0, failedCount: failed };
}

// Chat uploads happen before sendMessage, so validate the same job participant before touching Drive.
// Other upload categories keep their existing behavior for backward compatibility.
function uploadChatAttachment_(p) {
  var jobNo = String(p.jobNo || "").trim();
  if (!jobNo) return { ok: false, error: "missing_job_no" };
  var fileName = String(p.fileName || "file").trim();
  if (!fileName || fileName.length > 255) return { ok: false, error: "invalid_file_name" };
  var extMatch = fileName.toLowerCase().match(/\.([a-z0-9]+)$/);
  var allowedExt = { png: true, jpg: true, jpeg: true, webp: true, gif: true, pdf: true, txt: true, docx: true, xlsx: true };
  if (!extMatch || !allowedExt[extMatch[1]]) return { ok: false, error: "unsupported_file_type" };
  var mimeType = String(p.mimeType || "application/octet-stream").split(";")[0].trim().toLowerCase();
  var mimeByExt = {
    png: { "image/png": true, "application/octet-stream": true },
    jpg: { "image/jpeg": true, "image/jpg": true, "application/octet-stream": true },
    jpeg: { "image/jpeg": true, "image/jpg": true, "application/octet-stream": true },
    webp: { "image/webp": true, "application/octet-stream": true },
    gif: { "image/gif": true, "application/octet-stream": true },
    pdf: { "application/pdf": true, "application/octet-stream": true },
    txt: { "text/plain": true, "application/octet-stream": true },
    docx: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": true, "application/octet-stream": true },
    xlsx: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": true, "application/octet-stream": true }
  };
  if (!mimeByExt[extMatch[1]][mimeType]) return { ok: false, error: "file_type_mismatch" };
  var clientUploadId = String(p.clientUploadId || "").trim();
  if (clientUploadId && (clientUploadId.length > 128 || !/^[A-Za-z0-9._:-]+$/.test(clientUploadId))) {
    return { ok: false, error: "invalid_client_upload_id" };
  }
  var bytes;
  try { bytes = Utilities.base64Decode(String(p.base64 || "")); }
  catch (err) { return { ok: false, error: "invalid_file_data" }; }
  if (!bytes.length) return { ok: false, error: "empty_file" };
  if (bytes.length > 2 * 1024 * 1024) return { ok: false, error: "file_too_large", maxBytes: 2 * 1024 * 1024 };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var lockHeld = true;
  try {
    var ss = openSheet_(p);
    var auth = authorizeChat_(ss, p, jobNo);
    if (!auth.ok) return auth;
    var filesSh = ss.getSheetByName("FILES");
    if (!filesSh) return { ok: false, error: "files_sheet_not_found" };
    var fileValues = filesSh.getDataRange().getValues();
    if (!fileValues.length) return { ok: false, error: "files_schema_invalid" };
    var fileHeaders = fileValues[0];
    var cFileId = fileHeaders.indexOf("file_id"), cJob = fileHeaders.indexOf("job_no");
    var cName = fileHeaders.indexOf("file_name"), cMime = fileHeaders.indexOf("mime_type"), cSize = fileHeaders.indexOf("file_size");
    var cCategory = fileHeaders.indexOf("file_category"), cBy = fileHeaders.indexOf("uploaded_by");
    var cUploadedAt = fileHeaders.indexOf("uploaded_at"), cLink = fileHeaders.indexOf("web_view_link");
    var cClientUpload = fileHeaders.indexOf("client_upload_id");
    if (cFileId < 0 || cJob < 0 || cCategory < 0 || cBy < 0 || cLink < 0) return { ok: false, error: "files_schema_invalid" };
    if (clientUploadId && cClientUpload < 0) return { ok: false, error: "files_schema_outdated", message: "กรุณารัน setupSheets() เพื่อเพิ่ม client_upload_id" };
    if (clientUploadId) {
      for (var i = 1; i < fileValues.length; i++) {
        if (String(fileValues[i][cJob] || "") !== jobNo || normalizeChatEmail_(fileValues[i][cBy]) !== auth.actor.email ||
            String(fileValues[i][cCategory] || "") !== "chat-attachments" || String(fileValues[i][cClientUpload] || "") !== clientUploadId) continue;
        lock.releaseLock();
        lockHeld = false;
        var duplicateSharing = { complete: false, failedCount: 1 };
        try { duplicateSharing = grantChatFileViewers_(ss, DriveApp.getFileById(String(fileValues[i][cFileId] || "")), jobNo, auth.actor.email); }
        catch (shareRetryErr) {}
        return {
          ok: true, duplicate: true,
          fileId: String(fileValues[i][cFileId] || ""), link: safeDriveLink_(fileValues[i][cLink]),
          sharingComplete: duplicateSharing.complete,
          warnings: duplicateSharing.complete ? [] : ["viewer_grant_failed"],
          file: {
            fileId: String(fileValues[i][cFileId] || ""),
            fileName: cName < 0 ? "" : String(fileValues[i][cName] || ""),
            mimeType: cMime < 0 ? "" : String(fileValues[i][cMime] || ""),
            fileSize: cSize < 0 || !isFinite(Number(fileValues[i][cSize])) || Number(fileValues[i][cSize]) < 0 ? "" : Number(fileValues[i][cSize]),
            fileCategory: String(fileValues[i][cCategory] || ""), uploadedBy: auth.actor.email,
            uploadedAt: cUploadedAt < 0 ? "" : isoDate_(fileValues[i][cUploadedAt]),
            webViewLink: safeDriveLink_(fileValues[i][cLink])
          }
        };
      }
    }
    var folder = openFolder_(p);
    var jobFolder = findOrCreateFolder_(folder, jobNo);
    var catFolder = findOrCreateFolder_(jobFolder, "chat-attachments");
    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    var file = catFolder.createFile(blob);
    var now = new Date();
    try {
      var row = [];
      for (var c = 0; c < fileHeaders.length; c++) {
        switch (fileHeaders[c]) {
          case "file_id": row.push(file.getId()); break;
          case "job_no": row.push(jobNo); break;
          case "file_name": row.push(safeSheetText_(fileName)); break;
          case "mime_type": row.push(mimeType); break;
          case "file_size": row.push(bytes.length); break;
          case "file_category": row.push("chat-attachments"); break;
          case "uploaded_by": row.push(auth.actor.email); break;
          case "uploaded_at": row.push(now); break;
          case "web_view_link": row.push(file.getUrl()); break;
          case "client_upload_id": row.push(clientUploadId); break;
          default: row.push("");
        }
      }
      filesSh.getRange(filesSh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    } catch (sheetErr) {
      try { file.setTrashed(true); } catch (trashErr) {}
      return { ok: false, error: "file_registry_write_failed" };
    }
    lock.releaseLock();
    lockHeld = false;
    var sharing = { complete: false, failedCount: 1 };
    try { sharing = grantChatFileViewers_(ss, file, jobNo, auth.actor.email); }
    catch (shareErr) {}
    return {
      ok: true, duplicate: false,
      fileId: file.getId(),
      link: file.getUrl(),
      sharingComplete: sharing.complete,
      warnings: sharing.complete ? [] : ["viewer_grant_failed"],
      file: {
        fileId: file.getId(), fileName: fileName, mimeType: mimeType,
        fileSize: bytes.length, fileCategory: "chat-attachments", uploadedBy: auth.actor.email,
        uploadedAt: now.toISOString(), webViewLink: file.getUrl()
      }
    };
  } finally {
    if (lockHeld) lock.releaseLock();
  }
}

function findOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function listJobs_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName("JOBS");
  var values = sh.getDataRange().getValues();
  var headers = values.shift();
  var jobs = values.map(function (row) {
    var o = {};
    headers.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });

  // เรียงจากไอดีล่าสุดลงไปหาเก่าสุด เพื่อให้หน้า "คิวงาน" ขึ้นคำขอใหม่ล่าสุดไว้บนสุด
  // (ชีตเป็นแบบ append-only ลำดับแถวจึงเป็นเก่า->ใหม่ ถ้าไม่เรียงตรงนี้หน้าเว็บจะโชว์ใบที่เก่าที่สุดก่อน)
  // job_no รูปแบบตายตัว PRN-YYYYMM-NNNN ความยาวเท่ากันทุกใบ เทียบเป็นข้อความจึงได้ลำดับตามเวลาเลย
  // เผื่อกรณีไอดีผิดรูป/ซ้ำ ใช้ submitted_at ตัดสินอีกชั้น
  var jobTime_ = function (o) { var t = new Date(o.submitted_at).getTime(); return isNaN(t) ? 0 : t; };
  jobs.sort(function (a, b) {
    var x = String(a.job_no || ""), y = String(b.job_no || "");
    if (x !== y) return x < y ? 1 : -1;
    return jobTime_(b) - jobTime_(a);
  });

  var itemsSh = ss.getSheetByName("JOB_ITEMS");
  var items = [];
  if (itemsSh && itemsSh.getLastRow() > 1) {
    var iValues = itemsSh.getDataRange().getValues();
    var iHeaders = iValues.shift();
    items = iValues.map(function (row) {
      var o = {};
      iHeaders.forEach(function (h, i) { o[h] = row[i]; });
      return o;
    });
  }

  return { ok: true, jobs: jobs, items: items };
}

// อ่านไฟล์แนบทั้งหมดของงานหนึ่งใบจากแท็บ FILES (รวมสลิปการโอนเงิน category "payment-slip") — ใช้โดย popup รายละเอียดคำขอในหน้า "คิวงาน"
function listFilesForJob_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName("FILES");
  if (!sh) return { ok: true, files: [] };
  var values = sh.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, files: [] };
  var headers = values.shift();
  var jobIdx = headers.indexOf("job_no");
  if (jobIdx === -1) return { ok: true, files: [] };
  var files = values
    .filter(function (row) { return String(row[jobIdx]) === String(p.jobNo); })
    .map(function (row) {
      var o = {};
      headers.forEach(function (h, i) { o[h] = row[i]; });
      return o;
    });
  return { ok: true, files: files };
}

// ---------- แชตประจำงาน (หนึ่ง conversation ต่อ job_no) ----------
function normalizeChatEmail_(raw) {
  return String(raw == null ? "" : raw).trim().toLowerCase();
}

function chatEmailFrom_(p) {
  return normalizeChatEmail_(p.email || p.senderEmail || p.readerEmail || "");
}

function normalizeChatRole_(raw) {
  var role = String(raw == null ? "" : raw).trim().toUpperCase();
  if (!role) return "";
  if (role === "REQUESTER" || role === "REQ" || role === "USER") return "REQUESTER";
  if (role === "ADMIN") return "ADMIN";
  if (role === "EXECUTIVE" || role === "EXEC") return "EXECUTIVE";
  if (role === "STAFF") return "STAFF";
  return null;
}

function isValidChatEmail_(email) {
  return email.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function activeCell_(value, columnExists) {
  if (!columnExists || value === "" || value == null) return true;
  return value === true || value === 1 || value === "1" || String(value).toUpperCase() === "TRUE";
}

function isoDate_(value) {
  if (!value && value !== 0) return "";
  var d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? "" : d.toISOString();
}

function parseChatDate_(value) {
  if (value === undefined || value === null || value === "") return null;
  var d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function safeDriveLink_(value) {
  var link = String(value || "").trim();
  return /^https:\/\/(drive|docs)\.google\.com\//i.test(link) ? link : "";
}

function chatLimit_(raw, fallback) {
  var n = parseInt(raw, 10);
  if (!isFinite(n) || n < 1) n = fallback;
  return Math.min(CHAT_MAX_PAGE_SIZE_, n);
}

function chatCursor_(raw) {
  if (raw === undefined || raw === null || raw === "") return null;
  var n = parseInt(raw, 10);
  return isFinite(n) && n >= 0 ? n : null;
}

function parseArrayParam_(raw) {
  if (raw === undefined || raw === null || raw === "") return { ok: true, value: [] };
  var value = raw;
  if (typeof raw === "string") {
    try { value = JSON.parse(raw); }
    catch (err) { return { ok: false, error: "invalid_array" }; }
  }
  if (!Array.isArray(value)) return { ok: false, error: "invalid_array" };
  return { ok: true, value: value };
}

function readChatJobs_(ss) {
  var sh = ss.getSheetByName("JOBS");
  if (!sh) return { ok: false, error: "jobs_sheet_not_found" };
  var values = sh.getDataRange().getValues();
  if (!values.length) return { ok: false, error: "jobs_schema_invalid" };
  var h = values[0];
  var cNo = h.indexOf("job_no"), cEmail = h.indexOf("requester_email");
  var cName = h.indexOf("requester_name"), cStatus = h.indexOf("status"), cSubmitted = h.indexOf("submitted_at");
  if (cNo < 0 || cEmail < 0) return { ok: false, error: "jobs_schema_invalid" };
  var rows = [], byNo = {}, duplicates = {};
  for (var i = 1; i < values.length; i++) {
    var jobNo = String(values[i][cNo] || "").trim();
    if (!jobNo) continue;
    var job = {
      jobNo: jobNo,
      requesterEmail: normalizeChatEmail_(values[i][cEmail]),
      requesterName: cName < 0 ? "" : String(values[i][cName] || ""),
      status: cStatus < 0 ? "" : String(values[i][cStatus] || ""),
      submittedAt: cSubmitted < 0 ? "" : values[i][cSubmitted]
    };
    if (byNo[jobNo]) duplicates[jobNo] = true;
    else byNo[jobNo] = job;
    rows.push(job);
  }
  return { ok: true, rows: rows, byNo: byNo, duplicates: duplicates };
}

function findActiveChatStaff_(ss, email) {
  var sh = ss.getSheetByName("USERS");
  if (!sh || !email) return null;
  var values = sh.getDataRange().getValues();
  if (!values.length) return null;
  var h = values[0];
  var cEmail = h.indexOf("email"), cRole = h.indexOf("role");
  var cActive = h.indexOf("active"), cName = h.indexOf("full_name");
  if (cEmail < 0 || cRole < 0) return null;
  for (var i = 1; i < values.length; i++) {
    if (normalizeChatEmail_(values[i][cEmail]) !== email) continue;
    var role = normalizeChatRole_(values[i][cRole]);
    if (role !== "ADMIN" && role !== "EXECUTIVE") return null;
    if (!activeCell_(cActive < 0 ? "" : values[i][cActive], cActive >= 0)) return null;
    return { email: email, role: role, name: cName < 0 ? "" : String(values[i][cName] || "") };
  }
  return null;
}

function authorizeChat_(ss, p, jobNo, jobsData) {
  var email = chatEmailFrom_(p);
  if (!email) return { ok: false, error: "missing_email", message: "กรุณาระบุอีเมลผู้ใช้งาน" };
  if (!isValidChatEmail_(email)) return { ok: false, error: "invalid_email", message: "รูปแบบอีเมลไม่ถูกต้อง" };

  var rawRole = p.role || p.senderRole || p.readerRole || "";
  var roleHint = normalizeChatRole_(rawRole);
  if (rawRole && roleHint === null) return { ok: false, error: "invalid_role" };
  jobsData = jobsData || readChatJobs_(ss);
  if (!jobsData.ok) return jobsData;

  var job = null;
  if (jobNo) {
    jobNo = String(jobNo).trim();
    job = jobsData.byNo[jobNo];
    if (!job) return { ok: false, error: "job_not_found" };
    if (jobsData.duplicates[jobNo]) return { ok: false, error: "duplicate_job_no", message: "พบเลขที่งานซ้ำในแท็บ JOBS" };
  }

  if (roleHint === "REQUESTER") {
    if (job && job.requesterEmail !== email) return { ok: false, error: "job_forbidden" };
    // A requester with no jobs may still open chat and receive an empty thread list.
    // Specific-job actions remain ownership checked above.
    return { ok: true, actor: { email: email, role: "REQUESTER", name: job ? job.requesterName : "" }, job: job, jobsData: jobsData };
  }

  var staff = findActiveChatStaff_(ss, email);
  if (roleHint === "ADMIN" || roleHint === "EXECUTIVE" || roleHint === "STAFF") {
    if (!staff) return { ok: false, error: "staff_forbidden", message: "บัญชีเจ้าหน้าที่ไม่พบ ถูกระงับ หรือไม่มีสิทธิ์" };
    return { ok: true, actor: staff, job: job, jobsData: jobsData };
  }

  // If the caller omitted role, resolve it from the same existing sheets instead of trusting a client-side default.
  if (staff) return { ok: true, actor: staff, job: job, jobsData: jobsData };
  if (job) {
    if (job.requesterEmail !== email) return { ok: false, error: "job_forbidden" };
    return { ok: true, actor: { email: email, role: "REQUESTER", name: job.requesterName }, job: job, jobsData: jobsData };
  }
  for (var j = 0; j < jobsData.rows.length; j++) {
    if (jobsData.rows[j].requesterEmail === email) {
      return { ok: true, actor: { email: email, role: "REQUESTER", name: "" }, job: null, jobsData: jobsData };
    }
  }
  return { ok: false, error: "chat_forbidden" };
}

function safeSheetText_(text) {
  text = String(text == null ? "" : text);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function readChatMessages_(ss) {
  var sh = ss.getSheetByName("JOB_MESSAGES");
  if (!sh) return { ok: false, error: "chat_not_setup", message: "ไม่พบแท็บ JOB_MESSAGES กรุณารัน setupSheets()" };
  var range = sh.getDataRange();
  var values = range.getValues();
  if (!values.length) return { ok: false, error: "chat_schema_invalid" };
  var display = range.getDisplayValues ? range.getDisplayValues() : values;
  var h = values[0];
  var cols = {
    id: h.indexOf("message_id"), job: h.indexOf("job_no"), sequence: h.indexOf("sequence"),
    email: h.indexOf("sender_email"), role: h.indexOf("sender_role"), type: h.indexOf("message_type"),
    text: h.indexOf("message_text"), attachments: h.indexOf("attachment_file_ids"),
    created: h.indexOf("created_at"), clientId: h.indexOf("client_message_id")
  };
  for (var key in cols) if (cols.hasOwnProperty(key) && cols[key] < 0) return { ok: false, error: "chat_schema_invalid", missing: key };
  var rows = [], fallbackSequence = {};
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][cols.id] || "");
    var jobNo = String(values[i][cols.job] || "");
    if (!id || !jobNo) continue;
    fallbackSequence[jobNo] = (fallbackSequence[jobNo] || 0) + 1;
    var sequence = parseInt(values[i][cols.sequence], 10);
    if (!isFinite(sequence) || sequence < 1) sequence = fallbackSequence[jobNo];
    var parsedIds = parseArrayParam_(String(values[i][cols.attachments] || ""));
    var shownText = String(display[i][cols.text] == null ? "" : display[i][cols.text]);
    if (display === values && /^'[=+\-@]/.test(shownText)) shownText = shownText.slice(1);
    rows.push({
      messageId: id,
      jobNo: jobNo,
      sequence: sequence,
      senderEmail: normalizeChatEmail_(values[i][cols.email]),
      senderRole: String(values[i][cols.role] || ""),
      messageType: String(values[i][cols.type] || "USER"),
      text: shownText,
      attachmentFileIds: parsedIds.ok ? parsedIds.value.map(function (x) { return String(x); }) : [],
      createdAt: values[i][cols.created],
      clientMessageId: String(values[i][cols.clientId] || ""),
      rowNumber: i + 1
    });
  }
  return { ok: true, sheet: sh, headers: h, columns: cols, rows: rows };
}

function sortChatMessages_(rows) {
  rows.sort(function (a, b) {
    if (a.sequence !== b.sequence) return a.sequence - b.sequence;
    var at = parseChatDate_(a.createdAt), bt = parseChatDate_(b.createdAt);
    var diff = (at ? at.getTime() : 0) - (bt ? bt.getTime() : 0);
    if (diff) return diff;
    return a.rowNumber - b.rowNumber;
  });
  return rows;
}

function messagesByJob_(rows) {
  var out = {};
  rows.forEach(function (m) {
    if (!out[m.jobNo]) out[m.jobNo] = [];
    out[m.jobNo].push(m);
  });
  Object.keys(out).forEach(function (jobNo) { sortChatMessages_(out[jobNo]); });
  return out;
}

function readChatReads_(ss) {
  var sh = ss.getSheetByName("CHAT_READS");
  if (!sh) return { ok: false, error: "chat_not_setup", message: "ไม่พบแท็บ CHAT_READS กรุณารัน setupSheets()" };
  var values = sh.getDataRange().getValues();
  if (!values.length) return { ok: false, error: "chat_reads_schema_invalid" };
  var h = values[0];
  var cols = {
    job: h.indexOf("job_no"), email: h.indexOf("reader_email"), role: h.indexOf("reader_role"),
    readAt: h.indexOf("last_read_at"), messageId: h.indexOf("last_read_message_id"), updatedAt: h.indexOf("updated_at")
  };
  for (var key in cols) if (cols.hasOwnProperty(key) && cols[key] < 0) return { ok: false, error: "chat_reads_schema_invalid", missing: key };
  var byKey = {};
  for (var i = 1; i < values.length; i++) {
    var jobNo = String(values[i][cols.job] || ""), email = normalizeChatEmail_(values[i][cols.email]);
    if (!jobNo || !email) continue;
    byKey[jobNo + "\n" + email] = {
      jobNo: jobNo, readerEmail: email, readerRole: String(values[i][cols.role] || ""),
      lastReadAt: values[i][cols.readAt], lastReadMessageId: String(values[i][cols.messageId] || ""),
      updatedAt: values[i][cols.updatedAt], rowNumber: i + 1
    };
  }
  return { ok: true, sheet: sh, headers: h, columns: cols, byKey: byKey };
}

function fileMetadataMap_(ss, allowedJobs) {
  var out = {}, sh = ss.getSheetByName("FILES");
  if (!sh || sh.getLastRow() <= 1) return out;
  var values = sh.getDataRange().getValues(), h = values[0];
  var cId = h.indexOf("file_id"), cJob = h.indexOf("job_no"), cName = h.indexOf("file_name");
  var cMime = h.indexOf("mime_type"), cSize = h.indexOf("file_size"), cCat = h.indexOf("file_category");
  var cBy = h.indexOf("uploaded_by"), cAt = h.indexOf("uploaded_at"), cLink = h.indexOf("web_view_link");
  if (cId < 0 || cJob < 0) return out;
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][cId] || ""), jobNo = String(values[i][cJob] || "");
    if (!id || !jobNo || (allowedJobs && !allowedJobs[jobNo])) continue;
    out[jobNo + "\n" + id] = {
      fileId: id,
      fileName: cName < 0 ? "" : String(values[i][cName] || ""),
      mimeType: cMime < 0 ? "" : String(values[i][cMime] || ""),
      fileSize: cSize < 0 || !isFinite(Number(values[i][cSize])) || Number(values[i][cSize]) < 0 ? "" : Number(values[i][cSize]),
      fileCategory: cCat < 0 ? "" : String(values[i][cCat] || ""),
      uploadedBy: cBy < 0 ? "" : normalizeChatEmail_(values[i][cBy]),
      uploadedAt: cAt < 0 ? "" : isoDate_(values[i][cAt]),
      webViewLink: cLink < 0 ? "" : safeDriveLink_(values[i][cLink])
    };
  }
  return out;
}

function chatMessageOutput_(m, fileMap) {
  var attachments = [];
  (m.attachmentFileIds || []).forEach(function (id) {
    var meta = fileMap && fileMap[m.jobNo + "\n" + id];
    if (meta) attachments.push(meta);
  });
  return {
    messageId: String(m.messageId || ""),
    jobNo: String(m.jobNo || ""),
    sequence: Number(m.sequence) || 0,
    senderEmail: String(m.senderEmail || ""),
    senderRole: String(m.senderRole || ""),
    messageType: String(m.messageType || "USER"),
    text: String(m.text == null ? "" : m.text),
    attachmentFileIds: (m.attachmentFileIds || []).slice(),
    attachments: attachments,
    createdAt: isoDate_(m.createdAt),
    clientMessageId: String(m.clientMessageId || "")
  };
}

function unreadForJob_(messages, marker, readerEmail) {
  messages = messages || [];
  var start = 0, foundMarker = false;
  if (marker && marker.lastReadMessageId) {
    for (var i = 0; i < messages.length; i++) {
      if (messages[i].messageId === marker.lastReadMessageId) { start = i + 1; foundMarker = true; break; }
    }
  }
  var readAt = marker ? parseChatDate_(marker.lastReadAt) : null;
  var count = 0;
  for (var j = foundMarker ? start : 0; j < messages.length; j++) {
    if (!foundMarker && readAt) {
      var created = parseChatDate_(messages[j].createdAt);
      if (created && created.getTime() <= readAt.getTime()) continue;
    }
    var m = messages[j];
    var mType = String(m.messageType || m.message_type || "").toUpperCase();
    var sRole = String(m.senderRole || m.sender_role || "").toUpperCase();
    var sEmail = normalizeChatEmail_(m.senderEmail || m.sender_email);
    if (mType === "SYSTEM" || sRole === "SYSTEM" || sEmail === "system") continue;
    if (sEmail !== readerEmail) count++;
  }
  return count;
}

function appendChatMessage_(table, message) {
  var row = [];
  for (var i = 0; i < table.headers.length; i++) {
    switch (table.headers[i]) {
      case "message_id": row.push(message.messageId); break;
      case "job_no": row.push(message.jobNo); break;
      case "sequence": row.push(message.sequence); break;
      case "sender_email": row.push(message.senderEmail); break;
      case "sender_role": row.push(message.senderRole); break;
      case "message_type": row.push(message.messageType); break;
      case "message_text": row.push(safeSheetText_(message.text)); break;
      case "attachment_file_ids": row.push(JSON.stringify(message.attachmentFileIds || [])); break;
      case "created_at": row.push(message.createdAt); break;
      case "client_message_id": row.push(message.clientMessageId || ""); break;
      default: row.push("");
    }
  }
  table.sheet.getRange(table.sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function nextChatSequence_(rows, jobNo) {
  var max = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].jobNo === jobNo && rows[i].sequence > max) max = rows[i].sequence;
  }
  return max + 1;
}

function listChatThreads_(p) {
  var ss = openSheet_(p);
  var jobsData = readChatJobs_(ss);
  if (!jobsData.ok) return jobsData;
  var auth = authorizeChat_(ss, p, "", jobsData);
  if (!auth.ok) return auth;
  var messageTable = readChatMessages_(ss);
  if (!messageTable.ok) return messageTable;
  var reads = readChatReads_(ss);
  if (!reads.ok) return reads;

  var grouped = messagesByJob_(messageTable.rows), threads = [], totalUnread = 0;
  var seen = {}, actor = auth.actor;
  for (var i = 0; i < jobsData.rows.length; i++) {
    var job = jobsData.rows[i];
    if (seen[job.jobNo]) continue;
    seen[job.jobNo] = true;
    if (jobsData.duplicates[job.jobNo]) return { ok: false, error: "duplicate_job_no", jobNo: job.jobNo };
    if (actor.role === "REQUESTER" && job.requesterEmail !== actor.email) continue;
    if (String(job.status || "").toUpperCase() === "SERVICE_DONE") continue;
    var messages = grouped[job.jobNo] || [];
    var latest = messages.length ? messages[messages.length - 1] : null;
    var marker = reads.byKey[job.jobNo + "\n" + actor.email] || null;
    var unread = unreadForJob_(messages, marker, actor.email);
    totalUnread += unread;
    var userMessages = messages.filter(function (m) {
      var mType = String(m.messageType || m.message_type || "").toUpperCase();
      var sRole = String(m.senderRole || m.sender_role || "").toUpperCase();
      var sEmail = normalizeChatEmail_(m.senderEmail || m.sender_email);
      return mType !== "SYSTEM" && sRole !== "SYSTEM" && sEmail !== "system";
    });
    var latestUserMsg = userMessages.length ? userMessages[userMessages.length - 1] : null;
    var activity = latestUserMsg ? latestUserMsg.createdAt : job.submittedAt;
    threads.push({
      jobNo: job.jobNo,
      requesterEmail: job.requesterEmail,
      requesterName: job.requesterName,
      status: job.status,
      submittedAt: isoDate_(job.submittedAt),
      lastMessageAt: isoDate_(activity),
      lastSequence: latest ? latest.sequence : 0,
      lastMessage: latestUserMsg ? chatMessageOutput_(latestUserMsg, null) : null,
      unreadCount: unread,
      _activityMs: parseChatDate_(activity) ? parseChatDate_(activity).getTime() : 0
    });
  }
  threads.sort(function (a, b) {
    if (a._activityMs !== b._activityMs) return b._activityMs - a._activityMs;
    return a.jobNo < b.jobNo ? 1 : (a.jobNo > b.jobNo ? -1 : 0);
  });

  var updatedAfter = p.updatedAfter === undefined ? null : parseChatDate_(p.updatedAfter);
  if (p.updatedAfter !== undefined && p.updatedAfter !== "" && !updatedAfter) return { ok: false, error: "invalid_updated_after" };
  if (updatedAfter) {
    var afterMs = updatedAfter.getTime();
    threads = threads.filter(function (t) { return t._activityMs > afterMs; });
  }

  // By default return every accessible job, including jobs with no messages. Optional cursor/limit is for very large staff queues.
  var offset = chatCursor_(p.cursor);
  if (p.cursor !== undefined && p.cursor !== "" && offset === null) return { ok: false, error: "invalid_cursor" };
  offset = offset || 0;
  var useLimit = p.limit !== undefined && p.limit !== null && p.limit !== "";
  var limit = useLimit ? chatLimit_(p.limit, CHAT_MAX_PAGE_SIZE_) : Math.max(threads.length, 1);
  var page = threads.slice(offset, offset + limit), hasMore = offset + page.length < threads.length;
  page.forEach(function (t) { delete t._activityMs; });
  return {
    ok: true,
    threads: page,
    totalUnread: totalUnread,
    nextCursor: hasMore ? String(offset + page.length) : null,
    hasMore: hasMore,
    serverTime: new Date().toISOString()
  };
}

function listMessages_(p) {
  var jobNo = String(p.jobNo || "").trim();
  if (!jobNo) return { ok: false, error: "missing_job_no" };
  var ss = openSheet_(p);
  var auth = authorizeChat_(ss, p, jobNo);
  if (!auth.ok) return auth;
  var table = readChatMessages_(ss);
  if (!table.ok) return table;
  var all = table.rows.filter(function (m) { return m.jobNo === jobNo; });
  sortChatMessages_(all);
  var limit = chatLimit_(p.limit, 50), page = [], hasMore = false, nextCursor = null;
  var baseSequence = 0;

  if (p.afterSequence !== undefined && p.afterSequence !== null && p.afterSequence !== "") {
    var afterSequence = parseInt(p.afterSequence, 10);
    if (!isFinite(afterSequence) || afterSequence < 0) return { ok: false, error: "invalid_after_sequence" };
    baseSequence = afterSequence;
    var newer = all.filter(function (m) { return m.sequence > afterSequence; });
    page = newer.slice(0, limit);
    hasMore = page.length < newer.length;
  } else if (p.afterMessageId) {
    var afterIndex = -1;
    for (var i = 0; i < all.length; i++) if (all[i].messageId === String(p.afterMessageId)) { afterIndex = i; break; }
    if (afterIndex < 0) return { ok: false, error: "message_not_found" };
    baseSequence = all[afterIndex].sequence;
    page = all.slice(afterIndex + 1, afterIndex + 1 + limit);
    hasMore = afterIndex + 1 + page.length < all.length;
  } else if (p.after !== undefined && p.after !== null && p.after !== "") {
    var afterDate = parseChatDate_(p.after);
    if (!afterDate) return { ok: false, error: "invalid_after" };
    var newerByDate = all.filter(function (m) {
      var d = parseChatDate_(m.createdAt);
      return d && d.getTime() > afterDate.getTime();
    });
    page = newerByDate.slice(0, limit);
    hasMore = page.length < newerByDate.length;
  } else {
    var end = chatCursor_(p.cursor);
    if (p.cursor !== undefined && p.cursor !== "" && end === null) return { ok: false, error: "invalid_cursor" };
    if (end === null) end = all.length;
    end = Math.min(end, all.length);
    var start = Math.max(0, end - limit);
    page = all.slice(start, end);
    hasMore = start > 0;
    nextCursor = hasMore ? String(start) : null;
  }

  var allowed = {}; allowed[jobNo] = true;
  var files = fileMetadataMap_(ss, allowed);
  var outputs = page.map(function (m) { return chatMessageOutput_(m, files); });
  var lastSequence = outputs.length ? outputs[outputs.length - 1].sequence : baseSequence;
  return {
    ok: true,
    jobNo: jobNo,
    messages: outputs,
    lastSequence: lastSequence,
    nextCursor: nextCursor,
    hasMore: hasMore,
    serverTime: new Date().toISOString()
  };
}

function sendMessage_(p) {
  var jobNo = String(p.jobNo || "").trim();
  if (!jobNo) return { ok: false, error: "missing_job_no" };
  var clientMessageId = String(p.clientMessageId || "").trim();
  if (!clientMessageId) return { ok: false, error: "missing_client_message_id" };
  if (clientMessageId.length > 128 || !/^[A-Za-z0-9._:-]+$/.test(clientMessageId)) {
    return { ok: false, error: "invalid_client_message_id" };
  }

  var textValue = p.messageText !== undefined ? p.messageText : (p.text !== undefined ? p.text : p.message);
  var messageText = textValue == null ? "" : String(textValue);
  if (messageText.length > CHAT_MAX_MESSAGE_LENGTH_) {
    return { ok: false, error: "message_too_long", maxLength: CHAT_MAX_MESSAGE_LENGTH_ };
  }
  var parsedAttachments = parseArrayParam_(p.attachmentFileIds);
  if (!parsedAttachments.ok) return { ok: false, error: "invalid_attachment_file_ids" };
  if (parsedAttachments.value.length > CHAT_MAX_ATTACHMENTS_) {
    return { ok: false, error: "too_many_attachments", maxAttachments: CHAT_MAX_ATTACHMENTS_ };
  }
  var attachmentIds = [], attachmentSeen = {};
  for (var a = 0; a < parsedAttachments.value.length; a++) {
    var fileId = String(parsedAttachments.value[a] || "").trim();
    if (!fileId || fileId.length > 200 || !/^[A-Za-z0-9_-]+$/.test(fileId)) return { ok: false, error: "invalid_attachment_file_id" };
    if (!attachmentSeen[fileId]) { attachmentSeen[fileId] = true; attachmentIds.push(fileId); }
  }
  if (!messageText.trim() && !attachmentIds.length) return { ok: false, error: "empty_message" };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSheet_(p);
    var auth = authorizeChat_(ss, p, jobNo);
    if (!auth.ok) return auth;
    var table = readChatMessages_(ss);
    if (!table.ok) return table;

    for (var i = 0; i < table.rows.length; i++) {
      var existing = table.rows[i];
      if (existing.jobNo === jobNo && existing.senderEmail === auth.actor.email && existing.clientMessageId === clientMessageId) {
        var duplicateAllowed = {}; duplicateAllowed[jobNo] = true;
        return {
          ok: true,
          duplicate: true,
          message: chatMessageOutput_(existing, fileMetadataMap_(ss, duplicateAllowed)),
          serverTime: new Date().toISOString()
        };
      }
    }

    var allowed = {}; allowed[jobNo] = true;
    var fileMap = fileMetadataMap_(ss, allowed), missing = [], forbiddenAttachments = [];
    for (var f = 0; f < attachmentIds.length; f++) {
      var attachment = fileMap[jobNo + "\n" + attachmentIds[f]];
      if (!attachment) missing.push(attachmentIds[f]);
      else if (attachment.fileCategory !== "chat-attachments") forbiddenAttachments.push(attachmentIds[f]);
    }
    if (missing.length) return { ok: false, error: "attachment_not_found", fileIds: missing };
    if (forbiddenAttachments.length) return { ok: false, error: "attachment_category_not_allowed", fileIds: forbiddenAttachments };

    var now = new Date();
    var message = {
      messageId: Utilities.getUuid(),
      jobNo: jobNo,
      sequence: nextChatSequence_(table.rows, jobNo),
      senderEmail: auth.actor.email,
      senderRole: auth.actor.role,
      messageType: "USER",
      text: messageText,
      attachmentFileIds: attachmentIds,
      createdAt: now,
      clientMessageId: clientMessageId
    };
    appendChatMessage_(table, message);
    return { ok: true, duplicate: false, message: chatMessageOutput_(message, fileMap), serverTime: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function upsertChatRead_(reads, actor, jobNo, lastReadAt, messageId, now) {
  var key = jobNo + "\n" + actor.email;
  var old = reads.byKey[key] || null;
  var row = [];
  for (var i = 0; i < reads.headers.length; i++) {
    switch (reads.headers[i]) {
      case "job_no": row.push(jobNo); break;
      case "reader_email": row.push(actor.email); break;
      case "reader_role": row.push(actor.role); break;
      case "last_read_at": row.push(lastReadAt); break;
      case "last_read_message_id": row.push(messageId || ""); break;
      case "updated_at": row.push(now); break;
      default: row.push("");
    }
  }
  var rowNumber = old ? old.rowNumber : reads.sheet.getLastRow() + 1;
  reads.sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return {
    jobNo: jobNo, readerEmail: actor.email, readerRole: actor.role,
    lastReadAt: lastReadAt, lastReadMessageId: messageId || "", updatedAt: now, rowNumber: rowNumber
  };
}

function markChatRead_(p) {
  var jobNo = String(p.jobNo || "").trim();
  if (!jobNo) return { ok: false, error: "missing_job_no" };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSheet_(p);
    var auth = authorizeChat_(ss, p, jobNo);
    if (!auth.ok) return auth;
    var table = readChatMessages_(ss);
    if (!table.ok) return table;
    var reads = readChatReads_(ss);
    if (!reads.ok) return reads;
    var messages = table.rows.filter(function (m) { return m.jobNo === jobNo; });
    sortChatMessages_(messages);

    var target = null, requestedId = String(p.messageId || p.lastMessageId || p.lastReadMessageId || "");
    var requestedSequenceRaw = p.sequence !== undefined ? p.sequence : p.lastSequence;
    if (requestedId) {
      for (var i = 0; i < messages.length; i++) if (messages[i].messageId === requestedId) { target = messages[i]; break; }
      if (!target) return { ok: false, error: "message_not_found" };
    } else if (requestedSequenceRaw !== undefined && requestedSequenceRaw !== null && requestedSequenceRaw !== "") {
      var requestedSequence = parseInt(requestedSequenceRaw, 10);
      if (!isFinite(requestedSequence) || requestedSequence < 0) return { ok: false, error: "invalid_sequence" };
      for (var s = 0; s < messages.length; s++) if (messages[s].sequence <= requestedSequence) target = messages[s];
      if (requestedSequence > 0 && !target) return { ok: false, error: "message_not_found" };
    } else if (messages.length) target = messages[messages.length - 1];

    var old = reads.byKey[jobNo + "\n" + auth.actor.email] || null;
    if (old && old.lastReadMessageId && target) {
      var oldIndex = -1, targetIndex = -1;
      for (var j = 0; j < messages.length; j++) {
        if (messages[j].messageId === old.lastReadMessageId) oldIndex = j;
        if (messages[j].messageId === target.messageId) targetIndex = j;
      }
      if (oldIndex > targetIndex) target = messages[oldIndex];
    }
    var now = new Date();
    var lastReadAt = target ? (parseChatDate_(target.createdAt) || now) : now;
    var marker = upsertChatRead_(reads, auth.actor, jobNo, lastReadAt, target ? target.messageId : "", now);
    return {
      ok: true,
      jobNo: jobNo,
      lastReadAt: isoDate_(lastReadAt),
      lastReadMessageId: marker.lastReadMessageId,
      unreadCount: unreadForJob_(messages, marker, auth.actor.email),
      serverTime: now.toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}

function getUnreadCounts_(p) {
  var ss = openSheet_(p);
  var jobsData = readChatJobs_(ss);
  if (!jobsData.ok) return jobsData;
  var auth = authorizeChat_(ss, p, "", jobsData);
  if (!auth.ok) return auth;
  var table = readChatMessages_(ss);
  if (!table.ok) return table;
  var reads = readChatReads_(ss);
  if (!reads.ok) return reads;
  var grouped = messagesByJob_(table.rows), wanted = null;
  if (p.jobNos !== undefined && p.jobNos !== null && p.jobNos !== "") {
    var parsed = parseArrayParam_(p.jobNos);
    if (!parsed.ok || parsed.value.length > CHAT_MAX_PAGE_SIZE_) return { ok: false, error: "invalid_job_nos" };
    wanted = {};
    for (var w = 0; w < parsed.value.length; w++) wanted[String(parsed.value[w])] = true;
  }

  var byJob = {}, total = 0, seen = {};
  for (var i = 0; i < jobsData.rows.length; i++) {
    var job = jobsData.rows[i];
    if (seen[job.jobNo]) continue;
    seen[job.jobNo] = true;
    if (jobsData.duplicates[job.jobNo]) return { ok: false, error: "duplicate_job_no", jobNo: job.jobNo };
    if (auth.actor.role === "REQUESTER" && job.requesterEmail !== auth.actor.email) continue;
    if (String(job.status || "").toUpperCase() === "SERVICE_DONE") continue;
    if (wanted && !wanted[job.jobNo]) continue;
    var marker = reads.byKey[job.jobNo + "\n" + auth.actor.email] || null;
    var n = unreadForJob_(grouped[job.jobNo] || [], marker, auth.actor.email);
    byJob[job.jobNo] = n;
    total += n;
  }
  if (wanted) {
    var requested = Object.keys(wanted);
    for (var r = 0; r < requested.length; r++) if (!byJob.hasOwnProperty(requested[r])) return { ok: false, error: "job_forbidden", jobNo: requested[r] };
  }
  return { ok: true, totalUnread: total, byJob: byJob, counts: byJob, serverTime: new Date().toISOString() };
}

function appendStatusSystemMessage_(ss, jobNo, oldStatus, newStatus) {
  var table = readChatMessages_(ss);
  if (!table.ok) return false; // Keep the pre-chat updateStatus behavior if setupSheets has not been rerun yet.
  var oldText = String(oldStatus == null ? "" : oldStatus).slice(0, 200);
  var newText = String(newStatus == null ? "" : newStatus).slice(0, 200);
  var now = new Date();
  appendChatMessage_(table, {
    messageId: Utilities.getUuid(),
    jobNo: String(jobNo),
    sequence: nextChatSequence_(table.rows, String(jobNo)),
    senderEmail: "system",
    senderRole: "SYSTEM",
    messageType: "SYSTEM",
    text: "สถานะงานเปลี่ยนจาก " + (oldText || "-") + " เป็น " + (newText || "-"),
    attachmentFileIds: [],
    createdAt: now,
    clientMessageId: ""
  });
  return true;
}

function checkAdminUser_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName("USERS");
  if (!sh) return { ok: false, error: "users_sheet_not_found", message: "ไม่พบแท็บ USERS ใน Google Sheet" };
  var values = sh.getDataRange().getValues();
  if (values.length <= 1) return { ok: false, error: "no_users_found", message: "ยังไม่มีข้อมูลในแท็บ USERS (กรุณาเพิ่มอีเมลเจ้าหน้าที่ในแท็บ USERS)" };
  var headers = values.shift();
  var emailIdx = headers.indexOf("email");
  var roleIdx = headers.indexOf("role");
  var activeIdx = headers.indexOf("active");
  var nameIdx = headers.indexOf("full_name");
  var unitIdx = headers.indexOf("unit_code");
  var passIdx = headers.indexOf("password");

  var targetEmail = String(p.email || "").trim().toLowerCase();
  var targetPass = String(p.password || "").trim();

  // 1. ค้นหาผู้ใช้ในแท็บ USERS
  var matchedUser = null;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var em = String(row[emailIdx] || "").trim().toLowerCase();
    if (em === targetEmail) {
      matchedUser = row;
      break;
    }
  }

  if (!matchedUser) {
    return { ok: false, error: "user_not_found", message: "ไม่พบอีเมลนี้ในรายชื่อผู้ใช้ของระบบ (แท็บ USERS)" };
  }

  // 2. ตรวจสอบสถานะ active
  var rawActive = matchedUser[activeIdx];
  var isActive = rawActive === true || String(rawActive).toUpperCase() === "TRUE" || rawActive === 1 || rawActive === "1" || rawActive === "";
  if (!isActive) {
    return { ok: false, error: "user_inactive", message: "บัญชีนี้ถูกระงับการใช้งานในระบบ" };
  }

  // 3. ตรวจสอบบทบาท role
  var role = String(matchedUser[roleIdx] || "").trim().toUpperCase();
  if (role !== "ADMIN" && role !== "EXECUTIVE") {
    return { ok: false, error: "not_admin", message: "อีเมลนี้ไม่มีสิทธิ์ระดับเจ้าหน้าที่โรงพิมพ์ (สิทธิ์ปัจจุบัน: " + role + ")" };
  }

  // 4. ดึงรหัสผ่านที่ถูกต้องจาก Google Sheet (แท็บ USERS คอลัมน์ password หรือแท็บ SETTINGS คีย์ admin_password)
  var expectedPass = "";
  if (passIdx !== -1 && String(matchedUser[passIdx] || "").trim()) {
    expectedPass = String(matchedUser[passIdx]).trim();
  } else {
    var settingsSh = ss.getSheetByName("SETTINGS");
    if (settingsSh) {
      var sValues = settingsSh.getDataRange().getValues();
      for (var s = 1; s < sValues.length; s++) {
        var sKey = String(sValues[s][0] || "").trim().toLowerCase();
        if (sKey === "admin_password" || sKey === "admin_pin") {
          expectedPass = String(sValues[s][1] || "").trim();
          break;
        }
      }
      // ถ้ายังไม่มีแถว admin_password ใน SETTINGS ให้สร้างแถวเริ่มต้นให้อัตโนมัติ เพื่อให้ผู้ใช้เข้าไปดูและแก้ได้
      if (!expectedPass) {
        expectedPass = "admin1234";
        settingsSh.appendRow(["admin_password", expectedPass, "system", new Date(), "รหัสผ่านสำหรับเข้าสู่ระบบแอดมินโรงพิมพ์ (สามารถเปลี่ยนรหัสได้ที่นี่)"]);
      }
    }
  }

  // 5. ตรวจสอบรหัสผ่านที่ผู้ใช้ส่งมา
  if (!expectedPass || !passwordMatches_(targetPass, expectedPass)) {
    return { ok: false, error: "invalid_password", message: "รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง" };
  }

  return {
    ok: true,
    isAdmin: true,
    role: role,
    email: targetEmail,
    name: String(matchedUser[nameIdx] || targetEmail),
    unit: String(matchedUser[unitIdx] || "")
  };
}

function listUsers_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName("USERS");
  if (!sh) return { ok: false, error: "users_sheet_not_found" };
  var values = sh.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, users: [] };
  var headers = values.shift();
  var users = values.map(function (row) {
    var o = {};
    headers.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
  return { ok: true, users: users };
}

// หาคอลัมน์จากรายชื่อหัวตารางที่ยอมรับได้หลายแบบ (รองรับทั้งหัวตารางภาษาอังกฤษและภาษาไทยที่อาจกรอกไว้เอง)
function findColIdx_(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

// ---------- ตรวจจับอีเมลผู้ขอกับทะเบียนบุคลากร (หน้า gate ตอน "เริ่มใช้บริการ") ----------
// เทียบอีเมลที่กรอกกับคอลัมน์ requester_email ในแท็บ Personnel_Information ถ้าพบ ดึงชื่อ-สกุล และหน่วยงาน มาเติมฟอร์มให้อัตโนมัติ
function lookupPersonnel_(p) {
  var ss = openSheet_(p);
  var sh = ss.getSheetByName(PERSONNEL_SHEET_NAME);
  if (!sh) return { ok: false, error: "personnel_sheet_not_found", message: "ไม่พบแท็บ " + PERSONNEL_SHEET_NAME + " ใน Google Sheet" };

  var values = sh.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, found: false };

  var headers = values.shift();
  var emailIdx = findColIdx_(headers, ["requester_email", "email"]);
  var nameIdx = findColIdx_(headers, ["full_name", "ชื่อ-สกุล", "ชื่อ - สกุล", "name"]);
  var unitIdx = findColIdx_(headers, ["unit_name", "หน่วยงาน", "unit", "unit_code"]);
  if (emailIdx === -1) return { ok: false, error: "email_column_not_found", message: "ไม่พบคอลัมน์อีเมลในแท็บ " + PERSONNEL_SHEET_NAME };

  var target = String(p.email || "").trim().toLowerCase();
  if (!target) return { ok: true, found: false };

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var em = String(row[emailIdx] || "").trim().toLowerCase();
    if (em === target) {
      return {
        ok: true, found: true,
        email: String(row[emailIdx] || "").trim(),
        name: nameIdx !== -1 ? String(row[nameIdx] || "").trim() : "",
        unit: unitIdx !== -1 ? String(row[unitIdx] || "").trim() : ""
      };
    }
  }
  return { ok: true, found: false };
}

// ---------- ส่งเอกสารสรุปรายงานทางอีเมล (หน้า "เอกสารสรุปรายงาน" ของผู้บริหาร) ----------
// ผู้ใช้ต้องกรอกอีเมลปลายทางในหน้าเว็บทุกครั้งก่อนกดส่ง ไม่มีอีเมลตายตัวฝังอยู่ในระบบ

/**
 * รันฟังก์ชันนี้ "ครั้งเดียว" จากตัวแก้ไข Apps Script โดยตรง (เลือก authorizeMailSend จาก dropdown แล้วกด Run)
 * เพื่อขออนุมัติสิทธิ์ "ส่งอีเมลในนามของคุณ" ให้กับสคริปต์ — ต้องทำครั้งแรกก่อนใช้ปุ่ม "ส่งอีเมล" ในหน้าเว็บ
 * เพราะการเรียกผ่าน Web App (fetch จากหน้าเว็บ) ไม่สามารถเด้งหน้าต่างขอสิทธิ์ใหม่ให้กดอนุมัติได้เอง
 * รันแล้วเช็คอีเมลของบัญชีที่ deploy สคริปต์นี้ไว้ ควรได้รับอีเมลทดสอบ 1 ฉบับ
 * (ตั้งชื่อไม่มีขีดล่างต่อท้าย เพราะ Apps Script จะไม่แสดงฟังก์ชันที่ลงท้ายด้วย "_" ใน dropdown เลือกรัน)
 */
function authorizeMailSend() {
  var me = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(me, "ทดสอบสิทธิ์ส่งอีเมล — SPU Printing Service",
    "ถ้าคุณได้รับอีเมลฉบับนี้ แปลว่า Apps Script มีสิทธิ์ส่งอีเมลเรียบร้อยแล้ว ปุ่ม \"ส่งอีเมล\" ในหน้าเว็บใช้งานได้ตามปกติ");
  Logger.log("ส่งอีเมลทดสอบไปที่ " + me + " แล้ว — ถ้าไม่มี error แปลว่าสิทธิ์อนุมัติเรียบร้อย");
}

function sendReportEmail_(p) {
  var to = String(p.to || "").trim();
  if (!to) return { ok: false, error: "missing_to", message: "กรุณากรอกอีเมลปลายทาง" };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    return { ok: false, error: "invalid_email", message: "รูปแบบอีเมลไม่ถูกต้อง" };
  }

  var periodText = String(p.periodText || "");
  var subject = "รายงานสรุปการใช้บริการโรงพิมพ์" + (periodText ? " — " + periodText : "");
  var html = buildReportEmailHtml_(periodText, p.rows, p.revenue);

  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
  return { ok: true, to: to };
}

function buildReportEmailHtml_(periodText, rows, revenue) {
  rows = Array.isArray(rows) ? rows : [];
  revenue = Array.isArray(revenue) ? revenue : [];

  var rowsHtml = rows.map(function (r) {
    var headStyle = r.head ? "font-weight:700;background:#f2f2f2;" : "";
    var firstCell = '<td style="border:1px solid #ddd;padding:6px 8px;' + headStyle + '">' + escapeHtml_(r.name || "") + "</td>";
    var restCells = (r.cells || []).map(function (v) {
      return '<td style="border:1px solid #ddd;padding:6px 8px;text-align:right;' + headStyle + '">' + escapeHtml_(String(v == null ? "" : v)) + "</td>";
    }).join("");
    return "<tr>" + firstCell + restCells + "</tr>";
  }).join("");

  var revenueHtml = revenue.map(function (r) {
    return '<tr><td style="padding:5px 8px;border-bottom:1px solid #eee;">' + escapeHtml_(r.label || "") + '</td>'
      + '<td style="padding:5px 8px;border-bottom:1px solid #eee;text-align:right;font-variant-numeric:tabular-nums;">' + escapeHtml_(r.value || "") + "</td></tr>";
  }).join("");

  return ""
    + '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111;max-width:720px">'
    + '<h2 style="margin:0 0 4px">สรุปการขอใช้บริการ โรงพิมพ์</h2>'
    + '<div style="color:#555;margin-bottom:16px">' + escapeHtml_(periodText) + "</div>"
    + '<table style="border-collapse:collapse;width:100%;font-size:12.5px">' + rowsHtml + "</table>"
    + '<h3 style="margin:20px 0 8px">สรุปรายได้</h3>'
    + '<table style="border-collapse:collapse;width:100%;font-size:13px">' + revenueHtml + "</table>"
    + '<div style="margin-top:20px;color:#888;font-size:11.5px">อีเมลนี้ส่งจากระบบใช้บริการโรงพิมพ์ มหาวิทยาลัยศรีปทุม โดยอัตโนมัติ ข้อมูลคำนวณจาก Google Sheet ล่าสุด ณ เวลาที่ส่ง</div>'
    + "</div>";
}

function escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
