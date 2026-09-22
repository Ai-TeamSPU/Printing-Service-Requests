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
  JOB_ITEMS: ["item_id","job_no","service_code","variant_snapshot","quantity","unit","unit_price","estimated_amount","confirmed_amount","price_rule_id","price_rule_snapshot","override_reason","color_mode","paper_size","paper_type","sides","staple","own_paper","qty_original","qty_sets","exam_type","exam_subject","service_group"],
  PRICE_RULES: ["rule_id","service_code","condition","price_type","price","min_price","max_price","effective_from","effective_to","active","note"],
  UNITS: ["unit_code","unit_name","parent_group","display_order","active"],
  USERS: ["email","full_name","role","unit_code","phone","active","password"],
  FILES: ["file_id","job_no","file_name","mime_type","file_size","file_category","uploaded_by","uploaded_at","web_view_link"],
  STATUS_LOG: ["log_id","job_no","old_status","new_status","changed_by","changed_at","note","channel"],
  NOTIFY_LOG: ["notify_id","job_no","template_code","recipient","status","sent_at","error"],
  SETTINGS: ["key","value","updated_by","updated_at","note"]
};

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
    switch (p.action) {
      case "testSheet": return json_(testSheet_(p));
      case "testDrive": return json_(testDrive_(p));
      case "testAccess": return json_(testAccess_(p));
      case "submitJob": return json_(submitJob_(p));
      case "updateStatus": return json_(updateStatus_(p));
      case "importJobs": return json_(importJobs_(p));
      case "uploadFile": return json_(uploadFile_(p));
      case "listJobs": return json_(listJobs_(p));
      case "checkAdminUser": return json_(checkAdminUser_(p));
      case "listUsers": return json_(listUsers_(p));
      case "sendReportEmail": return json_(sendReportEmail_(p));
      case "lookupPersonnel": return json_(lookupPersonnel_(p));
      case "listPriceRules": return json_(listPriceRules_(p));
      case "togglePriceRule": return json_(togglePriceRule_(p));
      case "listFilesForJob": return json_(listFilesForJob_(p));
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
  return { ok: missing.length === 0, name: ss.getName(), tabs: names, missing: missing };
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

function logStatus_(ss, jobNo, oldStatus, newStatus, by, channel, note) {
  ss.getSheetByName("STATUS_LOG").appendRow([
    Utilities.getUuid(), jobNo, oldStatus || "", newStatus, by || "", new Date(), note || "", channel || "web"
  ]);
}

function submitJob_(p) {
  var ss = openSheet_(p);
  var jobsSh = ss.getSheetByName("JOBS");
  var itemsSh = ss.getSheetByName("JOB_ITEMS");
  var jobNo = nextJobNo_(jobsSh);

  // สร้างโฟลเดอร์งานบน Drive แบบ "ทำได้ก็ทำ" — ถ้ายังไม่ได้ตั้งค่า Drive (ไม่มี folderId) หรือเขียนไม่ได้
  // ให้ข้ามส่วนนี้ไป ไม่ทำให้การบันทึกแถวลง Sheet ล้มเหลวตามไปด้วย
  var jobFolderId = "";
  if (p.folderId) {
    try {
      var folder = openFolder_(p);
      var jobFolder = folder.createFolder(jobNo);
      // สร้างแค่ "requester-files" ล่วงหน้า (โฟลเดอร์เดียวที่ใช้จริงตอนส่งคำขอ) — admin-files/proof/final/payment-slip
      // ยังไม่สร้างตอนนี้ จะสร้างเองอัตโนมัติทีหลังตอนมีการอัปโหลดไฟล์เข้าหมวดนั้นจริง ๆ ผ่าน findOrCreateFolder_ กันโฟลเดอร์เปล่าคาอยู่ใน Drive
      jobFolder.createFolder("requester-files");
      jobFolderId = jobFolder.getId();
    } catch (err) {
      // เก็บงานลง Sheet ต่อไปได้ แม้ Drive จะยังเชื่อมต่อไม่ได้
    }
  }

  jobsSh.appendRow([
    jobNo, new Date(), p.email || "", p.name || "", p.position || "", p.phone || "",
    p.unit || "", p.needBy || "", p.purpose || "", "RECEIVED",
    p.estimatedAmount || "", "", "", jobFolderId, p.budgetSource || "", p.budgetAcct || ""
  ]);
  (p.items || []).forEach(function (it, i) {
    itemsSh.appendRow([
      jobNo + "-" + (i + 1), jobNo, it.serviceCode || "", "", // variant_snapshot: เลิกเขียนแล้ว (ทางเลือก A) ใช้คอลัมน์แยกด้านล่างแทน
      it.qty || "", it.unit || "", it.unitPrice || "", it.estimatedAmount || "",
      "", it.priceRuleId || "", it.priceRuleSnapshot || "", "",
      it.colorMode || "", it.paperSize || "", it.paperType || "", it.sides || "",
      it.staple || "", it.ownPaper || "", it.qtyOriginal || "", it.qtySets || "",
      it.examType || "", it.examSubject || "", it.serviceGroup || ""
    ]);
  });
  logStatus_(ss, jobNo, "", "RECEIVED", p.email || "system", "web", "ส่งคำขอใหม่ผ่านเว็บ");
  return { ok: true, jobNo: jobNo, folderId: jobFolderId };
}

function updateStatus_(p) {
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
  return { ok: true };
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
        [jobsSh, itemsSh, logSh].forEach(function (sh) {
          var n = sh.getLastRow() - 1;
          if (n > 0) sh.deleteRows(2, n);
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
        it.examType || "", it.examSubject || "", it.serviceGroup || ""]);
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

function uploadFile_(p) {
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
  if (!expectedPass || targetPass !== expectedPass) {
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
