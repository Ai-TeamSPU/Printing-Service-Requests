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
  JOB_ITEMS: ["item_id","job_no","service_code","variant_snapshot","quantity","unit","unit_price","estimated_amount","confirmed_amount","price_rule_id","price_rule_snapshot","override_reason"],
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
      case "uploadFile": return json_(uploadFile_(p));
      case "listJobs": return json_(listJobs_(p));
      case "checkAdminUser": return json_(checkAdminUser_(p));
      case "listUsers": return json_(listUsers_(p));
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
      ["requester-files", "admin-files", "proof", "final"].forEach(function (n) { jobFolder.createFolder(n); });
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
      jobNo + "-" + (i + 1), jobNo, it.serviceCode || "", it.variant || "",
      it.qty || "", it.unit || "", it.unitPrice || "", it.estimatedAmount || "",
      "", it.priceRuleId || "", it.priceRuleSnapshot || "", ""
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
