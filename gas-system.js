// =============================================
// COVA 帳號驗證 - Google Apps Script
// 功能：驗證登入、查詢案件、新增案件
// =============================================

// ① 請換成你的 Google 試算表 ID
var SHEET_ID = "1-84R1y0kePNQrE8B_m9ToZZnmEgIW9HhHQKaKGZqpJo";

// 工作表名稱
var USERS_SHEET  = "帳號";
var CASES_SHEET  = "案件";

// =============================================
// POST 入口
// =============================================
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;

    if (action === "login")      return handleLogin(payload);
    if (action === "getCases")   return handleGetCases(payload);
    if (action === "createCase") return handleCreateCase(payload);
    if (action === "updateCase") return handleUpdateCase(payload);

    throw new Error("未知的 action: " + action);

  } catch(err) {
    return json({ status: "error", message: err.message });
  }
}

function doGet(e) {
  return ContentService.createTextOutput("COVA 系統運作中 ✓").setMimeType(ContentService.MimeType.TEXT);
}

// =============================================
// 登入驗證
// =============================================
function handleLogin(payload) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(USERS_SHEET);
  var rows  = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (String(row[0]).trim() === payload.username &&
        String(row[1]).trim() === payload.password) {
      return json({
        status:   "ok",
        username: String(row[0]).trim(),
        role:     String(row[2]).trim(),
        name:     String(row[3]).trim(),
      });
    }
  }
  return json({ status: "error", message: "帳號或密碼錯誤" });
}

// =============================================
// 取得案件列表
// =============================================
function handleGetCases(payload) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CASES_SHEET);
  var rows  = sheet.getDataRange().getValues();

  if (rows.length <= 1) return json({ status: "ok", cases: [] });

  var headers = rows[0];
  var cases = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    obj["_row"] = i + 1;

    var role = payload.role;
    if (role === "admin") {
      cases.push(obj);
    } else if (role === "sales") {
      if (String(obj["負責業務"]).trim() === payload.name) cases.push(obj);
    } else if (role === "tech") {
      if (String(obj["派工師傅"]).trim() === payload.name) cases.push(obj);
    }
  }

  return json({ status: "ok", cases: cases });
}

// =============================================
// 產生案件ID（年月日 + 當日流水號）
// =============================================
function generateCaseId() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CASES_SHEET);
  var today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd");

  // 掃描現有案件，找今天最大流水號
  var lastRow = sheet.getLastRow();
  var maxSeq = 0;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var id = String(ids[i][0]);
      if (id.startsWith(today)) {
        var seq = parseInt(id.substring(8)) || 0;
        if (seq > maxSeq) maxSeq = seq;
      }
    }
  }

  var nextSeq = String(maxSeq + 1).padStart(3, "0");
  return today + "-" + nextSeq;
}

// =============================================
// 新增案件
// =============================================
function handleCreateCase(payload) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CASES_SHEET);
  var d     = payload.data;

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "案件ID","客戶姓名+安裝地區","連絡電話","安裝地址",
      "安裝家電類型","購買機型","預約安裝日期","需要攜帶","非標準安裝收費金額",
      "負責業務","派工師傅","案件進度","電梯","客源管道",
      "Google相簿連結","現場狀況與客訴紀錄","建立時間"
    ]);
  }

  var caseId = generateCaseId();
  var now    = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm");

  sheet.appendRow([
    caseId,
    d.name       || "",
    d.phone      || "",
    d.address    || "",
    (d.appliances || []).join("、"),
    d.model      || "",
    d.date       || "",
    d.items      || "",
    d.fee        || "",
    d.sales      || "",
    d.tech       || "",
    d.status     || "聯繫中",
    d.elevator   || "",
    d.source     || "",
    d.album      || "",
    d.note       || "",
    now
  ]);

  return json({ status: "ok", caseId: caseId });
}

// =============================================
// 更新案件
// =============================================
function handleUpdateCase(payload) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CASES_SHEET);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var d = payload.data;
  var rowNum = payload.rowNum;

  for (var key in d) {
    var colIdx = headers.indexOf(key);
    if (colIdx >= 0) {
      sheet.getRange(rowNum, colIdx + 1).setValue(d[key]);
    }
  }

  return json({ status: "ok" });
}

// =============================================
// 工具函式
// =============================================
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
