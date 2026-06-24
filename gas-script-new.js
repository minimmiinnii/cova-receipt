// =============================================
// COVA 簽收單 - Google Apps Script
// 功能：接收簽收單 PDF + 照片，同時寄信 + 備份到 Drive
// =============================================

var COMPANY_EMAIL      = "heidi.wang@covalife.com";
var EMAIL_SUBJECT_PREFIX = "【COVA 簽收單】";
var DRIVE_FOLDER_NAME  = "COVA簽收單";  // Drive 根資料夾名稱

function doPost(e) {
  try {
    Logger.log("收到請求：" + e.postData.contents.substring(0, 200));
    var payload = JSON.parse(e.postData.contents);

    var pdfBase64 = payload.pdfBase64;
    var filename  = payload.filename  || "簽收單.pdf";
    var photos    = payload.photos    || [];
    var svcno     = payload.svcno     || "—";
    var customer  = payload.customer  || "—";
    var tech      = payload.tech      || "—";
    var date      = payload.date      || "—";
    var total     = payload.total     || "0";

    // PDF blob
    var pdfBlob = Utilities.newBlob(
      Utilities.base64Decode(pdfBase64), "application/pdf", filename
    );

    // 照片 blobs
    var photoBlobs = photos.map(function(p, i) {
      var ext = p.name.split(".").pop() || "jpg";
      var photoName = svcno + "_安裝照片_" + String(i+1).padStart(2,"0") + "." + ext;
      return Utilities.newBlob(
        Utilities.base64Decode(p.base64), p.type || "image/jpeg", photoName
      );
    });

    var allAttachments = [pdfBlob].concat(photoBlobs);

    // ① 寄送 Email
    var emailError = null;
    try {
      var subject = EMAIL_SUBJECT_PREFIX + svcno + "｜" + customer;
      var body =
        "您好，\n\n" +
        "以下為本次安裝簽收紀錄，PDF 簽收單及安裝照片已附於此郵件。\n\n" +
        "━━━━━━━━━━━━━━━━━━━━\n" +
        "案件ID：" + svcno    + "\n" +
        "客戶姓名：" + customer  + "\n" +
        "安裝師傅：" + tech      + "\n" +
        "安裝日期：" + date      + "\n" +
        "總計金額：NT$ " + total + "\n" +
        "安裝照片：" + photos.length + " 張\n" +
        "━━━━━━━━━━━━━━━━━━━━\n\n" +
        "此郵件由系統自動發送，請勿回覆。\nCOVA";

      GmailApp.sendEmail(COMPANY_EMAIL, subject, body, {
        attachments: allAttachments,
        name: "COVA 簽收系統"
      });
    } catch(err) {
      emailError = err.message;
      Logger.log("寄信失敗：" + err.message);
    }

    // ② 備份到 Google Drive
    var driveError = null;
    try {
      // 找或建立根資料夾 "COVA簽收單"
      var rootFolder;
      var rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
      if (rootFolders.hasNext()) {
        rootFolder = rootFolders.next();
      } else {
        rootFolder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
      }

      // 建立案件子資料夾（用案件ID命名）
      var caseFolder;
      var caseFolders = rootFolder.getFoldersByName(svcno);
      if (caseFolders.hasNext()) {
        caseFolder = caseFolders.next();
      } else {
        caseFolder = rootFolder.createFolder(svcno);
      }

      // 存入 PDF
      caseFolder.createFile(pdfBlob);

      // 存入照片
      photoBlobs.forEach(function(blob) {
        caseFolder.createFile(blob);
      });

    } catch(err) {
      driveError = err.message;
      Logger.log("Drive 備份失敗：" + err.message);
    }

    // 回傳結果
    var status = "ok";
    var message = "";
    if (emailError && driveError) {
      status = "error";
      message = "寄信失敗：" + emailError + "；Drive備份失敗：" + driveError;
    } else if (emailError) {
      status = "partial";
      message = "已備份至 Drive，但寄信失敗：" + emailError;
    } else if (driveError) {
      status = "partial";
      message = "已寄信，但 Drive 備份失敗：" + driveError;
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: status, message: message }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput("COVA 簽收系統運作中 ✓")
    .setMimeType(ContentService.MimeType.TEXT);
}
