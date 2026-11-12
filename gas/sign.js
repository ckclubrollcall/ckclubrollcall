/**
 * 產出指定日期所有社團的簽名報表 (Google 文件格式)
 * 會彈出視窗要求輸入社課日期，接著在雲端硬碟建立一份精美的文件表格，彙整當天所有社團的簽名狀態與圖檔。
 */
// eslint-disable-next-line no-unused-vars
function renderClubReportDoc() {
  const ui = SpreadsheetApp.getUi();

  // 1. 彈出輸入對話框要求輸入日期
  const response = ui.prompt("產生簽名報表", "請輸入社課日期（格式範例：2026/6/7）：", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const targetDate = response.getResponseText().trim();
    if (!targetDate) {
      ui.alert("❌ 請輸入有效日期！");
      return;
    }

    // 2. 讀取「原始社團」建立社團對照字典 (社團名稱 -> 社團編號 & 指導老師)
    const clubMap = getClubInfoMap();

    // 3. 取得該日期各社團最初「新增」的老師簽名與最新「編輯」的修改者紀錄
    const clubRecords = getSignatureRecordsByDate(targetDate);

    // 4. 驗證該日期是否真的有點名資料
    const clubsFound = Object.keys(clubRecords);
    if (clubsFound.length === 0) {
      ui.alert("ℹ️ 找不到該日期（" + targetDate + "）的任何點名紀錄！");
      return;
    }

    // 依社團編號排序
    clubsFound.sort(function (a, b) {
      const idA = clubMap[a] ? String(clubMap[a].id) : "";
      const idB = clubMap[b] ? String(clubMap[b].id) : "";
      return idA.localeCompare(idB);
    });

    // 5. 建立 Google 文件報表
    const docName = "社團點名簽名報表 - " + targetDate.replace(/\//g, "-");
    const reportFolderId = getReportFolderId();
    const targetFolder = reportFolderId ? DriveApp.getFolderById(reportFolderId) : null;

    if (targetFolder) {
      const oldFiles = targetFolder.getFilesByName(docName);
      while (oldFiles.hasNext()) {
        oldFiles.next().setTrashed(true); // 刪除同名的舊文件，防止檔案雜亂
      }
    }

    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    // 設計文件的頁首與標題樣式
    body.appendParagraph("【建中社團點名系統】").setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("社團點名簽名報表").setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("社課日期：" + targetDate).setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("\n");

    // 建立排版表格與標題列
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    const headers = ["社團編號", "社團名稱", "指導老師姓名", "指導老師簽名", "最後編輯時間", "最後編輯者簽名"];

    headers.forEach(function (text) {
      const cell = headerRow.appendTableCell(text);
      cell.setBackgroundColor("#F1F5F9");
      cell.getChild(0).asParagraph().setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });

    // 6. 逐列填入各社團資料並自雲端硬碟讀取簽名圖片插入表格中
    clubsFound.forEach(function (clubName) {
      const createRecord = clubRecords[clubName].createRecord;
      const latestEdit = clubRecords[clubName].latestEdit;

      const info = clubMap[clubName] || { id: "-(未對應)-", teacher: "-(未對應)-" };

      const tableRow = table.appendTableRow();
      tableRow.appendTableCell(info.id).setBold(true).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      tableRow.appendTableCell(clubName);
      tableRow.appendTableCell(info.teacher);

      // 1. 指導老師簽名 (取自最初新增紀錄)
      const teacherSigCell = tableRow.appendTableCell();
      insertSigToTableCell(teacherSigCell, createRecord ? createRecord.signatureRaw : null, "（無簽名資料）");

      // 2. 最後編輯時間 (若無編輯則顯示「無」)
      const editTimeCell = tableRow.appendTableCell();
      if (latestEdit && latestEdit.timestamp && !isNaN(latestEdit.timestamp.getTime())) {
        editTimeCell.setText(formatDateTime(latestEdit.timestamp));
      } else {
        editTimeCell.setText("無");
      }
      editTimeCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      // 3. 最後編輯者簽名 (若有編輯則插入修改者簽名，若無則顯示「無」)
      const editSigCell = tableRow.appendTableCell();
      if (latestEdit) {
        insertSigToTableCell(editSigCell, latestEdit.signatureRaw, "無");
      } else {
        editSigCell.setText("無");
        editSigCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    });

    doc.saveAndClose();
    if (targetFolder) {
      DriveApp.getFileById(doc.getId()).moveTo(targetFolder); // 將產出的報告移動至指定雲端資料夾中
    }

    // 7. 彈出處理完成對話框與直接開啟文件連結
    const url = doc.getUrl();
    const htmlOutput = HtmlService.createHtmlOutput(
      "<div style=\"font-family: Arial, sans-serif; text-align: center; padding: 10px;\">" +
      "<p style=\"color: #2e7d32; font-weight: bold; font-size: 16px;\">報表成功產生！</p>" +
      "<p style=\"margin-top: 20px;\"><a href=\"" + url + "\" target=\"_blank\" style=\"padding: 10px 20px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2);\">開啟 Google 文件</a></p>" +
      "</div>"
    ).setWidth(360).setHeight(180);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "處理完成");
  }
}

/**
 * 輔助工具：將簽名圖檔插入表格儲存格中
 */
function insertSigToTableCell(cell, sigRaw, defaultText) {
  const sigId = parseDriveId(sigRaw);
  if (sigId) {
    try {
      const imageBlob = DriveApp.getFileById(sigId).getBlob();
      const inlineImg = cell.appendImage(imageBlob);
      inlineImg.setWidth(120);
      inlineImg.setHeight(50);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      return;
    } catch (e) {
      cell.setText("⚠️ 無法載入圖檔\n(可能無權限或檔案已被刪除)");
      logToSheet("載入簽名圖檔失敗，檔案 ID: " + sigId + "，錯誤訊息: " + e.message);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      return;
    }
  }
  cell.setText(defaultText || "無");
  cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

/**
 * 輔助工具：將日期時間格式化為 yyyy/MM/dd HH:mm
 */
function formatDateTime(d) {
  if (!d || isNaN(d.getTime())) { return "無"; }
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const hh = d.getHours();
  const mm = d.getMinutes();
  return y + "/" + (m < 10 ? "0" : "") + m + "/" + (day < 10 ? "0" : "") + day + " " + (hh < 10 ? "0" : "") + hh + ":" + (mm < 10 ? "0" : "") + mm;
}

/**
 * 輔助工具：從簽名連結、Google Drive URL 或檔案名稱解析出 Google Drive File ID
 */
// eslint-disable-next-line no-redeclare
function parseDriveId(input) {
  if (!input) { return null; }
  const str = String(input).trim();
  if (str === "無" || str === "無簽名" || str === "（無簽名資料）" || str === "@") { return null; }

  // 1. 若字串中包含 Google Drive 標準 25~50 碼 ID（如 URL 或純 ID）
  // 排除副檔名結尾的檔名（例如 .png 檔名可能大於 25 字元，但不應直接當作 fileId）
  if (!/\.png$/i.test(str)) {
    const match = str.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
    if (match) { return match[0]; }
  }

  // 2. 若為檔名（例如 2026-10-18_社團名_修改者簽名.png），優先至簽名資料夾搜尋，並取建立時間最晚（最新）的檔案
  if (/\.png$/i.test(str)) {
    try {
      const folderId = getSignFolderId();
      let files = null;
      if (folderId) {
        const folder = DriveApp.getFolderById(folderId);
        files = folder.getFilesByName(str);
      }
      if (!files || !files.hasNext()) {
        files = DriveApp.getFilesByName(str); // 資料夾內找不到時全域尋找
      }

      let latestFile = null;
      let latestTime = 0;
      while (files.hasNext()) {
        const f = files.next();
        const fTime = f.getDateCreated().getTime();
        if (!latestFile || fTime > latestTime) {
          latestFile = f;
          latestTime = fTime;
        }
      }

      if (latestFile) {
        return latestFile.getId();
      }
    } catch (e) {
      logToSheet("透過檔名尋找簽名 ID 失敗: " + e.toString());
    }
  }

  // 3. 一般後備比對
  const fallbackMatch = str.match(/[-\w]{25,}/);
  return fallbackMatch ? fallbackMatch[0] : null;
}

/**
 * 輔助工具：將多樣的日期物件或字串，標準化轉換為 yyyy/MM/dd 格式以利文字比對
 */
function normalizeDate(dateInput) {
  if (!dateInput) { return ""; }
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) {
    return String(dateInput).replace(/\s+/g, "");
  }
  return d.getFullYear() + "/" + (d.getMonth() + 1) + "/" + d.getDate();
}