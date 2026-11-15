// ==========================================
// 統計總表：由後台腳本寫入（取代原本的 LET 公式）
// A 社團編號 B 社團名稱 C 繳交狀況 D 填寫人 E 應到人數 F 實到人數 G 出席率
// H 老師簽名 I 最後修改者簽名 J 最後修改時間
// 社課日期由 L3 輸入；L3 被編輯時，由安裝式觸發器自動重新整理。
// ==========================================

const SUMMARY = {
  SHEET: "統計總表",
  DATE_CELL_ROW: 3,
  DATE_CELL_COL: 12,   // L3
  FIRST_DATA_ROW: 2,
  COLUMN_COUNT: 10,    // A:J
  SIGNATURE_ROW_HEIGHT: 60
};

/**
 * 將簽名檔轉成可嵌入儲存格的圖片（data URI，不需對外公開檔案）。
 * 沒有簽名時回傳 emptyText；讀取失敗時回傳錯誤提示文字。
 */
function buildSignatureCellValue(signatureRaw, emptyText, altTitle) {
  const fileId = parseDriveId(signatureRaw);
  if (!fileId) { return emptyText; }
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    const dataUri = "data:image/png;base64," + Utilities.base64Encode(blob.getBytes());
    return SpreadsheetApp.newCellImage()
      .setSourceUrl(dataUri)
      .setAltTextTitle(altTitle)
      .build();
  } catch (e) {
    logToSheet("統計總表載入簽名圖檔失敗，檔案 ID: " + fileId + "，錯誤訊息: " + e.message);
    return "⚠️ 無法載入圖檔";
  }
}

/**
 * 依 L3 的社課日期重新產生整張統計總表（A:J）。
 */
function refreshSummary() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    logToSheet("refreshSummary: 取得鎖定逾時，略過本次更新");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SUMMARY.SHEET);
    const dateValue = sheet.getRange(SUMMARY.DATE_CELL_ROW, SUMMARY.DATE_CELL_COL).getValue();

    // 先清空舊資料（同時移除舊的公式）
    const oldRows = Math.max(sheet.getLastRow() - SUMMARY.FIRST_DATA_ROW + 1, 1);
    sheet.getRange(SUMMARY.FIRST_DATA_ROW, 1, oldRows, SUMMARY.COLUMN_COUNT).clearContent();

    const clubSheet = ss.getSheetByName("原始社團");
    const clubLastRow = clubSheet.getLastRow();
    if (clubLastRow < 2) { return; }
    const clubs = clubSheet.getRange(2, 1, clubLastRow - 1, 3).getValues().filter(function (r) {
      return r[0] !== "" && String(r[1]).trim() !== "";
    });

    const hasDate = dateValue !== "" && dateValue !== null;
    const recordsByClub = hasDate ? getSignatureRecordsByDate(dateValue) : {};

    const rows = clubs.map(function (club) {
      const id = club[0];
      const name = String(club[1]).trim();
      const expected = club[2];
      const rec = recordsByClub[name];

      if (!rec) {
        return [id, name, "❌ 未繳交", "-", expected, "", "", "", "", ""];
      }

      const latest = rec.latestRecord;
      const actual = latest.presentCount;
      const rate = actual !== "" && expected ? actual / expected : "";
      const edit = rec.latestEdit;

      return [
        id,
        name,
        "✅ 已繳交",
        latest.fillerInfo,
        expected,
        actual,
        rate,
        buildSignatureCellValue(rec.createRecord.signatureRaw, "無", name + " 老師簽名"),
        edit ? buildSignatureCellValue(edit.signatureRaw, "無", name + " 最後修改者簽名") : "無",
        edit && !isNaN(edit.timestamp.getTime()) ? formatDateTime(edit.timestamp) : "無"
      ];
    });

    if (rows.length === 0) { return; }

    const target = sheet.getRange(SUMMARY.FIRST_DATA_ROW, 1, rows.length, SUMMARY.COLUMN_COUNT);
    target.setValues(rows);
    sheet.getRange(SUMMARY.FIRST_DATA_ROW, 7, rows.length, 1).setNumberFormat("0.0%");
    sheet.setRowHeights(SUMMARY.FIRST_DATA_ROW, rows.length, SUMMARY.SIGNATURE_ROW_HEIGHT);
  } catch (e) {
    logToSheet("refreshSummary 錯誤: " + e.toString());
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 安裝式 onEdit 觸發器的處理函式：只有 L3（社課日期）被編輯時才重新整理。
 * 需用 installSummaryTrigger() 安裝，因為讀取 Drive 檔案需要授權，簡易觸發器做不到。
 */
// eslint-disable-next-line no-unused-vars
function handleSummaryEdit(e) {
  if (!e || !e.range) { return; }
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SUMMARY.SHEET) { return; }
  if (e.range.getRow() !== SUMMARY.DATE_CELL_ROW || e.range.getColumn() !== SUMMARY.DATE_CELL_COL) { return; }
  refreshSummary();
}

/**
 * 安裝（或重新安裝）統計總表的自動更新觸發器。
 */
// eslint-disable-next-line no-unused-vars
function installSummaryTrigger() {
  const ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "handleSummaryEdit") { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger("handleSummaryEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  ui.alert("✅ 已安裝觸發器：修改統計總表 L3 的社課日期後會自動更新。");
}
