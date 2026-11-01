/** 建立試算表選單 */
// eslint-disable-next-line no-unused-vars
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("簽名報表")
    .addItem("產出指定日期簽名報表 (Google 文件)", "renderClubReportDoc")
    .addToUi();

  ui.createMenu("刪除資料")
    .addItem("刪除所有點名紀錄（本功能用於新學期）", "clearAllRecords")
    .addToUi();
}

/**讀取系統設定資料夾 */
// eslint-disable-next-line no-unused-vars, no-redeclare
function getReportFolderId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const raw = ss.getSheetByName("系統設定").getRange("B2").getValue();
    return parseDriveId(raw);
  } catch (e) {
    logToSheet("getReportFolderId 錯誤: " + e.toString());
    return null;
  }
}