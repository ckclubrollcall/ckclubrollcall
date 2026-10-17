/** 建立試算表選單。 */
function onOpen() {
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('簽名報表')
    .addItem('產出指定日期簽名報表 (Google 文件)','renderClubReportDoc')
    .addItem('初始化／檢查點名紀錄欄位','initializeAttendanceSchema')
    .addToUi();
  ui.createMenu('刪除資料')
    .addItem('刪除所有點名紀錄（本功能用於新學期）','clearAllRecords')
    .addToUi();
}

function getReportFolderId() {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('系統設定');
  if (!sheet) throw new Error('找不到系統設定');
  var raw=sheet.getRange('B2').getValue();
  return raw ? parseDriveId(raw) : null;
}
