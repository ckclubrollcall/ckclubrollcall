/** 新學期清空 */
function clearAllRecords() {
  var ui = SpreadsheetApp.getUi();

  // 跳出警告
  var response = ui.alert(
    '⚠️ 警告：即將清空所有點名紀錄',
    '刪除所有點名紀錄、簽名存檔、簽名表存檔，且無法還原。您確定要執行嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    // 輸入驗證碼
    var promptResponse = ui.prompt('請輸入「DELETE」以確認執行清空作業：');

    if (promptResponse.getResponseText() === 'DELETE') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var recordSheet = ss.getSheetByName("點名紀錄");

      // 刪除位置
      var lastRow = recordSheet.getLastRow();
      if (lastRow > 1) {
        recordSheet.deleteRows(2, lastRow - 1);

        //刪除簽名檔
        var FOLDER_ID = getSignFolderId();
        if (FOLDER_ID) {
          var folder = DriveApp.getFolderById(FOLDER_ID);
          var files = folder.getFiles();
          while (files.hasNext()) {
            files.next().setTrashed(true);
          }
        };

        var REPORT_FOLDER_ID = getReportFolderId();
        if (REPORT_FOLDER_ID) {
          var reportFolder = DriveApp.getFolderById(REPORT_FOLDER_ID);
          var reportFiles = reportFolder.getFiles();
          while (reportFiles.hasNext()) {
            reportFiles.next().setTrashed(true);
          }
        }

        ui.alert('✅ 已成功清空所有紀錄！');
      } else {
        ui.alert('ℹ️ 目前已經是空的，無需清空。');
      }
    } else {
      ui.alert('❌ 驗證碼錯誤，操作已取消。');
    }
  }
}