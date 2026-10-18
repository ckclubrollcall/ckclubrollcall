/**
 * 產出指定日期所有社團的簽名報表 (Google 文件格式)
 * 會彈出視窗要求輸入社課日期，接著在雲端硬碟建立一份精美的文件表格，彙整當天所有社團的簽名狀態與圖檔。
 */
function renderClubReportDoc() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. 彈出輸入對話框要求輸入日期
  var response = ui.prompt('產生簽名報表', '請輸入社課日期（格式範例：2026/6/7）：', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var targetDate = response.getResponseText().trim();
    if (!targetDate) {
      ui.alert('❌ 請輸入有效日期！');
      return;
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var recordSheet = ss.getSheetByName('點名紀錄');
    var summarySheet = ss.getSheetByName('統計總表');
    
    // 2. 讀取統計總表建立社團對照字典 (社團名稱 -> 社團編號 & 指導老師)
    var summaryData = summarySheet.getDataRange().getValues();
    var clubMap = {}; 
    for (var i = 1; i < summaryData.length; i++) {
      var clubId = summaryData[i][0];
      var clubName = summaryData[i][1];
      var teacherName = summaryData[i][7] || '未填寫'; 
      if (clubName) {
        clubMap[clubName] = { id: clubId, teacher: teacherName };
      }
    }
    
    // 3. 讀取點名紀錄，篩選出指定日期中各社團的最初「新增」老師簽名與最新的「編輯」修改者紀錄
    var recordRange = recordSheet.getDataRange();
    var recordData = recordRange.getValues();
    var recordRichTexts = recordRange.getRichTextValues();
    var clubRecords = {}; 
    
    for (var j = 1; j < recordData.length; j++) {
      var row = recordData[j];
      var timestampVal = row[0];   // 填寫時間
      var clubName = row[2];       // 社團名稱
      var rwdDateStr = row[3];     // 社課日期
      
      // 支援 Google Drive 檔案智慧方塊：先自 richText 提取 URL，無則使用儲存格值
      var sigRichText = (recordRichTexts && recordRichTexts[j] && recordRichTexts[j][10]) ? recordRichTexts[j][10].getLinkUrl() : null;
      var signatureRaw = sigRichText || row[10];
      var recordType = row[11] ? String(row[11]).trim() : '';
      
      // 標準化日期格式進行比對
      var recordDateFormatted = normalizeDate(rwdDateStr);
      var targetDateFormatted = normalizeDate(targetDate);
      
      if (recordDateFormatted === targetDateFormatted && clubName) {
        var currentTimestamp = new Date(timestampVal);
        if (!clubRecords[clubName]) {
          clubRecords[clubName] = [];
        }
        clubRecords[clubName].push({
          timestamp: currentTimestamp,
          signatureRaw: signatureRaw,
          recordType: recordType
        });
      }
    }
    
    // 4. 驗證該日期是否真的有點名資料
    var clubsFound = Object.keys(clubRecords);
    if (clubsFound.length === 0) {
      ui.alert('ℹ️ 找不到該日期（' + targetDate + '）的任何點名紀錄！');
      return;
    }
    
    // 依社團編號排序
    clubsFound.sort(function(a, b) {
      var idA = clubMap[a] ? clubMap[a].id : '';
      var idB = clubMap[b] ? clubMap[b].id : '';
      return idA.localeCompare(idB);
    });
    
    // 5. 建立 Google 文件報表
    var docName = '社團點名簽名報表 - ' + targetDate.replace(/\//g, '-');
    var reportFolderId = getReportFolderId();
    var targetFolder = reportFolderId ? DriveApp.getFolderById(reportFolderId) : null;
    
    if (targetFolder) {
      var oldFiles = targetFolder.getFilesByName(docName);
      while (oldFiles.hasNext()) {
        oldFiles.next().setTrashed(true); // 刪除同名的舊文件，防止檔案雜亂
      }
    }
 
    var doc = DocumentApp.create(docName);
    var body = doc.getBody();
    
    // 設計文件的頁首與標題樣式
    body.appendParagraph('【建中社團點名系統】').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('社團點名簽名報表').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('社課日期：' + targetDate).setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('\n');
    
    // 建立排版表格與標題列
    var table = body.appendTable();
    var headerRow = table.appendTableRow();
    var headers = ['社團編號', '社團名稱', '指導老師姓名', '指導老師簽名', '最後編輯時間', '最後編輯者簽名'];
    
    headers.forEach(function(text) {
      var cell = headerRow.appendTableCell(text);
      cell.setBackgroundColor('#F1F5F9'); 
      cell.getChild(0).asParagraph().setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });
    
    // 6. 逐列填入各社團資料並自雲端硬碟讀取簽名圖片插入表格中
    clubsFound.forEach(function(clubName) {
      var rows = clubRecords[clubName];
      // 依時間由舊到新排序
      rows.sort(function(x, y) { return x.timestamp - y.timestamp; });
      
      // 最初的「新增」紀錄為老師簽名
      var createRecord = rows.filter(function(r) { return r.recordType === '新增'; })[0] || rows[0];
      // 「編輯」紀錄清單
      var editRecords = rows.filter(function(r) { return r.recordType === '編輯'; });
      // 若無標註 recordType 但有多筆提交，後續筆數視為編輯紀錄（向後相容）
      if (editRecords.length === 0 && rows.length > 1) {
        editRecords = rows.slice(1);
      }
      var latestEdit = editRecords.length > 0 ? editRecords[editRecords.length - 1] : null;
      
      var info = clubMap[clubName] || { id: '-(未對應)-', teacher: '-(未對應)-' };
      
      var tableRow = table.appendTableRow();
      tableRow.appendTableCell(info.id).setBold(true).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      tableRow.appendTableCell(clubName);
      tableRow.appendTableCell(info.teacher);
      
      // 1. 指導老師簽名 (取自最初新增紀錄)
      var teacherSigCell = tableRow.appendTableCell();
      insertSigToTableCell(teacherSigCell, createRecord ? createRecord.signatureRaw : null, '（無簽名資料）');
      
      // 2. 最後編輯時間 (若無編輯則顯示「無」)
      var editTimeCell = tableRow.appendTableCell();
      if (latestEdit && latestEdit.timestamp && !isNaN(latestEdit.timestamp.getTime())) {
        editTimeCell.setText(formatDateTime(latestEdit.timestamp));
      } else {
        editTimeCell.setText('無');
      }
      editTimeCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      // 3. 最後編輯者簽名 (若有編輯則插入修改者簽名，若無則顯示「無」)
      var editSigCell = tableRow.appendTableCell();
      if (latestEdit) {
        insertSigToTableCell(editSigCell, latestEdit.signatureRaw, '無');
      } else {
        editSigCell.setText('無');
        editSigCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    });
    
    doc.saveAndClose();
    if (targetFolder) {
      DriveApp.getFileById(doc.getId()).moveTo(targetFolder); // 將產出的報告移動至指定雲端資料夾中
    }
    
    // 7. 彈出處理完成對話框與直接開啟文件連結
    var url = doc.getUrl();
    var htmlOutput = HtmlService.createHtmlOutput(
      '<div style="font-family: Arial, sans-serif; text-align: center; padding: 10px;">' +
      '<p style="color: #2e7d32; font-weight: bold; font-size: 16px;">報表成功產生！</p>' +
      '<p style="margin-top: 20px;"><a href="' + url + '" target="_blank" style="padding: 10px 20px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">開啟 Google 文件</a></p>' +
      '</div>'
    ).setWidth(360).setHeight(180);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '處理完成');
  }
}

/**
 * 輔助工具：將簽名圖檔插入表格儲存格中
 */
function insertSigToTableCell(cell, sigRaw, defaultText) {
  var sigId = parseDriveId(sigRaw);
  if (sigId) {
    try {
      var imageBlob = DriveApp.getFileById(sigId).getBlob();
      var inlineImg = cell.appendImage(imageBlob);
      inlineImg.setWidth(120);
      inlineImg.setHeight(50);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      return;
    } catch (e) {
      cell.setText('⚠️ 無法載入圖檔\n(可能無權限或檔案已被刪除)');
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      return;
    }
  }
  cell.setText(defaultText || '無');
  cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

/**
 * 輔助工具：將日期時間格式化為 yyyy/MM/dd HH:mm
 */
function formatDateTime(d) {
  if (!d || isNaN(d.getTime())) return '無';
  var y = d.getFullYear();
  var m = d.getMonth() + 1;
  var day = d.getDate();
  var hh = d.getHours();
  var mm = d.getMinutes();
  return y + '/' + (m < 10 ? '0' : '') + m + '/' + (day < 10 ? '0' : '') + day + ' ' + (hh < 10 ? '0' : '') + hh + ':' + (mm < 10 ? '0' : '') + mm;
}

/**
 * 輔助工具：從簽名連結或 Google Drive URL 內擷取唯一的 25~50 碼檔案 ID
 */
function parseDriveId(input) {
  if (!input) return null;
  var str = String(input).trim();
  if (str === '無' || str === '無簽名' || str === '（無簽名資料）') return null;
  var match = str.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  return match ? match[0] : (str.length >= 25 ? str : null);
}

/**
 * 輔助工具：將多樣的日期物件或字串，標準化轉換為 yyyy/MM/dd 格式以利文字比對
 */
function normalizeDate(dateInput) {
  if (!dateInput) return '';
  var d = new Date(dateInput);
  if (isNaN(d.getTime())) {
    return String(dateInput).replace(/\s+/g, ''); 
  }
  return d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
}