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
    // 依標題列定位欄位，「統計總表」日後新增欄位（如最新簽名圖片、最後修改日期）時不會錯位；
    // 找不到對應標題時退回原本的固定位置（A、B、H 欄）以維持相容。
    var summaryCols = resolveSummaryColumns(summaryData[0] || []);
    var clubMap = {}; 
    for (var i = 1; i < summaryData.length; i++) {
      var clubId = summaryData[i][summaryCols.clubId];
      var clubName = summaryData[i][summaryCols.clubName];
      var teacherName = summaryData[i][summaryCols.teacher] || '未填寫';
      if (clubName) {
        clubMap[clubName] = { id: clubId, teacher: teacherName };
      }
    }
    
    // 3. 讀取點名紀錄，篩選出指定日期中「每個社團最新提交」的簽名與狀態
    var recordData = recordSheet.getDataRange().getValues();
    var latestRecords = {}; 
    
    for (var j = 1; j < recordData.length; j++) {
      var row = recordData[j];
      var timestampStr = row[0];   // 填寫時間
      var clubName = row[2];       // 社團名稱
      var rwdDateStr = row[3];     // 社課日期
      var signatureRaw = row[10];  // 簽名連結
      
      // 標準化日期格式進行比對
      var recordDateFormatted = normalizeDate(rwdDateStr);
      var targetDateFormatted = normalizeDate(targetDate);
      
      if (recordDateFormatted === targetDateFormatted && clubName) {
        var currentTimestamp = new Date(timestampStr);
        
        // 若該社團還沒有紀錄，或者這筆紀錄的時間戳記更晚，就覆蓋
        if (!latestRecords[clubName] || currentTimestamp > latestRecords[clubName].timestamp) {
          latestRecords[clubName] = {
            timestamp: currentTimestamp,
            signatureRaw: signatureRaw
          };
        }
      }
    }
    
    // 4. 驗證該日期是否真的有點名資料
    var clubsFound = Object.keys(latestRecords);
    if (clubsFound.length === 0) {
      ui.alert('ℹ️ 找不到該日期（' + targetDate + '）的任何點名紀錄！');
      return;
    }
    
    // 依社團編號排序
    clubsFound.sort(function(a, b) {
      var idA = clubMap[a] ? String(clubMap[a].id) : '';
      var idB = clubMap[b] ? String(clubMap[b].id) : '';
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
    var headers = ['社團編號', '社團名稱', '指導老師姓名', '指導老師簽名'];
    
    headers.forEach(function(text) {
      var cell = headerRow.appendTableCell(text);
      cell.setBackgroundColor('#F1F5F9'); 
      cell.getChild(0).asParagraph().setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });
    
    // 6. 逐列填入各社團資料並自雲端硬碟讀取簽名圖片插入表格中
    clubsFound.forEach(function(clubName) {
      var record = latestRecords[clubName];
      var info = clubMap[clubName] || { id: '-(未對應)-', teacher: '-(未對應)-' };
      
      var row = table.appendTableRow();
      row.appendTableCell(String(info.id)).setBold(true).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      row.appendTableCell(clubName);
      row.appendTableCell(info.teacher);
      
      var sigCell = row.appendTableCell();
      var sigId = parseDriveId(record.signatureRaw); // 提取簽名檔案的 Google Drive ID
      
      if (sigId) {
        try {
          var imageBlob = DriveApp.getFileById(sigId).getBlob();
          var inlineImg = sigCell.appendImage(imageBlob);
          
          // 限制圖片寬高，防止排版崩潰
          inlineImg.setWidth(120);
          inlineImg.setHeight(50);
          sigCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        } catch(e) {
          sigCell.setText('⚠️ 無法載入圖檔\n(可能無權限或檔案已被刪除)');
        }
      } else {
        sigCell.setText('（無簽名資料）');
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
 * 輔助工具：依「統計總表」標題列找出社團編號、社團名稱、指導老師姓名所在欄位（0-based）
 * - 指導老師欄需排除含「簽名」「日期」字樣的欄位，避免誤抓到新增的「最新簽名圖片」等欄
 * - 找不到時沿用舊版固定位置：A=社團編號、B=社團名稱、H=指導老師
 */
function resolveSummaryColumns(headerRow) {
  var cols = { clubId: 0, clubName: 1, teacher: 7 };
  var found = { clubId: false, clubName: false, teacher: false };
  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c] || '').replace(/\s+/g, '');
    if (!h) continue;
    if (!found.clubId && h.indexOf('社團編號') !== -1) { cols.clubId = c; found.clubId = true; continue; }
    if (!found.clubName && h.indexOf('社團名稱') !== -1) { cols.clubName = c; found.clubName = true; continue; }
    if (!found.teacher && h.indexOf('指導老師') !== -1 &&
        h.indexOf('簽名') === -1 && h.indexOf('日期') === -1 &&
        h.indexOf('出席') === -1 && h.indexOf('請假') === -1) {
      cols.teacher = c; found.teacher = true;
    }
  }
  return cols;
}

/**
 * 輔助工具：從簽名連結或 Google Drive URL 內擷取唯一的 25~50 碼檔案 ID
 */
function parseDriveId(input) {
  if (!input) return null;
  var str = String(input);
  var match = str.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  return match ? match[0] : str;
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