/**
 * 當試算表發生編輯時的觸發器 (Google Sheets Trigger)
 * 當使用者在試算表的「查詢面板」工作表點選或輸入社團名稱時，會自動觸發並產出該社團的出缺席報告。
 */
function onEdit(e){
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName(); 
  
  if (sheetName === "查詢面板") {
    renderClubReport(e);
  } 
}

/**
 * 核心統計表產生器
 * 負責清空舊有報表，並依據選定的社團，在工作表上動態繪製「點名出缺席矩陣圖」、「全學期缺席明細表」、「填答與課程內容摘要」三大版塊。
 */
function renderClubReport(e) {
  var sheet = e.source.getActiveSheet();
  var clubName = e.value;
    
  // 1. 清空舊報表區域的內容與格式
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if (maxRows > 3 && maxCols > 1) {
    sheet.getRange(4, 2, maxRows - 3, maxCols - 1).clear();
  }
  
  if (!clubName) return;

  // 2. 顯示讀取中提示，避免使用者在此期間操作試算表
  sheet.getRange("B4").setValue("🔄️請稍候...並避免操作").setFontColor("#888888");
  SpreadsheetApp.flush(); 

  // 3. 讀取「點名紀錄」與「原始名單」的底層資料
  var ss = e.source;
  var recordSheet = ss.getSheetByName("點名紀錄");
  var memberSheet = ss.getSheetByName("原始名單");
    
  var recordRange = recordSheet.getDataRange();
  var allRecords = recordRange.getValues();
  var allRichTexts = recordRange.getRichTextValues();
  for (var k = 0; k < allRecords.length; k++) {
    var richUrl = (allRichTexts && allRichTexts[k] && allRichTexts[k][10]) ? allRichTexts[k][10].getLinkUrl() : null;
    allRecords[k]._signatureUrl = richUrl || allRecords[k][10];
  }
  var allMembers = memberSheet.getDataRange().getValues();

  var currentRow = 4; // 起始列
  var startCol = 2;   // 起始欄（第 B 欄）

  // ==========================================
  // 板塊 A：產生點名出缺席矩陣圖 (打勾與打叉圖表)
  // ==========================================
  var matrixData = buildAttendanceMatrix(clubName, allRecords, allMembers);
  if (matrixData && matrixData.length > 0) {
    var range1 = sheet.getRange(currentRow, startCol, matrixData.length, matrixData[0].length);
    range1.setValues(matrixData); 
      
    var bgColors = [];
    var fontColors = [];

    // 設定表格顏色：出席為淡綠底深綠字 ✔，缺席為淡紅底深紅字 ✖
    for (var r = 0; r < matrixData.length; r++) {
      var bgRow = [];
      var fontRow = [];
      for (var c = 0; c < matrixData[r].length; c++) {
        var val = matrixData[r][c];
          
        if (r === 0) { 
          bgRow.push("#d9ead3"); 
          fontRow.push("#000000");
        } else {
          if (val === "✔") {
            bgRow.push("#e6f4ea"); fontRow.push("#137333");
          } else if (val === "✖") {
            bgRow.push("#fce8e6"); fontRow.push("#c5221f");
          } else {
            bgRow.push("#ffffff"); fontRow.push("#000000");
          }
        }
       }
      bgColors.push(bgRow);
      fontColors.push(fontRow);
    }
      
    // 套用樣式
    range1.setBackgrounds(bgColors);
    range1.setFontColors(fontColors);
    range1.setBorder(true, true, true, true, true, true);
    range1.setHorizontalAlignment("center"); 
    sheet.getRange(currentRow, startCol, 1, matrixData[0].length).setFontWeight("bold");
      
    currentRow += matrixData.length + 2; // 移動下方繪製起點，留出空白行
  }

  // ==========================================
  // 板塊 B：產生全學期缺席明細 (依日期並排)
  // ==========================================
  sheet.getRange(currentRow, startCol).setValue("全學期缺席明細").setFontWeight("bold").setFontSize(12).setFontColor("#2f5496");
  currentRow += 1;
    
  var absentBlocks = buildAbsenteeBlocks(clubName, allRecords, allMembers);
  if (absentBlocks && absentBlocks.length > 0) {
    var maxRows = 0;
    var currentCols = startCol;

    absentBlocks.forEach(block => {
      var blockRows = block.data.length;
      if (blockRows > maxRows) maxRows = blockRows;

      if (currentCols + 4 > sheet.getMaxColumns()) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), 5); // 動態增加 5 欄
      }

      // 產生日期標題列
      sheet.getRange(currentRow, currentCols, 1, 4).merge()
          .setValue(block.date + " 缺席")
          .setHorizontalAlignment("center")
          .setBackground("#f4cccc")
          .setFontWeight("bold");

      if (blockRows > 0) {
        // 產生欄位名稱
        sheet.getRange(currentRow + 1, currentCols, 1, 4)
            .setValues([["班級", "座號", "姓名", "學號"]])
            .setBackground("#fce8e6")
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
          
        // 填入缺席學生名冊
        sheet.getRange(currentRow + 2, currentCols, blockRows, 4)
            .setValues(block.data)
            .setHorizontalAlignment("center");
          
        sheet.getRange(currentRow, currentCols, blockRows + 2, 4)
            .setBorder(true, true, true, true, true, true);
      } else {
        // 無人缺席時顯示「全勤」
        sheet.getRange(currentRow + 1, currentCols, 1, 4).merge()
            .setValue("全勤")
            .setHorizontalAlignment("center");
          
        sheet.getRange(currentRow, currentCols, 2, 4)
            .setBorder(true, true, true, true, true, true);
      }

      currentCols += 5; // 板塊向右橫向平移五欄
    });

    currentRow += (maxRows > 0 ? maxRows + 2 : 2) + 2; 
  }

  // ==========================================
  // 板塊 C：產生填答與課程內容摘要 (社課內容與老師簽章連結)
  // ==========================================
  sheet.getRange(currentRow, startCol).setValue("全學期社課內容與回覆紀錄").setFontWeight("bold").setFontSize(12).setFontColor("#2f5496");
  currentRow += 1;

  var courseData = [["日期", "填寫人", "社課內容", "老師出席", "代課老師", "老師簽名"]];
  var clubRecords = allRecords.filter(r => r[2] === clubName && r[3]).reverse();

  // 預先對應日期與紀錄
  var recordsWithFormattedDate = clubRecords.map(function(r) {
    var formatted = "";
    if (r[3]) {
      var d = (r[3] instanceof Date) ? r[3] : new Date(r[3]);
      if (!isNaN(d.getTime())) {
        formatted = formatMMdd(d);
      }
    }
    return {
      formattedDate: formatted,
      record: r
    };
  });

  var dates = [];
  var dateMap = {};
  var recordByDate = {};
  
  recordsWithFormattedDate.forEach(function(item) {
    if (item.formattedDate) {
      if (!dateMap[item.formattedDate]) {
        dateMap[item.formattedDate] = true;
        dates.push(item.formattedDate);
      }
      if (!recordByDate[item.formattedDate]) {
        recordByDate[item.formattedDate] = item.record;
      }
    }
  });

  dates.sort();

  dates.forEach(d => {
    var r = recordByDate[d];
    var sigRaw = (r._signatureUrl !== undefined) ? r._signatureUrl : r[10];
    var sigValue = sigRaw ? sigRaw.toString().trim() : "";
    var safeSigUrl = /^https:\/\/(?:drive\.google\.com|docs\.google\.com|drive\.usercontent\.google\.com|[a-z0-9.-]*googleusercontent\.com)\//i.test(sigValue);
    // 只有確認為 Google/Drive HTTPS 簽名連結時才產生公式，避免把任意字串拼進 HYPERLINK。
    var sigDisplay = safeSigUrl ? '=HYPERLINK("' + sigValue.replace(/"/g, '""') + '", "📄 查看簽名")' : sigValue;
    courseData.push([d, r[1], r[5], r[6], r[7], sigDisplay]);
  });
    
  if (courseData.length === 1) {
    courseData.push(["目前尚無資料", "", "", "", "", ""]);
  }

  var range3 = sheet.getRange(currentRow, startCol, courseData.length, courseData[0].length);
  range3.setValues(courseData);
  range3.setBorder(true, true, true, true, true, true);
  sheet.getRange(currentRow, startCol, 1, courseData[0].length).setBackground("#c9daf8").setFontWeight("bold");

}

/**
 * 輔助工具：將日期格式化為 MM/dd
 */
function formatMMdd(date) {
  var m = date.getMonth() + 1;
  var d = date.getDate();
  return (m < 10 ? "0" : "") + m + "/" + (d < 10 ? "0" : "") + d;
}

/**
 * 組合社課出缺席打勾矩陣的核心演算法
 */
function buildAttendanceMatrix(clubName, records, members) {
  var clubMembers = members.filter(r => r[6] === clubName); 
  var clubRecords = records.filter(r => r[2] === clubName).reverse(); 

  var recordsWithFormattedDate = clubRecords.map(function(r) {
    var formatted = "";
    if (r[3]) {
      var d = (r[3] instanceof Date) ? r[3] : new Date(r[3]);
      if (!isNaN(d.getTime())) {
        formatted = formatMMdd(d);
      }
    }
    return {
      formattedDate: formatted,
      record: r
    };
  });

  var dates = [];
  var dateMap = {};
  var recordByDate = {};
  
  recordsWithFormattedDate.forEach(function(item) {
    if (item.formattedDate) {
      if (!dateMap[item.formattedDate]) {
        dateMap[item.formattedDate] = true;
        dates.push(item.formattedDate);
      }
      if (!recordByDate[item.formattedDate]) {
        recordByDate[item.formattedDate] = item.record;
      }
    }
  });

  dates.sort();

  var header = ["班級", "座號", "姓名", "學號", ...dates];
  var matrix = [header];

  clubMembers.forEach(m => {
    var studentName = m[2]; 
    var row = [m[0], m[1], m[2], m[3]]; 
    
    dates.forEach(d => {
      var dayRecord = recordByDate[d];
      if (dayRecord) {
        var absentList = dayRecord[9] || ""; 
        var isAbsent = absentList.indexOf(studentName) !== -1;
        row.push(isAbsent ? "✖" : "✔"); // 缺席記為紅叉，出席為綠勾
      } else {
        row.push(""); 
      }
    });
    matrix.push(row); 
  });
  
  return matrix;
}

/**
 * 整理所有出缺席日期對應的缺席者完整學籍資料，用於板塊 B
 */
function buildAbsenteeBlocks(clubName, records, members) {
  var clubMembers = members.filter(r => r[6] === clubName);
  var clubRecords = records.filter(r => r[2] === clubName).reverse();
  
  var recordsWithFormattedDate = clubRecords.map(function(r) {
    var formatted = "";
    if (r[3]) {
      var d = (r[3] instanceof Date) ? r[3] : new Date(r[3]);
      if (!isNaN(d.getTime())) {
        formatted = formatMMdd(d);
      }
    }
    return {
      formattedDate: formatted,
      record: r
    };
  });

  var dates = [];
  var dateMap = {};
  var recordByDate = {};
  
  recordsWithFormattedDate.forEach(function(item) {
    if (item.formattedDate) {
      if (!dateMap[item.formattedDate]) {
        dateMap[item.formattedDate] = true;
        dates.push(item.formattedDate);
      }
      if (!recordByDate[item.formattedDate]) {
        recordByDate[item.formattedDate] = item.record;
      }
    }
  });

  // Sort dates from newest to oldest (reverse)
  dates.sort().reverse();

  // 建立學生名稱的 Map 快取以提升查詢效率
  var memberByName = {};
  clubMembers.forEach(function(m) {
    if (m[2]) {
      memberByName[m[2]] = m;
    }
  });

  var blocks = [];

  dates.forEach(d => {
    var latestDayRecord = recordByDate[d];
    
    var absentees = [];
    if (latestDayRecord) {
      var absentList = latestDayRecord[9]; 

      if (absentList && absentList !== "全勤") { 
        var lines = absentList.split("\n"); 
        
        lines.forEach(person => {
          var infoArray = person.trim().split(/\s+/);
          var stuName = infoArray.length >= 3 ? infoArray[2] : person.trim();
          var matchedMember = memberByName[stuName];
          var stuId = matchedMember ? matchedMember[3] : ""; 

          if (infoArray.length >= 3) {
            absentees.push([infoArray[0], infoArray[1], stuName, stuId]);
          } else {
            absentees.push(["", "", stuName, stuId]);
          }
        });
      }
    }
    blocks.push({
      date: d,
      data: absentees
    });
  });
  
  return blocks; 
}