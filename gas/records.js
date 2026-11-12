// ==========================================
// 點名紀錄查詢：依社課日期整理各社團的簽名紀錄
// 供「統計總表」與「簽名報表（Google 文件）」共用，確保兩邊判定規則一致。
// ==========================================

/** 「點名紀錄」工作表欄位索引（0-based） */
const RECORD_COL = {
  TIMESTAMP: 0,     // A: 時間戳記
  FILLER_INFO: 1,   // B: 身分資料
  CLUB: 2,          // C: 社團名稱
  DATE: 3,          // D: 社課日期
  FILLER_NAME: 4,   // E: 填寫人姓名
  PRESENT_COUNT: 8, // I: 實到人數
  SIGNATURE: 10,    // K: 簽名（Drive 智慧方塊）
  TYPE: 11,         // L: 紀錄類別（新增 / 編輯）
  FILE_ID: 12       // M: 簽名檔案 ID（提交時寫入）
};

/**
 * 讀取「原始社團」，回傳 { 社團名稱: { id, expected, teacher } }
 */
function getClubInfoMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始社團");
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { return map; }

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  values.forEach(function (row) {
    const name = String(row[1]).trim();
    if (!row[0] || !name) { return; }
    map[name] = {
      id: row[0],
      expected: row[2],
      teacher: row[3] ? String(row[3]).trim() : "未填寫"
    };
  });
  return map;
}

/**
 * 以 Sheets API 批次讀取 K 欄智慧方塊的原始 URI，回傳 { recordData 列索引: uri }
 * 未啟用 Sheets API 或讀取失敗時回傳空物件。
 */
function readSignatureChipUris(ss, sheetName, numRows) {
  const chipUrlMap = {};
  try {
    if (typeof Sheets === "undefined" || !Sheets.Spreadsheets || numRows <= 1) { return chipUrlMap; }

    const resp = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: ["'" + sheetName + "'!K2:K" + numRows],
      fields: "sheets(data(rowData(values(chipRuns))))"
    });
    const rowsData = resp.sheets && resp.sheets[0] && resp.sheets[0].data && resp.sheets[0].data[0].rowData;
    if (!rowsData) { return chipUrlMap; }

    rowsData.forEach(function (rowItem, rIdx) {
      const chipRuns = rowItem && rowItem.values && rowItem.values[0] && rowItem.values[0].chipRuns;
      if (!chipRuns) { return; }
      for (let c = 0; c < chipRuns.length; c++) {
        const props = chipRuns[c].chip && chipRuns[c].chip.richLinkProperties;
        if (props && props.uri) {
          chipUrlMap[rIdx + 1] = props.uri; // 對應 recordData 的列索引（標題列為 0）
          break;
        }
      }
    });
  } catch (e) {
    logToSheet("讀取智慧方塊 API 失敗，將使用後備方式解析: " + e.toString());
  }
  return chipUrlMap;
}

/**
 * 取得指定社課日期中，每個社團的簽名紀錄。
 *
 * 回傳 { 社團名稱: { createRecord, latestEdit, latestRecord } }
 *   createRecord: 最初的「新增」紀錄（老師簽名）
 *   latestEdit:   最新的「編輯」紀錄（最後修改者簽名），沒有編輯時為 null
 *   latestRecord: 該日最新的一筆紀錄（含編輯）
 * 每筆紀錄的格式為 { timestamp, fillerInfo, fillerName, presentCount, signatureRaw, recordType }。
 * signatureRaw 的取用順序：M 欄檔案 ID → 智慧方塊 URI → 儲存格連結 → K 欄文字（檔名），
 * 之後交由 parseDriveId 解析成 Drive 檔案 ID。
 */
function getSignatureRecordsByDate(targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("點名紀錄");
  const range = recordSheet.getDataRange();
  const data = range.getValues();
  const richTexts = range.getRichTextValues();
  const chipUrlMap = readSignatureChipUris(ss, "點名紀錄", data.length);
  const targetDateFormatted = normalizeDate(targetDate);

  const grouped = {};
  for (let j = 1; j < data.length; j++) {
    const row = data[j];
    const clubName = String(row[RECORD_COL.CLUB]).trim();
    if (!clubName || normalizeDate(row[RECORD_COL.DATE]) !== targetDateFormatted) { continue; }

    const fileIdCell = row[RECORD_COL.FILE_ID] ? String(row[RECORD_COL.FILE_ID]).trim() : "";
    const sigRichText = richTexts[j] && richTexts[j][RECORD_COL.SIGNATURE]
      ? richTexts[j][RECORD_COL.SIGNATURE].getLinkUrl() : null;
    const signatureRaw = (/^[-\w]{25,}$/.test(fileIdCell) ? fileIdCell : null)
      || chipUrlMap[j] || sigRichText || row[RECORD_COL.SIGNATURE];

    if (!grouped[clubName]) { grouped[clubName] = []; }
    grouped[clubName].push({
      timestamp: new Date(row[RECORD_COL.TIMESTAMP]),
      fillerInfo: row[RECORD_COL.FILLER_INFO],
      fillerName: row[RECORD_COL.FILLER_NAME],
      presentCount: row[RECORD_COL.PRESENT_COUNT],
      signatureRaw: signatureRaw,
      recordType: row[RECORD_COL.TYPE] ? String(row[RECORD_COL.TYPE]).trim() : ""
    });
  }

  const result = {};
  Object.keys(grouped).forEach(function (clubName) {
    const rows = grouped[clubName].sort(function (x, y) { return x.timestamp - y.timestamp; });

    const createRecord = rows.filter(function (r) { return r.recordType === "新增"; })[0] || rows[0];
    let editRecords = rows.filter(function (r) { return r.recordType === "編輯"; });
    // 若無標註紀錄類別但有多筆提交，後續筆數視為編輯紀錄（向後相容）
    if (editRecords.length === 0 && rows.length > 1) {
      editRecords = rows.slice(1);
    }

    result[clubName] = {
      createRecord: createRecord,
      latestEdit: editRecords.length > 0 ? editRecords[editRecords.length - 1] : null,
      latestRecord: rows[rows.length - 1]
    };
  });
  return result;
}
