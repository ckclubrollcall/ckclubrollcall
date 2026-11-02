// ==========================================
// 0. 系統規格設定與安全常數
// ==========================================
const SECURITY = {
  MAX_DESC_LEN: 200,   // 社課內容簡述最大字數
  MAX_TEACHER_LEN: 30,    // 代課老師姓名最大字數
  MAX_CLUB_LEN: 30,    // 社團名稱最大字數
  MAX_EMAIL_LEN: 100,   // 電子郵件最大字數
  SUBMIT_COOLDOWN_SEC: 60,   // 同一帳號兩次提交的最短間隔時間（防重複發送，單位：秒）
  ALLOWED_DATE_RANGE: 5,    // 允許點名補填或修改的天數範圍（往回算 5 天內）
};

// ==========================================
// 1. 基本安全與防錯工具 (Validation & Sanitization)
// ==========================================

/**
 * 清除字串中的危險字元
 * 過濾 HTML 標籤與防止公式注入，保護試算表安全。
 */
function sanitize(val, maxLen) {
  if (val === null || val === undefined) { return ""; }
  let s = String(val)
    // eslint-disable-next-line no-control-regex
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "") // 去除不可見的控制字元
    .replace(/[<>]/g, "")      // 去除大於小於符號，防範網頁腳本注入
    .replace(/[=+\-@|`]/g, function (c) {   // 若開頭為試算表公式字元，前面加零寬字元避免公式執行
      return "\u200B" + c;
    })
    .trim();
  if (maxLen && s.length > maxLen) { s = s.slice(0, maxLen); }
  return s;
}

/**
 * 檢查電子郵件信箱格式是否正確
 */
function isValidEmail(email) {
  if (!email || typeof email !== "string") { return false; }
  if (email.length > SECURITY.MAX_EMAIL_LEN) { return false; }
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/**
 * 檢查點名日期是否在允許的 5 天範圍內
 */
function isValidDate(dateStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) { return false; }
  const d = new Date(dateStr.replace(/-/g, "/"));
  if (isNaN(d.getTime())) { return false; }

  const today = new Date();
  today.setHours(23, 59, 59, 999); // 設定為今天結束前的最後一刻

  const minDate = new Date();
  minDate.setDate(minDate.getDate() - SECURITY.ALLOWED_DATE_RANGE);
  minDate.setHours(0, 0, 0, 0); // 設定為 5 天前的起始時刻

  return d >= minDate && d <= today;
}

/**
 * 判斷填寫者身份是否為「幹部」（只有幹部可以點名與修改）
 */
function isOfficer(remark) {
  return remark !== "社員" && remark !== "外部登入_社員";
}

/**
 * 限制同一用戶的點名頻率，限制 60 秒內不得重複提交
 */
function checkAndSetCooldown(email) {
  const cache = CacheService.getScriptCache();
  const key = "submit_ts_" + email.replace(/[^a-zA-Z0-9]/g, "_");
  const lastTs = cache.get(key);

  if (lastTs) {
    return false;
  }
  cache.put(key, true, 60);
  return true;
}

/**
 * 自動過濾並解析雲端硬碟資料夾的 URL 取得實體 ID
 */
// eslint-disable-next-line no-redeclare
function parseDriveId(rawValue) {
  if (!rawValue) { return null; }
  const match = rawValue.toString().trim().match(/folders\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : rawValue.toString().trim();
}

/**
 * 寫入錯誤日誌至 Google 試算表，方便管理者除錯
 */
// eslint-disable-next-line no-redeclare
function logToSheet(msg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("系統日誌");
    if (!sheet) {
      sheet = ss.insertSheet("系統日誌");
      sheet.appendRow(["時間", "日誌內容"]);
    }
    sheet.appendRow([new Date(), String(msg)]);
  } catch {
    // 靜默失敗，防止日誌功能出錯導致主程式崩潰
  }
}

// ==========================================
// 2. 身分驗證核心 (Google Token 與測試 Token)
// ==========================================

/**
 * 驗證 Google 登入回傳的身分 Token (IdToken)
 */
function verifyIdToken(idToken) {
  if (!idToken) {
    logToSheet("verifyIdToken: idToken 為空");
    return null;
  }
  try {
    const url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return null;
    }
    const payload = JSON.parse(response.getContentText());
    const expectedClientId = PropertiesService.getScriptProperties().getProperty("OAUTH_CLIENT_ID");
    if (payload.aud !== expectedClientId) {
      logToSheet("Client ID 不符。");
      return null;
    }
    if (payload.email_verified !== "true" && payload.email_verified !== true) {
      return null;
    }
    return payload.email;
  } catch (e) {
    logToSheet("verifyIdToken 錯誤: " + e.toString());
    return null;
  }
}

/**
 * 統一 API 請求的身分認證入口
 */
function authenticateRequest(token) {
  if (!token) { return { error: true, msg: "未提供憑證，請重新登入" }; }

  let email;

  // DEV ONLY
  const LT = PropertiesService.getScriptProperties().getProperty("STRESSTEST_SECRET");
  if (LT && token && token.indexOf("LT:" + LT + ":") === 0) {
    email = token.split(":")[2];
    const user = getLoginUserExternal(email);
    return user.error ? { error: true, msg: user.msg } : { error: false, user: user };
  }

  try {
    email = verifyIdToken(token);
    if (!email) {
      return { error: true, needLogout: true, msg: "Google 登入憑證已過期或無效，請重新登入" };
    }
  } catch (e) {
    logToSheet("authenticateRequest 錯誤: " + e.toString());
    return { error: true, needLogout: true, msg: "Google 登入憑證驗證失敗，請重新登入" };
  }

  const user = getLoginUserExternal(email);
  if (user.error) {
    return { error: true, needLogout: true, msg: user.msg || "使用者資料讀取失敗，請聯絡管理員" };
  }

  return { error: false, user: user };
}

// ==========================================
// 3. 雲端資料夾與設定讀取
// ==========================================

/**
 * 取得「系統設定」試算表 B1 格中所儲存的手寫簽名存檔雲端資料夾 ID
 */
// eslint-disable-next-line no-redeclare
function getSignFolderId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const raw = ss.getSheetByName("系統設定").getRange("B1").getValue();
    if (!raw) { throw new Error("「系統設定」工作表的 B1 儲存格不可以是空白！"); }
    return parseDriveId(raw);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ 系統設定錯誤：" + e.message);
    logToSheet("getSignFolderId 錯誤: " + e.toString());
    return null;
  }
}

// ==========================================
// 4. 學生與教師登入處理
// ==========================================

/**
 * 以 Google 信箱於試算表「原始名單」中檢索其學生資訊與角色
 */
function getLoginUserExternal(email) {
  if (!isValidEmail(email)) { return { error: true, msg: "無效的電子郵件地址" }; }

  // 強制檢查網域，只允許建中官方網域帳號 @gl.ck.tp.edu.tw 登入
  if (!email.toLowerCase().endsWith("@gl.ck.tp.edu.tw")) {
    logToSheet("拒絕非官方網域登入嘗試: " + email);
    return { error: true, msg: "非建中官方網域帳號，無法登入" };
  }

  let studentId = "";
  const match = email.match(/^ck(\d+)/i); // 建中信箱前綴常為 ck + 學號
  if (match) { studentId = match[1]; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始名單");
  const data = sheet.getDataRange().getValues();

  const COL = { CLASS: 0, NO: 1, NAME: 2, ID: 3, REMARK: 4, CLUB_ID: 5, CLUB: 6 };

  for (let i = 1; i < data.length; i++) { // FIXME: Low efficiency, can be optimized.
    if (String(data[i][COL.ID]) === studentId) {
      const remark = data[i][COL.REMARK] || "社員";
      return {
        email: email,
        id: studentId,
        class: data[i][COL.CLASS],
        no: data[i][COL.NO],
        name: data[i][COL.NAME],
        club: data[i][COL.CLUB],
        clubId: data[i][COL.CLUB_ID],
        remark: remark,
        fillerInfo: remark + "-" + data[i][COL.CLASS] + " " + data[i][COL.NO] + " " + data[i][COL.NAME]
      };
    }
  }

  this.logToSheet("未找到對應學生資訊: " + email + " (學號: " + studentId + ")");
  return { error: true, msg: "未找到對應學生資訊" };
}

// ==========================================
// 5. 點名與名單資料存取業務 (GAS Database Logic)
// ==========================================

/**
 * 載入指定社團的所有社員名單
 * 支援模糊名稱與編號自動對照。
 */
function getClubMembers(clubName) {
  const safeClub = sanitize(clubName, SECURITY.MAX_CLUB_LEN).trim();
  if (!safeClub) { return []; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始名單");
  const data = sheet.getDataRange().getValues();
  const students = [];

  for (let i = 1; i < data.length; i++) {
    const sheetClub = data[i][6] ? String(data[i][6]).trim() : "";
    const sheetClubId = data[i][5] ? String(data[i][5]).trim() : "";
    if (sheetClub === safeClub || sheetClubId === safeClub || sheetClub.toLowerCase() === safeClub.toLowerCase() || sheetClubId.toLowerCase() === safeClub.toLowerCase()) {
      students.push({
        class: data[i][0],
        no: data[i][1],
        name: data[i][2],
        id: data[i][3]
      });
    }
  }

  // 照班級與座號排序，利於幹部點名
  students.sort(function (a, b) {
    return a.class !== b.class ? a.class - b.class : a.no - b.no;
  });

  return students;
}

function formatDateLocal(date) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  const d = date.getDate();
  return y + "-" + (m < 10 ? "0" : "") + m + "-" + (d < 10 ? "0" : "") + d;
}

/**
 * 抓取指定社團在歷史上的所有最新點名紀錄行（過濾舊紀錄）
 */
function getLatestRowsByClub(club) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("點名紀錄");
  const range = sheet.getDataRange();
  const data = range.getValues();
  const richTexts = range.getRichTextValues();
  const latest = {}; // 鍵值為 yyyy-MM-dd，數值為該日期最新的紀錄行

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== club) { continue; }
    const dDate = data[i][3];
    if (!(dDate instanceof Date)) { continue; }
    const key = formatDateLocal(dDate);
    const ts = data[i][0] instanceof Date ? data[i][0].getTime() : 0;

    // 若同一日期有重覆提交，只保留時間戳記最新的一筆
    if (!latest[key] || ts > latest[key]._ts) {
      // 支援 Google Drive 檔案智慧方塊：先自 richText 提取 URL，無則使用儲存格值
      const sigRichText = (richTexts && richTexts[i] && richTexts[i][10]) ? richTexts[i][10].getLinkUrl() : null;
      const sigValue = sigRichText || (data[i][10] ? String(data[i][10]).trim() : "");

      latest[key] = {
        _ts: ts,
        date: key,
        fillerInfo: data[i][1],
        club: data[i][2],
        fillerName: data[i][4],
        desc: data[i][5],
        teacherPresent: data[i][6],
        subTeacher: data[i][7],
        presentCount: data[i][8],
        absentList: data[i][9],
        signatureUrl: sigValue,
        recordType: data[i][11] ? String(data[i][11]).trim() : "新增"
      };
    }
  }
  return latest;
}

/**
 * 判斷特定日期的點名是否依然在 5 天可修改範圍內
 */
function isWithinEditableRange(dateStr) {
  return isValidDate(dateStr);
}

/**
 * 取得該社團歷史上所有點名紀錄之列表概要，提供前台選單呈現
 */
function getAttendanceList(club) {
  const safeClub = sanitize(club, SECURITY.MAX_CLUB_LEN);
  if (!safeClub) { return []; }
  const latest = getLatestRowsByClub(safeClub);

  const list = Object.keys(latest).map(function (key) {
    const r = latest[key];
    return {
      date: r.date,
      fillerName: r.fillerName,
      teacherPresent: r.teacherPresent,
      presentCount: r.presentCount,
      editable: isWithinEditableRange(r.date) // 標示目前該日期是否可修改
    };
  });

  list.sort(function (a, b) { return a.date < b.date ? 1 : -1; }); // 新到舊排序
  return list;
}

/**
 * 取得指定社團於指定日期點名之詳細完整資訊
 */
function getAttendanceDetail(club, date) {
  const safeClub = sanitize(club, SECURITY.MAX_CLUB_LEN);
  const latest = getLatestRowsByClub(safeClub);
  const r = latest[date];
  if (!r) { return { status: "error", message: "找不到該日期的紀錄" }; }
  r.editable = isWithinEditableRange(r.date);
  return r;
}

/**
 * 供一般社員查詢個人專屬的出缺席記錄歷史
 */
function getMyAttendance(club, identity) {
  const safeClub = sanitize(club, SECURITY.MAX_CLUB_LEN);
  const safeIdentity = sanitize(identity, 30);
  if (!safeClub || !safeIdentity) { return []; }

  const latest = getLatestRowsByClub(safeClub);

  return Object.keys(latest).map(function (key) {
    const r = latest[key];
    const absentLines = String(r.absentList || "").split("\n");
    const isAbsent = absentLines.indexOf(safeIdentity) !== -1; // 檢查缺席清單是否有自己
    return { date: r.date, status: isAbsent ? "缺席" : "出席" };
  }).sort(function (a, b) { return a.date < b.date ? 1 : -1; });
}

/**
 * 將前端傳入的點名資訊進行安全性驗證，並上傳老師手寫簽名，最後寫入試算表新列
 */
function submitAttendance(data) {
  if (!data) { return { status: "error", message: "資料不完整" }; }

  try {
    let safeEmail = "";
    if (data.email && data.email !== "") {
      if (!isValidEmail(data.email)) { return { status: "error", message: "Email 格式不正確" }; }
      safeEmail = data.email;
    }

    // 1. 優先檢查提交頻率，避免短時間內重複送出。
    // 沿用原本的識別方式：優先使用 safeEmail，否則使用填寫者資訊。
    const cooldownKey = safeEmail || (data.fillerInfo || "anonymous");

    if (!checkAndSetCooldown(cooldownKey)) {
      return {
        status: "error",
        message: "送出過於頻繁，請稍後再試"
      };
    }


    // 2. 驗證日期。
    if (!isValidDate(data.date)) {
      return {
        status: "error",
        message: "日期格式不正確或超出允許範圍"
      };
    }

    // 3. 清洗文字內容，並限制各欄位長度。
    const safeDesc = sanitize(
      data.desc,
      SECURITY.MAX_DESC_LEN
    );

    const safeSubTeacher = sanitize(
      data.subTeacher,
      SECURITY.MAX_TEACHER_LEN
    );

    const safeClub = sanitize(
      data.club,
      SECURITY.MAX_CLUB_LEN
    );

    const safeFillerInfo = sanitize(data.fillerInfo, 60);
    const safeFillerName = sanitize(data.fillerName, 20);

    const safeTeacherPresent = (data.teacherPresent === "否") ? "否" : "是";


    // 4. 驗證必填欄位。
    if (!safeDesc || !safeClub) {
      return {
        status: "error",
        message: "資料不完整"
      };
    }


    // 5. 取得該社團的正式名單。
    const members = getClubMembers(safeClub);

    if (!members || !members.length) {
      return {
        status: "error",
        message: "找不到該社團正式名單"
      };
    }


    // 6. 建立正式名單索引。
    // 使用無原型物件，讓查找只依據實際加入的學生身分。
    const memberMap = Object.create(null);

    members.forEach(function (m) {
      const identity = sanitize(
        m.class + " " + m.no + " " + m.name,
        30
      );

      memberMap[identity] = true;
    });


    // 7. 處理缺席名單：限制行數、清洗、驗證及去重。
    const absentIdentities = [];
    const seen = Object.create(null);

    if (data.absentList && data.absentList !== "全勤") {
      // 同時處理 Windows、Unix 與舊式 Mac 換行。
      let absentLines = String(data.absentList).split(/\r\n|\n|\r/);

      // 不直接截斷，避免遺漏缺席學生後仍繼續計算出席人數。
      if (absentLines.length > 200) {
        return {
          status: "error",
          message: "缺席名單不可超過 200 行"
        };
      }

      absentLines = absentLines
        .map(function (line) {
          return sanitize(line, 30);
        })
        .filter(function (line) {
          return line.length > 0;
        });

      for (let i = 0; i < absentLines.length; i++) {
        const identity = absentLines[i];

        if (!memberMap[identity]) {
          return {
            status: "error",
            message: "缺席名單包含非本社團學生或格式不正確"
          };
        }

        if (!seen[identity]) {
          seen[identity] = true;
          absentIdentities.push(identity);
        }
      }
    }

    // 8. 產生最終缺席名單，並由正式名單推算出席人數。
    // 不採用前端傳入的 data.presentCount。
    const safeAbsentList = absentIdentities.length
      ? absentIdentities.join("\n")
      : "全勤";
    const safePresentCount = members.length - absentIdentities.length;

    // 9. 檢查既有紀錄以判定「紀錄類別」（新增 vs 編輯），防範前端偽造傳遞值
    const existingRecords = getLatestRowsByClub(safeClub);
    const isEdit = !!existingRecords[data.date];
    const recordType = isEdit ? "編輯" : "新增";

    // 10. 驗證簽名：徹底廢止 KEEP 邏輯，修改點名紀錄者必須重新簽名
    if (data.signature === "KEEP") {
      return { status: "error", message: "修改點名紀錄必須重新簽名，不可沿用原簽名" };
    }

    if (!data.signature || data.signature === "") {
      return { status: "error", message: isEdit ? "修改者必須完成簽名" : "指導老師必須完成簽名" };
    }

    if (data.signature.length > 200000) {
      return { status: "error", message: "簽名檔過大，請重新簽名" };
    }

    if (!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(data.signature)) {
      return { status: "error", message: "簽名格式不正確" };
    }

    // 11. 將 Base64 格式轉回 PNG 圖檔，並依類別命名後寫入指定 Google Drive 資料夾
    const folderId = getSignFolderId();
    const folder = DriveApp.getFolderById(folderId);
    const base64 = data.signature.split(",")[1];
    const fileName = data.date + "_" + safeClub + (isEdit ? "_修改者簽名.png" : "_老師簽名.png");
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), "image/png", fileName);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl(); // 取得檔案的公開/內部連結

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return { status: "error", message: "系統忙碌，請稍後再試" };
    }

    try {

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("點名紀錄");
      const formattedDate = new Date(data.date.replace(/-/g, "/"));

      // 確保點名紀錄標題列包含第 L 欄「紀錄類別」
      if (sheet.getLastRow() >= 1) {
        const headerL = sheet.getRange(1, 12).getValue();
        if (!headerL) {
          sheet.getRange(1, 12).setValue("紀錄類別");
        }
      }

      // 附加新列至點名紀錄試算表
      sheet.appendRow([
        new Date(),            // A: 填寫日期時間
        safeFillerInfo,        // B: 填答人身分類別
        safeClub,              // C: 社團官方名稱
        formattedDate,         // D: 點名社課日期
        safeFillerName,        // E: 填寫人姓名
        safeDesc,              // F: 社課內容簡述
        safeTeacherPresent,    // G: 指導老師出席與否
        safeSubTeacher,        // H: 代課老師姓名
        safePresentCount,      // I: 出席總人數
        safeAbsentList,        // J: 換行缺席名單
        fileUrl,               // K: 老師簽名 / 修改者簽名 (將嘗試轉換為 Google Drive 智慧方塊)
        recordType             // L: 紀錄類別 (新增 / 編輯)
      ]);

      const lastRow = sheet.getLastRow();

      // 12. 優先使用 Google Sheets API (v4 batchUpdate chipRuns) 產生 Google Drive 檔案智慧方塊
      try {
        if (typeof Sheets !== "undefined" && Sheets.Spreadsheets && fileUrl) {
          const spreadsheetId = ss.getId();
          const sheetId = sheet.getSheetId();
          const rowIndex = lastRow - 1; // 0-indexed
          const colIndex = 10;          // K 欄為第 11 欄 -> 0-indexed 為 10

          Sheets.Spreadsheets.batchUpdate({
            requests: [{
              updateCells: {
                rows: [{
                  values: [{
                    userEnteredValue: { stringValue: "@" },
                    chipRuns: [{
                      startIndex: 0,
                      chip: {
                        richLinkProperties: {
                          uri: fileUrl
                        }
                      }
                    }]
                  }]
                }],
                range: {
                  sheetId: sheetId,
                  startRowIndex: rowIndex,
                  endRowIndex: rowIndex + 1,
                  startColumnIndex: colIndex,
                  endColumnIndex: colIndex + 1
                },
                fields: "userEnteredValue,chipRuns"
              }
            }]
          }, spreadsheetId);
        }
      } catch (chipError) {
        logToSheet("Drive 檔案智慧方塊轉換失敗，已保留原始連結: " + chipError.toString());
        // 若 Sheets API 發生錯誤或未啟用，已保留原先寫入之 fileUrl 作為安全降級
      }

      return "SUCCESS";
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    logToSheet("submitAttendance 錯誤: " + err.toString());
    return { status: "error", message: "提交點名資料失敗，請稍後再試" };
  }
}

// ==========================================
// 6. API 路由器 (doPost / doGet)
// ==========================================

/**
 * 格式化與回傳 JSON 格式字串並附加 CORS 標頭，允許跨網域存取
 */
function jsonResponse(data, headers) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
    ;
  return output;
}

/**
 * 響應 HTTP GET 請求 (純提示)
 */
// eslint-disable-next-line no-unused-vars
function doGet(e) {
  return ContentService.createTextOutput("建中社團點名系統 API endpoint. 請使用 POST 方法存取。");
}

/**
 * 核心 Web API 控制器
 * 接收從網頁前端發過來的 POST 請求，並根據 action 指令轉發至對應處理函式。
 */
// eslint-disable-next-line no-unused-vars
function doPost(e) {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  if (!e || !e.postData) {
    return jsonResponse({ status: "error", message: "沒有收到任何資料" }, corsHeaders);
  }

  try {
    const req = JSON.parse(e.postData.contents);
    const action = req.action;
    let result;
    const token = (action === "submitAttendance") ? (req.data && req.data.token) : req.token;
    const auth = authenticateRequest(token);
    if (auth.error) {
      return jsonResponse(auth, corsHeaders);
    }

    // API Action 路由對照表
    switch (action) {
      case "getLoginUser":
        result = auth.user;
        break;
      case "getClubMembers":
        if (auth.user.club !== req.clubName && auth.user.remark !== "管理員") {
          return jsonResponse({ status: "error", message: "無權限查看此社團名單" }, corsHeaders);
        }
        result = getClubMembers(req.clubName);
        break;
      case "submitAttendance": {
        if (auth.user.club !== req.data.club) {
          return jsonResponse({ status: "error", message: "無權限為此社團提交點名資料" }, corsHeaders);
        }
        if (!isOfficer(auth.user.remark)) {
          return jsonResponse({ status: "error", message: "只有社團幹部才能提交或修改點名資料" }, corsHeaders);
        }
        req.data.email = auth.user.email;
        req.data.fillerInfo = auth.user.fillerInfo;
        req.data.fillerName = auth.user.name;

        result = submitAttendance(req.data);
        break;
      }
      case "getAttendanceList": {
        if (auth.user.club !== req.club) {
          return jsonResponse({ status: "error", message: "無權限查看此社團的紀錄" }, corsHeaders);
        }
        if (!isOfficer(auth.user.remark)) {
          return jsonResponse({ status: "error", message: "只有幹部才能查看社團點名列表" }, corsHeaders);
        }
        result = getAttendanceList(req.club);
        break;
      }
      case "getAttendanceDetail": {
        if (auth.user.club !== req.club) {
          return jsonResponse({ status: "error", message: "無權限查看此社團的紀錄詳情" }, corsHeaders);
        }
        if (!isOfficer(auth.user.remark)) {
          return jsonResponse({ status: "error", message: "無權限查看此紀錄詳情" }, corsHeaders);
        }
        result = getAttendanceDetail(req.club, req.date);
        break;
      }
      case "getMyAttendance": {
        if (auth.user.club !== req.club) {
          return jsonResponse({ status: "error", message: "無權限查看此社團的出缺席紀錄" }, corsHeaders);
        }
        const expectedIdentity = auth.user.class + " " + auth.user.no + " " + auth.user.name;
        if (expectedIdentity.trim() !== req.identity.trim()) {
          return jsonResponse({ status: "error", message: "無權限查看他人出缺席紀錄" }, corsHeaders);
        }
        result = getMyAttendance(req.club, req.identity);
        break;
      }
      default: {
        result = { status: "error", message: "未知的操作指令" };
      }
    }
    return jsonResponse(result, corsHeaders);
  } catch (err) {
    logToSheet("doPost 錯誤: " + err.toString());
    return jsonResponse({ status: "error", message: "請求處理失敗，請稍後再試" }, corsHeaders);
  }
}
