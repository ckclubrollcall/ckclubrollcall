// ==========================================
// 0. 系統規格設定與安全常數
// ==========================================
var SECURITY = {
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
  if (val === null || val === undefined) return "";
  var s = String(val)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "") // 去除不可見的控制字元
    .replace(/[<>]/g, "")      // 去除大於小於符號，防範網頁腳本注入
    .replace(/[=+\-@|`]/g, function (c) {   // 若開頭為試算表公式字元，前面加零寬字元避免公式執行
      return "\u200B" + c;
    })
    .trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

/**
 * 檢查電子郵件信箱格式是否正確
 */
function isValidEmail(email) {
  if (!email || typeof email !== "string") return false;
  if (email.length > SECURITY.MAX_EMAIL_LEN) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/**
 * 檢查點名日期是否在允許的 5 天範圍內
 */
function isValidDate(dateStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return false;
  var d = new Date(dateStr.replace(/-/g, "/"));
  if (isNaN(d.getTime())) return false;

  var today = new Date();
  today.setHours(23, 59, 59, 999); // 設定為今天結束前的最後一刻

  var minDate = new Date();
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
  var props = PropertiesService.getScriptProperties();
  var key = "submit_ts_" + email.replace(/[^a-zA-Z0-9]/g, "_");
  var lastTs = props.getProperty(key);

  if (lastTs) {
    var elapsed = (Date.now() - parseInt(lastTs, 10)) / 1000;
    if (elapsed < SECURITY.SUBMIT_COOLDOWN_SEC) {
      return false; // 仍在冷卻中
    }
  }
  props.setProperty(key, String(Date.now())); // 更新最後提交時間
  return true;
}

/**
 * 自動過濾並解析雲端硬碟資料夾的 URL 取得實體 ID
 */
function parseDriveId(rawValue) {
  if (!rawValue) return null;
  var match = rawValue.toString().trim().match(/folders\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : rawValue.toString().trim();
}

/**
 * 寫入錯誤日誌至 Google 試算表，方便管理者除錯
 */
function logToSheet(msg) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("系統日誌");
    if (!sheet) {
      sheet = ss.insertSheet("系統日誌");
      sheet.appendRow(["時間", "日誌內容"]);
    }
    sheet.appendRow([new Date(), String(msg)]);
  } catch (e) {
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
    var url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return null;
    }
    var payload = JSON.parse(response.getContentText());
    var expectedClientId = "311035173241-nuamoemc7al4bhlp5p0nploepd6rg23h.apps.googleusercontent.com";
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

// 金鑰：加密安全，防止偽造 Token
var TEST_TOKEN_SECRET = PropertiesService.getScriptProperties().getProperty("TEST_TOKEN_SECRET") || "ck-club-rollcall-fallback-secret-2026";

/**
 * SHA256 加密計算工具，用於密碼安全比對與 Token 簽章
 */
function sha256(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  var output = "";
  for (var i = 0; i < rawHash.length; i++) {
    var value = rawHash[i];
    if (value < 0) value += 256;
    var byteString = value.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    output += byteString;
  }
  return output;
}

/**
 * 建立測試登入的專屬加密 Token（時效 14 天）
 * 把登入時查到的完整身分資料都簽進 Token，
 * 之後不管重新驗證幾次，都能還原出「跟第一次登入時一模一樣」的資料，
 * 不需要再用假資料重建。
 */
function generateTestToken(userInfo) {
  var expiry = Date.now() + 14 * 24 * 60 * 60 * 1000;
  var payload = JSON.stringify({
    email: userInfo.email,
    club: userInfo.club,
    clubId: userInfo.clubId,
    name: userInfo.name,
    remark: userInfo.remark,
    exp: expiry
  });
  // 明確指定 UTF-8，確保中文字不會在 Base64 編碼時損毀
  var signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payload, TEST_TOKEN_SECRET, Utilities.Charset.UTF_8);
  var sigStr = Utilities.base64Encode(signature);
  return "TEST_" + Utilities.base64Encode(payload, Utilities.Charset.UTF_8) + "." + sigStr;
}

/**
 * 驗證並解密測試登入 Token
 */
function verifyTestToken(token) {
  if (!token || !token.startsWith("TEST_")) return null;
  try {
    var parts = token.substring(5).split(".");
    if (parts.length !== 2) return null;
    // 明確指定 UTF-8，與 generateTestToken 的編碼方式一致
    var payloadStr = Utilities.newBlob(Utilities.base64Decode(parts[0])).getDataAsString("UTF-8");
    var expectedSig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payloadStr, TEST_TOKEN_SECRET, Utilities.Charset.UTF_8);
    var expectedSigStr = Utilities.base64Encode(expectedSig);
    if (parts[1] !== expectedSigStr) return null; // 憑證遭人修改

    var obj = JSON.parse(payloadStr);
    if (Date.now() > obj.exp) return null; // 憑證過期
    return obj; // 內含 email / club / clubId / name / remark
  } catch (e) {
    return null;
  }
}

/**
 * 統一 API 請求的身分認證入口
 */
function authenticateRequest(token) {
  if (!token) return { error: true, msg: "未提供憑證，請重新登入" };

  var email = null;
  var isTestUser = false;
  var testPayload = null;

  if (token.startsWith("TEST_")) {
    testPayload = verifyTestToken(token);
    if (!testPayload) {
      return { error: true, needLogout: true, msg: "測試登入憑證已過期或無效，請重新登入" };
    }
    email = testPayload.email;
    isTestUser = true;
  } else {
    email = verifyIdToken(token);
    if (!email) {
      return { error: true, needLogout: true, msg: "Google 登入憑證已過期或無效，請重新登入" };
    }
  }

  var user = getLoginUserExternal(email);
  if (user.needManualLogin && !isTestUser) {
    return { error: false, user: { needManualLogin: true, defaultEmail: email } };
  }

  if (isTestUser) {
    // 直接還原 Token 裡簽署的完整資料，跟當初登入時一字不差，
    // 不再用「測試人員」這種假資料覆蓋掉真實姓名/社團代碼
    user = {
      email: testPayload.email,
      id: "TEST_USER",
      class: "000",
      no: "00",
      name: testPayload.name,
      club: testPayload.club,
      clubId: testPayload.clubId,
      remark: testPayload.remark,
      fillerInfo: testPayload.remark + "-" + testPayload.name
    };
    return { error: false, user: user };
  }

  user = getLoginUserExternal(email);
  if (user.needManualLogin && !isTestUser) {
    return { error: false, user: { needManualLogin: true, defaultEmail: email } };
  }

  return { error: false, user: user };
}

// ==========================================
// 3. 雲端資料夾與設定讀取
// ==========================================

/**
 * 取得「系統設定」試算表 B1 格中所儲存的手寫簽名存檔雲端資料夾 ID
 */
function getSignFolderId() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var raw = ss.getSheetByName("系統設定").getRange("B1").getValue();
    if (!raw) throw new Error("「系統設定」工作表的 B1 儲存格不可以是空白！");
    return parseDriveId(raw);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ 系統設定錯誤：" + e.message);
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
  if (!isValidEmail(email)) return { needManualLogin: true };

  // 強制檢查網域，只允許建中官方網域帳號 @gl.ck.tp.edu.tw 登入
  if (!email.toLowerCase().endsWith("@gl.ck.tp.edu.tw")) {
    logToSheet("拒絕非官方網域登入嘗試: " + email);
    return { needManualLogin: true };
  }

  var studentId = "";
  var match = email.match(/^ck(\d+)/i); // 建中信箱前綴常為 ck + 學號
  if (match) studentId = match[1];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始名單");
  var data = sheet.getDataRange().getValues();

  var COL = { CLASS: 0, NO: 1, NAME: 2, ID: 3, REMARK: 4, CLUB_ID: 5, CLUB: 6 };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.ID]) === studentId) {
      var remark = data[i][COL.REMARK] || "社員";
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

  return { needManualLogin: true, defaultEmail: email };
}

/**
 * 外部非官方帳號測試登入處理
 * 會自動規格化使用者輸入的社團名稱/編號，提高測試登入成功率。
 */
function verifyTestLogin(inputEmail, inputPassword, inputClub, inputIdentity) {
  if (!checkAndSetCooldown("login_try_" + inputEmail)) {
    return { error: true, msg: "嘗試登入過於頻繁，請等待 60 秒後再試" };
  }
  if (!isValidEmail(inputEmail)) return { error: true, msg: "Email 格式不正確" };
  if (!inputPassword || typeof inputPassword !== "string" || inputPassword.length > 100) {
    return { error: true, msg: "密碼格式不正確" };
  }
  var safeClub = sanitize(inputClub, SECURITY.MAX_CLUB_LEN);
  if (!safeClub) return { error: true, msg: "社團名稱不可為空" };

  var safeIdentity = (inputIdentity === "社員") ? "社員" : "幹部";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("系統設定");
  var data = sheet.getDataRange().getValues();

  for (var i = 4; i < data.length; i++) {
    var rowEmail = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    if (rowEmail === inputEmail.toLowerCase()) {
      var teacherName = data[i][1] ? data[i][1].toString().trim() : "測試人員";
      var correctPassword = data[i][2] ? data[i][2].toString().trim() : "";

      // 雜湊比對密碼
      var inputPasswordHash = sha256(inputPassword);
      if (inputPasswordHash !== correctPassword.toString().trim().toLowerCase()) {
        return { error: true, msg: "密碼錯誤！請重新確認。" };
      }

      var remarkVal = "外部登入_" + safeIdentity;

      // 尋找最相符的社團官方名稱與編號 (進行自動模糊容錯比對)
      var normalizedClub = safeClub;
      var normalizedClubId = "測試";
      try {
        var memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始名單");
        if (memberSheet) {
          var memberData = memberSheet.getDataRange().getValues();
          var foundExact = false;
          var foundPartial = false;
          var partialClub = "";
          var partialClubId = "";
          for (var j = 1; j < memberData.length; j++) {
            var sheetClub = memberData[j][6] ? String(memberData[j][6]).trim() : "";
            var sheetClubId = memberData[j][5] ? String(memberData[j][5]).trim() : "";
            if (sheetClub.toLowerCase() === safeClub.toLowerCase() || sheetClubId.toLowerCase() === safeClub.toLowerCase()) {
              normalizedClub = sheetClub;
              normalizedClubId = sheetClubId;
              foundExact = true;
              break;
            }
            if (!foundPartial && sheetClub && (sheetClub.toLowerCase().indexOf(safeClub.toLowerCase()) !== -1 || safeClub.toLowerCase().indexOf(sheetClub.toLowerCase()) !== -1)) {
              partialClub = sheetClub;
              partialClubId = sheetClubId;
              foundPartial = true;
            }
          }
          if (!foundExact && foundPartial) {
            normalizedClub = partialClub;
            normalizedClubId = partialClubId;
          }
        }
      } catch (err) {
        // 忽略錯誤，使用原本的輸入
      }

      var token = generateTestToken({
        email: inputEmail,
        club: normalizedClub,
        clubId: normalizedClubId,
        name: teacherName,
        remark: remarkVal
      });

      return {
        email: inputEmail,
        id: "TEST_USER",
        class: "000",
        no: "00",
        name: teacherName,
        club: normalizedClub,
        clubId: normalizedClubId,
        remark: remarkVal,
        fillerInfo: remarkVal + "-" + teacherName,
        token: token
      };
    }
  }
  return { error: true, msg: "您並非測試人員，請聯絡管理員建立測試帳號" };
}

// ==========================================
// 5. 點名與名單資料存取業務 (GAS Database Logic)
// ==========================================

/**
 * 載入指定社團的所有社員名單
 * 支援模糊名稱與編號自動對照。
 */
function getClubMembers(clubName) {
  var safeClub = sanitize(clubName, SECURITY.MAX_CLUB_LEN).trim();
  if (!safeClub) return [];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始名單");
  var data = sheet.getDataRange().getValues();
  var students = [];

  for (var i = 1; i < data.length; i++) {
    var sheetClub = data[i][6] ? String(data[i][6]).trim() : "";
    var sheetClubId = data[i][5] ? String(data[i][5]).trim() : "";
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
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  return y + "-" + (m < 10 ? "0" : "") + m + "-" + (d < 10 ? "0" : "") + d;
}

/**
 * 抓取指定社團在歷史上的所有最新點名紀錄行（過濾舊紀錄）
 */
function getLatestRowsByClub(club) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("點名紀錄");
  var data = sheet.getDataRange().getValues();
  var latest = {}; // 鍵值為 yyyy-MM-dd，數值為該日期最新的紀錄行

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== club) continue;
    var dDate = data[i][3];
    if (!(dDate instanceof Date)) continue;
    var key = formatDateLocal(dDate);
    var ts = data[i][0] instanceof Date ? data[i][0].getTime() : 0;

    // 若同一日期有重覆提交，只保留時間戳記最新的一筆
    if (!latest[key] || ts > latest[key]._ts) {
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
        signatureUrl: data[i][10]
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
  var safeClub = sanitize(club, SECURITY.MAX_CLUB_LEN);
  if (!safeClub) return [];
  var latest = getLatestRowsByClub(safeClub);

  var list = Object.keys(latest).map(function (key) {
    var r = latest[key];
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
  var safeClub = sanitize(club, SECURITY.MAX_CLUB_LEN);
  var latest = getLatestRowsByClub(safeClub);
  var r = latest[date];
  if (!r) return { status: "error", message: "找不到該日期的紀錄" };
  r.editable = isWithinEditableRange(r.date);
  return r;
}

/**
 * 供一般社員查詢個人專屬的出缺席記錄歷史
 */
function getMyAttendance(club, identity) {
  var safeClub = sanitize(club, SECURITY.MAX_CLUB_LEN);
  var safeIdentity = sanitize(identity, 30);
  if (!safeClub || !safeIdentity) return [];

  var latest = getLatestRowsByClub(safeClub);

  return Object.keys(latest).map(function (key) {
    var r = latest[key];
    var absentLines = String(r.absentList || "").split("\n");
    var isAbsent = absentLines.indexOf(safeIdentity) !== -1; // 檢查缺席清單是否有自己
    return { date: r.date, status: isAbsent ? "缺席" : "出席" };
  }).sort(function (a, b) { return a.date < b.date ? 1 : -1; });
}

/**
 * 將前端傳入的點名資訊進行安全性驗證，並上傳老師手寫簽名，最後寫入試算表新列
 */
function submitAttendance(data) {
  if (!data) return { status: "error", message: "資料不完整" };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { status: "error", message: "系統忙碌，請稍後再試" };
  }

  try {
    var safeEmail = "";
    if (data.email && data.email !== "") {
      if (!isValidEmail(data.email)) return { status: "error", message: "Email 格式不正確" };
      safeEmail = data.email;
    }

    // 1. 優先檢查提交頻率，避免短時間內重複送出。
    // 沿用原本的識別方式：優先使用 safeEmail，否則使用填寫者資訊。
    var cooldownKey = safeEmail || (data.fillerInfo || "anonymous");

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
    var safeDesc = sanitize(
      data.desc,
      SECURITY.MAX_DESC_LEN
    );

    var safeSubTeacher = sanitize(
      data.subTeacher,
      SECURITY.MAX_TEACHER_LEN
    );

    var safeClub = sanitize(
      data.club,
      SECURITY.MAX_CLUB_LEN
    );

    var safeFillerInfo = sanitize(data.fillerInfo, 60);
    var safeFillerName = sanitize(data.fillerName, 20);

    var safeTeacherPresent = (data.teacherPresent === "否") ? "否" : "是";


    // 4. 驗證必填欄位。
    if (!safeDesc || !safeClub) {
      return {
        status: "error",
        message: "資料不完整"
      };
    }


    // 5. 取得該社團的正式名單。
    var members = getClubMembers(safeClub);

    if (!members || !members.length) {
      return {
        status: "error",
        message: "找不到該社團正式名單"
      };
    }


    // 6. 建立正式名單索引。
    // 使用無原型物件，讓查找只依據實際加入的學生身分。
    var memberMap = Object.create(null);

    members.forEach(function (m) {
      var identity = sanitize(
        m.class + " " + m.no + " " + m.name,
        30
      );

      memberMap[identity] = true;
    });


    // 7. 處理缺席名單：限制行數、清洗、驗證及去重。
    var absentIdentities = [];
    var seen = Object.create(null);

    if (data.absentList && data.absentList !== "全勤") {
      // 同時處理 Windows、Unix 與舊式 Mac 換行。
      var absentLines = String(data.absentList).split(/\r\n|\n|\r/);

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

      for (var i = 0; i < absentLines.length; i++) {
        var identity = absentLines[i];

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
    var safeAbsentList = absentIdentities.length
      ? absentIdentities.join("\n")
      : "全勤";
    var safePresentCount = members.length - absentIdentities.length;

    var signatureData = "";
    if (data.signature === 'KEEP') {
      // 沿用原有簽名：從最新的一筆歷史紀錄拿回原本的雲端連結
      var existing = getLatestRowsByClub(safeClub);
      var existingRow = existing[data.date];
      signatureData = existingRow ? existingRow.signatureUrl : '無簽名';
    } else {
      if (data.signature && data.signature.length > 200000) {
        return { status: "error", message: "簽名檔過大，請重新簽名" };
      }
      if (data.signature && data.signature !== "") {
        if (!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(data.signature)) {
          return { status: "error", message: "簽名格式不正確" };
        }
        signatureData = data.signature;
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("點名紀錄");
    var formattedDate = new Date(data.date.replace(/-/g, "/"));

    // 老師手寫簽名上傳處理
    var fileUrl = "無簽名";
    if (signatureData === '無簽名' || signatureData === '') {
      fileUrl = "無簽名";
    } else if (signatureData.startsWith('https://')) {
      fileUrl = signatureData; // 沿用原本的連結
    } else {
      // 將 Base64 格式轉回 PNG 圖檔，並寫入指定 Google Drive 資料夾
      var folderId = getSignFolderId();
      var folder = DriveApp.getFolderById(folderId);
      var base64 = signatureData.split(",")[1];
      var fileName = data.date + "_" + safeClub + "_老師簽名.png";
      var blob = Utilities.newBlob(Utilities.base64Decode(base64), "image/png", fileName);
      var file = folder.createFile(blob);
      fileUrl = file.getUrl(); // 取得檔案的公開/內部連結
    }

    // 附加新列至點名紀錄試算表
    sheet.appendRow([
      new Date(),            // A: 填寫日期時間
      safeFillerInfo,        // B: 填答人學號資訊
      safeClub,              // C: 社團官方名稱
      formattedDate,         // D: 點名社課日期
      safeFillerName,        // E: 填寫人姓名
      safeDesc,              // F: 社課內容簡述
      safeTeacherPresent,    // G: 指導老師出席與否
      safeSubTeacher,        // H: 代課老師姓名
      safePresentCount,      // I: 出席總人數
      safeAbsentList,        // J: 換行缺席名單
      fileUrl                // K: 簽名檔 Drive 網址
    ]);

    return "SUCCESS";
  } finally {
    lock.releaseLock();
  }
  }

// ==========================================
// 6. API 路由器 (doPost / doGet)
// ==========================================

/**
 * 格式化與回傳 JSON 格式字串並附加 CORS 標頭，允許跨網域存取
 */
function jsonResponse(data, headers) {
    var output = ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
    return output;
  }

  /**
   * 響應 HTTP GET 請求 (純提示)
   */
  function doGet(e) {
    return ContentService.createTextOutput("建中社團點名系統 API endpoint. 請使用 POST 方法存取。");
  }

  /**
   * 核心 Web API 控制器
   * 接收從網頁前端發過來的 POST 請求，並根據 action 指令轉發至對應處理函式。
   */
  function doPost(e) {
    var corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    };

    if (!e || !e.postData) {
      return jsonResponse({ status: "error", message: "沒有收到任何資料" }, corsHeaders);
    }

    try {
      var req = JSON.parse(e.postData.contents);
      var action = req.action;
      var result;

      if (action === "verifyTestLogin") {
        result = verifyTestLogin(req.email, req.password, req.club, req.identity);
      } else {
        var token = (action === "submitAttendance") ? (req.data && req.data.token) : req.token;
        var auth = authenticateRequest(token);
        if (auth.error) {
          return jsonResponse(auth, corsHeaders);
        }

        // API Action 路由對照表
        if (action === "getLoginUser") {
          result = auth.user;
        } else if (action === "getClubMembers") {
          if (auth.user.club !== req.clubName && auth.user.remark !== "管理員") {
            return jsonResponse({ status: "error", message: "無權限查看此社團名單" }, corsHeaders);
          }
          result = getClubMembers(req.clubName);
        } else if (action === "submitAttendance") {
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
        } else if (action === "getAttendanceList") {
          if (auth.user.club !== req.club) {
            return jsonResponse({ status: "error", message: "無權限查看此社團的紀錄" }, corsHeaders);
          }
          if (!isOfficer(auth.user.remark)) {
            return jsonResponse({ status: "error", message: "只有幹部才能查看社團點名列表" }, corsHeaders);
          }
          result = getAttendanceList(req.club);
        } else if (action === "getAttendanceDetail") {
          if (auth.user.club !== req.club) {
            return jsonResponse({ status: "error", message: "無權限查看此社團的紀錄詳情" }, corsHeaders);
          }
          if (!isOfficer(auth.user.remark)) {
            return jsonResponse({ status: "error", message: "無權限查看此紀錄詳情" }, corsHeaders);
          }
          result = getAttendanceDetail(req.club, req.date);
        } else if (action === "getMyAttendance") {
          if (auth.user.club !== req.club) {
            return jsonResponse({ status: "error", message: "無權限查看此社團的出缺席紀錄" }, corsHeaders);
          }
          var expectedIdentity = auth.user.class + " " + auth.user.no + " " + auth.user.name;
          if (auth.user.id === "TEST_USER") {
            expectedIdentity = req.identity;
          }
          if (expectedIdentity.trim() !== req.identity.trim()) {
            return jsonResponse({ status: "error", message: "無權限查看他人出缺席紀錄" }, corsHeaders);
          }
          result = getMyAttendance(req.club, req.identity);
        } else {
          result = { status: "error", message: "未知的操作指令" };
        }
      }

      return jsonResponse(result, corsHeaders);

    } catch (err) {
      return jsonResponse({ status: "error", message: "請求處理失敗，請稍後再試" }, corsHeaders);
    }
  }

  /**
   * 系統驗證授權專用之測試函數
   */
  function forceAuth() {
    var res = UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo", { muteHttpExceptions: true });
    Logger.log("測試連線狀態碼: " + res.getResponseCode());
  }