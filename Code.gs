// ╔══════════════════════════════════════════════════════════════╗
// ║      Google Apps Script — 題庫系統  v9-6851                  ║
// ║      對應前端版本：quiz_final v1.6851                        ║
// ║                                                              ║
// ║  更新紀錄：                                                   ║
// ║  v9-6851 - Firebase 同步增加 rankingCaches/home，首頁排行走快取
// ║  v9-685 - 新增 Firebase 同步；題庫/設定可推送到 Firestore
// ║         學生名單雜湊改依表頭讀取，避免欄位順序造成登入失敗
// ║  v9-684 - 重複登入只踢舊視窗、新視窗保持有效；新增分析快取
// ║         班級分類/學生分類/題型/題目分析快取供後台快速讀取
// ║  v9-683 - 重複登入雙方作廢、今日練習快取分頁、錯題重做不計分
// ║         todayTotal + todayByClass（各班今日不重複練習人數） ║
// ║  v9-681 - studentHistory 加入 duration（作答秒數）
// ║  v9-68 - 踢出 + 留記錄
// ║         loginStudent：舊 token 標記「已踢出」（不是已取代）  ║
// ║         submitScore / submitAnswerDetail：交卷前驗證 token   ║
// ║         token 已踢出 → 拒絕寫入，回傳 status:"kicked"        ║
// ║  v9-67 - 每題作答秒數 + 認知類型統計（方案 B+C）
// ║         submitAnswerDetail：明細 M 欄新增作答秒數        ║
// ║         getDetailStats：回傳每題平均用時 + cogTypeStats   ║
// ║  v9-66 - 重複登入偵測（可舉證）
// ║         loginStudent：舊 token 標記「已取代」不踢出   ║
// ║         submitScore：新增 token + IP 欄位（K、L欄）   ║
// ║         新增 getDuplicateLoginReport action           ║
// ║         移除 verifySession（不再需要輪詢）            ║
// ║  v9-6 - 新增作答計時功能                                     ║
// ║         submitScore 新增 duration（作答秒數）欄位            ║
// ║         成績紀錄 Sheet 第 10 欄 = 作答秒數                   ║
// ║         getMyCompletion 回傳各分類平均每題用時               ║
// ║         getTeacherData 回傳各分類平均每題用時統計            ║
// ║  v9-5 - clearTopicCache 工具函式、checkDataStatus            ║
// ║  v9-4 - Session Token、Script Properties 排行快取            ║
// ║  v9-3 - 截止日倒數、學生成績總表自動更新                     ║
// ║  v9-2 - 後台題目分析補入新題目                               ║
// ║  v9  - GAS 語法相容修正                                      ║
// ║                                                              ║
// ║  成績紀錄 Sheet 欄位：                                        ║
// ║   A=時間戳記 B=學號 C=姓名 D=測驗單元 E=測驗模式            ║
// ║   F=第幾次  G=分數 H=答對題數 I=答錯題數 J=作答秒數（★新增）║
// ╚══════════════════════════════════════════════════════════════╝

const SHEET_QUESTIONS     = "題庫";
const SHEET_SCORES        = "成績紀錄";
const SHEET_DETAILS       = "題目作答明細";
const SHEET_STUDENTS      = "學生名單";
const SHEET_ADMINS        = "管理人名單";
const SHEET_SETTINGS      = "系統設定";
const SHEET_WRONG_IDX     = "WrongIndex";
const SHEET_TOPIC_CACHE   = "分類快取";
const SHEET_RANKING_CACHE = "排行快取";
const SHEET_TODAY_PRACTICE_CACHE = "今日練習快取";
const SHEET_CLASS_CATEGORY_ANALYSIS = "班級分類分析快取";
const SHEET_STUDENT_CATEGORY_ANALYSIS = "學生分類分析快取";
const SHEET_QUESTION_TYPE_ANALYSIS = "題型分析快取";
const SHEET_QUESTION_ANALYSIS = "題目分析快取";
const SHEET_SCORE_TABLE   = "學生成績總表";  // ★ v9-3
const SHEET_LOGIN_STATE   = "登入狀態";        // ★ v9-4

// ─────────────────────────────────────────────
// doGet：回傳分類清單＋標題
//   ★ v9：分類快取用題庫行數做版本判斷（比 getLastUpdated 更穩定）
//   ★ v8：學生名單雜湊快取（Script Properties）
// ─────────────────────────────────────────────
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_QUESTIONS);
    if (!sheet) throw new Error("找不到「" + SHEET_QUESTIONS + "」分頁");

    // 讀取標題（A1/J1）
    const firstRow   = sheet.getRange(1, 1, 1, 10).getValues()[0];
    const title      = firstRow[0] ? firstRow[0].toString().trim() : "動態題庫測驗";
    const titleColor = firstRow[9] ? firstRow[9].toString().trim() : "pink";

    // ── ★ 分類快取：用題庫目前行數做版本號 ──
    const currentRowCount = sheet.getLastRow();
    var topics = null;
    var cacheSheet = ss.getSheetByName(SHEET_TOPIC_CACHE);

    if (cacheSheet && cacheSheet.getLastRow() >= 2) {
      var cachedRowCount = cacheSheet.getRange("A1").getValue();
      if (cachedRowCount && parseInt(cachedRowCount) === currentRowCount) {
        // 快取有效：直接讀取
        var cacheRows = cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, 2).getValues();
        topics = [];
        for (var ci = 0; ci < cacheRows.length; ci++) {
          if (cacheRows[ci][0]) {
            topics.push({ name: cacheRows[ci][0].toString(), color: cacheRows[ci][1] ? cacheRows[ci][1].toString() : "red" });
          }
        }
      }
    }

    if (!topics) {
      // 快取過期或不存在：掃描題庫重建
      topics = buildTopicsAndUpdateCache(ss, sheet, currentRowCount);
    }

    // 學生名單雜湊快取
    const studentHashes = getStudentHashesCached(ss);

    // 完成度設定
    const completionSettings = readSettings(ss);

    // 班級清單
    var classArr = [];
    var stSheet  = ss.getSheetByName(SHEET_STUDENTS);
    if (stSheet && stSheet.getLastRow() > 1) {
      var stRows = stSheet.getDataRange().getValues();
      var classSet = {};
      for (var si = 1; si < stRows.length; si++) {
        var cls = stRows[si][2] ? stRows[si][2].toString().trim() : "";
        if (cls) classSet[cls] = true;
      }
      classArr = Object.keys(classSet).sort(function(a, b) { return a.localeCompare(b, "zh-TW"); });
    }

    return jsonResponse({ status: "success", title: title, titleColor: titleColor, topics: topics, studentHashes: studentHashes, completionSettings: completionSettings, allClassList: classArr, deadline: completionSettings.deadline || "" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// ─────────────────────────────────────────────
// Firebase v1.6851 同步：Google Sheet → Firestore 快取
// ─────────────────────────────────────────────
function normalizeAnswerTextV1685(rawAns, opts) {
  var raw = rawAns === null || rawAns === undefined ? "" : rawAns.toString().trim();
  var up = raw.toUpperCase();
  if (["A","B","C","D"].indexOf(up) !== -1) return opts[up.charCodeAt(0) - 65] || "";
  if (["1","2","3","4"].indexOf(up) !== -1) return opts[parseInt(up, 10) - 1] || "";
  return raw;
}

function readQuestionsForFirebaseV1685(ss) {
  var sheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return h ? h.toString().trim() : ""; });
  var iTop = findColIdx(headers, ["分類","category"]);
  var iQ = findColIdx(headers, ["題目","question","q"]);
  var iA = findColIdx(headers, ["選項A","選項1","optionA","a"]);
  var iB = findColIdx(headers, ["選項B","選項2","optionB","b"]);
  var iC = findColIdx(headers, ["選項C","選項3","optionC","c"]);
  var iD = findColIdx(headers, ["選項D","選項4","optionD","d"]);
  var iAns = findColIdx(headers, ["正確答案","答案","answer","ans"]);
  var iExp = findColIdx(headers, ["解析","explanation"]);
  var iType = findColIdx(headers, ["題型","type"]); if (iType === -1) iType = 10;
  var iImg = findColIdx(headers, ["圖片網址","圖片","imageUrl","img"]); if (iImg === -1) iImg = 11;
  var iId = findColIdx(headers, ["題目ID","ID","id"]); if (iId === -1) iId = 12;
  var iCog = findColIdx(headers, ["認知類型","cogType","認知"]); if (iCog === -1) iCog = 13;
  var version = "QB_" + sheet.getLastRow() + "_" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMddHHmmss");
  var out = [];
  var lastImgUrl = "";
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var qText = iQ !== -1 && r[iQ] ? r[iQ].toString().trim() : "";
    if (!qText) continue;
    var qType = iType !== -1 && r[iType] ? r[iType].toString().trim() : "";
    var isImg = qType === "圖片" || qType.toLowerCase() === "image";
    var imgUrl = iImg !== -1 && r[iImg] ? r[iImg].toString().trim() : "";
    if (isImg) { if (imgUrl) lastImgUrl = imgUrl; else imgUrl = lastImgUrl; }
    else lastImgUrl = "";
    var opts = [
      iA !== -1 && r[iA] ? r[iA].toString().trim() : "",
      iB !== -1 && r[iB] ? r[iB].toString().trim() : "",
      iC !== -1 && r[iC] ? r[iC].toString().trim() : "",
      iD !== -1 && r[iD] ? r[iD].toString().trim() : ""
    ].filter(Boolean);
    var qId = iId !== -1 && r[iId] ? r[iId].toString().trim() : "ROW_" + (i + 1);
    out.push({
      id: qId,
      top: iTop !== -1 && r[iTop] ? r[iTop].toString().trim() : "未分類",
      q: qText,
      options: opts,
      ans: normalizeAnswerTextV1685(iAns !== -1 ? r[iAns] : "", opts),
      exp: iExp !== -1 && r[iExp] ? r[iExp].toString().trim() : "尚無解析",
      color: r[9] ? r[9].toString().trim() : "red",
      questionType: qType,
      imgUrl: imgUrl,
      isImage: !!imgUrl || isImg,
      cogType: iCog !== -1 && r[iCog] ? r[iCog].toString().trim() : "",
      questionBankVersion: version,
      updatedAtText: localNow()
    });
  }
  return out;
}

function readStudentsForFirebaseV1685(ss) {
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0] || [];
  var idCol = findColIdx(headers, ["學號", "studentId", "student_id", "id"]); if (idCol === -1) idCol = 0;
  var nameCol = findColIdx(headers, ["姓名", "學生姓名", "name"]); if (nameCol === -1) nameCol = 1;
  var classCol = findColIdx(headers, ["班級", "修課班級", "class", "className"]); if (classCol === -1) classCol = 2;
  var emailCol = findColIdx(headers, ["Email", "email", "E-mail", "e-mail", "電子郵件", "信箱", "學校email", "學校Email", "Google帳號", "Google信箱"]);
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var sid = rows[i][idCol] ? rows[i][idCol].toString().trim() : "";
    if (!sid) continue;
    var email = emailCol !== -1 && rows[i][emailCol] ? rows[i][emailCol].toString().trim() : "";
    out.push({
      studentId: sid,
      name: rows[i][nameCol] ? rows[i][nameCol].toString().trim() : sid,
      className: rows[i][classCol] ? rows[i][classCol].toString().trim() : "未分班",
      email: email,
      emailLower: email.toLowerCase(),
      updatedAtText: localNow()
    });
  }
  return out;
}

function buildFirebasePayloadV1685() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!sheet) throw new Error("找不到「" + SHEET_QUESTIONS + "」分頁");
  var firstRow = sheet.getRange(1, 1, 1, 10).getValues()[0];
  var title = firstRow[0] ? firstRow[0].toString().trim() : "動態題庫測驗";
  var titleColor = firstRow[9] ? firstRow[9].toString().trim() : "pink";
  var topics = buildTopicsAndUpdateCache(ss, sheet, sheet.getLastRow());
  var questions = readQuestionsForFirebaseV1685(ss);
  var students = readStudentsForFirebaseV1685(ss);
  var settings = readSettings(ss);
  var rankingCache = buildRankingCacheForFirebaseV1685(ss);
  var classMap = {};
  students.forEach(function(s) { if (s.className) classMap[s.className] = true; });
  var allClassList = Object.keys(classMap).sort(function(a, b) { return a.localeCompare(b, "zh-TW"); });
  return {
    generatedAt: localNow(),
    settings: {
      title: title,
      titleColor: titleColor,
      topics: topics,
      studentHashes: getStudentHashesCached(ss),
      completionSettings: settings,
      allClassList: allClassList,
      deadline: settings.deadline || "",
      questionBankVersion: questions.length ? questions[0].questionBankVersion : "",
      updatedAtText: localNow()
    },
    rankingCache: rankingCache,
    questions: questions,
    students: students,
    counts: { questions: questions.length, students: students.length, topics: topics.length, classes: allClassList.length }
  };
}

function buildRankingCacheForFirebaseV1685(ss) {
  var cached = getRankingCacheProps(ss);
  if (!cached) return null;
  var todayCache = readTodayPracticeCache(ss);
  cached.todayTotal = todayCache.todayTotal;
  cached.todayByClass = todayCache.todayByClass;
  cached.todayDate = todayCache.todayDate;
  cached.todayUpdatedAt = todayCache.updatedAt;
  cached.updatedAtText = localNow();
  return cached;
}

function handleGetFirebaseBootstrap(payload) {
  return jsonResponse({ status: "ok", data: buildFirebasePayloadV1685() });
}

function firebaseSafeDocIdV1685(raw) {
  return encodeURIComponent((raw || "doc").toString()).replace(/\./g, "%2E");
}

function firebaseValueV1685(v) {
  if (v === null || v === undefined) return { nullValue: null };
  if (typeof v === "boolean") return { booleanValue: v };
  if (typeof v === "number") return Math.floor(v) === v ? { integerValue: String(v) } : { doubleValue: v };
  if (Array.isArray(v)) return { arrayValue: { values: v.map(firebaseValueV1685) } };
  if (typeof v === "object") {
    var fields = {};
    Object.keys(v).forEach(function(k) { fields[k] = firebaseValueV1685(v[k]); });
    return { mapValue: { fields: fields } };
  }
  return { stringValue: String(v) };
}

function firebaseFieldsV1685(obj) {
  var fields = {};
  Object.keys(obj || {}).forEach(function(k) { fields[k] = firebaseValueV1685(obj[k]); });
  return fields;
}

function firebaseJwtBase64V1685(objOrBytes) {
  var bytes = Array.isArray(objOrBytes) ? objOrBytes : Utilities.newBlob(JSON.stringify(objOrBytes)).getBytes();
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, "");
}

function firebaseAccessTokenV1685() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty("FIREBASE_CLIENT_EMAIL") || props.getProperty("FIREBASE_SERVICE_ACCOUNT_EMAIL");
  var key = props.getProperty("FIREBASE_PRIVATE_KEY");
  if (!email || !key) throw new Error("尚未設定 FIREBASE_CLIENT_EMAIL / FIREBASE_PRIVATE_KEY");
  key = key.replace(/\\n/g, "\n");
  var now = Math.floor(Date.now() / 1000);
  var unsigned = firebaseJwtBase64V1685({ alg: "RS256", typ: "JWT" }) + "." + firebaseJwtBase64V1685({
    iss: email,
    scope: "https://www.googleapis.com/auth/datastore",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now
  });
  var jwt = unsigned + "." + firebaseJwtBase64V1685(Utilities.computeRsaSha256Signature(unsigned, key));
  var res = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt },
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (!data.access_token) throw new Error("Firebase token 取得失敗：" + res.getContentText());
  return data.access_token;
}

function firestoreDocNameV1685(projectId, collection, id) {
  return "projects/" + projectId + "/databases/(default)/documents/" + collection + "/" + firebaseSafeDocIdV1685(id);
}

function firebaseBatchWriteV1685(projectId, token, writes) {
  if (!writes.length) return;
  var url = "https://firestore.googleapis.com/v1/projects/" + projectId + "/databases/(default)/documents:batchWrite";
  for (var i = 0; i < writes.length; i += 500) {
    var res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({ writes: writes.slice(i, i + 500) }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) throw new Error("Firestore 寫入失敗：" + res.getContentText());
  }
}

function handleSyncFirebaseV1685(payload) {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    return jsonResponse({ status: "needs_config", message: "尚未設定 FIREBASE_PROJECT_ID", data: buildFirebasePayloadV1685() });
  }
  var data = buildFirebasePayloadV1685();
  var token = firebaseAccessTokenV1685();
  var writes = [];
  writes.push({ update: { name: firestoreDocNameV1685(projectId, "system", "main"), fields: firebaseFieldsV1685(data.settings) } });
  if (data.rankingCache) {
    writes.push({ update: { name: firestoreDocNameV1685(projectId, "rankingCaches", "home"), fields: firebaseFieldsV1685(data.rankingCache) } });
  }
  data.questions.forEach(function(q) {
    writes.push({ update: { name: firestoreDocNameV1685(projectId, "questions", q.id), fields: firebaseFieldsV1685(q) } });
  });
  data.students.forEach(function(s) {
    writes.push({ update: { name: firestoreDocNameV1685(projectId, "students", s.studentId), fields: firebaseFieldsV1685(s) } });
  });
  firebaseBatchWriteV1685(projectId, token, writes);
  return jsonResponse({ status: "ok", message: "Firebase 同步完成", counts: data.counts, written: writes.length, generatedAt: data.generatedAt });
}

// ── 掃描題庫建立分類清單，並更新快取 Sheet ──
function buildTopicsAndUpdateCache(ss, sheet, currentRowCount) {
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return h.toString().trim(); });

  var iTop = -1;
  var topNames = ["分類","category"];
  for (var ni = 0; ni < topNames.length; ni++) {
    var idx = headers.indexOf(topNames[ni]);
    if (idx === -1) {
      for (var hi = 0; hi < headers.length; hi++) {
        if (headers[hi].toLowerCase() === topNames[ni].toLowerCase()) { idx = hi; break; }
      }
    }
    if (idx !== -1) { iTop = idx; break; }
  }
  var COLOR_COL = 9;

  var topicMap = {};
  var topicOrder = [];
  for (var i = 1; i < rows.length; i++) {
    var top   = iTop !== -1 ? (rows[i][iTop] ? rows[i][iTop].toString().trim() : "未分類") : "未分類";
    var color = rows[i][COLOR_COL] ? rows[i][COLOR_COL].toString().trim() : "red";
    if (top && !topicMap[top]) {
      topicMap[top] = color;
      topicOrder.push(top);
    }
  }

  var topics = [];
  for (var ti = 0; ti < topicOrder.length; ti++) {
    topics.push({ name: topicOrder[ti], color: topicMap[topicOrder[ti]] });
  }

  // 更新快取 Sheet
  var cacheSheet = ss.getSheetByName(SHEET_TOPIC_CACHE);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(SHEET_TOPIC_CACHE);
    cacheSheet.getRange("A1").setFontWeight("bold").setBackground("#fef9c3");
  }
  cacheSheet.clearContents();
  cacheSheet.getRange("A1").setValue(currentRowCount); // ★ 用行數當版本號
  if (topics.length > 0) {
    var cacheData = topics.map(function(t) { return [t.name, t.color]; });
    cacheSheet.getRange(2, 1, cacheData.length, 2).setValues(cacheData);
  }

  return topics;
}

// ── 學生名單雜湊快取（Script Properties）──
function getStudentHashesCached(ss) {
  var props  = PropertiesService.getScriptProperties();
  var sSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sSheet || sSheet.getLastRow() < 2) return [];

  var rows = sSheet.getDataRange().getValues();
  var headers = rows[0] || [];
  var idCol = findColIdx(headers, ["學號", "studentId", "student_id", "id"]);
  var nameCol = findColIdx(headers, ["姓名", "學生姓名", "name"]);
  if (idCol === -1) idCol = 0;
  if (nameCol === -1) nameCol = 1;

  var lastRow = sSheet.getLastRow();
  var lastSid = rows[lastRow - 1] && rows[lastRow - 1][idCol] ? rows[lastRow - 1][idCol].toString() : "";
  var cacheKey  = "HASH_VER_" + lastRow + "_" + lastSid;
  var cachedVer = props.getProperty("STUDENT_HASH_VER");
  var cachedVal = props.getProperty("STUDENT_HASHES");

  if (cachedVer === cacheKey && cachedVal) {
    try { return JSON.parse(cachedVal); } catch(e) {}
  }

  var hashes = [];
  for (var i = 1; i < rows.length; i++) {
    var sid  = rows[i][idCol] ? rows[i][idCol].toString().trim() : "";
    var name = rows[i][nameCol] ? rows[i][nameCol].toString().trim() : "";
    if (sid && name) hashes.push(hashString(sid + "|" + name));
  }

  props.setProperty("STUDENT_HASH_VER", cacheKey);
  props.setProperty("STUDENT_HASHES", JSON.stringify(hashes));
  return hashes;
}

// ─────────────────────────────────────────────
// doPost
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    if (action === "verifyStudent")      return handleVerifyStudent(payload);
    if (action === "adminLogin")         return handleAdminLogin(payload);
    if (action === "getTopicQuestions")  return handleGetTopicQuestions(payload);
    if (action === "batchJudge")         return handleBatchJudge(payload);
    if (action === "submitScore")        return handleSubmitScore(payload);
    if (action === "submitAnswerDetail") return handleSubmitDetail(payload);
    if (action === "saveSettings")       return handleSaveSettings(payload);
    if (action === "getMyCompletion")      return handleGetMyCompletion(payload);
    if (action === "getMyWrongQuestions")  return handleGetMyWrongQuestions(payload);
    if (action === "getCompletionRanking") return handleGetCompletionRanking(payload);
    if (action === "getTeacherData")        return handleGetTeacherData(payload);
    if (action === "getStudentScoreTable")  return handleGetStudentScoreTable(payload);
    if (action === "getDetailStats")         return handleGetDetailStats(payload);
    if (action === "getAnalysisCache")       return handleGetAnalysisCache(payload);
    if (action === "loginStudent")           return handleLoginStudent(payload);
    if (action === "verifySession")          return handleVerifySession(payload);
    if (action === "getClassStudents")       return handleGetClassStudents(payload);
    if (action === "getDuplicateLoginReport") return handleGetDuplicateLoginReport(payload);
    if (action === "getFirebaseBootstrap")   return handleGetFirebaseBootstrap(payload);
    if (action === "syncFirebaseV1685")      return handleSyncFirebaseV1685(payload);
    return jsonResponse({ status: "error", message: "未知的 action：" + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// ─────────────────────────────────────────────
// Action 0：getTopicQuestions
// ─────────────────────────────────────────────
function handleGetTopicQuestions(payload) {
  var topic     = payload.topic;
  var flashcard = payload.flashcard;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!sheet) return jsonResponse({ status: "error", message: "找不到題庫" });

  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return h.toString().trim(); });

  function ci(names) {
    for (var ni = 0; ni < names.length; ni++) {
      for (var hi = 0; hi < headers.length; hi++) {
        if (headers[hi] === names[ni] || headers[hi].toLowerCase() === names[ni].toLowerCase()) return hi;
      }
    }
    return -1;
  }

  var iId  = ci(["題目ID","ID","id"]);
  var iTop = ci(["分類","category"]);
  var iQ   = ci(["題目","question","q"]);
  var iA   = ci(["選項A","選項1","optionA","a"]);
  var iB   = ci(["選項B","選項2","optionB","b"]);
  var iC   = ci(["選項C","選項3","optionC","c"]);
  var iD   = ci(["選項D","選項4","optionD","d"]);
  var iAns = ci(["正確答案","答案","answer","ans"]);
  var iExp = ci(["解析","explanation"]);
  var COLOR_COL = 9;
  var TYPE_COL  = ci(["題型","type"]);
  if (TYPE_COL === -1) TYPE_COL = 10;
  var IMG_COL   = ci(["圖片網址","圖片","imageUrl","img"]);
  if (IMG_COL === -1) IMG_COL = 11;
  var COG_COL   = ci(["認知類型","cogType","認知"]);
  if (COG_COL === -1) COG_COL = 13;

  var rawData = [];
  var lastImgUrl = "";
  for (var i = 1; i < rows.length; i++) {
    var r   = rows[i];
    var q   = iQ !== -1 ? (r[iQ] ? r[iQ].toString().trim() : "") : "";
    if (!q) continue;
    var qTop = iTop !== -1 ? (r[iTop] ? r[iTop].toString().trim() : "未分類") : "未分類";
    if (topic !== "綜合練習" && qTop !== topic) continue;

    var qId   = (iId !== -1 && r[iId]) ? r[iId].toString().trim() : "ROW_" + (i + 1);
    var qType = r[TYPE_COL] ? r[TYPE_COL].toString().trim() : "";
    var isImg = qType === "圖片" || qType.toLowerCase() === "image";

    var imgUrl = r[IMG_COL] ? r[IMG_COL].toString().trim() : "";
    if (isImg) {
      if (imgUrl) lastImgUrl = imgUrl;
      else imgUrl = lastImgUrl;
    } else { lastImgUrl = ""; }

    var optA = iA !== -1 ? (r[iA] ? r[iA].toString().trim() : "") : "";
    var optB = iB !== -1 ? (r[iB] ? r[iB].toString().trim() : "") : "";
    var optC = iC !== -1 ? (r[iC] ? r[iC].toString().trim() : "") : "";
    var optD = iD !== -1 ? (r[iD] ? r[iD].toString().trim() : "") : "";
    var allOpts = [optA, optB, optC, optD];

    var rawAnsVal = iAns !== -1 ? (r[iAns] ? r[iAns].toString().trim() : "") : "";
    var ans = "";
    if (isImg) {
      var up = rawAnsVal.toUpperCase();
      if (["A","B","C","D"].indexOf(up) !== -1) ans = allOpts[up.charCodeAt(0) - 65] || optA;
      else if (["1","2","3","4"].indexOf(up) !== -1) ans = allOpts[parseInt(up) - 1] || optA;
      else ans = optA;
    } else {
      var up2 = rawAnsVal.toUpperCase();
      if (["A","B","C","D"].indexOf(up2) !== -1) ans = allOpts[up2.charCodeAt(0) - 65] || "";
      else if (["1","2","3","4"].indexOf(up2) !== -1) ans = allOpts[parseInt(up2) - 1] || "";
      else {
        var clean = rawAnsVal.replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase();
        for (var oi = 0; oi < allOpts.length; oi++) {
          if (allOpts[oi].replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase() === clean) { ans = allOpts[oi]; break; }
        }
        if (!ans) ans = rawAnsVal;
      }
    }

    var exp = iExp !== -1 ? (r[iExp] ? r[iExp].toString().trim() : "尚無解析") : "尚無解析";
    rawData.push({
      id: qId, top: qTop, q: q, ans: ans,
      rawOpts: isImg ? [] : allOpts.filter(function(o) { return o; }),
      allOpts: allOpts, exp: exp,
      color: r[COLOR_COL] ? r[COLOR_COL].toString().trim() : "red",
      isImage: isImg, imgUrl: imgUrl,
      questionType: qType,
      cogType: r[COG_COL] ? r[COG_COL].toString().trim() : "",
    });
  }

  // 圖片題組裝
  var imgGroupMap = {};
  rawData.forEach(function(q) {
    if (q.isImage && q.imgUrl) {
      if (!imgGroupMap[q.imgUrl]) imgGroupMap[q.imgUrl] = [];
      imgGroupMap[q.imgUrl].push({ id: q.id, ans: q.ans });
    }
  });
  var allImageAnsSet = {};
  rawData.forEach(function(q) { if (q.isImage && q.ans) allImageAnsSet[q.ans] = true; });
  var allImageAnsPool = Object.keys(allImageAnsSet);

  var data = rawData.map(function(q) {
    var options = q.rawOpts;
    if (q.isImage && q.imgUrl) {
      var selfAns = q.ans;
      var grp = imgGroupMap[q.imgUrl] || [];
      var sameGroupAns = [];
      grp.forEach(function(g) { if (g.id !== q.id && g.ans) sameGroupAns.push(g.ans); });
      var usedSet = {};
      usedSet[selfAns] = true;
      sameGroupAns.forEach(function(a) { usedSet[a] = true; });
      var fallback = allImageAnsPool.filter(function(a) { return !usedSet[a]; });
      var candidates = sameGroupAns.concat(fallback);
      for (var ci2 = candidates.length - 1; ci2 > 0; ci2--) {
        var j = Math.floor(Math.random() * (ci2 + 1));
        var tmp = candidates[ci2]; candidates[ci2] = candidates[j]; candidates[j] = tmp;
      }
      options = [selfAns].concat(candidates.slice(0, 3)).filter(Boolean);
    }
    var item = { id: q.id, top: q.top, q: q.q, options: options, color: q.color, isImage: q.isImage, questionType: q.questionType || "", cogType: q.cogType || "" };
    if (q.isImage && q.imgUrl) item.imgUrl = q.imgUrl;
    if (flashcard) { item.ans = q.ans; item.exp = q.exp; }
    return item;
  });

  return jsonResponse({ status: "ok", data: data });
}

// ─────────────────────────────────────────────
// Action 1：verifyStudent
// ─────────────────────────────────────────────
function handleVerifyStudent(payload) {
  var studentId = payload.studentId;
  var name = payload.name;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet || sheet.getLastRow() < 2)
    return jsonResponse({ status: "ok", verified: true });
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].toString().trim() === studentId.trim() &&
        rows[i][1] && rows[i][1].toString().trim() === name.trim())
      return jsonResponse({ status: "ok", verified: true });
  }
  return jsonResponse({ status: "ok", verified: false, message: "學號或姓名不符" });
}

// ─────────────────────────────────────────────
// Action 2：adminLogin
// ─────────────────────────────────────────────
function handleAdminLogin(payload) {
  var adminId = payload.adminId;
  var adminPassword = payload.adminPassword;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ADMINS);
  if (!sheet || sheet.getLastRow() < 2)
    return jsonResponse({ status: "error", message: "尚未建立「管理人名單」分頁" });
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].toString().trim() === adminId.trim() &&
        rows[i][1] && rows[i][1].toString().trim() === adminPassword.trim())
      return jsonResponse({ status: "ok", verified: true, adminName: adminId });
  }
  return jsonResponse({ status: "ok", verified: false, message: "帳號或密碼錯誤" });
}

// ─────────────────────────────────────────────
// Action 3：batchJudge
// ─────────────────────────────────────────────
function handleBatchJudge(payload) {
  var answers = payload.answers || [];
  if (answers.length === 0) return jsonResponse({ status: "ok", results: [] });

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEET_QUESTIONS);
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return h.toString().trim(); });

  var iId  = findColIdx(headers, ["題目ID","ID","id"]);
  var iAns = findColIdx(headers, ["正確答案","答案","answer","ans"]);
  var iExp = findColIdx(headers, ["解析","explanation"]);
  var iA   = findColIdx(headers, ["選項A","選項1","optionA","a"]);
  var iB   = findColIdx(headers, ["選項B","選項2","optionB","b"]);
  var iC   = findColIdx(headers, ["選項C","選項3","optionC","c"]);
  var iD   = findColIdx(headers, ["選項D","選項4","optionD","d"]);
  var TYPE_COL = findColIdx(headers, ["題型","type"]);
  if (TYPE_COL === -1) TYPE_COL = 10;

  var questionMap = {};
  for (var i = 1; i < rows.length; i++) {
    var r      = rows[i];
    var qId    = (iId !== -1 && r[iId]) ? r[iId].toString().trim() : "ROW_" + (i + 1);
    var rawAns = iAns !== -1 ? (r[iAns] ? r[iAns].toString().trim() : "") : "";
    var exp    = iExp !== -1 ? (r[iExp] ? r[iExp].toString().trim() : "尚無解析") : "尚無解析";
    var qType  = r[TYPE_COL] ? r[TYPE_COL].toString().trim() : "";
    var isImage = qType === "圖片" || qType.toLowerCase() === "image";

    var allOptions = [];
    [iA, iB, iC, iD].forEach(function(idx) {
      if (idx !== -1 && r[idx]) allOptions.push(r[idx].toString().trim());
    });

    var correctText = "";
    if (isImage) {
      var upImg = rawAns.toUpperCase();
      if (["A","B","C","D"].indexOf(upImg) !== -1) correctText = allOptions[upImg.charCodeAt(0) - 65] || "";
      else if (["1","2","3","4"].indexOf(upImg) !== -1) correctText = allOptions[parseInt(upImg) - 1] || "";
      else correctText = allOptions[0] || rawAns;
    } else {
      var upperAns = rawAns.toUpperCase();
      if (["A","B","C","D"].indexOf(upperAns) !== -1) {
        correctText = allOptions[upperAns.charCodeAt(0) - 65] || "";
      } else if (["1","2","3","4"].indexOf(upperAns) !== -1) {
        correctText = allOptions[parseInt(upperAns) - 1] || "";
      } else {
        var cleanAns = rawAns.replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase();
        for (var oi = 0; oi < allOptions.length; oi++) {
          if (allOptions[oi].replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase() === cleanAns) {
            correctText = allOptions[oi]; break;
          }
        }
        if (!correctText) correctText = rawAns;
      }
    }
    questionMap[qId] = { correctText: correctText, exp: exp };
  }

  var results = answers.map(function(ans) {
    var questionId  = ans.questionId;
    var selectedText = ans.selectedText;
    var shuffledOpts = ans.shuffledOpts;
    var q = questionMap[questionId];
    if (!q) return { questionId: questionId, correct: false, correctText: "", correctIndex: -1, exp: "找不到此題目" };

    var cleanSel  = (selectedText || "").replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase().replace(/\s+/g," ");
    var cleanCorr = q.correctText.replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase().replace(/\s+/g," ");
    var correct   = cleanSel === cleanCorr;

    var correctIndex = -1;
    if (shuffledOpts && Array.isArray(shuffledOpts)) {
      for (var si = 0; si < shuffledOpts.length; si++) {
        var cleanOpt = shuffledOpts[si].replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase().replace(/\s+/g," ");
        if (cleanOpt === cleanCorr) { correctIndex = si; break; }
      }
    }
    return { questionId: questionId, correct: correct, correctText: q.correctText, correctIndex: correctIndex, exp: q.exp };
  });

  return jsonResponse({ status: "ok", results: results });
}

// ─────────────────────────────────────────────
// Action 4：submitScore
// ─────────────────────────────────────────────
function handleSubmitScore(payload) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  // ★ v9-68 交卷前驗證 token
  if (payload.token && !isTokenValid(ss, payload.studentId, payload.token)) {
    return jsonResponse({ status: "kicked", message: "您的帳號已在其他裝置登入，本次成績未計入" });
  }
  if (payload.mode === "錯題重做") {
    return jsonResponse({ status: "ok", skipped: true, message: "錯題重做不計入成績紀錄" });
  }
  var sheet = ss.getSheetByName(SHEET_SCORES);
  if (!sheet) sheet = ss.insertSheet(SHEET_SCORES);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["時間戳記","學號","姓名","測驗單元","測驗模式","第幾次","分數","答對題數","答錯題數","作答秒數","Token","IP"]);
    sheet.getRange(1,1,1,12).setFontWeight("bold").setBackground("#fce7f3");
  }
  // ★ v9-7 新增 token + IP（重複登入舉證用）
  var duration = (payload.duration && payload.duration > 0) ? payload.duration : "";
  var token2   = payload.token || "";
  var ip2      = payload.ip    || "";
  sheet.appendRow([localNow(), payload.studentId, payload.name, payload.topic, payload.mode, payload.attempt, payload.score, payload.correctCount, payload.wrongCount, duration, token2, ip2]);
  // ★ v9-684 方案A：不在交卷時讓快取失效，排行由 autoUpdateScoreSheet 每小時更新
  return jsonResponse({ status: "ok" });
}

// ─────────────────────────────────────────────
// Action 5：submitAnswerDetail（含自動換頁）
// ─────────────────────────────────────────────
const MAX_DETAIL_ROWS = 20000;
const DETAIL_HEADER   = ["時間戳記","學號","姓名","題目ID","題目內容","單元","學生選項","正確答案","是否答對","測驗模式","第幾次","作答秒數","題型","認知類型"];

function getActiveDetailSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_DETAILS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DETAILS);
    sheet.appendRow(DETAIL_HEADER);
    sheet.getRange(1,1,1,DETAIL_HEADER.length).setFontWeight("bold").setBackground("#e0f2fe");
    return sheet;
  }
  ensureDetailHeader(sheet);
  if (sheet.getLastRow() - 1 >= MAX_DETAIL_ROWS) {
    var now     = new Date(new Date().getTime() + 8 * 60 * 60 * 1000);
    var yyyymm  = now.getUTCFullYear() + "-" + String(now.getUTCMonth()+1).padStart(2,"0");
    var archiveName = SHEET_DETAILS + "_" + yyyymm;
    var suffix = 1;
    while (ss.getSheetByName(archiveName)) archiveName = SHEET_DETAILS + "_" + yyyymm + "_" + (suffix++);
    sheet.setName(archiveName);
    sheet.setFrozenRows(1);
    sheet = ss.insertSheet(SHEET_DETAILS);
    sheet.appendRow(DETAIL_HEADER);
    sheet.getRange(1,1,1,DETAIL_HEADER.length).setFontWeight("bold").setBackground("#e0f2fe");
  }
  return sheet;
}

function ensureDetailHeader(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), DETAIL_HEADER.length)).getValues()[0];
  var changed = false;
  for (var i = 0; i < DETAIL_HEADER.length; i++) {
    if (!current[i] || current[i].toString().trim() !== DETAIL_HEADER[i]) {
      current[i] = DETAIL_HEADER[i];
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(1, 1, 1, DETAIL_HEADER.length).setValues([DETAIL_HEADER]);
    sheet.getRange(1,1,1,DETAIL_HEADER.length).setFontWeight("bold").setBackground("#e0f2fe");
  }
}

function handleSubmitDetail(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  // ★ v9-68 交卷前驗證 token
  if (payload.token && !isTokenValid(ss, payload.studentId, payload.token)) {
    return jsonResponse({ status: "kicked", message: "已在其他裝置登入，明細未記錄" });
  }
  var sheet = getActiveDetailSheet(ss);
  var rows  = (payload.details || []).map(function(d) {
    // ★ v9-67 M 欄 = 作答秒數（answerSec）
    var sec = (d.answerSec !== undefined && d.answerSec !== null) ? d.answerSec : "";
    return [localNow(), payload.studentId, payload.name, d.questionId, d.questionText, d.topic, d.selectedText, d.correctText, d.isCorrect ? "答對" : "答錯", payload.mode, payload.attempt, sec, d.questionType || "", d.cogType || ""];
  });
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, DETAIL_HEADER.length).setValues(rows);
  updateWrongIndex(ss, payload.studentId, payload.details || []);
  return jsonResponse({ status: "ok" });
}

// ─────────────────────────────────────────────
// WrongIndex 維護
// ─────────────────────────────────────────────
function getOrCreateWrongIndexSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_WRONG_IDX);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WRONG_IDX);
    sheet.appendRow(["學號","題目ID","題目分類","最新結果","最後作答時間"]);
    sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#fef3c7");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function updateWrongIndex(ss, studentId, details) {
  if (!details.length) return;
  var sheet   = getOrCreateWrongIndexSheet(ss);
  var lastRow = sheet.getLastRow();
  var keyToRow = {};
  if (lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var key = data[i][0] + "|" + data[i][1];
      keyToRow[key] = i + 2;
    }
  }
  var now     = localNow();
  var updates = [];
  var appends = [];
  details.forEach(function(d) {
    var key    = studentId + "|" + d.questionId;
    var result = d.isCorrect ? "答對" : "答錯";
    var vals   = [studentId, d.questionId, d.topic, result, now];
    if (keyToRow[key]) updates.push({ row: keyToRow[key], vals: vals });
    else               appends.push(vals);
  });
  updates.forEach(function(u) { sheet.getRange(u.row, 1, 1, 5).setValues([u.vals]); });
  if (appends.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 5).setValues(appends);
}

// ─────────────────────────────────────────────
// 【工具函式】重建 WrongIndex
// ─────────────────────────────────────────────
function rebuildWrongIndex() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets  = ss.getSheets();
  var detailSheets = allSheets.filter(function(s) {
    return s.getName() === SHEET_DETAILS || s.getName().indexOf(SHEET_DETAILS + "_") === 0;
  }).sort(function(a,b) { return a.getName().localeCompare(b.getName()); });

  var indexMap = {};
  for (var di = 0; di < detailSheets.length; di++) {
    var dSheet = detailSheets[di];
    if (dSheet.getLastRow() <= 1) continue;
    var rows = dSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var ts = rows[i][0], sid = rows[i][1], qid = rows[i][3], topic = rows[i][5], result = rows[i][8];
      if (!sid || !qid) continue;
      var key = sid.toString() + "|" + qid.toString();
      indexMap[key] = { sid: sid.toString(), qid: qid.toString(), topic: topic ? topic.toString() : "", result: result ? result.toString() : "", time: ts ? ts.toString() : "" };
    }
  }

  var sheet = ss.getSheetByName(SHEET_WRONG_IDX);
  if (sheet) ss.deleteSheet(sheet);
  sheet = getOrCreateWrongIndexSheet(ss);

  var vals = Object.keys(indexMap).map(function(k) {
    var v = indexMap[k];
    return [v.sid, v.qid, v.topic, v.result, v.time];
  });
  if (vals.length > 0) sheet.getRange(2, 1, vals.length, 5).setValues(vals);
  Logger.log("✅ WrongIndex 重建完成，共 " + vals.length + " 筆");
}

// ─────────────────────────────────────────────
// 【工具函式】手動封存超量分頁
// ─────────────────────────────────────────────
function manualArchiveDetailSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DETAILS);
  if (!sheet) { Logger.log("找不到「題目作答明細」分頁"); return; }
  var rowCount = sheet.getLastRow() - 1;
  Logger.log("目前列數：" + rowCount);
  if (rowCount <= MAX_DETAIL_ROWS) { Logger.log("列數未超過上限，不需封存"); return; }
  var now  = new Date(new Date().getTime() + 8 * 60 * 60 * 1000);
  var yyyymm = now.getUTCFullYear() + "-" + String(now.getUTCMonth()+1).padStart(2,"0");
  var archiveName = SHEET_DETAILS + "_" + yyyymm;
  var suffix = 1;
  while (ss.getSheetByName(archiveName)) archiveName = SHEET_DETAILS + "_" + yyyymm + "_" + (suffix++);
  sheet.setName(archiveName);
  sheet.setFrozenRows(1);
  Logger.log("✅ 已封存為「" + archiveName + "」");
  var newSheet = ss.insertSheet(SHEET_DETAILS);
  newSheet.appendRow(DETAIL_HEADER);
  newSheet.getRange(1,1,1,DETAIL_HEADER.length).setFontWeight("bold").setBackground("#e0f2fe");
  Logger.log("✅ 已建立新的「" + SHEET_DETAILS + "」分頁");
}

// ─────────────────────────────────────────────
// Action：saveSettings
// ─────────────────────────────────────────────
function handleSaveSettings(payload) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var pass    = parseInt(payload.passScore || "80");
  var topics  = payload.completionTopics  || [];
  var classes = payload.completionClasses || [];
  writeSettings(ss, pass, topics, classes);
  invalidateRankingCache(ss);
  return jsonResponse({ status: "ok" });
}

// ─────────────────────────────────────────────
// Action：getMyWrongQuestions（使用 WrongIndex）
// ─────────────────────────────────────────────
function handleGetMyWrongQuestions(payload) {
  var studentId = payload.studentId;
  var topic     = payload.topic;
  var hours     = payload.hours;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cutoff = (hours && hours > 0) ? new Date(new Date().getTime() - hours * 60 * 60 * 1000) : null;

  var wiSheet = ss.getSheetByName(SHEET_WRONG_IDX);
  if (!wiSheet || wiSheet.getLastRow() <= 1) return jsonResponse({ status: "ok", questions: [] });

  var rows = wiSheet.getDataRange().getValues();
  var wrongQids = [];
  for (var i = 1; i < rows.length; i++) {
    var sid    = rows[i][0] ? rows[i][0].toString() : "";
    var qid    = rows[i][1] ? rows[i][1].toString() : "";
    var qtopic = rows[i][2] ? rows[i][2].toString() : "";
    var result = rows[i][3] ? rows[i][3].toString() : "";
    var ts     = rows[i][4];
    if (sid !== studentId) continue;
    if (result !== "答錯") continue;
    if (topic !== "綜合練習" && qtopic !== topic) continue;
    if (cutoff && ts) {
      var rowDate = ts instanceof Date ? ts : new Date(ts);
      if (!isNaN(rowDate) && rowDate < cutoff) continue;
    }
    wrongQids.push(qid);
  }

  if (!wrongQids.length) return jsonResponse({ status: "ok", questions: [] });
  var wrongSet = {};
  wrongQids.forEach(function(q) { wrongSet[q] = true; });

  var qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!qSheet) return jsonResponse({ status: "ok", questions: [] });

  var qRows    = qSheet.getDataRange().getValues();
  var qHeaders = qRows[0].map(function(h) { return h.toString().trim(); });

  function ci2(names) {
    for (var ni = 0; ni < names.length; ni++) {
      for (var hi = 0; hi < qHeaders.length; hi++) {
        if (qHeaders[hi] === names[ni] || qHeaders[hi].toLowerCase() === names[ni].toLowerCase()) return hi;
      }
    }
    return -1;
  }

  var iId2  = ci2(["題目ID","ID","id"]);
  var iTop2 = ci2(["分類","category"]);
  var iQ2   = ci2(["題目","question","q"]);
  var iA2   = ci2(["選項A","選項1","optionA","a"]);
  var iB2   = ci2(["選項B","選項2","optionB","b"]);
  var iC2   = ci2(["選項C","選項3","optionC","c"]);
  var iD2   = ci2(["選項D","選項4","optionD","d"]);
  var iAns2 = ci2(["正確答案","答案","answer","ans"]);
  var iExp2 = ci2(["解析","explanation"]);
  var COLOR2 = 9;
  var TYPE2 = ci2(["題型","type"]); if (TYPE2 === -1) TYPE2 = 10;
  var IMG2 = ci2(["圖片網址","圖片","imageUrl","img"]); if (IMG2 === -1) IMG2 = 11;
  var COG2 = ci2(["認知類型","cogType","認知"]); if (COG2 === -1) COG2 = 13;

  // 建立 imgUrlMap
  var imgUrlMap = {};
  var lastImgUrl2 = "";
  for (var i2 = 1; i2 < qRows.length; i2++) {
    var r2   = qRows[i2];
    var qid2 = (iId2 !== -1 && r2[iId2]) ? r2[iId2].toString().trim() : "ROW_" + (i2 + 1);
    var qType2 = r2[TYPE2] ? r2[TYPE2].toString().trim() : "";
    var isImg2 = qType2 === "圖片" || qType2.toLowerCase() === "image";
    var imgUrl2 = r2[IMG2] ? r2[IMG2].toString().trim() : "";
    if (isImg2) {
      if (imgUrl2) lastImgUrl2 = imgUrl2;
      else imgUrl2 = lastImgUrl2;
    } else { lastImgUrl2 = ""; }
    imgUrlMap[qid2] = { isImg: isImg2, imgUrl: imgUrl2 };
  }

  // 圖片題組資料
  var imgGroupMap2 = {};
  var allImgAnsPool2 = [];
  for (var i3 = 1; i3 < qRows.length; i3++) {
    var r3   = qRows[i3];
    var qid3 = (iId2 !== -1 && r3[iId2]) ? r3[iId2].toString().trim() : "ROW_" + (i3 + 1);
    var info3 = imgUrlMap[qid3] || {};
    if (!info3.isImg || !info3.imgUrl) continue;
    var rawAns3 = iAns2 !== -1 ? (r3[iAns2] ? r3[iAns2].toString().trim() : "") : "";
    var oA3 = iA2 !== -1 ? (r3[iA2] ? r3[iA2].toString().trim() : "") : "";
    var oB3 = iB2 !== -1 ? (r3[iB2] ? r3[iB2].toString().trim() : "") : "";
    var oC3 = iC2 !== -1 ? (r3[iC2] ? r3[iC2].toString().trim() : "") : "";
    var oD3 = iD2 !== -1 ? (r3[iD2] ? r3[iD2].toString().trim() : "") : "";
    var ans3 = "";
    var up3 = rawAns3.toUpperCase();
    if (["A","B","C","D"].indexOf(up3) !== -1) ans3 = [oA3,oB3,oC3,oD3][up3.charCodeAt(0)-65] || "";
    else if (["1","2","3","4"].indexOf(up3) !== -1) ans3 = [oA3,oB3,oC3,oD3][parseInt(up3)-1] || "";
    else ans3 = oA3;
    if (!imgGroupMap2[info3.imgUrl]) imgGroupMap2[info3.imgUrl] = [];
    imgGroupMap2[info3.imgUrl].push({ qid: qid3, ans: ans3 });
    if (ans3) allImgAnsPool2.push(ans3);
  }

  var questions = [];
  for (var i4 = 1; i4 < qRows.length; i4++) {
    var r4   = qRows[i4];
    var qid4 = (iId2 !== -1 && r4[iId2]) ? r4[iId2].toString().trim() : "ROW_" + (i4 + 1);
    if (!wrongSet[qid4]) continue;
    var q4 = iQ2 !== -1 ? (r4[iQ2] ? r4[iQ2].toString().trim() : "") : "";
    if (!q4) continue;

    var oA4 = iA2 !== -1 ? (r4[iA2] ? r4[iA2].toString().trim() : "") : "";
    var oB4 = iB2 !== -1 ? (r4[iB2] ? r4[iB2].toString().trim() : "") : "";
    var oC4 = iC2 !== -1 ? (r4[iC2] ? r4[iC2].toString().trim() : "") : "";
    var oD4 = iD2 !== -1 ? (r4[iD2] ? r4[iD2].toString().trim() : "") : "";
    var info4 = imgUrlMap[qid4] || { isImg: false, imgUrl: "" };
    var rawAns4 = iAns2 !== -1 ? (r4[iAns2] ? r4[iAns2].toString().trim() : "") : "";
    var ans4 = "";
    var up4 = rawAns4.toUpperCase();
    if (["A","B","C","D"].indexOf(up4) !== -1) ans4 = [oA4,oB4,oC4,oD4][up4.charCodeAt(0)-65] || "";
    else if (["1","2","3","4"].indexOf(up4) !== -1) ans4 = [oA4,oB4,oC4,oD4][parseInt(up4)-1] || "";
    else {
      var clean4 = rawAns4.replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase();
      var opts4  = [oA4,oB4,oC4,oD4].filter(Boolean);
      for (var oi4 = 0; oi4 < opts4.length; oi4++) {
        if (opts4[oi4].replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase() === clean4) { ans4 = opts4[oi4]; break; }
      }
      if (!ans4) ans4 = rawAns4;
    }

    var opts4Final;
    if (info4.isImg && info4.imgUrl) {
      var selfAns4 = ans4;
      var grp4 = imgGroupMap2[info4.imgUrl] || [];
      var sameGrp4 = [];
      grp4.forEach(function(g) { if (g.qid !== qid4 && g.ans) sameGrp4.push(g.ans); });
      var usedSet4 = {};
      usedSet4[selfAns4] = true;
      sameGrp4.forEach(function(a) { usedSet4[a] = true; });
      var fallback4 = allImgAnsPool2.filter(function(a) { return !usedSet4[a]; });
      var cands4 = sameGrp4.concat(fallback4);
      for (var ci4 = cands4.length - 1; ci4 > 0; ci4--) {
        var j4 = Math.floor(Math.random() * (ci4 + 1));
        var tmp4 = cands4[ci4]; cands4[ci4] = cands4[j4]; cands4[j4] = tmp4;
      }
      opts4Final = [selfAns4].concat(cands4.slice(0, 3)).filter(Boolean);
    } else {
      opts4Final = [oA4,oB4,oC4,oD4].filter(Boolean);
    }

    questions.push({
      id:      qid4,
      top:     iTop2 !== -1 ? (r4[iTop2] ? r4[iTop2].toString().trim() : "未分類") : "未分類",
      q:       q4,
      options: opts4Final,
      ans:     ans4,
      exp:     iExp2 !== -1 ? (r4[iExp2] ? r4[iExp2].toString().trim() : "尚無解析") : "尚無解析",
      color:   r4[COLOR2] ? r4[COLOR2].toString().trim() : "red",
      questionType: r4[TYPE2] ? r4[TYPE2].toString().trim() : "",
      cogType: r4[COG2]   ? r4[COG2].toString().trim()   : "",
      isImage: info4.isImg,
      imgUrl:  info4.isImg ? info4.imgUrl : "",
    });
  }

  return jsonResponse({ status: "ok", questions: questions });
}

// ─────────────────────────────────────────────
// Action：getMyCompletion
// ─────────────────────────────────────────────
function handleGetMyCompletion(payload) {
  var studentId = payload.studentId;
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var settings  = readSettings(ss);
  var passScore = settings.passScore;
  var reqTopics = settings.completionTopics;

  if (!reqTopics.length) return jsonResponse({ status: "ok", passScore: passScore, completionTopics: [], details: [] });

  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  var topicBest  = {};
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var sRows = scoreSheet.getDataRange().getValues();
    for (var i = 1; i < sRows.length; i++) {
      var sid   = sRows[i][1] ? sRows[i][1].toString() : "";
      var topic = sRows[i][3] ? sRows[i][3].toString() : "";
      var mode  = sRows[i][4] ? sRows[i][4].toString() : "";
      var score = Number(sRows[i][6]);
      if (sid !== studentId) continue;
      if (mode === "錯題重做") continue;
      if (score > (topicBest[topic] || 0)) topicBest[topic] = score;
    }
  }

  // ★ v9-6：計算各分類平均每題用時
  var topicTimeMap = {};  // topic → { totalSec, totalQ }
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRows2 = scoreSheet.getDataRange().getValues();
    for (var i2 = 1; i2 < scRows2.length; i2++) {
      var sid2    = scRows2[i2][1] ? scRows2[i2][1].toString() : "";
      var topic2  = scRows2[i2][3] ? scRows2[i2][3].toString() : "";
      var mode2   = scRows2[i2][4] ? scRows2[i2][4].toString() : "";
      var correct2= Number(scRows2[i2][7]) || 0;
      var wrong2  = Number(scRows2[i2][8]) || 0;
      var dur2    = Number(scRows2[i2][9]) || 0;
      if (sid2 !== studentId || mode2 === "錯題重做" || topic2 === "綜合練習") continue;
      if (dur2 <= 0) continue;
      var qCount = correct2 + wrong2;
      if (qCount <= 0) continue;
      if (!topicTimeMap[topic2]) topicTimeMap[topic2] = { totalSec: 0, totalQ: 0 };
      topicTimeMap[topic2].totalSec += dur2;
      topicTimeMap[topic2].totalQ   += qCount;
    }
  }

  var details = reqTopics.map(function(t) {
    var avgSec = null;
    if (topicTimeMap[t] && topicTimeMap[t].totalQ > 0) {
      avgSec = Math.round(topicTimeMap[t].totalSec / topicTimeMap[t].totalQ);
    }
    return { topic: t, best: (topicBest[t] !== undefined ? topicBest[t] : null), passed: (topicBest[t] || 0) >= passScore, avgSec: avgSec };
  });
  return jsonResponse({ status: "ok", passScore: passScore, completionTopics: reqTopics, details: details });
}

// ─────────────────────────────────────────────
// Action：getCompletionRanking（含排行快取）
// ─────────────────────────────────────────────
function handleGetCompletionRanking(payload) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var settings = readSettings(ss);
  var passScore  = settings.passScore;
  var reqTopics  = settings.completionTopics;
  var reqClasses = settings.completionClasses || [];
  var fullMode   = payload.full === true; // ★ full:true 略過快取，回傳含 students 的完整資料

  if (!reqTopics.length) return jsonResponse({ status: "ok", passScore: passScore, completionTopics: [], ranking: [] });

  // ★ v9-4：非 full 模式才讀快取（改用 Script Properties）
  if (!fullMode) {
    var cached = getRankingCacheProps(ss);
    if (cached) {
      var todayCache = readTodayPracticeCache(ss);
      cached.todayTotal = todayCache.todayTotal;
      cached.todayByClass = todayCache.todayByClass;
      cached.todayDate = todayCache.todayDate;
      cached.todayUpdatedAt = todayCache.updatedAt;
      return jsonResponse(cached);
    }
  }

  // 重新計算
  var stuSheet   = ss.getSheetByName(SHEET_STUDENTS);
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);

  var studentInfoMap = {};
  if (stuSheet && stuSheet.getLastRow() > 1) {
    var stRows = stuSheet.getDataRange().getValues();
    for (var i = 1; i < stRows.length; i++) {
      var sid = stRows[i][0] ? stRows[i][0].toString().trim() : "";
      if (sid) studentInfoMap[sid] = { name: stRows[i][1] ? stRows[i][1].toString().trim() : "", class: stRows[i][2] ? stRows[i][2].toString().trim() : "未分班" };
    }
  }

  var studentTopicBest = {};
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRows = scoreSheet.getDataRange().getValues();
    for (var i2 = 1; i2 < scRows.length; i2++) {
      var sid2   = scRows[i2][1] ? scRows[i2][1].toString() : "";
      var topic2 = scRows[i2][3] ? scRows[i2][3].toString() : "";
      var mode2  = scRows[i2][4] ? scRows[i2][4].toString() : "";
      var score2 = Number(scRows[i2][6]);
      if (!sid2 || mode2 === "錯題重做") continue;
      if (!studentTopicBest[sid2]) studentTopicBest[sid2] = {};
      if (score2 > (studentTopicBest[sid2][topic2] || 0)) studentTopicBest[sid2][topic2] = score2;
    }
  }

  var classMap = {};
  Object.keys(studentInfoMap).forEach(function(sid) {
    var info = studentInfoMap[sid];
    var cls  = info.class;
    if (reqClasses.length > 0 && reqClasses.indexOf(cls) === -1) return;
    if (!classMap[cls]) classMap[cls] = { students: {} };
    var topicBest = studentTopicBest[sid] || {};
    var completed = reqTopics.filter(function(t) { return (topicBest[t] || 0) >= passScore; }).length;
    var details   = reqTopics.map(function(t) { return { topic: t, best: topicBest[t] || null, passed: (topicBest[t] || 0) >= passScore }; });
    classMap[cls].students[sid] = { name: info.name, completed: completed, total: reqTopics.length, details: details };
  });

  var ranking = Object.keys(classMap).map(function(cls) {
    var stuArr    = Object.keys(classMap[cls].students);
    var totalComp = stuArr.reduce(function(s, sid) { return s + classMap[cls].students[sid].completed; }, 0);
    var avgComp   = stuArr.length > 0 ? (totalComp / stuArr.length) : 0;
    var allDone   = stuArr.filter(function(sid) { return classMap[cls].students[sid].completed === reqTopics.length; }).length;
    var students  = stuArr.map(function(sid) {
      var v = classMap[cls].students[sid];
      return { sid: sid, name: v.name, completed: v.completed, total: v.total, details: v.details };
    }).sort(function(a, b) { return b.completed - a.completed; });
    return { class: cls, studentCount: stuArr.length, avgCompleted: Math.round(avgComp * 10) / 10, allDoneCount: allDone, pct: reqTopics.length > 0 ? Math.round((avgComp / reqTopics.length) * 100) : 0, students: students };
  }).sort(function(a, b) { return b.pct - a.pct || a.class.localeCompare(b.class, "zh-TW"); });

  var result = { status: "ok", passScore: passScore, completionTopics: reqTopics, ranking: ranking };

  var todayCache2 = readTodayPracticeCache(ss);
  result.todayTotal = todayCache2.todayTotal;
  result.todayByClass = todayCache2.todayByClass;
  result.todayDate = todayCache2.todayDate;
  result.todayUpdatedAt = todayCache2.updatedAt;

  // 快取只存班級摘要（不含 students）
  var rankingForCache = ranking.map(function(r) {
    return { class: r.class, studentCount: r.studentCount, avgCompleted: r.avgCompleted, allDoneCount: r.allDoneCount, pct: r.pct };
  });
  // ★ v9-4：改用 Script Properties 分班存放
    setRankingCacheProps(passScore, reqTopics, ranking);

  return jsonResponse(result);
}

// ── 排行快取工具函式 ──
function getRankingCache(ss) {
  var sheet = ss.getSheetByName(SHEET_RANKING_CACHE);
  if (!sheet) return null;
  try {
    var vals  = sheet.getRange("A1:A2").getValues();
    var valid = vals[0][0] ? vals[0][0].toString() : "";
    var json  = vals[1][0] ? vals[1][0].toString() : "";
    if (valid !== "VALID" || !json) return null;
    var data = JSON.parse(json);
    data.status = "ok";
    return data;
  } catch(e) { return null; }
}

function setRankingCache(ss, data) {
  var sheet = ss.getSheetByName(SHEET_RANKING_CACHE);
  if (!sheet) sheet = ss.insertSheet(SHEET_RANKING_CACHE);
  sheet.getRange("A1:A2").setValues([["VALID"], [JSON.stringify(data)]]);
  sheet.getRange("A1").setBackground("#dcfce7").setFontWeight("bold");
}

function invalidateRankingCache(ss) {
  var sheet = ss.getSheetByName(SHEET_RANKING_CACHE);
  if (sheet) sheet.getRange("A1").setValue("INVALID");
  invalidateRankingCacheProps(); // ★ v9-4 同步失效 Script Properties 快取
}

// ─────────────────────────────────────────────
// 【工具函式】強制重建排行快取
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// 【工具函式】清空分類快取（GAS 編輯器執行）
//   當分類顯示不完整時使用，重整頁面後自動重建
// ─────────────────────────────────────────────
function clearTopicCache() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TOPIC_CACHE);
  if (sheet) {
    sheet.clearContents();
    Logger.log("✅ 分類快取已清空，下次載入頁面會自動重建");
  } else {
    Logger.log("⚠️ 找不到分類快取 Sheet");
  }
}

function forceRebuildRankingCache() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  invalidateRankingCache(ss);
  Logger.log("已將快取標記為 INVALID");
  var settings = readSettings(ss);
  Logger.log("設定 - passScore: " + settings.passScore);
  Logger.log("設定 - completionTopics: " + (settings.completionTopics || []).join(", "));
  Logger.log("設定 - completionClasses: " + ((settings.completionClasses || []).join(", ") || "全部"));
  handleGetCompletionRanking({});
  var sheet = ss.getSheetByName(SHEET_RANKING_CACHE);
  if (sheet) {
    var a1 = sheet.getRange("A1").getValue() ? sheet.getRange("A1").getValue().toString() : "";
    var a2 = sheet.getRange("A2").getValue() ? sheet.getRange("A2").getValue().toString() : "";
    Logger.log("A1 狀態：" + a1);
    Logger.log("A2 長度：" + a2.length + " 字元");
    if (a2.length > 0) {
      try {
        var parsed = JSON.parse(a2);
        Logger.log("✅ 快取重建成功，班級數：" + (parsed.ranking ? parsed.ranking.length : 0));
      } catch(e) { Logger.log("❌ JSON 解析失敗：" + e.message); }
    } else {
      Logger.log("❌ A2 仍是空的，請先到後台設定完成度分類");
    }
  }
}

// ─────────────────────────────────────────────
// Action 6：getTeacherData
// ─────────────────────────────────────────────
function handleGetTeacherData(payload) {
  // ★ v9-4 輕量版：只讀題庫+學生名單+成績紀錄，不讀題目作答明細
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 從題庫建立對照表
  var cogTypeMap     = {};
  var correctTextMap = {};
  var qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (qSheet && qSheet.getLastRow() > 1) {
    var qRows    = qSheet.getDataRange().getValues();
    var qHeaders = qRows[0].map(function(h) { return h.toString().trim(); });
    var iId_t  = findColIdx(qHeaders, ["題目ID","ID","id"]);
    var iAns_t = findColIdx(qHeaders, ["正確答案","答案","answer","ans"]);
    var iA_t   = findColIdx(qHeaders, ["選項A","選項1","optionA","a"]);
    var iB_t   = findColIdx(qHeaders, ["選項B","選項2","optionB","b"]);
    var iC_t   = findColIdx(qHeaders, ["選項C","選項3","optionC","c"]);
    var iD_t   = findColIdx(qHeaders, ["選項D","選項4","optionD","d"]);
    var COG_T  = findColIdx(qHeaders, ["認知類型","cogType","認知"]);
    if (COG_T === -1) COG_T = 13;
    for (var qi = 1; qi < qRows.length; qi++) {
      var r_t = qRows[qi];
      var qid_t = (iId_t !== -1 && r_t[iId_t]) ? r_t[iId_t].toString().trim() : "ROW_" + (qi + 1);
      cogTypeMap[qid_t] = r_t[COG_T] ? r_t[COG_T].toString().trim() : "";
      var rawAns_t = iAns_t !== -1 ? (r_t[iAns_t] ? r_t[iAns_t].toString().trim() : "") : "";
      var opts_t   = [
        iA_t !== -1 ? (r_t[iA_t] ? r_t[iA_t].toString().trim() : "") : "",
        iB_t !== -1 ? (r_t[iB_t] ? r_t[iB_t].toString().trim() : "") : "",
        iC_t !== -1 ? (r_t[iC_t] ? r_t[iC_t].toString().trim() : "") : "",
        iD_t !== -1 ? (r_t[iD_t] ? r_t[iD_t].toString().trim() : "") : "",
      ];
      var ct = "";
      var up_t = rawAns_t.toUpperCase();
      if (["A","B","C","D"].indexOf(up_t) !== -1) ct = opts_t[up_t.charCodeAt(0) - 65] || "";
      else if (["1","2","3","4"].indexOf(up_t) !== -1) ct = opts_t[parseInt(up_t) - 1] || "";
      else {
        var cl_t = rawAns_t.replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase();
        for (var oi_t = 0; oi_t < opts_t.length; oi_t++) {
          if (opts_t[oi_t].replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase() === cl_t) { ct = opts_t[oi_t]; break; }
        }
        if (!ct) ct = rawAns_t;
      }
      correctTextMap[qid_t] = ct;
    }
  }

  var stuSheet = ss.getSheetByName(SHEET_STUDENTS);
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  var studentInfoMap = {};
  if (stuSheet && stuSheet.getLastRow() > 1) {
    var sRows2 = stuSheet.getDataRange().getValues();
    for (var si2 = 1; si2 < sRows2.length; si2++) {
      var sid2 = sRows2[si2][0] ? sRows2[si2][0].toString().trim() : "";
      if (sid2) studentInfoMap[sid2] = { name: sRows2[si2][1] ? sRows2[si2][1].toString().trim() : "", class: sRows2[si2][2] ? sRows2[si2][2].toString().trim() : "未分班" };
    }
  }

  var studentHistory      = {};
  var classStats          = {};

  // ★ v9-4 班級統計改從成績紀錄計算（不讀題目作答明細）
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRows2 = scoreSheet.getDataRange().getValues();
    for (var i6 = 1; i6 < scRows2.length; i6++) {
      var sid6    = scRows2[i6][1] ? scRows2[i6][1].toString() : "";
      if (!sid6) continue;
      var stuInfo6 = studentInfoMap[sid6] || {};
      if (!studentHistory[sid6]) studentHistory[sid6] = { name: stuInfo6.name || sid6, class: stuInfo6.class || "未分班", attempts: [] };
      studentHistory[sid6].attempts.push({
        date:     scRows2[i6][0] ? scRows2[i6][0].toString() : "",
        topic:    scRows2[i6][3] ? scRows2[i6][3].toString() : "",
        mode:     scRows2[i6][4] ? scRows2[i6][4].toString() : "",
        attempt:  Number(scRows2[i6][5]),
        score:    Number(scRows2[i6][6]),
        correct:  Number(scRows2[i6][7]),
        wrong:    Number(scRows2[i6][8]),
        duration: Number(scRows2[i6][9]) || 0,  // ★ v9-681 J欄 作答秒數
        isRetry:  scRows2[i6][4] ? scRows2[i6][4].toString() === "錯題重做" : false,
      });
    }
  }

  // 從 studentHistory 計算 classList
  Object.keys(studentHistory).forEach(function(sid) {
    var cls = studentHistory[sid].class || "未分班";
    if (!classStats[cls]) classStats[cls] = { correct: 0, total: 0, studentSet: {} };
    classStats[cls].studentSet[sid] = true;
    studentHistory[sid].attempts.forEach(function(a) {
      if (!a.isRetry) {
        classStats[cls].total  += (a.correct || 0) + (a.wrong || 0);
        classStats[cls].correct += (a.correct || 0);
      }
    });
  });

  // ★ v9-6：計算各分類平均每題用時（從成績紀錄）
  var topicTimeStats = {};  // topic → { totalSec, totalQ, count }
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRowsT = scoreSheet.getDataRange().getValues();
    for (var iT = 1; iT < scRowsT.length; iT++) {
      var topicT = scRowsT[iT][3] ? scRowsT[iT][3].toString() : "";
      var modeT  = scRowsT[iT][4] ? scRowsT[iT][4].toString() : "";
      var corrT  = Number(scRowsT[iT][7]) || 0;
      var wronT  = Number(scRowsT[iT][8]) || 0;
      var durT   = Number(scRowsT[iT][9]) || 0;
      if (modeT === "錯題重做" || topicT === "綜合練習" || durT <= 0) continue;
      var qCountT = corrT + wronT;
      if (qCountT <= 0) continue;
      if (!topicTimeStats[topicT]) topicTimeStats[topicT] = { totalSec: 0, totalQ: 0, count: 0 };
      topicTimeStats[topicT].totalSec += durT;
      topicTimeStats[topicT].totalQ   += qCountT;
      topicTimeStats[topicT].count++;
    }
  }

  var classList2 = Object.keys(classStats).map(function(cls) {
    var s = classStats[cls];
    return { class: cls, correct: s.correct, total: s.total,
      rate: s.total > 0 ? Math.round((s.correct / s.total) * 100) : null,
      studentCount: Object.keys(s.studentSet).length };
  }).sort(function(a,b) { return a.class.localeCompare(b.class, "zh-TW"); });

  // 各分類平均每題用時
  var topicTimeList = Object.keys(topicTimeStats).map(function(t) {
    var s = topicTimeStats[t];
    return { topic: t, avgSec: Math.round(s.totalSec / s.totalQ), sessionCount: s.count };
  }).sort(function(a,b) { return a.topic.localeCompare(b.topic, "zh-TW"); });

  return jsonResponse({ status: "ok", studentHistory: studentHistory, classList: classList2, studentInfoMap: studentInfoMap, topicTimeList: topicTimeList });
}

// ─────────────────────────────────────────────
// 工具函式
// ─────────────────────────────────────────────
function readSettings(ss) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  var defaults = { passScore: 80, completionTopics: [], completionClasses: [] };
  if (!sheet || sheet.getLastRow() < 2) return defaults;
  var rows = sheet.getDataRange().getValues();
  var map  = {};
  for (var i = 1; i < rows.length; i++) {
    var key = rows[i][0] ? rows[i][0].toString().trim() : "";
    var val = rows[i][1] ? rows[i][1].toString().trim() : "";
    if (key) map[key] = val;
  }
  return {
    passScore:         parseInt(map["completion_pass_score"] || "80"),
    completionTopics:  map["completion_topics"]  ? map["completion_topics"].split(",").map(function(s) { return s.trim(); }).filter(Boolean) : [],
    completionClasses: map["completion_classes"] ? map["completion_classes"].split(",").map(function(s) { return s.trim(); }).filter(Boolean) : [],
    deadline:          map["deadline"] || "",
  };
}

function writeSettings(ss, passScore, completionTopics, completionClasses) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.appendRow(["設定名稱", "值"]);
    sheet.getRange(1,1,1,2).setFontWeight("bold").setBackground("#f3e8ff");
  }
  var rows   = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues() : [];
  var keyMap = {};
  rows.forEach(function(r, i) { if (r[0]) keyMap[r[0].toString().trim()] = i + 2; });
  function upsert(key, val) {
    if (keyMap[key]) sheet.getRange(keyMap[key], 2).setValue(val);
    else sheet.appendRow([key, val]);
  }
  upsert("completion_pass_score", passScore);
  upsert("completion_topics",     completionTopics.join(","));
  upsert("completion_classes",    completionClasses.join(","));
}

// ★ v9-68 驗證 token 是否有效（未被踢出）
function isTokenValid(ss, studentId, token) {
  if (!token) return true;  // 無 token 時不驗證（相容舊資料）
  var sheet = ss.getSheetByName(SHEET_LOGIN_STATE);
  if (!sheet || sheet.getLastRow() <= 1) return true;
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 8).getValues();
  for (var i = data.length-1; i >= 0; i--) {
    if (data[i][0].toString() !== studentId) continue;
    if (data[i][2].toString() !== token)     continue;
    var status = data[i][7].toString();
    return status === "active";
  }
  return false;
}

function localNow() {
  var now = new Date();
  var tw  = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  function pad(n) { return String(n).padStart(2, "0"); }
  return tw.getUTCFullYear() + "/" + pad(tw.getUTCMonth()+1) + "/" + pad(tw.getUTCDate()) + " " + pad(tw.getUTCHours()) + ":" + pad(tw.getUTCMinutes()) + ":" + pad(tw.getUTCSeconds());
}

function dateKeyTW(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, "Asia/Taipei", "yyyy/MM/dd");
  }
  return value.toString().slice(0, 10).replace(/-/g, "/");
}

function buildStudentClassMap(ss) {
  var map = {};
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet || sheet.getLastRow() <= 1) return map;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var sid = rows[i][0] ? rows[i][0].toString().trim() : "";
    if (!sid) continue;
    map[sid] = rows[i][2] ? rows[i][2].toString().trim() : "未分班";
  }
  return map;
}

function computeTodayPracticeCounts(ss) {
  var todayStr = localNow().slice(0, 10);
  var studentClassMap = buildStudentClassMap(ss);
  var todaySet = {};
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);

  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var rows = scoreSheet.getRange(2, 1, scoreSheet.getLastRow() - 1, 4).getValues();
    rows.forEach(function(r) {
      var date = dateKeyTW(r[0]);
      var sid = r[1] ? r[1].toString().trim() : "";
      if (date === todayStr && sid && !todaySet[sid]) {
        todaySet[sid] = studentClassMap[sid] || "未分班";
      }
    });
  }

  var todayByClass = {};
  Object.keys(todaySet).forEach(function(sid) {
    var cls = todaySet[sid] || "未分班";
    todayByClass[cls] = (todayByClass[cls] || 0) + 1;
  });

  return {
    todayDate: todayStr,
    todayTotal: Object.keys(todaySet).length,
    todayByClass: todayByClass,
    updatedAt: localNow()
  };
}

function writeTodayPracticeCache(ss) {
  var data = computeTodayPracticeCounts(ss);
  var sheet = ss.getSheetByName(SHEET_TODAY_PRACTICE_CACHE);
  if (!sheet) sheet = ss.insertSheet(SHEET_TODAY_PRACTICE_CACHE);
  sheet.clearContents();

  var rows = [
    ["項目", "值"],
    ["更新時間", data.updatedAt],
    ["日期", data.todayDate],
    ["今日總人數", data.todayTotal],
    ["", ""],
    ["班級", "今日練習人數"]
  ];
  Object.keys(data.todayByClass).sort(function(a, b) {
    return a.localeCompare(b, "zh-TW");
  }).forEach(function(cls) {
    rows.push([cls, data.todayByClass[cls]]);
  });

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#dcfce7");
  sheet.setFrozenRows(1);
  return data;
}

function readTodayPracticeCache(ss) {
  var empty = { todayDate: localNow().slice(0, 10), todayTotal: 0, todayByClass: {}, updatedAt: "" };
  var sheet = ss.getSheetByName(SHEET_TODAY_PRACTICE_CACHE);
  if (!sheet || sheet.getLastRow() < 4) return writeTodayPracticeCache(ss);

  var values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  var result = { todayDate: "", todayTotal: 0, todayByClass: {}, updatedAt: "" };
  for (var i = 0; i < values.length; i++) {
    var key = values[i][0] ? values[i][0].toString() : "";
    var val = values[i][1];
    if (key === "更新時間") result.updatedAt = val ? val.toString() : "";
    if (key === "日期") result.todayDate = val ? val.toString() : "";
    if (key === "今日總人數") result.todayTotal = Number(val) || 0;
    if (i >= 6 && key) result.todayByClass[key] = Number(val) || 0;
  }
  if (!result.todayDate) result.todayDate = empty.todayDate;
  if (result.todayDate !== empty.todayDate) return writeTodayPracticeCache(ss);
  return result;
}

function buildQuestionMetaMap(ss) {
  var map = {};
  var sheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!sheet || sheet.getLastRow() <= 1) return map;
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return h ? h.toString().trim() : ""; });
  var iId = findColIdx(headers, ["題目ID","ID","id"]);
  var iTop = findColIdx(headers, ["分類","category"]);
  var iQ = findColIdx(headers, ["題目","question","q"]);
  var iType = findColIdx(headers, ["題型","type"]);
  if (iType === -1) iType = 10;
  var iCog = findColIdx(headers, ["認知類型","cogType","認知"]);
  if (iCog === -1) iCog = 13;

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var qText = iQ !== -1 && r[iQ] ? r[iQ].toString().trim() : "";
    if (!qText) continue;
    var qid = iId !== -1 && r[iId] ? r[iId].toString().trim() : "ROW_" + (i + 1);
    map[qid] = {
      questionId: qid,
      topic: iTop !== -1 && r[iTop] ? r[iTop].toString().trim() : "未分類",
      questionText: qText,
      questionType: r[iType] ? r[iType].toString().trim() : "未分類",
      cogType: r[iCog] ? r[iCog].toString().trim() : "未分類"
    };
  }
  return map;
}

function buildStudentInfoMap(ss) {
  var map = {};
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet || sheet.getLastRow() <= 1) return map;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var sid = rows[i][0] ? rows[i][0].toString().trim() : "";
    if (!sid) continue;
    map[sid] = {
      name: rows[i][1] ? rows[i][1].toString().trim() : "",
      class: rows[i][2] ? rows[i][2].toString().trim() : "未分班"
    };
  }
  return map;
}

function getDetailSheets(ss) {
  return ss.getSheets().filter(function(s) {
    return s.getName() === SHEET_DETAILS || s.getName().indexOf(SHEET_DETAILS + "_") === 0;
  }).sort(function(a, b) { return a.getName().localeCompare(b.getName()); });
}

function pct(correct, total) {
  return total > 0 ? Math.round((correct / total) * 100) : null;
}

function avgSec(secSum, secCount) {
  return secCount > 0 ? Math.round(secSum / secCount) : null;
}

function bumpAgg(map, key, item, sid) {
  if (!map[key]) map[key] = {
    correct: 0, total: 0, secSum: 0, secCount: 0, students: {},
    meta: item.meta || {}, optionCounts: {}
  };
  map[key].total++;
  if (item.isCorrect) map[key].correct++;
  if (item.answerSec > 0 && item.answerSec < 600) {
    map[key].secSum += item.answerSec;
    map[key].secCount++;
  }
  if (sid) map[key].students[sid] = true;
  if (!item.isCorrect && item.selectedText) {
    map[key].optionCounts[item.selectedText] = (map[key].optionCounts[item.selectedText] || 0) + 1;
  }
}

function commonWrongOption(optionCounts) {
  var arr = Object.keys(optionCounts || {}).map(function(k) { return [k, optionCounts[k]]; });
  arr.sort(function(a, b) { return b[1] - a[1]; });
  return arr.length ? arr[0][0] + "（" + arr[0][1] + "次）" : "";
}

function writeCacheSheet(ss, name, header, rows, color) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  var values = [header].concat(rows);
  sheet.getRange(1, 1, Math.max(values.length, 1), header.length).setValues(values);
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground(color || "#f0f9ff");
  sheet.setFrozenRows(1);
  if (header.length >= 3) sheet.setFrozenColumns(1);
}

function buildAndSaveAnalysisCaches(ss) {
  var updatedAt = localNow();
  var qMeta = buildQuestionMetaMap(ss);
  var studentInfo = buildStudentInfoMap(ss);
  var classCategory = {};
  var studentCategory = {};
  var questionType = {};
  var questionStats = {};

  getDetailSheets(ss).forEach(function(sheet) {
    if (sheet.getLastRow() <= 1) return;
    ensureDetailHeader(sheet);
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return h ? h.toString().trim() : ""; });
    var cTime = findColIdx(headers, ["時間戳記"]);
    var cSid = findColIdx(headers, ["學號"]);
    var cName = findColIdx(headers, ["姓名"]);
    var cQid = findColIdx(headers, ["題目ID"]);
    var cQtext = findColIdx(headers, ["題目內容"]);
    var cTopic = findColIdx(headers, ["單元","分類"]);
    var cSel = findColIdx(headers, ["學生選項"]);
    var cResult = findColIdx(headers, ["是否答對"]);
    var cMode = findColIdx(headers, ["測驗模式"]);
    var cSec = findColIdx(headers, ["作答秒數"]);
    var cType = findColIdx(headers, ["題型"]);
    var cCog = findColIdx(headers, ["認知類型"]);

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var mode = cMode !== -1 && r[cMode] ? r[cMode].toString() : "";
      if (mode === "錯題重做") continue;
      var sid = cSid !== -1 && r[cSid] ? r[cSid].toString().trim() : "";
      var qid = cQid !== -1 && r[cQid] ? r[cQid].toString().trim() : "";
      if (!sid || !qid) continue;
      var meta = qMeta[qid] || {};
      var info = studentInfo[sid] || {};
      var cls = info.class || "未分班";
      var name = info.name || (cName !== -1 && r[cName] ? r[cName].toString().trim() : "");
      var topic = cTopic !== -1 && r[cTopic] ? r[cTopic].toString().trim() : (meta.topic || "未分類");
      var qText = cQtext !== -1 && r[cQtext] ? r[cQtext].toString().trim() : (meta.questionText || "");
      var qType = cType !== -1 && r[cType] ? r[cType].toString().trim() : (meta.questionType || "未分類");
      var cog = cCog !== -1 && r[cCog] ? r[cCog].toString().trim() : (meta.cogType || "未分類");
      var selectedText = cSel !== -1 && r[cSel] ? r[cSel].toString().trim() : "";
      var result = cResult !== -1 && r[cResult] ? r[cResult].toString() : "";
      var answerSec = cSec !== -1 && r[cSec] !== "" ? Number(r[cSec]) || 0 : 0;
      var isCorrect = result === "答對";

      var item = { isCorrect: isCorrect, answerSec: answerSec, selectedText: selectedText, meta: {
        className: cls, sid: sid, name: name, topic: topic, questionId: qid,
        questionText: qText, questionType: qType || "未分類", cogType: cog || "未分類"
      }};
      bumpAgg(classCategory, cls + "|" + topic, item, sid);
      bumpAgg(studentCategory, cls + "|" + sid + "|" + topic, item, sid);
      bumpAgg(questionType, cls + "|" + (qType || "未分類"), item, sid);
      bumpAgg(questionStats, qid, item, sid);
    }
  });

  var classRows = Object.keys(classCategory).map(function(key) {
    var s = classCategory[key], m = s.meta;
    return [updatedAt, m.className, m.topic, Object.keys(s.students).length, s.total, s.correct, pct(s.correct, s.total), avgSec(s.secSum, s.secCount)];
  }).sort(function(a,b) { return a[1].localeCompare(b[1],"zh-TW") || a[2].localeCompare(b[2],"zh-TW"); });

  var studentRows = Object.keys(studentCategory).map(function(key) {
    var s = studentCategory[key], m = s.meta;
    return [updatedAt, m.className, m.sid, m.name, m.topic, s.total, s.correct, pct(s.correct, s.total), avgSec(s.secSum, s.secCount)];
  }).sort(function(a,b) { return a[1].localeCompare(b[1],"zh-TW") || a[2].localeCompare(b[2]) || a[4].localeCompare(b[4],"zh-TW"); });

  var typeRows = Object.keys(questionType).map(function(key) {
    var s = questionType[key], m = s.meta;
    return [updatedAt, m.className, m.questionType, Object.keys(s.students).length, s.total, s.correct, pct(s.correct, s.total), avgSec(s.secSum, s.secCount)];
  }).sort(function(a,b) { return a[1].localeCompare(b[1],"zh-TW") || a[2].localeCompare(b[2],"zh-TW"); });

  var questionRows = Object.keys(questionStats).map(function(qid) {
    var s = questionStats[qid], m = s.meta;
    return [updatedAt, qid, m.topic, m.questionType, m.cogType, m.questionText, Object.keys(s.students).length, s.total, s.correct, pct(s.correct, s.total), avgSec(s.secSum, s.secCount), commonWrongOption(s.optionCounts)];
  }).sort(function(a,b) { return a[2].localeCompare(b[2],"zh-TW") || (a[9] || 101) - (b[9] || 101); });

  writeCacheSheet(ss, SHEET_CLASS_CATEGORY_ANALYSIS, ["更新時間","班級","分類","作答人數","作答題數","答對題數","答對率","平均作答秒數"], classRows, "#dcfce7");
  writeCacheSheet(ss, SHEET_STUDENT_CATEGORY_ANALYSIS, ["更新時間","班級","學號","姓名","分類","作答題數","答對題數","答對率","平均作答秒數"], studentRows, "#e0f2fe");
  writeCacheSheet(ss, SHEET_QUESTION_TYPE_ANALYSIS, ["更新時間","班級","題型","作答人數","作答題數","答對題數","答對率","平均作答秒數"], typeRows, "#fef3c7");
  writeCacheSheet(ss, SHEET_QUESTION_ANALYSIS, ["更新時間","題目ID","分類","題型","認知類型","題目","作答人數","作答題數","答對題數","答對率","平均作答秒數","常錯選項"], questionRows, "#ede9fe");
  return { classCategory: classRows.length, studentCategory: studentRows.length, questionType: typeRows.length, questionStats: questionRows.length, updatedAt: updatedAt };
}

function readSheetObjects(ss, name, limit) {
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var lastRow = sheet.getLastRow();
  var rowCount = Math.min(lastRow - 1, limit || 5000);
  var values = sheet.getRange(1, 1, rowCount + 1, sheet.getLastColumn()).getValues();
  var headers = values[0].map(function(h) { return h ? h.toString().trim() : ""; });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[i][c];
    out.push(obj);
  }
  return out;
}

function handleGetAnalysisCache(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var force = payload && payload.force === true;
  var classSheet = ss.getSheetByName(SHEET_CLASS_CATEGORY_ANALYSIS);
  if (force || !classSheet || classSheet.getLastRow() <= 1) buildAndSaveAnalysisCaches(ss);
  return jsonResponse({
    status: "ok",
    classCategory: readSheetObjects(ss, SHEET_CLASS_CATEGORY_ANALYSIS, 5000),
    studentCategory: readSheetObjects(ss, SHEET_STUDENT_CATEGORY_ANALYSIS, 12000),
    questionType: readSheetObjects(ss, SHEET_QUESTION_TYPE_ANALYSIS, 5000),
    questionStats: readSheetObjects(ss, SHEET_QUESTION_ANALYSIS, 5000)
  });
}

function hashString(str) {
  var hash = 2166136261;
  for (var i = 0; i < str.length; i++) {
    hash ^= str.charCodeAt(i);
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

function findColIdx(headers, names) {
  for (var ni = 0; ni < names.length; ni++) {
    for (var hi = 0; hi < headers.length; hi++) {
      if (headers[hi] === names[ni] || headers[hi].toLowerCase() === names[ni].toLowerCase()) return hi;
    }
  }
  return -1;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ╔══════════════════════════════════════════════════════════════╗
// ║  ★ v9-3 新增功能                                            ║
// ╚══════════════════════════════════════════════════════════════╝

// ─────────────────────────────────────────────
// 核心計算函式：讀取所有成績，計算每位學生每個分類最高分
// 同時計算班級排行，一次計算兩用
// ─────────────────────────────────────────────
function calcStudentTopicScores(ss) {
  var settings   = readSettings(ss);
  var passScore  = settings.passScore;
  var reqTopics  = settings.completionTopics;
  var reqClasses = settings.completionClasses || [];

  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  var stuSheet   = ss.getSheetByName(SHEET_STUDENTS);

  // 學號 → { name, class }
  var studentInfoMap = {};
  if (stuSheet && stuSheet.getLastRow() > 1) {
    var stRows = stuSheet.getDataRange().getValues();
    for (var si = 1; si < stRows.length; si++) {
      var sid = stRows[si][0] ? stRows[si][0].toString().trim() : "";
      if (sid) studentInfoMap[sid] = {
        name:  stRows[si][1] ? stRows[si][1].toString().trim() : "",
        class: stRows[si][2] ? stRows[si][2].toString().trim() : "未分班"
      };
    }
  }

  // 學號 → 分類 → 最高分
  var studentTopicBest = {};
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRows = scoreSheet.getDataRange().getValues();
    for (var i = 1; i < scRows.length; i++) {
      var sid2   = scRows[i][1] ? scRows[i][1].toString().trim() : "";
      var topic2 = scRows[i][3] ? scRows[i][3].toString().trim() : "";
      var mode2  = scRows[i][4] ? scRows[i][4].toString().trim() : "";
      var score2 = Number(scRows[i][6]);
      if (!sid2 || mode2 === "錯題重做") continue;
      if (!studentTopicBest[sid2]) studentTopicBest[sid2] = {};
      if (score2 > (studentTopicBest[sid2][topic2] || 0)) studentTopicBest[sid2][topic2] = score2;
    }
  }

  return {
    settings:          settings,
    passScore:         passScore,
    reqTopics:         reqTopics,
    reqClasses:        reqClasses,
    studentInfoMap:    studentInfoMap,
    studentTopicBest:  studentTopicBest,
  };
}

// ─────────────────────────────────────────────
// 更新學生成績總表 + 排行快取（合併計算）
// ─────────────────────────────────────────────
function buildAndSaveScoreTable(ss) {
  var data = calcStudentTopicScores(ss);
  var settings         = data.settings;
  var passScore        = data.passScore;
  var reqTopics        = data.reqTopics;
  var reqClasses       = data.reqClasses;
  var studentInfoMap   = data.studentInfoMap;
  var studentTopicBest = data.studentTopicBest;

  // ── 取得所有分類清單（從題庫）──
  var qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  var allTopics = [];
  if (qSheet && qSheet.getLastRow() > 1) {
    var qRows    = qSheet.getDataRange().getValues();
    var qHeaders = qRows[0].map(function(h) { return h.toString().trim(); });
    var iTop = findColIdx(qHeaders, ["分類","category"]);
    var topicSet = {}, topicOrder = [];
    for (var qi = 1; qi < qRows.length; qi++) {
      var t = iTop !== -1 ? (qRows[qi][iTop] ? qRows[qi][iTop].toString().trim() : "") : "";
      if (t && !topicSet[t]) { topicSet[t] = true; topicOrder.push(t); }
    }
    allTopics = topicOrder;
  }

  // ── 寫入學生成績總表 ──
  var tableSheet = ss.getSheetByName(SHEET_SCORE_TABLE);
  if (!tableSheet) {
    tableSheet = ss.insertSheet(SHEET_SCORE_TABLE);
  }
  tableSheet.clearContents();

  // 標題列
  var header = ["班級","學號","姓名"].concat(allTopics).concat(["完成度","最後更新"]);
  tableSheet.getRange(1, 1, 1, header.length).setValues([header]);
  tableSheet.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#e0e7ff");
  tableSheet.setFrozenRows(1);
  tableSheet.setFrozenColumns(3);

  // 學生資料列（依班級→學號排序）
  var sids = Object.keys(studentInfoMap).sort(function(a, b) {
    var ca = studentInfoMap[a].class, cb = studentInfoMap[b].class;
    if (ca !== cb) return ca.localeCompare(cb, "zh-TW");
    return a.localeCompare(b);
  });

  var now = localNow();
  var rows = [];
  sids.forEach(function(sid) {
    var info      = studentInfoMap[sid];
    var topicBest = studentTopicBest[sid] || {};
    var scores    = allTopics.map(function(t) { return topicBest[t] !== undefined ? topicBest[t] : ""; });
    var completed = reqTopics.length > 0
      ? reqTopics.filter(function(t) { return (topicBest[t] || 0) >= passScore; }).length
      : "";
    var compText  = reqTopics.length > 0 ? (completed + "/" + reqTopics.length) : "";
    rows.push([info.class, sid, info.name].concat(scores).concat([compText, now]));
  });

  if (rows.length > 0) {
    tableSheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // ── 同時更新排行快取 ──
  if (reqTopics.length > 0) {
    var classMap = {};
    sids.forEach(function(sid) {
      var info  = studentInfoMap[sid];
      var cls   = info.class;
      if (reqClasses.length > 0 && reqClasses.indexOf(cls) === -1) return;
      if (!classMap[cls]) classMap[cls] = { students: {} };
      var topicBest = studentTopicBest[sid] || {};
      var completed = reqTopics.filter(function(t) { return (topicBest[t] || 0) >= passScore; }).length;
      var details   = reqTopics.map(function(t) {
        return { topic: t, best: topicBest[t] || null, passed: (topicBest[t] || 0) >= passScore };
      });
      classMap[cls].students[sid] = { name: info.name, completed: completed, total: reqTopics.length, details: details };
    });

    var ranking = Object.keys(classMap).map(function(cls) {
      var stuArr    = Object.keys(classMap[cls].students);
      var totalComp = stuArr.reduce(function(s, sid) { return s + classMap[cls].students[sid].completed; }, 0);
      var avgComp   = stuArr.length > 0 ? (totalComp / stuArr.length) : 0;
      var allDone   = stuArr.filter(function(sid) { return classMap[cls].students[sid].completed === reqTopics.length; }).length;
      var students  = stuArr.map(function(sid) {
        var v = classMap[cls].students[sid];
        return { sid: sid, name: v.name, completed: v.completed, total: v.total, details: v.details };
      }).sort(function(a, b) { return b.completed - a.completed; });
      return {
        class: cls, studentCount: stuArr.length,
        avgCompleted: Math.round(avgComp * 10) / 10,
        allDoneCount: allDone,
        pct: reqTopics.length > 0 ? Math.round((avgComp / reqTopics.length) * 100) : 0,
        students: students
      };
    }).sort(function(a, b) { return b.pct - a.pct || a.class.localeCompare(b.class, "zh-TW"); });

    var rankingForCache = ranking.map(function(r) {
      return { class: r.class, studentCount: r.studentCount, avgCompleted: r.avgCompleted, allDoneCount: r.allDoneCount, pct: r.pct };
    });
    // ★ v9-4：改用 Script Properties 分班存放
    setRankingCacheProps(passScore, reqTopics, ranking);
    Logger.log("✅ 排行快取已更新，班級數：" + ranking.length);
  }

  Logger.log("✅ 學生成績總表已更新，共 " + rows.length + " 位學生，" + allTopics.length + " 個分類");
  return rows.length;
}

// ─────────────────────────────────────────────
// 【定時觸發】每小時自動執行
//   沒有新作答時跳過，不更新
//   在 GAS 觸發器設定：函式=autoUpdateScoreSheet，每小時
// ─────────────────────────────────────────────
function autoUpdateScoreSheet() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var scoreSheet  = ss.getSheetByName(SHEET_SCORES);
  writeTodayPracticeCache(ss);

  if (!scoreSheet || scoreSheet.getLastRow() <= 1) {
    Logger.log("⏭ 成績紀錄是空的，跳過更新");
    return;
  }

  // 檢查最後一筆成績的時間
  var lastRow      = scoreSheet.getLastRow();
  var lastTimeVal  = scoreSheet.getRange(lastRow, 1).getValue();
  var lastTime     = lastTimeVal instanceof Date ? lastTimeVal : new Date(lastTimeVal);

  if (isNaN(lastTime)) {
    Logger.log("⏭ 無法讀取最後作答時間，跳過");
    return;
  }

  // 距現在超過 2 小時沒有新作答 → 跳過
  var diffHours = (new Date().getTime() - lastTime.getTime()) / (1000 * 60 * 60);
  if (diffHours > 2) {
    Logger.log("⏭ 距上次作答 " + diffHours.toFixed(1) + " 小時，無新作答，跳過更新");
    return;
  }

  Logger.log("⏱ 距上次作答 " + diffHours.toFixed(1) + " 小時，開始更新...");
  buildAndSaveScoreTable(ss);
  buildAndSaveAnalysisCaches(ss);
}

// ─────────────────────────────────────────────
// Action：getStudentScoreTable（手動立即更新）
// ─────────────────────────────────────────────
function handleGetStudentScoreTable(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var count = buildAndSaveScoreTable(ss);
  writeTodayPracticeCache(ss);
  var analysis = buildAndSaveAnalysisCaches(ss);
  return jsonResponse({ status: "ok", message: "✅ 學生成績總表與分析快取已更新，共 " + count + " 位學生", count: count, analysis: analysis });
}

// ─────────────────────────────────────────────
// 【工具函式】手動立即更新（GAS 編輯器執行）
// ─────────────────────────────────────────────
function manualUpdateScoreTable() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var count = buildAndSaveScoreTable(ss);
  writeTodayPracticeCache(ss);
  var analysis = buildAndSaveAnalysisCaches(ss);
  Logger.log("✅ 完成！共 " + count + " 位學生；分析快取：" + JSON.stringify(analysis));
}

// ╔══════════════════════════════════════════════════════════════╗
// ║  ★ v9-4 新增功能                                            ║
// ╚══════════════════════════════════════════════════════════════╝

// ─────────────────────────────────────────────
// 登入狀態 Sheet 工具
// 欄位：A=學號 B=姓名 C=token D=登入時間 E=IP F=裝置 G=瀏覽器 H=狀態
// ─────────────────────────────────────────────
function getOrCreateLoginStateSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_LOGIN_STATE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOGIN_STATE);
    sheet.appendRow(["學號","姓名","token","登入時間","IP","裝置","瀏覽器","狀態"]);
    sheet.getRange(1,1,1,8).setFontWeight("bold").setBackground("#e0e7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ─────────────────────────────────────────────
// 產生 Token
// ─────────────────────────────────────────────
function generateToken(studentId) {
  var now   = new Date().getTime().toString(36);
  var rand  = Math.random().toString(36).slice(2, 8);
  return studentId + "_" + now + "_" + rand;
}

// ─────────────────────────────────────────────
// Action：loginStudent
//   學生登入時呼叫，產生 token，記錄 IP/裝置/瀏覽器
//   若同帳號已有 active token → 舊 token 標記為 kicked，新 token 保持 active
// ─────────────────────────────────────────────
function handleLoginStudent(payload) {
  var studentId = payload.studentId ? payload.studentId.toString().trim() : "";
  var name      = payload.name      ? payload.name.toString().trim()      : "";
  var ip        = payload.ip        ? payload.ip.toString()               : "未知";
  var device    = payload.device    ? payload.device.toString()           : "未知";
  var browser   = payload.browser   ? payload.browser.toString()         : "未知";

  if (!studentId || !name) return jsonResponse({ status: "error", message: "學號或姓名不得為空" });

  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateLoginStateSheet(ss);
    var token = generateToken(studentId);
    var activeCount = 0;

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0].toString() === studentId && data[i][7].toString() === "active") {
          activeCount++;
          sheet.getRange(i + 2, 8).setValue("kicked");
        }
      }
    }

    sheet.appendRow([studentId, name, token, localNow(), ip, device, browser, "active"]);
    return jsonResponse({ status: "ok", token: token, isDuplicate: activeCount > 0, kickedOldSessions: activeCount });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// Action：verifySession
//   前端每 30 秒呼叫一次，驗證 token 是否仍有效
//   回傳 { valid: true } 或 { valid: false, reason: "kicked" }
// ─────────────────────────────────────────────
function handleVerifySession(payload) {
  var studentId = payload.studentId ? payload.studentId.toString().trim() : "";
  var token     = payload.token     ? payload.token.toString().trim()     : "";

  if (!studentId || !token) return jsonResponse({ valid: false, reason: "missing" });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LOGIN_STATE);
  if (!sheet || sheet.getLastRow() <= 1) return jsonResponse({ valid: false, reason: "no_data" });

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

  // 從最新往舊找（最後登入的優先）
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    if (row[0].toString() !== studentId) continue;
    if (row[2].toString() !== token)     continue;

    var status = row[7].toString();
    if (status === "active")  return jsonResponse({ valid: true });
    if (status === "kicked")  return jsonResponse({ valid: false, reason: "kicked" });
    return jsonResponse({ valid: false, reason: "invalid" });
  }

  return jsonResponse({ valid: false, reason: "not_found" });
}

// ─────────────────────────────────────────────
// ★ v9-4：排行快取改用 Script Properties 分班存
//   RANKING_VALID    = "true" / "false"
//   RANKING_SUMMARY  = 班級摘要 JSON（首頁排行）
//   RANKING_CLASS_護516 = 該班完整 students JSON
// ─────────────────────────────────────────────
function getRankingCacheProps(ss) {
  var props = PropertiesService.getScriptProperties();
  try {
    var valid = props.getProperty("RANKING_VALID");
    if (valid !== "true") return null;
    var summaryJson = props.getProperty("RANKING_SUMMARY");
    if (!summaryJson) return null;
    var summary = JSON.parse(summaryJson);
    // 回傳摘要版本（不含 students），點班級時再讀各班
    return { status: "ok", passScore: summary.passScore, completionTopics: summary.completionTopics, ranking: summary.ranking };
  } catch(e) { return null; }
}

function getClassStudentsCacheProps(className) {
  var props = PropertiesService.getScriptProperties();
  try {
    var key  = "RANKING_CLASS_" + className;
    var json = props.getProperty(key);
    if (!json) return null;
    return JSON.parse(json);
  } catch(e) { return null; }
}

function setRankingCacheProps(passScore, completionTopics, ranking) {
  var props = PropertiesService.getScriptProperties();
  try {
    // 存班級摘要（不含 students）
    var rankingSummary = ranking.map(function(r) {
      return { class: r.class, studentCount: r.studentCount, avgCompleted: r.avgCompleted, allDoneCount: r.allDoneCount, pct: r.pct };
    });
    props.setProperty("RANKING_SUMMARY", JSON.stringify({ passScore: passScore, completionTopics: completionTopics, ranking: rankingSummary }));

    // 分班存完整 students 資料
    ranking.forEach(function(r) {
      var key = "RANKING_CLASS_" + r.class;
      props.setProperty(key, JSON.stringify(r.students || []));
    });

    props.setProperty("RANKING_VALID", "true");
    Logger.log("✅ 排行快取已存入 Script Properties，班級數：" + ranking.length);
  } catch(e) {
    Logger.log("⚠️ Script Properties 寫入失敗：" + e.message);
  }
}

function invalidateRankingCacheProps() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("RANKING_VALID", "false");
}

// ─────────────────────────────────────────────
// 新增 Action：getClassStudents
//   前端點班級時呼叫，取得該班完整學生資料
// ─────────────────────────────────────────────
// 路由已在 doPost 加入（需再補）

function handleGetClassStudents(payload) {
  var className = payload.className ? payload.className.toString() : "";
  if (!className) return jsonResponse({ status: "error", message: "缺少 className" });

  // 先從 Script Properties 快取取
  var cached = getClassStudentsCacheProps(className);
  if (cached) return jsonResponse({ status: "ok", students: cached, fromCache: true });

  // 快取不存在 → 即時計算
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var data = calcStudentTopicScores(ss);
  var students = [];
  Object.keys(data.studentInfoMap).forEach(function(sid) {
    var info = data.studentInfoMap[sid];
    if (info.class !== className) return;
    var topicBest = data.studentTopicBest[sid] || {};
    var completed = data.reqTopics.filter(function(t) { return (topicBest[t] || 0) >= data.passScore; }).length;
    var details   = data.reqTopics.map(function(t) {
      return { topic: t, best: topicBest[t] || null, passed: (topicBest[t] || 0) >= data.passScore };
    });
    students.push({ sid: sid, name: info.name, completed: completed, total: data.reqTopics.length, details: details });
  });
  students.sort(function(a, b) { return b.completed - a.completed; });
  return jsonResponse({ status: "ok", students: students, fromCache: false });
}

// ─────────────────────────────────────────────
// ★ v9-4 handleGetDetailStats（重量版，懶載入）
//   只有前端點「題目難度分析」或「單元統計」時才呼叫
//   讀：題庫 + 題目作答明細（所有分頁）
//   回傳：questionStats、topicStats、studentWrongDetails
// ─────────────────────────────────────────────
function handleGetDetailStats(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 從題庫建立對照表
  var cogTypeMap     = {};
  var correctTextMap = {};
  var qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (qSheet && qSheet.getLastRow() > 1) {
    var qRows    = qSheet.getDataRange().getValues();
    var qHeaders = qRows[0].map(function(h) { return h.toString().trim(); });
    var iId_t  = findColIdx(qHeaders, ["題目ID","ID","id"]);
    var iAns_t = findColIdx(qHeaders, ["正確答案","答案","answer","ans"]);
    var iA_t   = findColIdx(qHeaders, ["選項A","選項1","optionA","a"]);
    var iB_t   = findColIdx(qHeaders, ["選項B","選項2","optionB","b"]);
    var iC_t   = findColIdx(qHeaders, ["選項C","選項3","optionC","c"]);
    var iD_t   = findColIdx(qHeaders, ["選項D","選項4","optionD","d"]);
    var COG_T  = findColIdx(qHeaders, ["認知類型","cogType","認知"]);
    if (COG_T === -1) COG_T = 13;
    for (var qi = 1; qi < qRows.length; qi++) {
      var r_t   = qRows[qi];
      var qid_t = (iId_t !== -1 && r_t[iId_t]) ? r_t[iId_t].toString().trim() : "ROW_" + (qi + 1);
      cogTypeMap[qid_t] = r_t[COG_T] ? r_t[COG_T].toString().trim() : "";
      var rawAns_t = iAns_t !== -1 ? (r_t[iAns_t] ? r_t[iAns_t].toString().trim() : "") : "";
      var opts_t   = [
        iA_t !== -1 ? (r_t[iA_t] ? r_t[iA_t].toString().trim() : "") : "",
        iB_t !== -1 ? (r_t[iB_t] ? r_t[iB_t].toString().trim() : "") : "",
        iC_t !== -1 ? (r_t[iC_t] ? r_t[iC_t].toString().trim() : "") : "",
        iD_t !== -1 ? (r_t[iD_t] ? r_t[iD_t].toString().trim() : "") : "",
      ];
      var ct = "", up_t = rawAns_t.toUpperCase();
      if (["A","B","C","D"].indexOf(up_t) !== -1) ct = opts_t[up_t.charCodeAt(0) - 65] || "";
      else if (["1","2","3","4"].indexOf(up_t) !== -1) ct = opts_t[parseInt(up_t) - 1] || "";
      else {
        var cl_t = rawAns_t.replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase();
        for (var oi_t = 0; oi_t < opts_t.length; oi_t++) {
          if (opts_t[oi_t].replace(/^([1-4]|[A-D])[.\-、\s]*/i,"").trim().toLowerCase() === cl_t) { ct = opts_t[oi_t]; break; }
        }
        if (!ct) ct = rawAns_t;
      }
      correctTextMap[qid_t] = ct;
    }
  }

  // 讀題目作答明細（所有分頁）
  var questionStats       = {};
  var topicStats          = {};
  var studentWrongDetails = {};

  var allSheets = ss.getSheets();
  var detailSheets = allSheets.filter(function(s) {
    return s.getName() === SHEET_DETAILS || s.getName().indexOf(SHEET_DETAILS + "_") === 0;
  }).sort(function(a,b) { return a.getName().localeCompare(b.getName()); });

  for (var di = 0; di < detailSheets.length; di++) {
    var dSheet = detailSheets[di];
    if (dSheet.getLastRow() <= 1) continue;
    var dRows = dSheet.getDataRange().getValues();
    for (var i = 1; i < dRows.length; i++) {
      var sid     = dRows[i][1] ? dRows[i][1].toString() : "";
      var qid     = dRows[i][3] ? dRows[i][3].toString() : "";
      var qtext   = dRows[i][4] ? dRows[i][4].toString() : "";
      var topic   = dRows[i][5] ? dRows[i][5].toString() : "未分類";
      var selOpt  = dRows[i][6] ? dRows[i][6].toString() : "未作答";
      var corrOpt = dRows[i][7] ? dRows[i][7].toString() : "";
      var result  = dRows[i][8] ? dRows[i][8].toString() : "";
      if (!sid || !qid) continue;
      var isCorrect = result === "答對";

      if (!questionStats[qid]) {
        questionStats[qid] = { text: qtext, topic: topic, correct: 0, total: 0, optionCounts: {}, correctText: correctTextMap[qid] || corrOpt, cogType: cogTypeMap[qid] || "" };
      }
      if (correctTextMap[qid]) questionStats[qid].correctText = correctTextMap[qid];
      questionStats[qid].total++;
      if (isCorrect) questionStats[qid].correct++;
      if (selOpt && selOpt !== "未作答") {
        questionStats[qid].optionCounts[selOpt] = (questionStats[qid].optionCounts[selOpt] || 0) + 1;
      }
      // ★ v9-67 M 欄 = 作答秒數
      var answerSec = (dRows[i][11] !== undefined && dRows[i][11] !== "") ? Number(dRows[i][11]) : null;
      if (answerSec !== null && answerSec > 0 && answerSec < 600) {
        if (!questionStats[qid].secSum)   questionStats[qid].secSum   = 0;
        if (!questionStats[qid].secCount) questionStats[qid].secCount = 0;
        questionStats[qid].secSum   += answerSec;
        questionStats[qid].secCount++;
      }

      if (!topicStats[topic]) topicStats[topic] = { correct: 0, total: 0 };
      topicStats[topic].total++;
      if (isCorrect) topicStats[topic].correct++;

      if (!isCorrect) {
        if (!studentWrongDetails[sid]) studentWrongDetails[sid] = {};
        studentWrongDetails[sid][qid] = { qid: qid, qtext: qtext, topic: topic, selectedText: selOpt, correctText: corrOpt };
      } else {
        if (studentWrongDetails[sid]) delete studentWrongDetails[sid][qid];
      }
    }
  }

  // 補入題庫中尚無作答的新題目
  if (qSheet && qSheet.getLastRow() > 1) {
    var qRowsFull = qSheet.getDataRange().getValues();
    var qHdrFull  = qRowsFull[0].map(function(h) { return h.toString().trim(); });
    var iId_f  = findColIdx(qHdrFull, ["題目ID","ID","id"]);
    var iTop_f = findColIdx(qHdrFull, ["分類","category"]);
    var iQ_f   = findColIdx(qHdrFull, ["題目","question","q"]);
    for (var qfi = 1; qfi < qRowsFull.length; qfi++) {
      var r_f   = qRowsFull[qfi];
      var qid_f = (iId_f !== -1 && r_f[iId_f]) ? r_f[iId_f].toString().trim() : "ROW_" + (qfi + 1);
      var qt_f  = iQ_f   !== -1 ? (r_f[iQ_f]   ? r_f[iQ_f].toString().trim()   : "") : "";
      var top_f = iTop_f !== -1 ? (r_f[iTop_f]  ? r_f[iTop_f].toString().trim() : "未分類") : "未分類";
      if (!qt_f) continue;
      if (!questionStats[qid_f]) {
        questionStats[qid_f] = { text: qt_f, topic: top_f, correct: 0, total: 0, optionCounts: {}, correctText: correctTextMap[qid_f] || "", cogType: cogTypeMap[qid_f] || "" };
      } else {
        if (!questionStats[qid_f].text || questionStats[qid_f].text === qid_f) {
          questionStats[qid_f].text  = qt_f;
          questionStats[qid_f].topic = top_f;
        }
      }
    }
  }

  var questionList = Object.keys(questionStats).map(function(id) {
    var s = questionStats[id];
    var avgSec = (s.secCount && s.secCount > 0) ? Math.round(s.secSum / s.secCount) : null;
    return { id: id, text: s.text, topic: s.topic, correct: s.correct, total: s.total,
      rate: s.total > 0 ? Math.round((s.correct / s.total) * 100) : null,
      optionCounts: s.optionCounts, correctText: s.correctText || "", cogType: s.cogType || "",
      avgSec: avgSec };  // ★ v9-67 每題平均用時
  }).sort(function(a,b) {
    if (a.topic < b.topic) return -1;
    if (a.topic > b.topic) return 1;
    return (a.rate !== null ? a.rate : 100) - (b.rate !== null ? b.rate : 100);
  });

  var topicList = Object.keys(topicStats).map(function(t) {
    var s = topicStats[t];
    return { topic: t, correct: s.correct, total: s.total, rate: s.total > 0 ? Math.round((s.correct / s.total) * 100) : null };
  });

  var wrongFmt = {};
  Object.keys(studentWrongDetails).forEach(function(sid) {
    wrongFmt[sid] = Object.keys(studentWrongDetails[sid]).map(function(qid) { return studentWrongDetails[sid][qid]; });
  });

  // ★ v9-67 認知類型統計（方案 C）
  var cogTypeStats = {};
  questionList.forEach(function(q) {
    var cog = q.cogType || "未分類";
    if (!cogTypeStats[cog]) cogTypeStats[cog] = { correct: 0, total: 0, secSum: 0, secCount: 0 };
    cogTypeStats[cog].total   += q.total;
    cogTypeStats[cog].correct += q.correct;
    if (q.avgSec !== null) {
      cogTypeStats[cog].secSum   += q.avgSec * (questionStats[q.id] ? (questionStats[q.id].secCount||0) : 0);
      cogTypeStats[cog].secCount += questionStats[q.id] ? (questionStats[q.id].secCount||0) : 0;
    }
  });
  var cogTypeList = Object.keys(cogTypeStats).map(function(cog) {
    var s = cogTypeStats[cog];
    return {
      cogType: cog,
      correct: s.correct,
      total:   s.total,
      rate:    s.total > 0 ? Math.round((s.correct / s.total) * 100) : null,
      avgSec:  s.secCount > 0 ? Math.round(s.secSum / s.secCount) : null
    };
  }).sort(function(a,b) { return a.cogType.localeCompare(b.cogType, "zh-TW"); });

  return jsonResponse({ status: "ok", questionStats: questionList, topicStats: topicList, studentWrongDetails: wrongFmt, cogTypeStats: cogTypeList });
}

// ─────────────────────────────────────────────
// 【工具函式】checkDataStatus（GAS 編輯器執行）
//   確認各 Sheet 目前資料狀況
// ─────────────────────────────────────────────
function checkDataStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Sheet 名稱：" + ss.getName());
  Logger.log("Sheet ID："   + ss.getId());

  var sheets = [
    SHEET_QUESTIONS, SHEET_SCORES, SHEET_DETAILS,
    SHEET_STUDENTS,  SHEET_WRONG_IDX, SHEET_SCORE_TABLE,
    SHEET_LOGIN_STATE, SHEET_SETTINGS
  ];
  sheets.forEach(function(name) {
    var s = ss.getSheetByName(name);
    Logger.log((s ? "✅ " : "❌ ") + name + "：" + (s ? (s.getLastRow()-1) + " 筆" : "不存在"));
  });

  // 封存明細分頁
  var allSheets = ss.getSheets();
  var archived  = allSheets.filter(function(s) {
    return s.getName().indexOf(SHEET_DETAILS + "_") === 0;
  });
  if (archived.length) {
    Logger.log("封存明細分頁：" + archived.length + " 個");
    archived.forEach(function(s) { Logger.log("  - " + s.getName() + "（" + (s.getLastRow()-1) + " 列）"); });
  }

  // Script Properties 排行快取
  var props  = PropertiesService.getScriptProperties();
  var valid  = props.getProperty("RANKING_VALID") || "無";
  var sumLen = (props.getProperty("RANKING_SUMMARY") || "").length;
  Logger.log("排行快取 VALID：" + valid + "，SUMMARY 長度：" + sumLen);
}

// ─────────────────────────────────────────────
// 【診斷工具】diagnoseTopics（GAS 編輯器執行）
//   顯示題庫標題列和前3筆資料，確認欄位偵測是否正確
// ─────────────────────────────────────────────
function diagnoseTopics() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!sheet) { Logger.log("❌ 找不到題庫"); return; }
  var rows = sheet.getDataRange().getValues();
  Logger.log("題庫總列數：" + (rows.length - 1));
  Logger.log("標題列：" + JSON.stringify(rows[0]));
  if (rows.length > 1) Logger.log("第1筆：" + JSON.stringify(rows[1]));
  if (rows.length > 2) Logger.log("第2筆：" + JSON.stringify(rows[2]));

  // 確認分類欄位
  var headers = rows[0].map(function(h) { return h.toString().trim(); });
  var iTop = -1;
  ["分類","category"].forEach(function(name) {
    if (iTop !== -1) return;
    var idx = headers.indexOf(name);
    if (idx === -1) {
      for (var i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase() === name.toLowerCase()) { idx = i; break; }
      }
    }
    if (idx !== -1) iTop = idx;
  });
  Logger.log("分類欄位索引（iTop）：" + iTop + (iTop !== -1 ? "（B欄=1，0-based）" : "（❌ 找不到！）"));

  // 列出所有分類
  var topics = {};
  for (var i = 1; i < rows.length; i++) {
    var t = iTop !== -1 ? (rows[i][iTop] ? rows[i][iTop].toString().trim() : "") : "";
    if (t) topics[t] = true;
  }
  Logger.log("找到的分類（共" + Object.keys(topics).length + "個）：" + Object.keys(topics).join(", "));
}

// ─────────────────────────────────────────────
// ★ v9-66 getDuplicateLoginReport
//   分析「登入狀態」和「成績紀錄」交叉比對
//   找出同帳號使用多個 token 送出成績的紀錄
// ─────────────────────────────────────────────
function handleGetDuplicateLoginReport(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 讀取登入狀態 Sheet ──
  var loginSheet = ss.getSheetByName(SHEET_LOGIN_STATE);
  var loginRows  = [];
  if (loginSheet && loginSheet.getLastRow() > 1) {
    loginRows = loginSheet.getRange(2, 1, loginSheet.getLastRow()-1, 8).getValues();
  }

  // 整理：studentId → [{ token, time, ip, device, browser, status }]
  var loginMap = {};
  loginRows.forEach(function(r) {
    var sid     = r[0] ? r[0].toString() : "";
    var name    = r[1] ? r[1].toString() : "";
    var token   = r[2] ? r[2].toString() : "";
    var time    = r[3] ? r[3].toString() : "";
    var ip      = r[4] ? r[4].toString() : "";
    var device  = r[5] ? r[5].toString() : "";
    var browser = r[6] ? r[6].toString() : "";
    var status  = r[7] ? r[7].toString() : "";
    if (!sid || !token) return;
    if (!loginMap[sid]) loginMap[sid] = { name: name, logins: [] };
    loginMap[sid].logins.push({ token: token, time: time, ip: ip, device: device, browser: browser, status: status });
  });

  // ── 讀取成績紀錄 Sheet（K欄=token, L欄=IP）──
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  var scoreMap   = {};  // studentId → [{ token, ip, time, topic, score }]
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var sRows = scoreSheet.getRange(2, 1, scoreSheet.getLastRow()-1, 12).getValues();
    sRows.forEach(function(r) {
      var sid   = r[1] ? r[1].toString() : "";
      var token = r[10] ? r[10].toString() : "";
      var ip    = r[11] ? r[11].toString() : "";
      if (!sid) return;
      if (!scoreMap[sid]) scoreMap[sid] = [];
      scoreMap[sid].push({
        time:  r[0] ? r[0].toString() : "",
        topic: r[3] ? r[3].toString() : "",
        mode:  r[4] ? r[4].toString() : "",
        score: Number(r[6]),
        token: token,
        ip:    ip
      });
    });
  }

  // ── 找出有重複登入證據的學生 ──
  var suspects = [];
  Object.keys(loginMap).forEach(function(sid) {
    var info   = loginMap[sid];
    var logins = info.logins;

    var hasKicked = logins.some(function(l) { return l.status === "kicked"; });
    if (!hasKicked) return;

    // 找出這個學生用了哪些不同的 token 送出成績
    var scores    = scoreMap[sid] || [];
    var tokenSet  = {};
    scores.forEach(function(s) {
      if (s.token) {
        if (!tokenSet[s.token]) tokenSet[s.token] = [];
        tokenSet[s.token].push(s);
      }
    });
    var uniqueTokens = Object.keys(tokenSet).length;

    suspects.push({
      sid:          sid,
      name:         info.name,
      loginCount:   logins.length,
      replacedCount: logins.filter(function(l) { return l.status === "kicked"; }).length,
      uniqueTokensInScore: uniqueTokens,
      logins:       logins,
      scoresByToken: tokenSet
    });
  });

  // 依被踢出次數排序，最可疑的在前
  suspects.sort(function(a,b) { return b.replacedCount - a.replacedCount; });

  return jsonResponse({ status: "ok", suspects: suspects, total: suspects.length });
}
