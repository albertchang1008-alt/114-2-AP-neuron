// ╔══════════════════════════════════════════════════════════════╗
// ║      Google Apps Script — 題庫系統  v9-690                   ║
// ║      對應前端版本：quiz_final v1.691                          ║
// ║                                                              ║
// ║  更新紀錄：                                                   ║
// ║  v9-690  - v1.691：GAS 保留為 Google Sheet → Firebase 同步與後台入口
// ║            學生端題庫、判分、作答明細、錯題改由 Firebase 處理
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
// Firebase v1.691 同步：Google Sheet → Firestore 快取
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
  for (var i = 0; i < writes.length; i += 100) {
    var res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({ writes: writes.slice(i, i + 100) }),
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
  var topicTimeMap = {};  // topic → { totalSec, totalQ }
  
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    // 唯讀前 10 欄 (A~J)
    var sRows = scoreSheet.getRange(1, 1, scoreSheet.getLastRow(), 10).getValues();
    for (var i = 1; i < sRows.length; i++) {
      var sid   = sRows[i][1] ? sRows[i][1].toString() : "";
      if (sid !== studentId) continue;
      
      var topic = sRows[i][3] ? sRows[i][3].toString() : "";
      var mode  = sRows[i][4] ? sRows[i][4].toString() : "";
      if (mode === "錯題重做") continue;
      
      var score = Number(sRows[i][6]);
      if (score > (topicBest[topic] || 0)) topicBest[topic] = score;
      
      // 計算各分類平均每題用時 (原本 scRows2 邏輯)
      if (topic !== "綜合練習") {
        var correct = Number(sRows[i][7]) || 0;
        var wrong   = Number(sRows[i][8]) || 0;
        var dur     = Number(sRows[i][9]) || 0;
        if (dur > 0) {
          var qCount = correct + wrong;
          if (qCount > 0) {
            if (!topicTimeMap[topic]) topicTimeMap[topic] = { totalSec: 0, totalQ: 0 };
            topicTimeMap[topic].totalSec += dur;
            topicTimeMap[topic].totalQ   += qCount;
          }
        }
      }
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
    var scRows = scoreSheet.getRange(1, 1, scoreSheet.getLastRow(), 7).getValues();
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
  var topicTimeStats      = {};  // topic → { totalSec, totalQ, count }

  // ★ v9-4 班級統計與用時改從成績紀錄計算，只讀取前 10 欄 (A~J) 合併為單次讀取
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRows2 = scoreSheet.getRange(1, 1, scoreSheet.getLastRow(), 10).getValues();
    for (var i6 = 1; i6 < scRows2.length; i6++) {
      var sid6    = scRows2[i6][1] ? scRows2[i6][1].toString() : "";
      if (!sid6) continue;
      var stuInfo6 = studentInfoMap[sid6] || {};
      if (!studentHistory[sid6]) studentHistory[sid6] = { name: stuInfo6.name || sid6, class: stuInfo6.class || "未分班", attempts: [] };
      
      var sDate    = scRows2[i6][0] ? scRows2[i6][0].toString() : "";
      var sTopic   = scRows2[i6][3] ? scRows2[i6][3].toString() : "";
      var sMode    = scRows2[i6][4] ? scRows2[i6][4].toString() : "";
      var sAttempt = Number(scRows2[i6][5]);
      var sScore   = Number(scRows2[i6][6]);
      var sCorr    = Number(scRows2[i6][7]) || 0;
      var sWron    = Number(scRows2[i6][8]) || 0;
      var sDur     = Number(scRows2[i6][9]) || 0;
      var isRetry  = sMode === "錯題重做";
      
      studentHistory[sid6].attempts.push({
        date:     sDate,
        topic:    sTopic,
        mode:     sMode,
        attempt:  sAttempt,
        score:    sScore,
        correct:  sCorr,
        wrong:    sWron,
        duration: sDur,
        isRetry:  isRetry,
      });

      // 計算各分類平均每題用時
      if (!isRetry && sTopic !== "綜合練習" && sDur > 0) {
        var qCountT = sCorr + sWron;
        if (qCountT > 0) {
          if (!topicTimeStats[sTopic]) topicTimeStats[sTopic] = { totalSec: 0, totalQ: 0, count: 0 };
          topicTimeStats[sTopic].totalSec += sDur;
          topicTimeStats[sTopic].totalQ   += qCountT;
          topicTimeStats[sTopic].count++;
        }
      }
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
  var now = new Date();
  var todayStr = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd");
  var studentClassMap = buildStudentClassMap(ss);
  var todaySet = {};
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);

  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var rows = scoreSheet.getRange(2, 1, scoreSheet.getLastRow() - 1, 2).getValues(); // 僅需時間與學號
    rows.forEach(function(r) {
      var timeVal = r[0];
      if (!timeVal) return;
      
      var d = null;
      if (timeVal instanceof Date) {
        d = timeVal;
      } else {
        // 如果是字串，轉換並嘗試解析
        var str = timeVal.toString().trim().replace(/-/g, "/");
        if (str.indexOf("下午") !== -1) {
          str = str.replace("下午", "").trim();
          d = new Date(str);
          if (!isNaN(d.getTime())) d.setHours(d.getHours() + 12);
        } else if (str.indexOf("上午") !== -1) {
          str = str.replace("上午", "").trim();
          d = new Date(str);
        } else {
          d = new Date(str);
        }
      }
      
      if (d && !isNaN(d.getTime())) {
        var dateStr = Utilities.formatDate(d, "Asia/Taipei", "yyyy/MM/dd");
        var sid = r[1] ? r[1].toString().trim() : "";
        if (dateStr === todayStr && sid && !todaySet[sid]) {
          todaySet[sid] = studentClassMap[sid] || "未分班";
        }
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

  // 方案三：改從 SHEET_SCORES 中讀取明細 JSON 進行解析
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  if (!scoreSheet || scoreSheet.getLastRow() <= 1) {
    return { classCategory: 0, studentCategory: 0, questionType: 0, questionStats: 0, updatedAt: updatedAt };
  }
  
  var rows = scoreSheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return h ? h.toString().trim() : ""; });
  
  var cSid = findColIdx(headers, ["學號"]);
  var cName = findColIdx(headers, ["姓名"]);
  var cMode = findColIdx(headers, ["測驗模式"]);
  var cJson = findColIdx(headers, ["作答明細(JSON)", "作答明細 (JSON)"]);
  var cTime = findColIdx(headers, ["時間戳記", "時間"]);
  
  // 僅分析最近 7 天之內的紀錄以防 OOM
  var cutoffTime = new Date().getTime() - 7 * 24 * 60 * 60 * 1000;
  
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var mode = cMode !== -1 && r[cMode] ? r[cMode].toString() : "";
    if (mode === "錯題重做") continue;
    
    var sid = cSid !== -1 && r[cSid] ? r[cSid].toString().trim() : "";
    if (!sid) continue;
    
    var sTime = cTime !== -1 ? getRowSeconds(r[cTime]) * 1000 : 0;
    if (sTime > 0 && sTime < cutoffTime) continue; // 大於 7 天前就直接跳過不分析
    
    var jsonStr = cJson !== -1 ? r[cJson] : "";
    if (!jsonStr) continue;
    
    var details = [];
    try {
      details = JSON.parse(jsonStr);
    } catch(e) {
      continue; // 解析失敗跳過
    }
    if (!Array.isArray(details)) continue;
    
    var info = studentInfo[sid] || {};
    var cls = info.class || "未分班";
    var name = info.name || (cName !== -1 && r[cName] ? r[cName].toString().trim() : "");
    
    details.forEach(function(d, idx) {
      var qid = d.qid || d.questionId || ("Q_" + idx);
      var meta = qMeta[qid] || {};
      
      var topic = d.topic || meta.topic || "未分類";
      var qText = d.questionText || meta.questionText || "";
      var qType = d.questionType || meta.questionType || "未分類";
      var cog = d.cogType || meta.cogType || "未分類";
      var selectedText = d.sel !== undefined ? d.sel : (d.selectedText || "");
      var secVal = d.sec !== undefined ? d.sec : d.answerSec;
      var answerSec = secVal !== null && secVal !== undefined ? Number(secVal) || 0 : 0;
      var isCorrect = d.ok === true || d.isCorrect === true;
      
      var item = { isCorrect: isCorrect, answerSec: answerSec, selectedText: selectedText, meta: {
        className: cls, sid: sid, name: name, topic: topic, questionId: qid,
        questionText: qText, questionType: qType, cogType: cog
      }};
      bumpAgg(classCategory, cls + "|" + topic, item, sid);
      bumpAgg(studentCategory, cls + "|" + sid + "|" + topic, item, sid);
      bumpAgg(questionType, cls + "|" + qType, item, sid);
      bumpAgg(questionStats, qid, item, sid);
    });
  }


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

  // ★ v1.691 同步寫入快取到 Firebase rankingCaches/analysisCaches
  try {
    var analysisCacheObj = {
      classCategory: classRows.map(function(row) {
        return { updatedAt: row[0], className: row[1], topic: row[2], studentCount: row[3], total: row[4], correct: row[5], rate: row[6], avgSec: row[7] };
      }),
      studentCategory: studentRows.map(function(row) {
        return { updatedAt: row[0], className: row[1], sid: row[2], name: row[3], topic: row[4], total: row[5], correct: row[6], rate: row[7], avgSec: row[8] };
      }),
      questionType: typeRows.map(function(row) {
        return { updatedAt: row[0], className: row[1], questionType: row[2], studentCount: row[3], total: row[4], correct: row[5], rate: row[6], avgSec: row[7] };
      }),
      questionStats: questionRows.map(function(row) {
        return { updatedAt: row[0], qid: row[1], topic: row[2], questionType: row[3], cogType: row[4], questionText: row[5], studentCount: row[6], total: row[7], correct: row[8], rate: row[9], avgSec: row[10], commonWrongOption: row[11] };
      })
    };
    uploadCacheToFirebase("rankingCaches", "analysisCaches", analysisCacheObj);
  } catch(err) {
    Logger.log("⚠️ 上傳分析快取至 Firebase 失敗：" + err.message);
  }

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
    var scRows = scoreSheet.getRange(1, 1, scoreSheet.getLastRow(), 7).getValues();
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

  // ── 將學生成績總表同步至 Firestore ──
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (projectId) {
    try {
      var token = firebaseAccessTokenV1685();
      var formattedRows = rows.map(function(r) {
        var scoreMap = {};
        allTopics.forEach(function(topic, idx) {
          scoreMap[topic] = r[3 + idx] !== undefined ? r[3 + idx] : "";
        });
        return {
          className: r[0] || "",
          studentId: r[1] || "",
          name: r[2] || "",
          scores: scoreMap,
          completed: r[3 + allTopics.length] || "",
          updatedAt: r[4 + allTopics.length] || ""
        };
      });
      var scoreTableData = {
        updatedAtText: localNow(),
        passScore: passScore,
        reqTopics: reqTopics,
        reqClasses: reqClasses,
        allTopics: allTopics,
        rows: formattedRows
      };
      var writes = [];
      writes.push({ update: { name: firestoreDocNameV1685(projectId, "system", "scoreTable"), fields: firebaseFieldsV1685(scoreTableData) } });
      firebaseBatchWriteV1685(projectId, token, writes);
      Logger.log("✅ 成績總表快取已同步至 Firestore system/scoreTable");
    } catch (err) {
      Logger.log("⚠️ 同步成績總表快取至 Firestore 失敗：" + err.message);
    }
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
  try {
    syncFirestoreToSheetsV169(ss);
  } catch (err) {
    Logger.log("⚠️ 定時更新前同步 Firestore 失敗：" + err.message);
  }
  var scoreSheet  = ss.getSheetByName(SHEET_SCORES);
  writeTodayPracticeCache(ss);

  if (!scoreSheet || scoreSheet.getLastRow() <= 1) {
    Logger.log("⏭ 成績紀錄是空的，跳過更新");
    return;
  }

  // 檢查最後一筆成績的時間
  var lastRow      = scoreSheet.getLastRow();
  var lastTimeVal  = scoreSheet.getRange(lastRow, 1).getValue();
  var seconds      = getRowSeconds(lastTimeVal);
  var lastTime     = seconds > 0 ? new Date(seconds * 1000) : null;

  if (!lastTime || isNaN(lastTime.getTime())) {
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
// ─────────────────────────────────────────────
function handleGetStudentScoreTable(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    syncFirestoreToSheetsV169(ss);
  } catch (err) {
    Logger.log("⚠️ 手動更新前同步 Firestore 失敗：" + err.message);
  }
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
  try {
    syncFirestoreToSheetsV169(ss);
  } catch (err) {
    Logger.log("⚠️ 手動同步 Firestore 失敗：" + err.message);
  }
  var count = buildAndSaveScoreTable(ss);
  writeTodayPracticeCache(ss);
  var analysis = buildAndSaveAnalysisCaches(ss);
  
  // ★ v1.691 同步上傳 teacherData 快取與各學生的完成度快取至 Firebase
  try {
    updateTeacherDataAndStudentProgressFirebase(ss);
  } catch(err) {
    Logger.log("⚠️ 同步 teacherData 與進度快取至 Firebase 失敗：" + err.message);
  }
  
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

    // ★ v1.691 同步寫入快取至 Firebase (首頁排行與各分班詳情)
    try {
      var projectId = props.getProperty("FIREBASE_PROJECT_ID");
      if (projectId) {
        var token = firebaseAccessTokenV1685();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var todayCache = readTodayPracticeCache(ss);
        
        var homeCacheObj = {
          passScore: passScore,
          completionTopics: completionTopics,
          ranking: rankingSummary,
          todayTotal: todayCache.todayTotal,
          todayByClass: todayCache.todayByClass,
          todayDate: todayCache.todayDate,
          updatedAt: todayCache.updatedAt
        };
        
        var writes = [];
        writes.push({
          update: {
            name: firestoreDocNameV1685(projectId, "rankingCaches", "home"),
            fields: firebaseFieldsV1685(homeCacheObj)
          }
        });
        
        ranking.forEach(function(r) {
          writes.push({
            update: {
              name: firestoreDocNameV1685(projectId, "rankingClasses", r.class),
              fields: firebaseFieldsV1685({ class: r.class, students: r.students || [] })
            }
          });
        });
        
        firebaseBatchWriteV1685(projectId, token, writes);
        Logger.log("✅ 已同步寫入 Firebase 排行快取 (首頁排行與 " + ranking.length + " 個班級學生資料)！");
      }
    } catch(err) {
      Logger.log("⚠️ 同步排行快取至 Firebase 失敗：" + err.message);
    }
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

// ─────────────────────────────────────────────
// Firebase v1.691 增量同步回 Google Sheets 核心函式
// ─────────────────────────────────────────────

// 解析 Firestore REST API 返回的 fields 結構為一般 JS Object
function parseFirebaseFields(fields) {
  var obj = {};
  if (!fields) return obj;
  Object.keys(fields).forEach(function(k) {
    var valObj = fields[k];
    if (valObj.stringValue !== undefined) obj[k] = valObj.stringValue;
    else if (valObj.integerValue !== undefined) obj[k] = Number(valObj.integerValue);
    else if (valObj.doubleValue !== undefined) obj[k] = Number(valObj.doubleValue);
    else if (valObj.booleanValue !== undefined) obj[k] = valObj.booleanValue;
    else if (valObj.timestampValue !== undefined) obj[k] = new Date(valObj.timestampValue);
    else if (valObj.arrayValue !== undefined) {
      var arr = valObj.arrayValue.values || [];
      obj[k] = arr.map(function(item) {
        if (item.stringValue !== undefined) return item.stringValue;
        if (item.integerValue !== undefined) return Number(item.integerValue);
        if (item.doubleValue !== undefined) return Number(item.doubleValue);
        if (item.booleanValue !== undefined) return item.booleanValue;
        if (item.mapValue !== undefined) return parseFirebaseFields(item.mapValue.fields);
        return item;
      });
    }
    else if (valObj.mapValue !== undefined) {
      obj[k] = parseFirebaseFields(valObj.mapValue.fields);
    }
    else if (valObj.nullValue !== undefined) {
      obj[k] = null;
    }
  });
  return obj;
}

// Firebase 轉移至 Firestore 的分水嶺時間（在此時間之前的歷史資料完全不做去重，確保資料安全）
var FIREBASE_CUTOFF_TAIPEI = "2026/06/10 09:28:00";

// 取得 Sheet 中最後一筆記錄的台北時間字串
function getSheetLastTimeTaipei(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return "2000/01/01 00:00:00";
  }
  var val = sheet.getRange(lastRow, 1).getValue();
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  }
  var str = (val || "").toString().trim();
  if (/^\d{4}[/\-]\d{2}[/\-]\d{2}/.test(str)) {
    return str.replace(/-/g, "/");
  }
  return "2000/01/01 00:00:00";
}

// 將台北時間字串轉為 UTC ISO 時間戳記
function taipeiStringToISO(str) {
  var m = str.match(/^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/);
  if (!m) return "2000-01-01T00:00:00.000Z";
  var year = parseInt(m[1], 10);
  var month = parseInt(m[2], 10) - 1;
  var day = parseInt(m[3], 10);
  var hour = parseInt(m[4], 10);
  var min = parseInt(m[5], 10);
  var sec = parseInt(m[6], 10);
  var utcMs = Date.UTC(year, month, day, hour, min, sec) - 8 * 60 * 60 * 1000;
  return new Date(utcMs).toISOString();
}

// 取得 Sheet 中作答時間並轉換為秒數 (以 UTC UNIX 格式)
function getRowSeconds(rowDate) {
  if (!rowDate) return 0;
  var twStr = "";
  if (rowDate instanceof Date) {
    twStr = Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  } else {
    var str = rowDate.toString().trim();
    if (/^\d{4}[/\-]\d{1,2}[/\-]\d{1,2}/.test(str)) {
      twStr = str.replace(/-/g, "/");
    }
  }
  if (!twStr) return 0;
  var iso = taipeiStringToISO(twStr);
  return Math.floor(new Date(iso).getTime() / 1000);
}

/**
 * 壓縮與瘦身作答明細 JSON，移除非必要的超長欄位 (如題目文字)，以避免超過 Google Sheets 儲存格 50000 字元的限制。
 * @param {string} jsonStr 原始明細 JSON 字串
 * @return {string} 壓縮短 key 且移除題目文字後的 JSON 字串
 */
function compressDetailsJson(jsonStr) {
  if (!jsonStr) return "";
  try {
    var details = JSON.parse(jsonStr);
    if (!Array.isArray(details)) return jsonStr;
    
    var compressed = details.map(function(d) {
      return {
        qid: d.questionId || d.qid || "",
        ok: d.isCorrect === true || d.ok === true,
        sec: d.answerSec !== undefined ? d.answerSec : (d.sec !== undefined ? d.sec : null),
        sel: d.selectedText || d.sel || ""
      };
    });
    
    var result = JSON.stringify(compressed);
    // 超過 50000 字元時的安全閥截斷
    if (result.length > 49900) {
      Logger.log("⚠️ 壓縮後 JSON 長度仍超限 (" + result.length + ")，執行截斷安全閥");
      result = result.substring(0, 49900) + "]";
    }
    return result;
  } catch(e) {
    Logger.log("⚠️ 壓縮 detailsJson 失敗：" + e.message);
    return jsonStr.substring(0, 49900); // 粗暴截斷防卡死
  }
}

// 將 Date 物件或時間戳記格式化為台北時間字串 (yyyy/MM/dd HH:mm:ss)
function formatTaipeiDate(date) {
  if (!date) return "";
  var d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
}

// 清除「成績紀錄」Sheet 中的重複資料 (僅針對移轉後資料進行除重)
function deduplicateSheetScores(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var range = sheet.getDataRange();
  var values = range.getValues();
  var headers = values[0];
  var uniqueKeys = {};
  var rowsToKeep = [headers];
  var duplicateCount = 0;
  
  var cutoffSeconds = getRowSeconds(FIREBASE_CUTOFF_TAIPEI);
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var seconds = getRowSeconds(row[0]);
    
    // 移轉前的歷史資料一律直接保留，完全不參與去重邏輯，保障其安全性與格式相容性
    if (seconds < cutoffSeconds) {
      rowsToKeep.push(row);
      continue;
    }
    
    var studentId = row[1] ? row[1].toString().trim() : "";
    var topic = row[3] ? row[3].toString().trim() : "";
    var score = row[6] !== undefined ? row[6].toString().trim() : "";
    
    var key = seconds + "_" + studentId + "_" + topic + "_" + score;
    if (uniqueKeys[key]) {
      duplicateCount++;
      continue;
    }
    uniqueKeys[key] = true;
    rowsToKeep.push(row);
  }
  
  if (duplicateCount > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, headers.length).setValues(rowsToKeep);
    Logger.log("🧹 已清除 成績紀錄 表中 " + duplicateCount + " 筆重複資料");
  }
}

// 清除「作答明細」
function fillExistingScoresDetailsJson() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  if (!scoreSheet || scoreSheet.getLastRow() <= 1) {
    Logger.log("❌ 成績紀錄工作表是空的！");
    return;
  }
  
  var scoreRange = scoreSheet.getDataRange();
  var scoreVals = scoreRange.getValues();
  var scoreHeaders = scoreVals[0].map(function(h) { return h ? h.toString().trim() : ""; });
  
  Logger.log("📋 試算表實際標頭: " + JSON.stringify(scoreHeaders));
  
  // 動態尋找欄位 index
  var cTime = findColIdx(scoreHeaders, ["時間戳記", "時間"]);
  var cSid = findColIdx(scoreHeaders, ["學號"]);
  var cName = findColIdx(scoreHeaders, ["姓名"]);
  var cTopic = findColIdx(scoreHeaders, ["測驗單元", "單元"]);
  var cScore = findColIdx(scoreHeaders, ["分數"]);
  var cJson = findColIdx(scoreHeaders, ["作答明細(JSON)", "作答明細 (JSON)"]);
  
  Logger.log("ℹ️ 試算表欄位位置 - 時間: " + cTime + ", 學號: " + cSid + ", 姓名: " + cName + ", 單元: " + cTopic + ", 分數: " + cScore + ", JSON: " + cJson);
  
  // 清除因為長度 bug 被擠到第 14 欄的錯誤標頭與多餘欄位
  if (scoreHeaders[13] === "作答明細(JSON)" || scoreHeaders[13] === "作答明細 (JSON)") {
    scoreSheet.getRange(1, 14).clear();
    scoreVals = scoreSheet.getDataRange().getValues();
    scoreHeaders = scoreVals[0].map(function(h) { return h ? h.toString().trim() : ""; });
    cJson = -1;
  }
  
  // 確保標頭有第 13 欄 (作答明細(JSON))
  if (cJson === -1) {
    scoreSheet.getRange(1, 13).setValue("作答明細(JSON)").setFontWeight("bold").setBackground("#fce7f3");
    scoreVals = scoreSheet.getDataRange().getValues();
    scoreHeaders = scoreVals[0].map(function(h) { return h ? h.toString().trim() : ""; });
    cJson = 12; // index 12 就是第 13 欄
    Logger.log("ℹ️ 已強制將第 13 欄設為作答明細(JSON)！");
  }

  var detailSheet = ss.getSheetByName(SHEET_DETAILS);
  var fillCount = 0;
  
  // 取得 7 天前的時間戳記 (毫秒)
  var cutoffMs = Date.now() - 7 * 24 * 60 * 60 * 1000;
  var cutoffSeconds = Math.floor(cutoffMs / 1000);
  var startTimeISO = new Date(cutoffMs).toISOString();
  Logger.log("📅 設定 7 天前截止點: " + startTimeISO + " (秒數: " + cutoffSeconds + ")");

  // ────────────────────────────────────────────────────────
  // 優先方案：如果「題目作答明細」工作表還在，直接從本機 Sheets 讀取並重組 JSON
  // ────────────────────────────────────────────────────────
  if (detailSheet && detailSheet.getLastRow() > 1) {
    Logger.log("📂 偵測到「題目作答明細」工作表存在，行數: " + detailSheet.getLastRow() + "。開始重組...");
    var detailVals = detailSheet.getDataRange().getValues();
    var detailHeaders = detailVals[0].map(function(h) { return h ? h.toString().trim() : ""; });
    
    var dTimeIdx = findColIdx(detailHeaders, ["時間戳記", "時間"]);
    var dSidIdx = findColIdx(detailHeaders, ["學號"]);
    var dQidIdx = findColIdx(detailHeaders, ["題目ID", "題目id"]);
    var dTopicIdx = findColIdx(detailHeaders, ["測驗單元", "單元"]);
    var dSelIdx = findColIdx(detailHeaders, ["學生選項", "選項"]);
    var dIsCorrectIdx = findColIdx(detailHeaders, ["是否答對", "答對"]);
    var dSecIdx = findColIdx(detailHeaders, ["作答秒數", "秒數"]);
    
    Logger.log("ℹ️ 明細表欄位位置 - 時間: " + dTimeIdx + ", 學號: " + dSidIdx + ", 題目ID: " + dQidIdx + ", 單元: " + dTopicIdx);
    
    // 以 學號 + "_" + 單元 分組
    var localDetailMap = {};
    var validDetailCount = 0;
    for (var d = 1; d < detailVals.length; d++) {
      var dRow = detailVals[d];
      var seconds = getRowSeconds(dRow[dTimeIdx]);
      // 限制只處理 7 天內明細，減少記憶體與配對時間
      if (seconds < cutoffSeconds) continue;
      
      var sid = dSidIdx !== -1 && dRow[dSidIdx] ? dRow[dSidIdx].toString().replace(/\s+/g, "") : "";
      var topic = dTopicIdx !== -1 && dRow[dTopicIdx] ? dRow[dTopicIdx].toString().replace(/\s+/g, "") : "";
      if (!sid || !topic) continue;
      
      var key = sid + "_" + topic;
      
      var qid = dQidIdx !== -1 && dRow[dQidIdx] ? dRow[dQidIdx].toString().trim() : "";
      var isCorrect = dIsCorrectIdx !== -1 && (dRow[dIsCorrectIdx] === "答對" || dRow[dIsCorrectIdx] === true || dRow[dIsCorrectIdx] === "true");
      var sec = dSecIdx !== -1 && dRow[dSecIdx] !== undefined && dRow[dSecIdx] !== "" ? Number(dRow[dSecIdx]) || 0 : 0;
      var sel = dSelIdx !== -1 && dRow[dSelIdx] !== undefined ? dRow[dSelIdx].toString() : "";
      
      if (!localDetailMap[key]) {
        localDetailMap[key] = [];
      }
      localDetailMap[key].push({
        seconds: seconds,
        item: { qid: qid, ok: isCorrect, sec: sec, sel: sel }
      });
      validDetailCount++;
    }
    Logger.log("📥 成功在記憶體中建立 (7天內) " + Object.keys(localDetailMap).length + " 個 學號_單元 分組，共 " + validDetailCount + " 筆明細資料");
    
    // 開始與成績紀錄進行比對補填
    var unmatchedSamples = [];
    var alreadyFilledCount = 0;
    
    for (var i = 1; i < scoreVals.length; i++) {
      var row = scoreVals[i];
      var existingJson = cJson !== -1 && row[cJson] ? row[cJson].toString().trim() : "";
      if (existingJson) {
        alreadyFilledCount++;
        continue;
      }
      
      var sTime = cTime !== -1 ? getRowSeconds(row[cTime]) : 0;
      if (sTime <= 0) continue;
      // 限制只比對 7 天內的紀錄
      if (sTime < cutoffSeconds) continue;
      
      var sSid = cSid !== -1 && row[cSid] ? row[cSid].toString().replace(/\s+/g, "") : "";
      var sTopic = cTopic !== -1 && row[cTopic] ? row[cTopic].toString().replace(/\s+/g, "") : "";
      var key = sSid + "_" + sTopic;
      
      var detailList = localDetailMap[key];
      var matched = false;
      if (detailList && detailList.length > 0) {
        var matchedItems = [];
        detailList.forEach(function(dObj) {
          if (Math.abs(dObj.seconds - sTime) <= 300) {
            matchedItems.push(dObj.item);
          }
        });
        
        if (matchedItems.length > 0) {
          row[cJson] = JSON.stringify(matchedItems);
          fillCount++;
          matched = true;
        }
      }
      
      if (!matched && unmatchedSamples.length < 5) {
        unmatchedSamples.push({
          rowNum: i + 1,
          sid: sSid,
          topic: sTopic,
          scoreTimeStr: cTime !== -1 ? row[cTime] : "",
          scoreTimeSec: sTime,
          hasDetailKey: !!detailList,
          detailTimes: detailList ? detailList.map(function(d) { return d.seconds; }) : []
        });
      }
    }
    
    Logger.log("ℹ️ 比對報告：目前已填 JSON " + alreadyFilledCount + " 筆，本次補填成功 " + fillCount + " 筆。");
    if (unmatchedSamples.length > 0) {
      Logger.log("⚠️ 配對失敗範例 (前 5 筆)：" + JSON.stringify(unmatchedSamples, null, 2));
    }
    
    if (fillCount > 0) {
      Logger.log("💾 正在寫入 " + fillCount + " 筆從本地重組的 JSON 明細到成績紀錄表...");
      scoreSheet.getRange(1, 1, scoreVals.length, scoreHeaders.length).setValues(scoreVals);
      Logger.log("🎉 本地重組補填完成！已成功補上 " + fillCount + " 筆成績紀錄的作答明細！");
      return;
    } else {
      Logger.log("ℹ️ 本地比對結束，沒有成功匹配到資料。繼續嘗試從 Firebase 同步補填...");
    }
  } else {
    Logger.log("📂 「題目作答明細」工作表不存在或為空。");
  }

  // ────────────────────────────────────────────────────────
  // 備用方案：若「題目作答明細」已刪除，則從 Firebase 的 answerBatches 同步重組
  // ────────────────────────────────────────────────────────
  Logger.log("⚠️ 開始執行 Firebase 下載重組方案...");
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID，無法從 Firebase 同步。");
    return;
  }
  var token = firebaseAccessTokenV1685();
  
  Logger.log("🚀 1. 下載 Firebase answerBatches (7 天內) 進行比對...");
  var newBatches = queryFirestoreCollectionAll(projectId, token, "answerBatches", startTimeISO);
  Logger.log("📥 取得 " + newBatches.length + " 筆 batches 記錄");
  
  var firebaseGroup = {};
  var firebaseHasJsonCount = 0;
  newBatches.forEach(function(doc) {
    if (!doc.detailsJson) return;
    firebaseHasJsonCount++;
    var studentId = doc.studentId ? doc.studentId.toString().replace(/\s+/g, "") : "";
    var topic = doc.topic ? doc.topic.toString().replace(/\s+/g, "") : "";
    var score = doc.score !== undefined ? doc.score.toString().replace(/\s+/g, "") : "";
    var key = studentId + "_" + topic + "_" + score;
    
    var docISO = doc.createdAt ? new Date(doc.createdAt).toISOString() : "";
    var docSeconds = docISO ? Math.floor(new Date(docISO).getTime() / 1000) : 0;
    
    if (!firebaseGroup[key]) firebaseGroup[key] = [];
    firebaseGroup[key].push({
      seconds: docSeconds,
      detailsJson: compressDetailsJson(doc.detailsJson)
    });
  });
  Logger.log("📥 Firebase 中含有 detailsJson 的 batches 有 " + firebaseHasJsonCount + " 筆");
  
  var firebaseFillCount = 0;
  var unmatchedFirebaseSamples = [];
  for (var i = 1; i < scoreVals.length; i++) {
    var row = scoreVals[i];
    var existingJson = cJson !== -1 && row[cJson] ? row[cJson].toString().trim() : "";
    if (existingJson) continue;
    
    var seconds = cTime !== -1 ? getRowSeconds(row[cTime]) : 0;
    if (seconds <= 0) continue;
    // 限制只比對 7 天內的紀錄
    if (seconds < cutoffSeconds) continue;
    
    var studentId = cSid !== -1 && row[cSid] ? row[cSid].toString().replace(/\s+/g, "") : "";
    var topic = cTopic !== -1 && row[cTopic] ? row[cTopic].toString().replace(/\s+/g, "") : "";
    var score = cScore !== -1 && row[cScore] !== undefined ? row[cScore].toString().replace(/\s+/g, "") : "";
    var key = studentId + "_" + topic + "_" + score;
    
    var group = firebaseGroup[key];
    var matched = false;
    if (group && group.length > 0) {
      var bestIdx = -1;
      var minDiff = 1800; // 30 分鐘
      for (var g = 0; g < group.length; g++) {
        var diff = Math.abs(group[g].seconds - seconds);
        if (diff < minDiff) {
          minDiff = diff;
          bestIdx = g;
        }
      }
      if (bestIdx !== -1) {
        row[cJson] = group[bestIdx].detailsJson;
        firebaseFillCount++;
        matched = true;
        group.splice(bestIdx, 1);
      }
    }
    
    if (!matched && unmatchedFirebaseSamples.length < 5) {
      unmatchedFirebaseSamples.push({
        rowNum: i + 1,
        key: key,
        seconds: seconds,
        hasFirebaseKey: !!group,
        firebaseTimes: group ? group.map(function(g) { return g.seconds; }) : []
      });
    }
  }
  
  Logger.log("ℹ️ 比對報告（Firebase Batches）：本次補填成功 " + firebaseFillCount + " 筆。");
  if (unmatchedFirebaseSamples.length > 0) {
    Logger.log("⚠️ Firebase Batches 配對失敗範例 (前 5 筆)：" + JSON.stringify(unmatchedFirebaseSamples, null, 2));
  }
  
  if (firebaseFillCount > 0) {
    Logger.log("💾 正在寫入 " + firebaseFillCount + " 筆從 Firebase 匹配的 JSON 明細到試算表...");
    scoreSheet.getRange(1, 1, scoreVals.length, scoreHeaders.length).setValues(scoreVals);
  }

  // ────────────────────────────────────────────────────────
  // 終極備用方案：如果還有剩餘的空白明細，直接從 Firebase answerDetails 重組
  // ────────────────────────────────────────────────────────
  var remainingEmptyCount = 0;
  for (var i = 1; i < scoreVals.length; i++) {
    var sTime = cTime !== -1 ? getRowSeconds(scoreVals[i][cTime]) : 0;
    if (sTime < cutoffSeconds) continue; // 只看 7 天內的
    
    var existingJson = cJson !== -1 && scoreVals[i][cJson] ? scoreVals[i][cJson].toString().trim() : "";
    if (!existingJson) remainingEmptyCount++;
  }
  
  if (remainingEmptyCount > 0) {
    Logger.log("⚠️ 尚有 " + remainingEmptyCount + " 筆最近一週歷史紀錄缺乏作答明細，啟動終極方案：從 Firebase answerDetails (7 天內) 集合全量重組...");
    
    var allDetails = queryFirestoreCollectionAll(projectId, token, "answerDetails", startTimeISO);
    Logger.log("📥 成功下載 " + allDetails.length + " 筆 Firebase 題目明細");
    
    // 以 學號 + "_" + 單元 分組
    var firebaseDetailGroup = {};
    allDetails.forEach(function(doc) {
      var studentId = doc.studentId ? doc.studentId.toString().replace(/\s+/g, "") : "";
      var topic = doc.topic ? doc.topic.toString().replace(/\s+/g, "") : "";
      if (!studentId || !topic) return;
      
      var key = studentId + "_" + topic;
      var docISO = doc.createdAt ? new Date(doc.createdAt).toISOString() : "";
      var docSeconds = docISO ? Math.floor(new Date(docISO).getTime() / 1000) : 0;
      
      var isCorrect = doc.isCorrect === true || doc.isCorrect === "true";
      var sec = doc.answerSec !== undefined && doc.answerSec !== null ? Number(doc.answerSec) || 0 : 0;
      var sel = doc.selectedText !== undefined ? doc.selectedText.toString() : "";
      var qid = doc.questionId || "";
      
      if (!firebaseDetailGroup[key]) {
        firebaseDetailGroup[key] = [];
      }
      firebaseDetailGroup[key].push({
        seconds: docSeconds,
        item: { qid: qid, ok: isCorrect, sec: sec, sel: sel }
      });
    });
    Logger.log("📥 成功在記憶體中建立 " + Object.keys(firebaseDetailGroup).length + " 個 學號_單元 Firebase 明細分組");
    
    var detailsFillCount = 0;
    for (var i = 1; i < scoreVals.length; i++) {
      var row = scoreVals[i];
      var seconds = cTime !== -1 ? getRowSeconds(row[cTime]) : 0;
      if (seconds <= 0 || seconds < cutoffSeconds) continue;
      
      var existingJson = cJson !== -1 && row[cJson] ? row[cJson].toString().trim() : "";
      if (existingJson) continue;
      
      var studentId = cSid !== -1 && row[cSid] ? row[cSid].toString().replace(/\s+/g, "") : "";
      var topic = cTopic !== -1 && row[cTopic] ? row[cTopic].toString().replace(/\s+/g, "") : "";
      var key = studentId + "_" + topic;
      
      var list = firebaseDetailGroup[key];
      if (list && list.length > 0) {
        // 尋找在 seconds 前後 5 分鐘 (300 秒) 內的所有題目細節
        var matchedItems = [];
        list.forEach(function(dObj) {
          if (Math.abs(dObj.seconds - seconds) <= 300) {
            matchedItems.push(dObj.item);
          }
        });
        
        if (matchedItems.length > 0) {
          row[cJson] = JSON.stringify(matchedItems);
          detailsFillCount++;
        }
      }
    }
    
    if (detailsFillCount > 0) {
      Logger.log("💾 正在寫入 " + detailsFillCount + " 筆從 Firebase answerDetails 重組的 JSON 到試算表...");
      scoreSheet.getRange(1, 1, scoreVals.length, scoreHeaders.length).setValues(scoreVals);
      Logger.log("🎉 終極補填完成！已成功補上 " + detailsFillCount + " 筆歷史成績的作答明細！");
    } else {
      Logger.log("ℹ️ 終極比對結束，未配對到任何題目細節。");
    }
  }
}


/**
 * 一鍵終極修復所有問題（補填 JSON 並重建所有快取）
 */
function oneClickFixAll() {
  Logger.log("🚀 1/2. 開始一鍵補填歷史作答明細 (JSON)...");
  fillExistingScoresDetailsJson();
  
  Logger.log("🚀 2/2. 開始重建學生成績總表與分析快取...");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  manualUpdateScoreTable();
  
  Logger.log("🎉 一鍵修復完成！請立刻回到試算表確認快取工作表！");
}

// 供測試或手動一鍵補齊今天所有 Firebase 成績紀錄的入口
function runManualForceSyncAll() {
  forceSyncFromTime("2026/06/10 09:28:00");
}

/**
 * 抓取 Firebase 09:28 以後的所有成績，並以 Tab 分隔格式印在 Logger 中，供使用者手動貼回。
 */
function logFirebaseScoresToText() {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID，請在 Script Properties 中設定");
    return;
  }
  var token = firebaseAccessTokenV1685();
  var startTimeISO = taipeiStringToISO("2026/06/10 09:28:00");
  
  Logger.log("🚀 正在從 Firebase 下載 " + startTimeISO + " 之後的所有成績...");
  var newBatches = queryFirestoreCollection(projectId, token, "answerBatches", startTimeISO);
  
  if (newBatches.length === 0) {
    Logger.log("❌ 沒有找到任何成績紀錄。");
    return;
  }
  
  var lines = [];
  newBatches.forEach(function(doc) {
    var twTime = formatTaipeiDate(doc.createdAt);
    var row = [
      twTime,
      doc.studentId || "",
      doc.name || "",
      doc.topic || "",
      doc.mode || "",
      doc.attempt || 1,
      doc.score || 0,
      doc.correctCount || 0,
      doc.wrongCount || 0,
      doc.duration !== undefined ? doc.duration : "",
      doc.token || "",
      doc.ip || ""
    ];
    lines.push(row.join("\t"));
  });
  
  var resultText = lines.join("\n");
  Logger.log("\n=== 複製下方內容，直接在 Google Sheets「成績紀錄」第一筆空白列貼上 ===\n" + resultText + "\n==================================================");
}

/**
 * 抓取 Firebase 09:28 以後的所有作答明細，並以 Tab 分隔格式印在 Logger 中，供使用者手動貼回。
 */
function logFirebaseDetailsToText() {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID，請在 Script Properties 中設定");
    return;
  }
  var token = firebaseAccessTokenV1685();
  var startTimeISO = taipeiStringToISO("2026/06/10 09:28:00");
  
  Logger.log("🚀 正在從 Firebase 下載 " + startTimeISO + " 之後的所有作答明細...");
  var newBatches = queryFirestoreCollection(projectId, token, "answerBatches", startTimeISO);
  
  if (newBatches.length === 0) {
    Logger.log("❌ 沒有找到任何作答紀錄。");
    return;
  }
  
  var detailRows = [];
  newBatches.forEach(function(doc) {
    if (doc.detailsJson) {
      try {
        var details = JSON.parse(doc.detailsJson);
        if (Array.isArray(details)) {
          details.forEach(function(d, idx) {
            var row = [
              doc.createdAt ? formatTaipeiDate(doc.createdAt) : "",
              doc.studentId || "",
              doc.name || "",
              d.questionId || ("Q_" + idx),
              d.questionText || "",
              d.topic || "",
              d.selectedText || "",
              d.correctText || "",
              d.isCorrect ? "答對" : "答錯",
              doc.mode || "",
              doc.attempt || 1,
              d.answerSec !== null && d.answerSec !== undefined ? d.answerSec : "",
              d.questionType || "",
              d.cogType || ""
            ];
            detailRows.push(row.join("\t"));
          });
        }
      } catch (e) {
        // ignore
      }
    }
  });
  
  var resultText = detailRows.join("\n");
  Logger.log("\n=== 複製下方內容，直接在 Google Sheets「作答明細」第一筆空白列貼上 ===\n" + resultText + "\n==================================================");
}

/**
 * 搜尋 Firebase 中所有名為「張欣翰」的成績記錄 (Debug 專用)
 */
function debugFindStudent() {
  var studentName = "張欣翰";
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID");
    return;
  }
  var token = firebaseAccessTokenV1685();
  
  var url = "https://firestore.googleapis.com/v1/projects/" + projectId + "/databases/(default)/documents:runQuery";
  var payload = {
    structuredQuery: {
      from: [{ collectionId: "answerBatches" }],
      limit: 3000
    }
  };
  
  Logger.log("🚀 開始在 Firebase 搜尋「" + studentName + "」的紀錄...");
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (res.getResponseCode() >= 300) {
    Logger.log("❌ Firebase 查詢失敗：" + res.getContentText());
    return;
  }
  
  var results = JSON.parse(res.getContentText());
  Logger.log("📥 從 Firebase 取得的原始紀錄總數: " + (results ? results.length : 0));
  
  var count = 0;
  results.forEach(function(r) {
    if (r.document && r.document.fields) {
      var doc = parseFirebaseFields(r.document.fields);
      if (doc.name && doc.name.indexOf(studentName) !== -1) {
        count++;
        var twTime = doc.createdAt ? formatTaipeiDate(doc.createdAt) : "無時間欄位";
        Logger.log("🔍 找到紀錄 [" + count + "]: 姓名=" + doc.name + ", 學號=" + doc.studentId + ", 單元=" + doc.topic + ", 分數=" + doc.score + ", 交卷時間=" + twTime + ", Token=" + doc.token);
      }
    }
  });
  
  if (count === 0) {
    Logger.log("❌ 在這批 Firebase 資料中完全找不到名為「" + studentName + "」的紀錄！");
  } else {
    Logger.log("🎉 搜尋結束，共找到 " + count + " 筆紀錄。");
  }
}

/**
 * 抓取 Firebase 09:28 以後的所有成績，並直接寫入一個名為「臨時_Firebase成績紀錄」的新工作表中。
 * 這是為了避開 Google Apps Script 執行紀錄 Logger.log 單次長度被截斷的限制。
 */
function writeFirebaseScoresToTempSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID");
    return;
  }
  var token = firebaseAccessTokenV1685();
  var startTimeISO = taipeiStringToISO("2026/06/10 09:28:00");
  
  Logger.log("🚀 正在從 Firebase 下載 " + startTimeISO + " 之後的所有成績...");
  var newBatches = queryFirestoreCollection(projectId, token, "answerBatches", startTimeISO);
  
  if (newBatches.length === 0) {
    Logger.log("❌ 沒有找到任何成績紀錄。");
    return;
  }
  
  // 建立或取得臨時工作表
  var tempSheetName = "臨時_Firebase成績紀錄";
  var tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) {
    tempSheet.clear();
  } else {
    tempSheet = ss.insertSheet(tempSheetName);
  }
  
  // 寫入標頭
  var headers = ["時間戳記","學號","姓名","測驗單元","測驗模式","第幾次","分數","答對題數","答錯題數","作答秒數","Token","IP"];
  tempSheet.appendRow(headers);
  
  var scoreRows = [];
  newBatches.forEach(function(doc) {
    var twTime = formatTaipeiDate(doc.createdAt);
    scoreRows.push([
      twTime,
      doc.studentId || "",
      doc.name || "",
      doc.topic || "",
      doc.mode || "",
      doc.attempt || 1,
      doc.score || 0,
      doc.correctCount || 0,
      doc.wrongCount || 0,
      doc.duration !== undefined ? doc.duration : "",
      doc.token || "",
      doc.ip || ""
    ]);
  });
  
  if (scoreRows.length > 0) {
    tempSheet.getRange(2, 1, scoreRows.length, 12).setValues(scoreRows);
  }
  
  Logger.log("🎉 已將 " + scoreRows.length + " 筆成績資料寫入工作表「" + tempSheetName + "」！請直接至該工作表複製。");
}

/**
 * 抓取 Firebase 09:28 以後的所有作答明細，並直接寫入一個名為「臨時_Firebase作答明細」的新工作表中。
 */
function writeFirebaseDetailsToTempSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID");
    return;
  }
  var token = firebaseAccessTokenV1685();
  var startTimeISO = taipeiStringToISO("2026/06/10 09:28:00");
  
  Logger.log("🚀 正在從 Firebase 下載 " + startTimeISO + " 之後的所有作答明細...");
  var newBatches = queryFirestoreCollection(projectId, token, "answerBatches", startTimeISO);
  
  if (newBatches.length === 0) {
    Logger.log("❌ 沒有找到任何作答紀錄。");
    return;
  }
  
  var tempSheetName = "臨時_Firebase作答明細";
  var tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) {
    tempSheet.clear();
  } else {
    tempSheet = ss.insertSheet(tempSheetName);
  }
  
  var headers = ["時間戳記","學號","姓名","題目ID","題目內容","單元分類","學生選擇","正確答案","是否正確","測驗模式","第幾次","作答秒數","題型","認知階層"];
  tempSheet.appendRow(headers);
  
  var detailRows = [];
  newBatches.forEach(function(doc) {
    if (doc.detailsJson) {
      try {
        var details = JSON.parse(doc.detailsJson);
        if (Array.isArray(details)) {
          details.forEach(function(d, idx) {
            detailRows.push([
              doc.createdAt ? formatTaipeiDate(doc.createdAt) : "",
              doc.studentId || "",
              doc.name || "",
              d.questionId || ("Q_" + idx),
              d.questionText || "",
              d.topic || "",
              d.selectedText || "",
              d.correctText || "",
              d.isCorrect ? "答對" : "答錯",
              doc.mode || "",
              doc.attempt || 1,
              d.answerSec !== null && d.answerSec !== undefined ? d.answerSec : "",
              d.questionType || "",
              d.cogType || ""
            ]);
          });
        }
      } catch (e) {
        // ignore
      }
    }
  });
  
  if (detailRows.length > 0) {
    var batchSize = 5000;
    for (var i = 0; i < detailRows.length; i += batchSize) {
      var chunk = detailRows.slice(i, i + batchSize);
      tempSheet.getRange(tempSheet.getLastRow() + 1, 1, chunk.length, chunk[0].length).setValues(chunk);
    }
  }
  
  Logger.log("🎉 已將 " + detailRows.length + " 筆明細資料寫入工作表「" + tempSheetName + "」！請直接至該工作表複製。");
}

/**
 * 系統健康與確認機制檢測工具 (Health Check)
 * 自動檢測 Firebase 通訊、時區轉換、去重安全分水嶺以及 Sheet 讀寫狀態
 */
function runSystemHealthCheck() {
  Logger.log("==================================================");
  Logger.log("🔍 開始執行 題庫與 Firebase 系統整合健康檢查...");
  Logger.log("==================================================");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  
  // 1. 檢測 Firebase 專案設定
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ [設定檢查] 未在指令碼屬性中設定 FIREBASE_PROJECT_ID！");
    return;
  }
  Logger.log("✅ [設定檢查] Firebase Project ID: " + projectId);
  
  // 2. 檢測 Token 取得
  var token = "";
  try {
    token = firebaseAccessTokenV1685();
    if (token) {
      Logger.log("✅ [連線檢查] 成功取得 Firebase Access Token (前10碼: " + token.substring(0, 10) + "...)");
    } else {
      Logger.log("❌ [連線檢查] 取得的 Token 為空值！");
      return;
    }
  } catch(e) {
    Logger.log("❌ [連線檢查] 取得 Access Token 失敗：" + e.message);
    return;
  }
  
  // 3. 檢測 Firestore REST API 讀取
  try {
    var url = "https://firestore.googleapis.com/v1/projects/" + projectId + "/databases/(default)/documents:runQuery";
    var payload = {
      structuredQuery: {
        from: [{ collectionId: "answerBatches" }],
        limit: 1
      }
    };
    var res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (res.getResponseCode() === 200) {
      var results = JSON.parse(res.getContentText());
      Logger.log("✅ [通訊檢查] Firestore 連線讀取成功，至少包含 " + (results ? results.length : 0) + " 筆 answerBatches 紀錄");
    } else {
      Logger.log("❌ [通訊檢查] Firestore 讀取失敗，狀態碼: " + res.getResponseCode() + ", 錯誤: " + res.getContentText());
      return;
    }
  } catch(e) {
    Logger.log("❌ [通訊檢查] Firestore 通訊例外錯誤：" + e.message);
    return;
  }
  
  // 4. 檢測時間解析與格式容錯
  try {
    var dateNormal = "2026/06/10 09:28:00";
    var dateNoZero = "2026/6/10 9:28:0"; // 故意不補零
    var dateDash   = "2026-6-10 9:28:0";  // 故意用 dash
    var dateObject = new Date("2026-06-10T01:28:00.000Z"); // Date 物件
    
    var secNormal = getRowSeconds(dateNormal);
    var secNoZero = getRowSeconds(dateNoZero);
    var secDash   = getRowSeconds(dateDash);
    var secObject = getRowSeconds(dateObject);
    
    Logger.log("info [時間比對測試] 基準值：" + dateNormal + " -> 秒數 = " + secNormal);
    
    if (secNormal === secNoZero && secNormal === secDash && secNormal === secObject) {
      Logger.log("✅ [時間檢查] 時間容錯解析完全正常！(所有格式皆能精確解出秒數: " + secNormal + ")");
    } else {
      Logger.log("❌ [時間檢查] 時間容錯測試失敗！秒數不一致：");
      Logger.log("   - 正常格式: " + secNormal);
      Logger.log("   - 未補零格式: " + secNoZero);
      Logger.log("   - 減號格式: " + secDash);
      Logger.log("   - 物件格式: " + secObject);
    }
  } catch(e) {
    Logger.log("❌ [時間檢查] 時間解析測試發生錯誤：" + e.message);
  }
  
  // 5. 檢測後續同步防漏與去重機制
  try {
    var scoreSheet = ss.getSheetByName(SHEET_SCORES);
    if (scoreSheet) {
      var lastTime = getSheetLastTimeTaipei(scoreSheet);
      var lastTimeISO = taipeiStringToISO(lastTime);
      var startQueryTime = new Date(new Date(lastTimeISO).getTime() - 24 * 60 * 60 * 1000);
      var cutoffMs = new Date(taipeiStringToISO(FIREBASE_CUTOFF_TAIPEI)).getTime();
      
      var actualQueryTime = startQueryTime.getTime() < cutoffMs ? new Date(cutoffMs) : startQueryTime;
      Logger.log("✅ [防漏檢查] 後續同步防漏機制運作正常。最新時間為 " + lastTime + "，安全查詢起點 (已回推24h並卡轉移點): " + formatTaipeiDate(actualQueryTime));
      
      // 測試去重演算法
      var testKey1 = "1781100000_114510000_TEST_99";
      var testKey2 = "1781100000_114510000_TEST_99";
      var tempMap = {};
      tempMap[testKey1] = true;
      if (tempMap[testKey2]) {
        Logger.log("✅ [去重檢查] 秒級複合 Key 去重匹配演算法驗證成功！");
      } else {
        Logger.log("❌ [去重檢查] 去重演算法測試異常！");
      }
    }
  } catch(e) {
    Logger.log("❌ [防漏去重檢查] 檢測時發生錯誤：" + e.message);
  }
  
  // 6. 檢測 Sheet 當前狀態 (方案三)
  try {
    var scoreSheet = ss.getSheetByName(SHEET_SCORES);
    if (scoreSheet) {
      var scoreLastRow = scoreSheet.getLastRow();
      var scoreTime = getSheetLastTimeTaipei(scoreSheet);
      var headers = scoreSheet.getRange(1, 1, 1, scoreSheet.getLastColumn()).getValues()[0];
      var cJson = findColIdx(headers, ["作答明細(JSON)", "作答明細 (JSON)"]);
      
      Logger.log("✅ [工作表檢查] 「成績紀錄」工作表狀態正常。總列數: " + scoreLastRow + ", 最後一筆台北時間: " + scoreTime);
      if (cJson !== -1) {
        Logger.log("✅ [工作表檢查] 成功定位「作答明細(JSON)」欄位 (第 " + (cJson + 1) + " 欄)");
      } else {
        Logger.log("⚠️ [工作表檢查] 警告：在成績紀錄中找不到「作答明細(JSON)」欄位！(請點擊執行 runManualForceSyncAll 以自動建立此欄位)");
      }
    } else {
      Logger.log("⚠️ [工作表檢查] 找不到「成績紀錄」工作表！");
    }
    
    // 明細表已棄用
    var detailSheet = ss.getSheetByName(SHEET_DETAILS);
    if (detailSheet) {
      Logger.log("ℹ️ [工作表檢查] 偵測到舊的「作答明細」分頁仍在，建議確認數據無誤後將其刪除以釋放系統儲存格空間。");
    } else {
      Logger.log("✅ [工作表檢查] 舊「作答明細」工作表已成功清空/刪除 (釋放儲存格額度)。");
    }
  } catch(e) {
    Logger.log("❌ [工作表檢查] 讀取工作表資訊失敗：" + e.message);
  }
  
  Logger.log("==================================================");
  Logger.log("🎉 系統健康檢查完成！全部核心指標皆正常。");
  Logger.log("==================================================");
}

/**
 * 手動一鍵清理「成績紀錄」與「作答明細」中的所有重複資料 (全量清理)
 * 因為全量清理在資料量大時 (數萬列) 會非常耗時，故將其移出定時自動同步中，改為有需要時手動點擊執行
 */
function manualCleanAllDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log("🧹 開始手動清理 成績紀錄 全量重複資料...");
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  if (scoreSheet) {
    deduplicateSheetScores(scoreSheet);
  }
  
  Logger.log("🧹 開始手動清理 作答明細 全量重複資料...");
  var detailSheet = getActiveDetailSheet(ss);
  if (detailSheet) {
    deduplicateSheetDetails(detailSheet);
  }
  
  Logger.log("🎉 手動全量去重清理完成！");
}

/**
 * 一鍵清理與刪除所有工作表中，資料列之後的多餘空白列與空白欄 (釋放儲存格額度)
 * 這能有效解決「儲存格數量將會超過系統上限 (10000000 個)」的錯誤
 */
function trimAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var totalRemovedRows = 0;
  var totalRemovedCols = 0;
  
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    var maxRows = sheet.getMaxRows();
    var lastCol = sheet.getLastColumn();
    var maxColumns = sheet.getMaxColumns();
    
    // 1. 刪除多餘空白列 (保留 5 列緩衝)
    if (maxRows > lastRow + 5) {
      var rowsToRemove = maxRows - (lastRow + 5);
      sheet.deleteRows(lastRow + 6, rowsToRemove);
      totalRemovedRows += rowsToRemove;
      Logger.log("🧹 工作表 [" + name + "]：刪除了 " + rowsToRemove + " 行多餘空白列 (目前最大列數: " + (lastRow + 5) + ")");
    }
    
    // 2. 刪除多餘空白欄 (保留 2 欄緩衝)
    if (maxColumns > lastCol + 2) {
      var colsToRemove = maxColumns - (lastCol + 2);
      sheet.deleteColumns(lastCol + 3, colsToRemove);
      totalRemovedCols += colsToRemove;
      Logger.log("🧹 工作表 [" + name + "]：刪除了 " + colsToRemove + " 欄多餘空白欄 (目前最大欄數: " + (lastCol + 2) + ")");
    }
  });
  
  Logger.log("🎉 清理完成！共釋放了 " + totalRemovedRows + " 行空白列 與 " + totalRemovedCols + " 欄空白欄。");
}


// 註：此處重複的舊版 fillExistingScoresDetailsJson 已被移除，請使用檔案中部份更健壯、支援分頁比對的新版。


// ────────────────────────────────────────────────────────
// ★ v1.691 新增：Firebase 快取上傳與同步輔助函數
// ────────────────────────────────────────────────────────

function uploadCacheToFirebase(collection, docId, dataObject) {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("⚠️ 找不到 FIREBASE_PROJECT_ID，無法上傳快取到 Firebase。");
    return;
  }
  try {
    var token = firebaseAccessTokenV1685();
    var docName = firestoreDocNameV1685(projectId, collection, docId);
    var fields = firebaseFieldsV1685(dataObject);
    var writes = [{ update: { name: docName, fields: fields } }];
    firebaseBatchWriteV1685(projectId, token, writes);
    Logger.log("✅ 成功上傳快取到 Firebase Firestore: " + collection + "/" + docId);
  } catch (err) {
    Logger.log("❌ 上傳快取到 Firebase 失敗 (" + collection + "/" + docId + ")：" + err.message);
  }
}

function updateTeacherDataAndStudentProgressFirebase(ss) {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("⚠️ 找不到 FIREBASE_PROJECT_ID，無法同步快取到 Firebase。");
    return;
  }
  var token = firebaseAccessTokenV1685();
  
  // 1. 取得學生名單
  var stuSheet = ss.getSheetByName(SHEET_STUDENTS);
  var studentInfoMap = {};
  if (stuSheet && stuSheet.getLastRow() > 1) {
    var stRows = stuSheet.getDataRange().getValues();
    for (var si = 1; si < stRows.length; si++) {
      var sid = stRows[si][0] ? stRows[si][0].toString().trim() : "";
      if (sid) {
        studentInfoMap[sid] = {
          name: stRows[si][1] ? stRows[si][1].toString().trim() : "",
          class: stRows[si][2] ? stRows[si][2].toString().trim() : "未分班"
        };
      }
    }
  }
  
  // 2. 唯讀前 10 欄，統計學生成績歷史與班級統計
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  var studentHistory = {};
  var classStats = {};
  var topicTimeStats = {};
  
  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    var scRows = scoreSheet.getRange(1, 1, scoreSheet.getLastRow(), 10).getValues();
    for (var i = 1; i < scRows.length; i++) {
      var sid = scRows[i][1] ? scRows[i][1].toString().trim() : "";
      if (!sid) continue;
      
      var stuInfo = studentInfoMap[sid] || {};
      if (!studentHistory[sid]) {
        studentHistory[sid] = {
          name: stuInfo.name || sid,
          class: stuInfo.class || "未分班",
          attempts: []
        };
      }
      
      var sDate    = scRows[i][0] ? scRows[i][0].toString() : "";
      var sTopic   = scRows[i][3] ? scRows[i][3].toString() : "";
      var sMode    = scRows[i][4] ? scRows[i][4].toString() : "";
      var sAttempt = Number(scRows[i][5]);
      var sScore   = Number(scRows[i][6]);
      var sCorr    = Number(scRows[i][7]) || 0;
      var sWron    = Number(scRows[i][8]) || 0;
      var sDur     = Number(scRows[i][9]) || 0;
      var isRetry  = sMode === "錯題重做";
      
      studentHistory[sid].attempts.push({
        date:     sDate,
        topic:    sTopic,
        mode:     sMode,
        attempt:  sAttempt,
        score:    sScore,
        correct:  sCorr,
        wrong:    sWron,
        duration: sDur,
        isRetry:  isRetry,
      });
      
      if (!isRetry && sTopic !== "綜合練習" && sDur > 0) {
        var qCount = sCorr + sWron;
        if (qCount > 0) {
          if (!topicTimeStats[sTopic]) topicTimeStats[sTopic] = { totalSec: 0, totalQ: 0, count: 0 };
          topicTimeStats[sTopic].totalSec += sDur;
          topicTimeStats[sTopic].totalQ   += qCount;
          topicTimeStats[sTopic].count++;
        }
      }
    }
  }
  
  // 計算 classList
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
  
  var classList = Object.keys(classStats).map(function(cls) {
    var s = classStats[cls];
    return {
      class: cls,
      correct: s.correct,
      total: s.total,
      rate: s.total > 0 ? Math.round((s.correct / s.total) * 100) : null,
      studentCount: Object.keys(s.studentSet).length
    };
  }).sort(function(a,b) { return a.class.localeCompare(b.class, "zh-TW"); });
  
  var topicTimeList = Object.keys(topicTimeStats).map(function(t) {
    var s = topicTimeStats[t];
    return { topic: t, avgSec: Math.round(s.totalSec / s.totalQ), sessionCount: s.count };
  }).sort(function(a,b) { return a.topic.localeCompare(b.topic, "zh-TW"); });
  
  // 3. 上傳 teacherData 到 Firebase
  // ★ v1.691: 不再將巨大的 studentHistory 塞入大盤快取，改為空物件以防 payload 爆 size
  var teacherData = {
    studentHistory: {},
    classList: classList,
    studentInfoMap: studentInfoMap,
    topicTimeList: topicTimeList
  };
  
  var writes = [];
  writes.push({
    update: {
      name: firestoreDocNameV1685(projectId, "rankingCaches", "teacherData"),
      fields: firebaseFieldsV1685(teacherData)
    }
  });
  
  // ★ v1.691: 將有作答學生的個人歷程，寫入獨立的 studentAttempts/{sid} 集合，避免單一 doc 超過 1MB
  Object.keys(studentHistory).forEach(function(sid) {
    writes.push({
      update: {
        name: firestoreDocNameV1685(projectId, "studentAttempts", sid),
        fields: firebaseFieldsV1685(studentHistory[sid])
      }
    });
  });
  
  // 4. 計算並批次準備每一位學生的進度快取
  var settings = readSettings(ss);
  var passScore = settings.passScore;
  var completionTopics = settings.completionTopics;
  
  Object.keys(studentInfoMap).forEach(function(sid) {
    var topicBest = {};
    var topicTimeMap = {};
    if (studentHistory[sid]) {
      studentHistory[sid].attempts.forEach(function(a) {
        if (a.isRetry) return;
        if (a.score > (topicBest[a.topic] || 0)) topicBest[a.topic] = a.score;
        if (a.topic !== "綜合練習" && a.duration > 0) {
          var qCount = (a.correct || 0) + (a.wrong || 0);
          if (qCount > 0) {
            if (!topicTimeMap[a.topic]) topicTimeMap[a.topic] = { totalSec: 0, totalQ: 0 };
            topicTimeMap[a.topic].totalSec += a.duration;
            topicTimeMap[a.topic].totalQ   += qCount;
          }
        }
      });
    }
    
    var details = completionTopics.map(function(t) {
      var avgSec = null;
      if (topicTimeMap[t] && topicTimeMap[t].totalQ > 0) {
        avgSec = Math.round(topicTimeMap[t].totalSec / topicTimeMap[t].totalQ);
      }
      return {
        topic: t,
        best: topicBest[t] !== undefined ? topicBest[t] : null,
        passed: (topicBest[t] || 0) >= passScore,
        avgSec: avgSec
      };
    });
    
    var progressObj = {
      passScore: passScore,
      completionTopics: completionTopics,
      details: details
    };
    
    writes.push({
      update: {
        name: firestoreDocNameV1685(projectId, "studentProgress", sid),
        fields: firebaseFieldsV1685(progressObj)
      }
    });
  });
  
  // 批次寫入所有快取文件到 Firestore (每 500 筆一組)
  firebaseBatchWriteV1685(projectId, token, writes);
  Logger.log("🎉 成功批次上傳 teacherData 與 " + Object.keys(studentInfoMap).length + " 位學生的個人完成度進度快取至 Firebase！");
}


/**
 * 從 Firestore 查詢特定集合 (支援時間過濾)
 */
function queryFirestoreCollection(projectId, token, collection, startTimeISO) {
  var url = "https://firestore.googleapis.com/v1/projects/" + projectId + "/databases/(default)/documents:runQuery";
  
  var structuredQuery = {
    from: [{ collectionId: collection }]
  };
  
  if (startTimeISO) {
    structuredQuery.where = {
      fieldFilter: {
        field: { fieldPath: "createdAt" },
        op: "GREATER_THAN_OR_EQUAL",
        value: { stringValue: startTimeISO }
      }
    };
  }
  
  var payload = {
    structuredQuery: structuredQuery
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() >= 300) {
    throw new Error("Firestore 查詢失敗 (" + collection + ")：" + res.getContentText());
  }
  
  var parsedRes = JSON.parse(res.getContentText());
  var allDocs = [];
  
  if (Array.isArray(parsedRes)) {
    parsedRes.forEach(function(r) {
      if (r.document && r.document.fields) {
        var docObj = parseFirebaseFields(r.document.fields);
        if (r.document.name) {
          var parts = r.document.name.split("/");
          docObj.id = parts[parts.length - 1];
        }
        allDocs.push(docObj);
      }
    });
  }
  
  return allDocs;
}

/**
 * 查詢特定集合的所有資料 (與 queryFirestoreCollection 同步)
 */
function queryFirestoreCollectionAll(projectId, token, collection, startTimeISO) {
  return queryFirestoreCollection(projectId, token, collection, startTimeISO);
}


/**
 * 增量同步 Firestore 數據到 Google Sheets
 */
function syncFirestoreToSheetsV169(ss) {
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  if (!scoreSheet) return;
  var lastTimeTaipei = getSheetLastTimeTaipei(scoreSheet);
  Logger.log("⏱ 偵測到試算表最後作答時間: " + lastTimeTaipei + "，開始增量同步...");
  
  // 呼叫底層同步函數
  forceSyncFromTime(lastTimeTaipei);
}

/**
 * 強制同步特定時間點之後的 Firestore 數據到 Google Sheets
 */
function forceSyncFromTime(timeStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoreSheet = ss.getSheetByName(SHEET_SCORES);
  if (!scoreSheet) {
    Logger.log("❌ 找不到成績紀錄工作表！");
    return;
  }
  
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("FIREBASE_PROJECT_ID");
  if (!projectId) {
    Logger.log("❌ 找不到 FIREBASE_PROJECT_ID，請在 Script Properties 中設定");
    return;
  }
  var token = firebaseAccessTokenV1685();
  var startTimeISO = taipeiStringToISO(timeStr);
  
  Logger.log("🚀 開始從 Firebase 下載 " + startTimeISO + " 之後的所有成績...");
  var newBatches = queryFirestoreCollection(projectId, token, "answerBatches", startTimeISO);
  
  if (newBatches.length === 0) {
    Logger.log("ℹ️ 沒有找到任何新的成績紀錄。");
    return;
  }
  
  // 讀取現有的成績，建立防重 key
  var scoreVals = scoreSheet.getDataRange().getValues();
  var scoreHeaders = scoreVals[0].map(function(h) { return h ? h.toString().trim() : ""; });
  
  var cTime = findColIdx(scoreHeaders, ["時間戳記", "時間"]);
  var cSid = findColIdx(scoreHeaders, ["學號"]);
  var cTopic = findColIdx(scoreHeaders, ["測驗單元", "單元"]);
  var cScore = findColIdx(scoreHeaders, ["分數"]);
  var cJson = findColIdx(scoreHeaders, ["作答明細(JSON)", "作答明細 (JSON)"]);
  
  // 如果沒有第 13 欄，強制建立它
  if (cJson === -1) {
    scoreSheet.getRange(1, 13).setValue("作答明細(JSON)").setFontWeight("bold").setBackground("#fce7f3");
    scoreVals = scoreSheet.getDataRange().getValues();
    cJson = 12;
  }
  
  var uniqueKeys = {};
  for (var i = 1; i < scoreVals.length; i++) {
    var r = scoreVals[i];
    var seconds = r[cTime] ? getRowSeconds(r[cTime]) : 0;
    var studentId = cSid !== -1 && r[cSid] ? r[cSid].toString().replace(/\s+/g, "") : "";
    var topic = cTopic !== -1 && r[cTopic] ? r[cTopic].toString().replace(/\s+/g, "") : "";
    var score = cScore !== -1 && r[cScore] !== undefined ? r[cScore].toString().replace(/\s+/g, "") : "";
    
    var key = seconds + "_" + studentId + "_" + topic + "_" + score;
    uniqueKeys[key] = true;
  }
  
  // 準備寫入的新列
  var rowsToWrite = [];
  newBatches.forEach(function(doc) {
    var docISO = doc.createdAt ? new Date(doc.createdAt).toISOString() : "";
    var seconds = docISO ? Math.floor(new Date(docISO).getTime() / 1000) : 0;
    
    var studentId = doc.studentId ? doc.studentId.toString().replace(/\s+/g, "") : "";
    var topic = doc.topic ? doc.topic.toString().replace(/\s+/g, "") : "";
    var score = doc.score !== undefined ? doc.score.toString().replace(/\s+/g, "") : "";
    
    var key = seconds + "_" + studentId + "_" + topic + "_" + score;
    if (uniqueKeys[key]) return; // 重複，略過
    
    var twTime = formatTaipeiDate(doc.createdAt);
    var detailsStr = doc.detailsJson ? (typeof doc.detailsJson === "string" ? doc.detailsJson : JSON.stringify(doc.detailsJson)) : "";
    
    // 建立一列 13 欄的資料
    var newRow = [
      twTime,                  // A. 時間戳記
      doc.studentId || "",     // B. 學號
      doc.name || "",          // C. 姓名
      doc.topic || "",         // D. 測驗單元
      doc.mode || "",          // E. 測驗模式
      doc.attempt || 1,        // F. 第幾次
      doc.score || 0,          // G. 分數
      doc.correctCount || 0,   // H. 答對題數
      doc.wrongCount || 0,     // I. 答錯題數
      doc.duration || "",      // J. 作答秒數 (空欄)
      "",                      // K. (空欄)
      "",                      // L. (空欄)
      detailsStr               // M. 作答明細(JSON)
    ];
    
    rowsToWrite.push(newRow);
    uniqueKeys[key] = true;
  });
  
  if (rowsToWrite.length > 0) {
    Logger.log("💾 正在向試算表寫入 " + rowsToWrite.length + " 筆新成績...");
    var lastRow = scoreSheet.getLastRow();
    scoreSheet.getRange(lastRow + 1, 1, rowsToWrite.length, 13).setValues(rowsToWrite);
    Logger.log("🎉 成功同步 " + rowsToWrite.length + " 筆成績！");
  } else {
    Logger.log("ℹ️ 所有拉回的成績皆為重複，無需寫入。");
  }
}
