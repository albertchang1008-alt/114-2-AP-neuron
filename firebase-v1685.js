(function () {
  "use strict";

  var cfg = window.FIREBASE_V18_CONFIG || {};
  var app = null;
  var db = null;
  var auth = null;
  var boot = null;
  var queueKey = "quiz_v1685_firebase_queue";

  function enabled() {
    var c = cfg.firebaseConfig || {};
    return !!(cfg.enabled && window.firebase && c.apiKey && c.projectId && c.authDomain && c.appId);
  }

  function init() {
    if (!enabled()) return false;
    if (app && db && auth) return true;
    app = window.firebase.apps && window.firebase.apps.length
      ? window.firebase.app()
      : window.firebase.initializeApp(cfg.firebaseConfig);
    db = window.firebase.firestore(app);
    auth = window.firebase.auth(app);
    return true;
  }

  function docPath(path) {
    var parts = String(path || "").split("/").filter(Boolean);
    if (parts.length % 2 !== 0) throw new Error("Firestore 文件路徑不正確：" + path);
    var ref = db.collection(parts[0]).doc(parts[1]);
    for (var i = 2; i < parts.length; i += 2) ref = ref.collection(parts[i]).doc(parts[i + 1]);
    return ref;
  }

  function normalizeQuestion(doc) {
    var q = doc.data ? doc.data() : doc;
    return {
      id: q.id || doc.id || "",
      top: q.top || q.category || "未分類",
      q: q.q || q.question || "",
      options: Array.isArray(q.options) ? q.options : [q.optionA, q.optionB, q.optionC, q.optionD].filter(Boolean),
      ans: q.ans || q.answer || "",
      exp: q.exp || q.explanation || "尚無解析",
      color: q.color || "red",
      questionType: q.questionType || q.type || "",
      imgUrl: q.imgUrl || q.imageUrl || "",
      isImage: !!(q.isImage || q.imgUrl || q.imageUrl),
      cogType: q.cogType || "",
      source: q.source || "firebase",
      questionBankVersion: q.questionBankVersion || q.version || ""
    };
  }

  function uniqueTopics(questions) {
    var map = {};
    questions.forEach(function (q) {
      var name = q.top || "未分類";
      if (!map[name]) map[name] = { name: name, color: q.color || "red", count: 0 };
      map[name].count += 1;
    });
    return Object.keys(map).sort(function (a, b) { return a.localeCompare(b, "zh-TW"); }).map(function (k) { return map[k]; });
  }

  async function loadBootstrap() {
    if (!init()) return null;
    if (boot) return boot;
    var c = cfg.collections || {};
    var snap = await db.collection(c.questions || "questions").get();
    var questions = [];
    snap.forEach(function (doc) {
      var q = normalizeQuestion(doc);
      if (q.id && q.q) questions.push(q);
    });
    if (!questions.length) return null;

    var settings = {};
    var rankingCache = null;
    try {
      var s = await docPath(c.settings || "system/main").get();
      if (s.exists) settings = s.data() || {};
    } catch (err) {
      console.warn("[v1.6851] Firebase 設定讀取失敗，略過：", err);
    }
    try {
      var r = await docPath(c.homeRanking || "rankingCaches/home").get();
      if (r.exists) rankingCache = r.data() || null;
    } catch (err) {
      console.warn("[v1.6851] Firebase 排行讀取失敗，略過：", err);
    }

    boot = {
      status: "success",
      source: "firebase",
      title: settings.title || "動態題庫測驗",
      titleColor: settings.titleColor || "sky",
      topics: settings.topics || uniqueTopics(questions),
      questions: questions,
      studentHashes: settings.studentHashes || [],
      completionSettings: settings.completionSettings || settings,
      allClassList: settings.allClassList || [],
      deadline: settings.deadline || "",
      rankingCache: rankingCache,
      questionBankVersion: settings.questionBankVersion || ""
    };
    return boot;
  }

  function currentUserEmail() {
    return auth && auth.currentUser ? (auth.currentUser.email || "") : "";
  }

  async function ensureSignedIn() {
    if (!init()) throw new Error("Firebase 尚未啟用");
    if (auth.currentUser) return auth.currentUser;
    return auth.signInAnonymously().then(function (cred) {
      return cred.user;
    });
  }

  function nowField() {
    return window.firebase.firestore.FieldValue.serverTimestamp();
  }

  function safeDocId(raw) {
    return String(raw || "doc").replace(/[^\w.-]/g, "_").slice(0, 150);
  }

  function readQueue() {
    try { return JSON.parse(localStorage.getItem(queueKey) || "[]"); }
    catch (e) { return []; }
  }

  function writeQueue(items) {
    localStorage.setItem(queueKey, JSON.stringify(items.slice(-500)));
  }

  function enqueue(payload) {
    var items = readQueue();
    items.push({ createdAt: new Date().toISOString(), payload: payload });
    writeQueue(items);
  }

  async function submitAttempt(payload) {
    if (!init()) throw new Error("Firebase 尚未啟用");
    await ensureSignedIn();
    var c = cfg.collections || {};
    var batchId = payload.batchId || ("B_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8));
    var email = currentUserEmail();
    var batch = {
      batchId: batchId,
      studentId: payload.studentId || "",
      name: payload.name || "",
      email: email,
      topic: payload.topic || "",
      mode: payload.mode || "",
      attempt: Number(payload.attempt) || 1,
      score: Number(payload.score) || 0,
      correctCount: Number(payload.correctCount) || 0,
      wrongCount: Number(payload.wrongCount) || 0,
      duration: Number(payload.duration) || 0,
      isRetryMode: !!payload.isRetryMode,
      token: payload.token || "",
      ip: payload.ip || "",
      questionBankVersion: payload.questionBankVersion || "",
      settingsVersion: payload.settingsVersion || "",
      createdAt: nowField(),
      clientCreatedAt: new Date().toISOString(),
      source: "firebase-v1.6851"
    };
    var details = Array.isArray(payload.details) ? payload.details : [];
    var writer = db.batch();
    writer.set(db.collection(c.answerBatches || "answerBatches").doc(batchId), batch, { merge: true });
    details.forEach(function (d, idx) {
      var qid = d.questionId || ("Q_" + idx);
      var detailId = safeDocId(batchId + "_" + qid + "_" + idx);
      var detail = {
        batchId: batchId,
        studentId: batch.studentId,
        name: batch.name,
        email: email,
        mode: batch.mode,
        attempt: batch.attempt,
        questionId: qid,
        questionText: d.questionText || "",
        topic: d.topic || "",
        selectedText: d.selectedText || "",
        correctText: d.correctText || "",
        isCorrect: !!d.isCorrect,
        answerSec: d.answerSec === null || d.answerSec === undefined ? null : Number(d.answerSec),
        questionType: d.questionType || "",
        cogType: d.cogType || "",
        createdAt: nowField(),
        clientCreatedAt: new Date().toISOString(),
        source: "firebase-v1.6851"
      };
      writer.set(db.collection(c.answerDetails || "answerDetails").doc(detailId), detail, { merge: true });

      var progressId = safeDocId(batch.studentId + "_" + qid);
      writer.set(db.collection("studentProgress").doc(progressId), {
        studentId: batch.studentId,
        name: batch.name,
        email: email,
        questionId: qid,
        topic: detail.topic,
        questionType: detail.questionType,
        cogType: detail.cogType,
        lastBatchId: batchId,
        lastMode: batch.mode,
        lastIsCorrect: detail.isCorrect,
        lastAnswerSec: detail.answerSec,
        lastAnsweredAt: nowField(),
        updatedAt: nowField(),
        answerCount: window.firebase.firestore.FieldValue.increment(1),
        correctCount: window.firebase.firestore.FieldValue.increment(detail.isCorrect ? 1 : 0),
        wrongCount: window.firebase.firestore.FieldValue.increment(detail.isCorrect ? 0 : 1)
      }, { merge: true });

      var wrongId = progressId;
      if (!detail.isCorrect) {
        writer.set(db.collection("wrongQuestions").doc(wrongId), {
          studentId: batch.studentId,
          name: batch.name,
          email: email,
          questionId: qid,
          questionText: detail.questionText,
          topic: detail.topic,
          correctText: detail.correctText,
          selectedText: detail.selectedText,
          lastWrongAt: nowField(),
          lastBatchId: batchId,
          active: true,
          source: "firebase-v1.6851"
        }, { merge: true });
      } else {
        writer.set(db.collection("wrongQuestions").doc(wrongId), {
          studentId: batch.studentId,
          email: email,
          questionId: qid,
          active: false,
          masteredAt: nowField(),
          lastBatchId: batchId
        }, { merge: true });
      }
    });
    await writer.commit();
    return { status: "ok", batchId: batchId, writtenDetails: details.length };
  }

  async function submitAttemptWithFallback(payload) {
    try {
      return await submitAttempt(payload);
    } catch (err) {
      console.warn("[v1.6851] Firebase 作答寫入失敗，已暫存：", err);
      enqueue(payload);
      return { status: "queued", message: err.message };
    }
  }

  async function flushQueue() {
    if (!init()) return { status: "skip" };
    await ensureSignedIn();
    var items = readQueue();
    if (!items.length) return { status: "ok", flushed: 0 };
    var remain = [];
    var flushed = 0;
    for (var i = 0; i < items.length; i++) {
      try {
        await submitAttempt(items[i].payload);
        flushed += 1;
      } catch (err) {
        remain.push(items[i]);
      }
    }
    writeQueue(remain);
    return { status: "ok", flushed: flushed, remaining: remain.length };
  }

  window.Firebase1685 = {
    init: init,
    isEnabled: init,
    loadBootstrap: loadBootstrap,
    submitAttempt: submitAttempt,
    submitAttemptWithFallback: submitAttemptWithFallback,
    flushQueue: flushQueue,
    ensureSignedIn: ensureSignedIn,
    queueCount: function () { return readQueue().length; }
  };
})();
