/**
 * 🚀 Anti-Gravity 2 教師後台分析資料處理層 (admin-firebase.js)
 * 
 * 此腳本負責從 Firebase Firestore (answerBatches) 直接撈取原始資料，
 * 並在客戶端 Browser 內執行 Transform (重組/聚合) 邏輯，
 * 產出與 `admin.html` 預期 100% 相容的 `teacherData` 結構，
 * 以徹底取代 Google Apps Script (GAS) 緩慢且易超時的後端運算。
 */

const AdminFirebase = {
  
  /**
   * 核心進入點：拉取並重組分析資料，並與舊版快取融合 (Merge)
   * @param {firebase.firestore.Firestore} db - 已初始化的 Firestore 實例
   * @param {Object} existingCache - (Optional) 舊版的 rankingCaches/teacherData (包含 studentInfoMap 與歷史 studentHistory)
   * @returns {Promise<Object>} 回傳完整的 teacherData 物件
   */
  async fetchAndBuildTeacherData(db, existingCache = {}) {
    try {
      console.log("🚀 [AdminFirebase] 開始從 Firestore 拉取所有 answerBatches 與融合舊資料...");
      const startTime = Date.now();
      
      // 1. Fetch: 一次性拉取所有批次
      const snapshot = await db.collection("answerBatches").get();
      const batches = [];
      snapshot.forEach(doc => {
        batches.push({ id: doc.id, ...doc.data() });
      });
      console.log(`📦 [AdminFirebase] 成功拉取 ${batches.length} 筆批次資料，耗時 ${Date.now() - startTime}ms`);

      // 2. 準備 Transform 所需的資料結構，並直接從現有快取繼承
      const studentInfoMap = existingCache.studentInfoMap || {}; // sid -> { name, class }
      
      // ★ 融合策略 (Merge Strategy)：把舊有的 studentHistory (如果有) 當作底層
      const studentHistory = existingCache.studentHistory || {}; // sid -> { name, class, best: {}, last: {}, attempts: [] }
      
      const studentWrongDetails = {}; // sid -> [ { topic, questionText... } ]
      
      const qStatsMap = {}; // questionId -> { text, correct, wrong, topic, type, cogType, totalSec, count }
      const tStatsMap = {}; // topic -> { count, totalScore, totalSec }
      const cogStatsMap = {}; // cogType -> { correct, total }
      
      const topicTimeMap = {}; // topic -> { intervals: { [hourTimestamp]: count } }

      // 3. 遍歷所有的 batches
      batches.forEach(batch => {
        const sid = batch.studentId ? String(batch.studentId).trim() : "";
        if (!sid) return;

        const name = batch.name || sid;
        const topic = (batch.topic || "").trim();
        const score = Number(batch.score) || 0;
        const createdAt = batch.createdAt || batch.clientCreatedAt;
        const mode = batch.mode || "";
        
        // 跳過錯題重做模式不列入統計
        if (mode === "錯題重做") return;

        // 維護基礎學生名單
        if (!studentInfoMap[sid]) {
          studentInfoMap[sid] = { name: name, class: "未分班" };
        }
        const stuClass = studentInfoMap[sid].class || "未分班";

        // 初始化學生歷程 (若舊快取中沒有該學生)
        if (!studentHistory[sid]) {
          studentHistory[sid] = { name: name, class: stuClass, best: {}, last: {}, attempts: [] };
        }
        if (!studentHistory[sid].best) studentHistory[sid].best = {};
        if (!studentHistory[sid].last) studentHistory[sid].last = {};
        if (!studentHistory[sid].attempts) studentHistory[sid].attempts = [];

        // 紀錄所有作答歷程
        studentHistory[sid].attempts.push({
          date: createdAt,
          topic: topic,
          mode: mode,
          attempt: Number(batch.attempt) || 1,
          score: score,
          correct: Number(batch.correctCount) || 0,
          wrong: Number(batch.wrongCount) || 0,
          duration: Number(batch.duration) || 0,
          isRetry: false
        });

        // 更新各單元最高分與最後作答時間
        if (topic) {
          if (studentHistory[sid].best[topic] === undefined || score > studentHistory[sid].best[topic]) {
            studentHistory[sid].best[topic] = score;
          }
          // 簡單紀錄最後一次作答資訊
          if (!studentHistory[sid].last[topic] || new Date(createdAt) > new Date(studentHistory[sid].last[topic].time || 0)) {
            studentHistory[sid].last[topic] = { score: score, time: createdAt, dateText: new Date(createdAt).toLocaleString('zh-TW') };
          }
        }

        // 時間軸趨勢統計 (以小時為單位)
        if (createdAt && topic) {
          const ts = new Date(createdAt).getTime();
          if (!isNaN(ts)) {
            const hourTs = Math.floor(ts / 3600000) * 3600000;
            if (!topicTimeMap[topic]) topicTimeMap[topic] = { min: ts, max: ts, intervals: {} };
            
            if (ts < topicTimeMap[topic].min) topicTimeMap[topic].min = ts;
            if (ts > topicTimeMap[topic].max) topicTimeMap[topic].max = ts;
            
            topicTimeMap[topic].intervals[hourTs] = (topicTimeMap[topic].intervals[hourTs] || 0) + 1;
          }
        }

        // 解析作答明細 (detailsJson)
        if (batch.detailsJson) {
          try {
            const details = JSON.parse(batch.detailsJson);

            details.forEach(d => {
              const isCorrect = !!d.isCorrect;
              const qid = d.questionFirebaseId || d.questionId;
              const qTopic = d.topic || topic || "未分類";
              const cogType = d.cogType || "未分類";

              // 累計單元總和 (用於 topicStats 與 topicTimeList)
              if (!tStatsMap[qTopic]) tStatsMap[qTopic] = { correct: 0, total: 0, totalSec: 0, count: 0 };
              tStatsMap[qTopic].total += 1;
              if (isCorrect) tStatsMap[qTopic].correct += 1;
              
              if (d.answerSec !== null && d.answerSec !== undefined) {
                  tStatsMap[qTopic].totalSec += Number(d.answerSec);
                  tStatsMap[qTopic].count += 1;
              }

              // 統計每題答錯
              if (!isCorrect) {
                if (!studentWrongDetails[sid]) studentWrongDetails[sid] = [];
                studentWrongDetails[sid].push({
                  topic: qTopic,
                  questionText: d.questionText || "",
                  selectedText: d.selectedText || "",
                  correctText: d.correctText || "",
                  answerSec: d.answerSec,
                  mode: mode,
                  timestamp: createdAt
                });
              }

              // 統計題目 (questionStats)
              if (qid) {
                if (!qStatsMap[qid]) {
                  qStatsMap[qid] = { 
                    qid: qid, text: d.questionText, topic: qTopic, 
                    type: d.questionType, cogType: cogType, 
                    correct: 0, wrong: 0, total: 0, totalSec: 0, count: 0 
                  };
                }
                qStatsMap[qid].total += 1;
                if (isCorrect) qStatsMap[qid].correct += 1;
                else qStatsMap[qid].wrong += 1;
                
                if (d.answerSec !== null && d.answerSec !== undefined) {
                  qStatsMap[qid].totalSec += Number(d.answerSec);
                  qStatsMap[qid].count += 1;
                }
              }

              // 統計認知階層 (cogTypeStats)
              if (cogType) {
                if (!cogStatsMap[cogType]) cogStatsMap[cogType] = { type: cogType, correct: 0, total: 0 };
                cogStatsMap[cogType].total += 1;
                if (isCorrect) cogStatsMap[cogType].correct += 1;
              }
            });
          } catch (err) {
            console.warn(`⚠️ [AdminFirebase] 解析批次 ${batch.id} 的 detailsJson 失敗:`, err);
          }
        }
      });

      // 4. Export: 將 Map 轉換為前端所需的 Array 格式
      
      // -- questionStats --
      const questionStats = Object.values(qStatsMap).map(q => {
        const correctRate = q.total > 0 ? Math.round((q.correct / q.total) * 100) : 0;
        const avgSec = q.count > 0 ? Number((q.totalSec / q.count).toFixed(1)) : 0;
        return {
          topic: q.topic,
          id: q.qid,
          text: q.text,
          correct: q.correct,
          wrong: q.wrong,
          total: q.total,
          rate: correctRate,
          type: q.type || "單選題",
          cogType: q.cogType || "未分類",
          avgSec: avgSec
        };
      }).sort((a, b) => a.rate - b.rate); // 預設由答對率低到高排序

      // -- topicStats --
      const topicStats = Object.entries(tStatsMap).map(([t, stats]) => {
        return {
          topic: t,
          total: stats.total,
          correct: stats.correct,
          rate: stats.total > 0 ? Math.round((stats.correct / stats.total) * 100) : 0
        };
      }).sort((a, b) => b.total - a.total);

      // -- cogTypeStats --
      const cogTypeStats = Object.values(cogStatsMap).map(c => {
        return {
          type: c.type,
          correct: c.correct,
          total: c.total,
          rate: c.total > 0 ? Math.round((c.correct / c.total) * 100) : 0
        };
      }).sort((a, b) => b.total - a.total);

      // -- topicTimeList --
      const topicTimeList = Object.entries(tStatsMap).map(([t, stats]) => {
        return { 
          topic: t, 
          avgSec: stats.count > 0 ? Math.round(stats.totalSec / stats.count) : 0, 
          sessionCount: stats.count 
        };
      }).sort((a, b) => a.topic.localeCompare(b.topic, "zh-TW"));

      // -- classList (班級統計) --
      const classMap = {};
      Object.keys(studentHistory).forEach(sid => {
        const stu = studentHistory[sid];
        const cls = stu.class || "未分班";
        if (!classMap[cls]) classMap[cls] = { class: cls, studentCount: 0, correct: 0, total: 0, completedSum: 0, allDoneCount: 0, students: [] };
        
        classMap[cls].studentCount += 1;

        // 計算該學生的歷程加總
        stu.attempts.forEach(a => {
          if (!a.isRetry) {
            classMap[cls].total += (a.correct || 0) + (a.wrong || 0);
            classMap[cls].correct += (a.correct || 0);
          }
        });

        // 取得所有出現過的 topic (包含歷史快取與最新資料)
        const mergedTopics = new Set(Object.keys(tStatsMap));
        (existingCache.topicStats || []).forEach(t => mergedTopics.add(t.topic));
        const reqTopics = Array.from(mergedTopics);
        
        const completed = reqTopics.filter(t => (stu.best[t] || 0) >= 80).length; // 預設及格分 80
        
        classMap[cls].completedSum += completed;
        if (completed === reqTopics.length && reqTopics.length > 0) classMap[cls].allDoneCount += 1;
        
        classMap[cls].students.push({
          sid: sid,
          name: stu.name,
          completed: completed,
          totalTopics: reqTopics.length,
          details: reqTopics.map(t => ({ topic: t, best: stu.best[t] || null, passed: (stu.best[t] || 0) >= 80 }))
        });
      });

      const classList = Object.values(classMap).map(c => {
        c.rate = c.total > 0 ? Math.round((c.correct / c.total) * 100) : null;
        c.avgCompleted = c.studentCount > 0 ? Math.round((c.completedSum / c.studentCount) * 10) / 10 : 0;
        const totalReq = Object.keys(tStatsMap).length;
        c.pct = totalReq > 0 ? Math.round((c.avgCompleted / totalReq) * 100) : 0;
        // 學生以完成度排序
        c.students.sort((a, b) => b.completed - a.completed);
        return c;
      }).sort((a, b) => a.class.localeCompare(b.class, "zh-TW"));

      console.log("✅ [AdminFirebase] Transform 處理完成，總耗時:", Date.now() - startTime, "ms");

      return {
        studentInfoMap: studentInfoMap,
        studentHistory: studentHistory,
        studentWrongDetails: studentWrongDetails,
        questionStats: questionStats,
        topicStats: topicStats,
        cogTypeStats: cogTypeStats,
        topicTimeList: topicTimeList,
        classList: classList,
        // 保留給現有 admin.html 相容的設定值
        passScore: 80,
        reqTopics: Object.keys(tStatsMap)
      };

    } catch (err) {
      console.error("❌ [AdminFirebase] fetchAndBuildTeacherData 發生錯誤:", err);
      throw err;
    }
  }
};

// 若於瀏覽器環境執行，暴露至全域變數供 admin.html 呼叫
if (typeof window !== 'undefined') {
  window.AdminFirebase = AdminFirebase;
}
