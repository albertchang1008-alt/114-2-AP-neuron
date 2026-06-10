// Firebase v1.69 設定檔
// 1. 到 Firebase Console 建立 Web App。
// 2. 將 Firebase SDK config 貼到 firebaseConfig。
// 3. v1.69 學生端固定使用 Firebase 題庫、判分、作答紀錄與錯題資料。
// 4. GAS/Google Sheet 保留為題庫、名單、設定與後台同步入口。
window.FIREBASE_V18_CONFIG = {
  enabled: true,
  firebaseConfig: {
    apiKey: "AIzaSyDC1YmFlsTUzfghgdiMTD9Rm0fR4vGNYiQ",
    authDomain: "ap-neuron.firebaseapp.com",
    projectId: "ap-neuron",
    storageBucket: "ap-neuron.firebasestorage.app",
    messagingSenderId: "905876942421",
    appId: "1:905876942421:web:a78fc9f04ba7b09230aedc",
    measurementId: "G-TEW1PL1N46"
  },
  collections: {
    questions: "questions",
    students: "students",
    settings: "system/main",
    homeRanking: "rankingCaches/home",
    answerBatches: "answerBatches",
    answerDetails: "answerDetails",
    studentProgress: "studentProgress",
    wrongQuestions: "wrongQuestions"
  }
};
