// Firebase v1.685 設定檔
// 1. 到 Firebase Console 建立 Web App。
// 2. 將 Firebase SDK config 貼到 firebaseConfig。
// 3. 將 enabled 改成 true 後，學生端會優先讀 Firebase；失敗時仍會回退 GAS。
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
    answerDetails: "answerDetails"
  }
};
