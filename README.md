# 題庫系統 v1.69 佈建說明

## 版本定位

v1.69 是由 `題庫系統修正版-2026-06-05-v1.685` 升級而來，改成接近 v1.81 的資料模式：

- Google Sheet / GAS：題庫、學生名單、系統設定、首頁排行快取的管理與同步入口。
- Firebase Firestore：學生端題庫讀取、作答判分、交卷批次紀錄、每題明細、學生進度、錯題資料。
- 學生作答中不逐題寫入 Firebase；完成一次練習或閃卡測驗後才批次送出。

## 專案注意

這一版目前設定的 Firebase project 是：

```text
ap-neuron
```

不要把本資料夾部署到其他 Firebase project。若要確認目前登入帳號看得到哪些 project：

```bash
"/Users/HHC/Documents/New project/.tools/node/bin/npx" -y firebase-tools@latest projects:list
```

若清單沒有 `ap-neuron`，代表目前 Firebase CLI 登入的 Google 帳號不是此專案的成員。

## 資料流

1. 老師在 Google Sheet 維護題庫、學生名單與設定。
2. 後台按「同步到 Firebase」，或在 GAS 執行同步函式。
3. GAS 將以下資料推送到 Firestore：
   - `questions`
   - `students`
   - `system/main`
   - `rankingCaches/home`
4. 學生端載入時直接讀 Firestore，不再用 GAS 即時計算題庫。
5. 學生完成一次作答後，前端用 Firebase 題庫答案本機判分，再批次寫入：
   - `answerBatches`
   - `answerDetails`
   - `studentProgress`
   - `wrongQuestions`

## 必要部署順序

1. 先確認 `firebase-config.js` 仍指向正確 project。
2. 部署 Firestore rules / indexes。
3. 部署或上傳前端檔案。
4. 到後台或 GAS 執行 Google Sheet → Firebase 同步。
5. 確認 Firestore 裡有 `questions` 資料後，再讓學生使用。

## 部署指令

在本資料夾執行：

```bash
cd "/Users/HHC/Documents/New project/題庫系統修正版-2026-06-05-v1.69"
"/Users/HHC/Documents/New project/.tools/node/bin/npx" -y firebase-tools@latest deploy --only firestore:rules,firestore:indexes --project ap-neuron
```

若要部署 Firebase Hosting：

```bash
cd "/Users/HHC/Documents/New project/題庫系統修正版-2026-06-05-v1.69"
"/Users/HHC/Documents/New project/.tools/node/bin/npx" -y firebase-tools@latest deploy --only hosting --project ap-neuron
```

若實際入口仍使用 GitHub Pages，則把以下前端檔案放到 GitHub Pages 對應位置：

- `index.html`
- `admin.html`
- `firebase-config.js`
- `firebase-v1685.js`

## 答案判分原則

答案以 Firebase 中由 Google Sheet 同步過來的題庫為準。前端不再把 `A`、`B`、`C`、`D` 或 `1`、`2`、`3`、`4` 從選項文字前方移除，避免把「A型肝炎」、「B型肝炎」這類醫學選項誤判。

若出現「正確答案不在選項中」，請檢查 Google Sheet 題庫的「正確答案」欄是否能對應到四個選項之一，然後重新同步到 Firebase。

