# 水電工程記帳系統 — 部署指南

這個專案已經把資料改成存在 **Firebase Firestore**（真正的雲端資料庫），並設計成可以放到
**GitHub Pages** 上，變成一個手機瀏覽器就能打開的網站，不再需要透過 Claude。

整個設定分成三大步驟：**① 建立 Firebase 專案 → ② 建立 GitHub 專案並上傳程式碼 → ③ 開啟 GitHub Pages**。
每一步都不需要寫程式，跟著畫面點就可以。

---

## ① 建立 Firebase 專案（資料庫）

1. 開啟 <https://console.firebase.google.com>，用 Google 帳號登入。
2. 按「新增專案」，取個名字（例如 `water-electric-ledger`），一路「繼續」到建立完成（Google Analytics 可以關掉，不需要）。
3. 進入專案後，左側選單找「**Firestore Database**」→ 按「建立資料庫」。
   - 位置選離你最近的（例如 `asia-east1`）。
   - 安全性規則先選「**測試模式**」（等一下第 6 步會換成正式的規則）。
4. 左側選單找「**Authentication**」→ 按「開始使用」→「Sign-in method」分頁 → 啟用「**匿名**」登入方式。
   - 這一步是讓網頁能自動、無聲地取得使用權限，你完全不需要輸入帳號密碼。
5. 回到「專案總覽」→ 按網頁圖示 `</>`「新增應用程式」→ 取個暱稱（例如 `web`）→「註冊應用程式」。
   - 這時畫面會顯示一段 `firebaseConfig = { apiKey: "...", ... }`，**先不要關掉這個畫面**。
6. 打開這個專案裡的 `src/firebase.js`，把裡面的 `firebaseConfig` 整段換成 Firebase 剛剛給你的那一段（`apiKey`、`authDomain`、`projectId`...）。存檔。
7. 回到 Firestore Database → 「規則」分頁，把內容整個換成：

   ```
   rules_version = '2';
   service cloud.firestore {
     match /databases/{database}/documents {
       match /wel-data/{document} {
         allow read, write: if request.auth != null;
       }
     }
   }
   ```

   按「發布」。這代表：**只有經過（匿名）登入的使用者才能讀寫資料**，比預設的「測試模式」（任何人都能讀寫）安全許多。

到這裡，資料庫已經準備好了。

---

## ② 建立 GitHub 專案並上傳程式碼

1. 到 <https://github.com/new> 建立一個新的 repository（例如取名 `water-electric-ledger`），設為 **Public**（GitHub Pages 的免費方案需要 Public repo），不要勾選「Add a README file」。
2. 如果你電腦上還沒裝過 Git，先安裝：<https://git-scm.com/downloads>
3. 打開終端機（Mac 是「終端機」App，Windows 是 Git Bash），依序執行（把 `你的帳號` 換成你的 GitHub 帳號）：

   ```bash
   cd 這個專案資料夾的路徑
   git init
   git add .
   git commit -m "初始版本"
   git branch -M main
   git remote add origin https://github.com/你的帳號/water-electric-ledger.git
   git push -u origin main
   ```

4. 如果你的 repo 名稱**不是** `water-electric-ledger`，要打開 `vite.config.js`，把 `base: "/water-electric-ledger/"` 改成 `base: "/你的repo名稱/"`（記得改完要重新 commit + push）。

---

## ③ 開啟 GitHub Pages（讓網站正式上線）

1. 到 GitHub 上你剛剛建立的 repo → 「Settings」→ 左側「Pages」。
2. 「Build and deployment」的「Source」選擇「**GitHub Actions**」。
3. 回到 repo 首頁 → 「Actions」分頁，應該會看到一個叫 `Deploy to GitHub Pages` 的流程正在跑（第一次 push 就會自動觸發，因為專案裡已經幫你寫好 `.github/workflows/deploy.yml`）。等它跑完（打勾）大概 1-2 分鐘。
4. 跑完之後回到「Settings」→「Pages」，最上面會出現一個網址，長得像：

   ```
   https://你的帳號.github.io/water-electric-ledger/
   ```

   這就是正式網址了，**手機打開瀏覽器貼上這個網址就能用**。

### 手機上加到主畫面（像 App 一樣）
- **iPhone（Safari）**：打開網址 → 按下方「分享」圖示 → 「加入主畫面」。
- **Android（Chrome）**：打開網址 → 右上角選單（⋮）→ 「新增至主畫面」。

之後就會在手機桌面出現一個圖示，點開就直接是這個記帳系統，不需要每次打開瀏覽器輸入網址。

---

## 之後要怎麼更新網站？

以後每次你請人（或自己）修改程式碼、想要更新網站內容，只要：

```bash
git add .
git commit -m "說明這次改了什麼"
git push
```

`push` 之後，GitHub Actions 會自動重新建置並更新網站，不用再手動做任何事，大約 1-2 分鐘後網址上就會是最新版本。

---

## 本機開發（進階，非必要）

如果想在自己電腦上先預覽再上傳：

```bash
npm install     # 第一次要先安裝套件
npm run dev     # 啟動本機預覽，瀏覽器打開 http://localhost:5173
npm run build   # 打包成正式版本（GitHub Actions 會自動做這步，通常不用手動執行）
```

---

## 資料儲存方式說明

- 所有資料（叫貨、發包、師傅出工、甲方收款…）都存在 Firestore 的 `wel-data` collection 底下，一種資料一份文件（例如 `wel-data/orders`、`wel-data/contracts`）。
- 這是**共用資料庫**：只要是用同一個 Firebase 專案設定打開這個網站的人，看到的都是同一份資料（跟原本 Claude 版本的「共用」模式一樣的概念），適合你自己一人或跟同事一起使用同一份帳目。
- 如果之後想要「不同案場團隊各自獨立資料庫、互相看不到」，可以另外建立多個 Firebase 專案、各自對應不同的 `firebase.js` 設定，這部分需要再客製化，有需要可以再提出。

## 常見問題

**Q: 打開網站畫面一直轉圈圈，沒有反應？**
A: 通常是 `src/firebase.js` 裡的設定值還沒換成你自己的（還是 `YOUR_API_KEY` 那種預設文字），或是 Firestore／匿名登入還沒啟用，回頭檢查第①大步驟。

**Q: 想要有登入畫面、不同人要輸入帳號密碼才能用，可以嗎？**
A: 可以，但需要額外開發登入頁面與帳號管理，目前這版是設計成「知道網址就能用」（用匿名登入只是為了符合 Firestore 安全規則，不是真的帳號系統）。如果需要真正的帳號登入保護，跟我說一聲，我可以再加上去。
