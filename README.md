# 我的存摺 · 個人存款管理

一個可以放在 GitHub Pages 上的個人存款管理網站。資料存在你自己的 Google 試算表，
只有知道你的 Apps Script 網址 + 密鑰的人才能讀寫。

## 檔案結構

```
finance-tracker/
├── index.html          ← 網站本體，就是這個要放到 GitHub 的檔案
├── apps-script/
│   └── Code.gs          ← 貼到 Google Apps Script 的後端程式碼
└── README.md
```

---

## 第一步：建立 Google 試算表 + Apps Script 後端

1. 到 https://sheets.google.com 新開一份試算表（隨便命名，例如「我的存款管理」）。
2. 上方選單 **擴充功能 → Apps Script**，會開一個新分頁的程式碼編輯器。
3. 把 `apps-script/Code.gs` 的內容全部貼進去，取代原本的內容，記得存檔（Ctrl+S）。
4. 左側「專案設定」（齒輪圖示）→ 往下找「指令碼屬性」→ 新增屬性：
   - 屬性：`SECRET_TOKEN`
   - 值：自己隨便打一串英數混合的密碼（例如 `a8x!K391qzLm`），這就是之後網站的登入密鑰。
5. 右上角「部署」→「新增部署作業」：
   - 類型選「網頁應用程式」
   - 說明隨意
   - 執行身分：**我**
   - 誰能存取：**所有人**（這不代表資料公開，沒有密鑰還是讀不到，見下方說明）
   - 按「部署」，第一次會要求你「授權」，照畫面點下去允許即可（會跳出「未驗證應用程式」的警告，因為是你自己寫的小工具，點「進階」→「前往（不安全）」繼續授權）。
6. 部署完成後會給你一個網址，長得像：
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
   把這個網址記下來，之後要填到網站的「設定」頁。

> 之後只要修改 `Code.gs`，都要回到「部署 → 管理部署作業 → 編輯（鉛筆圖示）→ 版本選新版本 → 部署」，
> 網址會保持不變，但要重新部署新版本，改動才會生效。

---

## 第二步：上傳到 GitHub

### 你的資料需要公開嗎？

**不用。** `index.html` 裡面完全不含你的任何財務資料，資料都在 Google 試算表裡，
而且 Apps Script 有密鑰保護，沒有正確密鑰打不進去。所以：

- GitHub repo 設 **public** 也沒關係（GitHub Pages 免費版最簡單的用法）。
- 如果你還是不放心，也可以把 repo 設成 **private**，GitHub 免費帳號目前也支援
  private repo 開啟 GitHub Pages（在 repo 的 Settings → Pages 設定），多一層保護。
- 密鑰（token）**不要**寫死在程式碼裡直接推上 GitHub——這個範本已經設計成密鑰只存在
  瀏覽器的 localStorage（你自己那台電腦/手機），不會出現在 GitHub 上的檔案內容中。

### 從電腦操作到 Git 推上去的完整步驟

假設你已經安裝好 [Git](https://git-scm.com/) 且有 GitHub 帳號。

1. 到 GitHub 網站建立一個新 repository（例如取名 `my-finance-tracker`），
   建立時**不要**勾選自動加 README（我們本地已經有檔案了）。

2. 打開電腦的終端機（Windows 用 Git Bash，Mac/Linux 用 Terminal），
   切換到你剛剛下載的 `finance-tracker` 資料夾：

   ```bash
   cd 路徑/到/finance-tracker
   ```

3. 初始化 Git 並第一次提交：

   ```bash
   git init
   git add .
   git commit -m "first commit: 個人存款管理網站"
   git branch -M main
   ```

4. 連結到你在 GitHub 建立的 repo（把網址換成你自己的）：

   ```bash
   git remote add origin https://github.com/你的帳號/my-finance-tracker.git
   git push -u origin main
   ```

   第一次 push 可能會要求你登入 GitHub（建議用 GitHub 官方引導設定
   Personal Access Token 或透過瀏覽器登入授權）。

5. 到 GitHub 該 repo 的 **Settings → Pages**：
   - Source 選 `main` 分支、根目錄 `/ (root)`
   - 存檔後等 1-2 分鐘，頁面會顯示一個網址，例如：
     ```
     https://你的帳號.github.io/my-finance-tracker/
     ```
   這就是你的網站了。

6. 打開這個網址 → 點「設定」分頁 → 貼上第一步拿到的 **Apps Script 網址** 和 **密鑰** →
   按「儲存設定並重新連線」。連線成功後就可以開始新增帳戶、負債、利息紀錄了。

### 之後修改網站要怎麼更新？

每次改完 `index.html`（或任何檔案）後，在同一個資料夾執行：

```bash
git add .
git commit -m "說明你改了什麼"
git push
```

GitHub Pages 會自動在幾分鐘內更新成最新版本，不用重新設定。

---

## 功能說明

- **總覽**：依帳戶類型（現金／銀行／行動支付／悠遊卡／投資）加總，並顯示未繳負債與淨資產。
- **帳戶**：新增、編輯、刪除各類帳戶，銀行帳戶可填年利率，會自動估算下期利息（僅供參考）。
- **利息紀錄**：銀行真的撥款給你時，在這裡新增一筆實際入帳金額，系統會自動加進該帳戶餘額，
  同時保留歷史紀錄可回顧。
- **負債**：管理國民年金未繳款等負債項目，可標記「已繳／未繳」、設定繳費期限。
- **設定**：填入 Apps Script 網址與密鑰，僅存在瀏覽器本機（localStorage），不會被推上 GitHub。

## 安全性補充

這個方案對「只有自己用」的個人小工具來說已經足夠：資料放在私人試算表、
API 有密鑰保護、GitHub 上不含任何真實資料。如果未來想更嚴謹，
可以考慮改用「Google 帳號登入驗證」取代單純密鑰比對，但對個人記帳用途通常不必要。
