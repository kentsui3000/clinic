# 部署指南 — 從零到上線

照著做，大約 15–20 分鐘可完成。

---

## Step 1：取得 Gemini API Key

1. 開啟 https://aistudio.google.com/apikey （要登入 Google 帳號）
2. 點 **「Create API Key」** → 選一個 GCP 專案（沒有就讓它建立新的）
3. 拷貝產生的 Key（形如 `AIzaSyXXX…`），等一下要用

> 💡 Gemini 2.5 Flash / Pro 的免費額度足以個人使用（每分鐘 15 次、每天 1500 次）

---

## Step 2：建立 Google 試算表

1. 到 https://sheets.google.com/ 新增一張試算表
2. 把它命名成例如 `bonus calculation`
3. **拷貝試算表 ID**：URL 中 `/d/` 後面那一段
   - 範例：`https://docs.google.com/spreadsheets/d/1buSauwDrkIp9cJNA_ZCBgHF344wXYZyFaXAT5zO_geA/edit`
   - 試算表 ID = `1buSauwDrkIp9cJNA_ZCBgHF344wXYZyFaXAT5zO_geA`

> 也可以沿用你既有的 `bonus calculation` 試算表 — 程式只會新增 `YYYY-M` 分頁，不會動到舊資料。

---

## Step 3：部署 Apps Script 後端

1. 開啟剛剛建立的試算表 → 上方選單 **「擴充功能 → Apps Script」**
2. 在編輯器中，把預設 `Code.gs` 內容**全部刪除**，貼上本專案 `Code.gs` 的全部內容
3. 把 `index.html`（本專案的）也加進來：
   - 左側檔案區 **「+」→ HTML 檔** → 命名 `index` → 把本專案 `index.html` 內容貼上
4. 在 `Code.gs` 上方找這兩行，**改成你自己的值**：
   ```js
   const GEMINI_API_KEY = 'AIzaSy...你的key';
   const WEBAPP_SECRET  = '自訂任意密碼例如abc123!@#';
   ```
5. 按 **儲存** 💾
6. 右上角 **「部署 → 新增部署作業」**
   - 類型：選 **「網頁應用程式」**
   - 說明：`Clinic Bonus OCR API`
   - 執行身分：**「我（your-email@gmail.com）」**
   - 誰可以存取：**「任何人」**（Apps Script 仍會用密碼驗證）
   - 按 **「部署」**
7. 第一次會跳授權視窗 → 同意全部權限
8. 部署完成後，**拷貝「網頁應用程式 URL」**（形如 `https://script.google.com/macros/s/AKfy.../exec`）

> 💡 如果之後改 `Code.gs`，記得回到「部署 → 管理部署作業 → ✏️ → 版本：新版本 → 部署」才會生效。

---

## Step 4：發佈靜態前端 (webapp/)

選一個你習慣的免費平台。**Cloudflare Pages** 對自有網域支援最好。

### 4A. Cloudflare Pages（推薦）

1. 把整個 `clinic` 專案推上 GitHub（你已經在 `claude/ocr-chinese-handwriting-app-hM5P2` 分支了）
2. 到 https://dash.cloudflare.com/ → **Workers & Pages → Create → Pages → Connect to Git**
3. 選你的 repo
4. 設定：
   - **Build command**：留空（純靜態）
   - **Build output directory**：`webapp`
5. **Save and Deploy**
6. 部署完成後會給你一個 `xxx.pages.dev` 的網址，先用它測試
7. 把它接到自己的網域：
   - **Custom domains → Set up a custom domain**
   - 填 `bonus.drkentsui.com`（或你想要的子網域）
   - Cloudflare 會自動加 DNS CNAME（如果 drkentsui.com 已在 Cloudflare）；不在的話，到你的 DNS 商加一筆 CNAME 指向 `xxx.pages.dev`

### 4B. Vercel

```
npm i -g vercel
cd webapp
vercel --prod
```

之後在 Vercel Dashboard → Settings → Domains 加 `bonus.drkentsui.com`。

### 4C. Netlify

把 `webapp/` 拖進 https://app.netlify.com/drop 即可，再到 Site settings → Domain 設定自有網域。

### 4D. GitHub Pages

```
git subtree push --prefix webapp origin gh-pages
```

到 GitHub repo → Settings → Pages → Source 選 `gh-pages` → Custom domain 填 `bonus.drkentsui.com`。

---

## Step 5：設定 App

1. 用手機（或電腦）開啟 `https://bonus.drkentsui.com`
2. 第一次會自動跳到 ⚙️ 設定頁
3. 填入：
   - **Apps Script Web App URL**：Step 3 拷貝的 URL
   - **共用密碼**：Step 3 寫進 `WEBAPP_SECRET` 的字串
   - **試算表 ID**：Step 2 的 ID（如果 Apps Script 是綁在這張試算表上，可以留空）
4. 按 **🔌 測試連線** → 應該看到「連線成功：pong」
5. 按 **儲存設定**

---

## Step 6：加到手機主畫面 (PWA)

### iPhone (Safari)
1. 開啟 `https://bonus.drkentsui.com`
2. 下方分享按鈕 → **「加入主畫面」**
3. 點 **「新增」** → 主畫面就有 App 圖示了

### Android (Chrome)
1. 開啟網址
2. 右上角 ⋮ → **「加到主畫面」**

---

## Step 7：實際拍一張試試

1. 點 📷 拍一張帳本（重點清楚拍到上方日期 + 值日生）
2. **開始辨識** → 等 5–15 秒
3. 校對日期、班別、值日生、員工分紅
4. 選擇要套用「主欄 / 次欄 / 主+次」哪個金額
5. **儲存到試算表**
6. 回到 Google Sheet，會看到新建的 `2026-04` 分頁，相應日期欄位自動填入

---

## 故障排除

| 症狀 | 可能原因 | 處理 |
|---|---|---|
| 測試連線「Failed to fetch」 | 還沒部署成 Web App / URL 拷錯 | 重新檢查「部署 → 管理部署」中的 URL，URL 結尾必為 `/exec` |
| 測試連線「密碼錯誤」 | 設定頁密碼 ≠ Code.gs 的 `WEBAPP_SECRET` | 兩邊改成相同 |
| 辨識「API 錯誤 (400)」 | Gemini Key 失效或圖片過大 | 換 Key、或重新拍小一點 |
| 辨識結果亂碼 | 模型版本太舊 | `Code.gs` 把 URL 改成 `gemini-2.5-pro` 或 `gemini-2.5-flash` |
| 寫入「未提供 spreadsheetId 且 Apps Script 未綁定試算表」 | 沒設定試算表 ID 且 Apps Script 不是 container-bound | 在設定頁填入試算表 ID |
| 改了 Code.gs 但行為沒變 | Apps Script Web App 是版本化的 | 部署 → 管理部署 → 編輯 → 版本選「新版本」→ 重新部署 |

---

## 之後想改進什麼？

- 在 Apps Script 加 `getEmployeeAliases()` 做姓名別名（例：`郭雅如` → `雅如`），讓矩陣分頁的列名跟你既有試算表一致
- 在 webapp 加離線佇列（Service Worker），訊號不好時可以先存著、之後同步
- 加批次模式，一次拍 3 張（早/午/晚）連續處理
