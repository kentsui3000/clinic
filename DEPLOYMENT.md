# 完整部署指南 (Vercel + GitHub + Google Apps Script)

> **這份指南已根據實際程式碼逐步驗證過。**
> 跟著做完，你就能在手機上拍照辨識帳本，自動寫入 Google 試算表。
> **預估時間：20–30 分鐘**

---

## 開始前的準備

| 你需要 | 用途 |
|---|---|
| Google 帳號 | Apps Script、試算表、Gemini API |
| GitHub 帳號 | 已有 repo `kentsui3000/clinic` |
| Vercel 帳號 | 用 GitHub 一鍵登入即可，免費 |
| `drkentsui.com` 的 DNS 後台 | 接自有網域才用得到，可選 |

---

## 階段 1：拿 Gemini API 金鑰（2 分鐘）

1. 開 https://aistudio.google.com/apikey
2. 用你的 Google 帳號登入
3. 右上角 **「Create API Key」**
4. 跳出視窗 → 選一個 GCP 專案（沒有就讓它自動建立）
5. 複製產生的金鑰，**開記事本貼住**：
   ```
   AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
   ```
6. ✅ **驗證**：金鑰開頭一定是 `AIzaSy`

> 💡 免費額度足夠個人使用（每分鐘 15 次、每天 1500 次）

---

## 階段 2：準備 Google 試算表（1 分鐘）

**A 案：用你既有的 `bonus calculation` 試算表**（推薦）
- 程式只會新增 `YYYY-M` 分頁，**不會動到舊資料**

**B 案：建一張新的**
- 到 https://sheets.google.com/ → 空白 → 命名

不論哪一案，**複製試算表 ID**（網址中 `/d/` 後的那段）：
```
https://docs.google.com/spreadsheets/d/【試算表 ID 在這裡】/edit
```
範例：`1buSauwDrkIp9cJNA_ZCBgHF344wXYZyFaXAT5zO_geA`

✅ **驗證**：試算表 ID 約 44 個英數字

---

## 階段 3：部署 Apps Script 後端（最重要，10 分鐘）

### 3.1 開啟 Apps Script 編輯器

1. 開啟你選定的試算表
2. 上方選單 **「擴充功能」→「Apps Script」**
3. 會在新分頁開啟 Apps Script 編輯器
4. 左側預設有一個 `Code.gs` 檔案

### 3.2 貼入新版 `Code.gs`

1. 點 `Code.gs`
2. **Ctrl+A 全選 → Delete** 把預設內容清空
3. 開啟 https://github.com/kentsui3000/clinic/blob/claude/ocr-chinese-handwriting-app-hM5P2/Code.gs
4. 點右上角 **「Raw」** → **Ctrl+A 全選 → Ctrl+C 複製**
5. 切回 Apps Script 編輯器 → **Ctrl+V 貼上**

### 3.3 修改兩個常數（⚠️ 這一步不做就不能用）

捲到檔案最上方，找到第 7 行與後面的 `WEBAPP_SECRET`：

```js
const GEMINI_API_KEY = 'YOUR_KEY_HERE';   // ← 換成階段 1 的金鑰
```

捲到 `// 行動版 Web App` 那段，找到：

```js
const WEBAPP_SECRET = 'change-me-before-deploy';   // ← 改成你自訂的密碼
```

把兩個字串都換成你的值。**`WEBAPP_SECRET` 之後手機要輸入相同字串，請記住**。

範例：
```js
const GEMINI_API_KEY = 'AIzaSyAbCd1234567890XyZ';
const WEBAPP_SECRET  = 'Kent2026!@#';
```

### 3.4 儲存

按 **Ctrl + S**（或點工具列的軟碟片圖示 💾）

✅ **驗證**：頁面標題不再有「未儲存」字樣

### 3.5 第一次手動執行（觸發授權）

1. 編輯器上方的下拉選單，選函式 **`doGet`**
2. 點 **「執行」** 按鈕
3. 跳出「需要授權」→ **「審查權限」**
4. 選你的 Google 帳號
5. 看到「Google 尚未驗證這個應用程式」 → 點 **「進階」**
6. 點 **「前往 [專案名稱]（不安全）」**
7. 點 **「允許」** （它需要 Sheets / 外部 URL 兩個權限）
8. 看到「執行已完成」即可

✅ **驗證**：底部「執行記錄」顯示 `info` 一筆，沒有紅字

### 3.6 部署成「網頁應用程式」

1. 右上角 **「部署」→「新增部署作業」**
2. 在「選取類型」旁的 **齒輪 ⚙️** → 選 **「網頁應用程式」**
3. 設定：
   | 欄位 | 填什麼 |
   |---|---|
   | 說明 | `Clinic Bonus OCR API`（隨便填） |
   | 執行身分 | **「我 (your-email@gmail.com)」** |
   | 誰可以存取 | **「任何人」** ⚠️ |
4. 按 **「部署」**
5. 部署完成後會顯示一個 **「網頁應用程式 URL」**：
   ```
   https://script.google.com/macros/s/AKfycb.................../exec
   ```
6. **整段複製貼到記事本**

✅ **驗證**：URL 結尾必為 `/exec`

> ⚠️ **「誰可以存取」一定要選「任何人」**，否則手機呼叫會被擋下。
> Apps Script 內部仍有 `WEBAPP_SECRET` 把關，不會被陌生人寫入。

---

## 階段 4：部署前端到 Vercel（5 分鐘）

### 4.1 登入 Vercel

1. 開 https://vercel.com/
2. 右上角 **「Sign Up」** 或 **「Log In」**
3. 選 **「Continue with GitHub」** → 授權

### 4.2 匯入 repo

1. Dashboard 右上角 **「Add New」→「Project」**
2. 找到 `kentsui3000/clinic` → 按 **「Import」**

### 4.3 ⚠️ 三個關鍵設定（很多人做錯）

匯入畫面會出現「Configure Project」：

| 欄位 | 填什麼 |
|---|---|
| **Project Name** | 隨便，例如 `clinic-bonus-ocr` |
| **Framework Preset** | **`Other`** |
| **Root Directory** | 點 **「Edit」→ 選 `webapp`** ⚠️ |
| Build Command | 留空 |
| Output Directory | 留空 |
| Install Command | 留空 |

按 **「Deploy」**。

### 4.4 ⚠️ 切換部署分支（如果你的程式碼不在 `main` 分支）

目前程式碼在 `claude/ocr-chinese-handwriting-app-hM5P2` 分支，不是 `main`。
Vercel 預設只部署 `main`，所以你要：

**方法 A：把分支併到 main**（推薦）
1. GitHub repo → **「Pull requests」→「New pull request」**
2. base: `main` ← compare: `claude/ocr-chinese-handwriting-app-hM5P2`
3. 建立 PR → Merge
4. Vercel 偵測到 main 更新後會自動部署

**方法 B：直接讓 Vercel 部署這個分支**
1. Vercel 專案 → **Settings → Git**
2. 「Production Branch」改成 `claude/ocr-chinese-handwriting-app-hM5P2`
3. 回到 **Deployments** → 右上角 **「Redeploy」**

### 4.5 取得網址

- 部署完成（約 30 秒）會給你一個網址：
  ```
  https://clinic-bonus-ocr.vercel.app
  ```
- 用手機開這個網址 → 應該看到綠色標題列「📋 分紅 OCR」

✅ **驗證**：能看到頁面 = Vercel 部分成功

---

## 階段 5：接上自有網域（可選，3 分鐘）

> 不做也可以，直接用 `xxx.vercel.app` 也行。要做就照下面：

### 5.1 在 Vercel 加網域

1. 專案 → **Settings → Domains**
2. 輸入 **`bonus.drkentsui.com`** → **Add**
3. Vercel 會顯示要加的 DNS 紀錄，類似：
   ```
   Type:  CNAME
   Name:  bonus
   Value: cname.vercel-dns.com
   ```

### 5.2 在你的網域 DNS 後台加紀錄

1. 到你買 `drkentsui.com` 的地方（GoDaddy / Namecheap / Cloudflare DNS / Google Domains）
2. 找 DNS 設定 → 新增紀錄：
   - 類型：**CNAME**
   - 名稱（Host）：**`bonus`**
   - 值（Value/Target）：**`cname.vercel-dns.com`**（以 Vercel 顯示的為準）
   - TTL：自動 / 1 hour
3. 儲存

### 5.3 等 DNS 生效（5–30 分鐘）

- 回到 Vercel 的 Domains 頁面，刷新看是否變成 ✅ Valid Configuration
- 開 https://bonus.drkentsui.com/ 看是否能進入 App

✅ **驗證**：網址列有 🔒 鎖頭（HTTPS 由 Vercel 自動簽發）

---

## 階段 6：手機首次設定（3 分鐘）

### 6.1 開啟 App

用手機（iPhone Safari 或 Android Chrome）開：
- `https://bonus.drkentsui.com/`（已接網域）
- 或 `https://你的專案.vercel.app`（沒接網域）

第一次會自動跳到 ⚙️ **設定頁**。

### 6.2 填三個欄位

| 欄位 | 填入 |
|---|---|
| Apps Script Web App URL | 階段 3.6 複製的 `.../exec` |
| 共用密碼 (Secret) | 階段 3.3 設定的 `WEBAPP_SECRET` |
| 試算表 ID | 階段 2 複製的 ID |

### 6.3 測試連線

1. 點 **「🔌 測試連線」**
2. 應該看到綠色提示：**「連線成功：pong」**
3. 看到後再點 **「儲存設定」**

如果失敗，看本指南最後的「故障排除」對照表。

### 6.4 加到主畫面（變成 App）

**iPhone（Safari）**
1. 下方分享按鈕 ⬆️
2. 滑動找 **「加入主畫面」**
3. 改名稱（可保留「分紅 OCR」）→ **「新增」**

**Android（Chrome）**
1. 右上角 ⋮
2. **「新增至主畫面」** 或 **「Install app」**
3. 確認

主畫面會出現綠色 `$` 圖示，點開就是全螢幕 App。

---

## 階段 7：第一次拍照辨識（2 分鐘）

1. 開 App → 點 📷 上傳區 → 拍一張清楚的帳本
   - **重點：上方的日期、班別、值日生印章一定要拍清楚**
2. 預覽出來後 → **「🔍 開始辨識」**
3. 等 5–15 秒
4. 進入校對頁，**逐欄檢查**：
   - 日期、班別（早/午/晚）正確嗎？
   - 值日生有抓到嗎？多位用半形逗號分隔
   - 員工姓名、金額有錯就直接點欄位修改
   - 不要的列點 ✕ 刪掉、漏掉的點「+ 新增一行」補上
5. 選 **「套用金額：次欄」**（這對應你既有試算表的數字）
6. 點 **「💾 儲存到試算表」**
7. 看到 ✅ **儲存成功**，回到 Google 試算表會有新建的 `2026-X` 分頁

---

## 階段 8：日常使用流程

| 時機 | 動作 |
|---|---|
| 早班結束 | 拍照 → 辨識 → 校對 → 儲存 |
| 午班結束 | 同上 |
| 晚班結束 | 同上 |
| 月底 | 開 App → 右上 📊 → 選月份 → 看每位員工總額 |

---

## 之後改程式碼怎麼辦？

### 改前端 (`webapp/`)

1. 在 GitHub 上修改檔案 (或本機 git push)
2. Vercel **自動偵測** → 重新部署 → 約 30 秒生效
3. 手機用「下拉刷新」即可看到新版

### 改後端 (`Code.gs`)

⚠️ **Apps Script 不會自動部署**，每次都要手動：

1. Apps Script 編輯器 → 改完 → **Ctrl+S**
2. 右上角 **「部署」→「管理部署作業」**
3. 點現有部署右側的 **鉛筆 ✏️**
4. **「版本」下拉 → 選「新版本」**
5. 填說明 → **「部署」**
6. ✅ URL 不會變，手機端不用改設定

---

## 故障排除

| 症狀 | 原因 | 解法 |
|---|---|---|
| 測試連線「Failed to fetch」 | URL 拷錯 / Apps Script 沒部署 | 檢查 URL 結尾是否為 `/exec`；重做階段 3.6 |
| 測試連線「密碼錯誤」 | 設定頁密碼 ≠ Code.gs 內 `WEBAPP_SECRET` | 兩邊改成完全相同 |
| 測試連線「Unauthorized」/ 401 | Apps Script「誰可以存取」沒選「任何人」 | 重新部署，正確選擇 |
| 辨識「API 錯誤 (400)」 | Gemini 金鑰失效或圖片過大 | 重貼金鑰；換較小、較清楚的照片 |
| 辨識「API 錯誤 (404)」 | 模型名稱已被 Google 下架 | 改 `Code.gs` 第 8 行 URL 為 `gemini-2.5-flash` 或 `gemini-2.5-pro` |
| 辨識結果完全空白 | 照片過暗 / 模糊 / 反光 | 重拍，或在 Apps Script 換成 `gemini-2.5-pro` |
| 辨識結果亂碼 | 模型版本太舊 | 同上，換 `gemini-2.5-pro` |
| 寫入「未提供 spreadsheetId 且 Apps Script 未綁定試算表」 | Apps Script 不是從 Sheet 開的 | 設定頁填試算表 ID（階段 2） |
| 改了 Code.gs 但行為沒變 | 沒重新「部署新版本」 | 看「之後改程式碼怎麼辦？」 |
| 試算表寫入位置不對 | 班別 → 列 對應不符你的習慣 | 改 `Code.gs` 內 `SHIFT_TO_BONUS_ROW` / `SHIFT_TO_DUTY_ROW` |
| Vercel 沒看到我的更新 | 你 push 到的分支不是 Production Branch | Vercel → Settings → Git → 改 Production Branch |

---

## 安全性與資料

- **Gemini API 金鑰** 只存在 Apps Script，前端不可見
- **`WEBAPP_SECRET`** 只存在 Apps Script + 你的手機 localStorage，URL 不會傳
- **試算表權限** 不對外開放，所有寫入都以「你」的身分執行
- **照片** 不會留在伺服器，OCR 後只保留辨識結果
- **localStorage** 上的設定可在 ⚙️ → 「匯出設定」備份

---

## 之後想加的功能（給未來的你）

- 姓名別名表（`郭雅如` ↔ `雅如`），讓矩陣分頁的列名跟既有試算表一致
- 離線佇列（Service Worker），訊號不好時先存著
- 一次拍三張（早/午/晚）連續處理
- 自動推 LINE / Email 月結通知
