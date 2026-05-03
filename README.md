# 診所分紅 OCR 助手 (Clinic Bonus & Duty OCR)

一個用手機即可拍照、辨識手寫帳本、並自動寫入 Google 試算表的小工具。

```
┌─────────────┐    ┌─────────────────┐    ┌──────────────────┐    ┌────────────────┐
│  iPhone /   │───▶│  webapp/        │───▶│  Apps Script     │───▶│  Google Sheet  │
│  Android    │    │  (靜態網頁)      │    │  doPost()        │    │  矩陣分頁       │
│  拍照        │    │  drkentsui.com  │    │  Gemini Vision   │    │  YYYY-M        │
└─────────────┘    └─────────────────┘    └──────────────────┘    └────────────────┘
```

## 功能

- 📷 手機拍照 / 從相簿選圖（前端會自動壓縮到 1800px 寬）
- 🤖 透過 **Gemini 2.5 Pro / Vision** 辨識：
  - 日期 (`2026.04.06 W一(早)` → `2026-04-06` + `morning`)
  - 班別（早 = morning / 午 = afternoon / 晚 = night）
  - 值日生（含印章 `江芯儀RN` → `江芯儀`）
  - 每位員工的姓名、主分紅、次分紅、備註
- ✏️ 辨識後可在手機上即時校對／修正
- 💾 一鍵寫入 Google 試算表的矩陣式月份分頁
  - 每位員工 5 列：`AM Duty / PM Duty / A / B / C`
  - 自動建立月份分頁、自動填入日期、星期、SUM 公式
- 📊 隨時查詢「本月每位員工分紅 / 值日 / 總計」
- 📱 支援 PWA — 加到主畫面就像原生 App
- 🔒 透過共用密碼保護，避免任意人寫入你的試算表

## 專案結構

```
clinic/
├── Code.gs                 # Apps Script 後端（OCR + 寫入試算表）
├── index.html              # 既有：Apps Script 側邊欄 UI（仍可使用）
├── webapp/                 # 🆕 行動網頁（部署到 drkentsui.com）
│   ├── index.html          #     主頁
│   ├── app.js              #     前端邏輯
│   ├── style.css           #     樣式
│   ├── manifest.webmanifest#     PWA manifest
│   └── icon.svg            #     圖示
├── README.md               # 本文件
└── DEPLOYMENT.md           # 詳細部署步驟（強烈建議照做一次）
```

## 快速部署 (5 分鐘)

詳見 [`DEPLOYMENT.md`](./DEPLOYMENT.md)。簡述如下：

1. **Apps Script**
   - 把整個 `Code.gs` 貼到 https://script.google.com/
   - 把 `GEMINI_API_KEY` 換成你的 Gemini API Key（`https://aistudio.google.com/apikey`）
   - 把 `WEBAPP_SECRET` 改成你自己的密碼
   - 「部署」→「新增部署」→ 類型「網頁應用程式」→ 執行身分「我」→ 存取權限「任何人」
   - 拷貝部署後得到的 `https://script.google.com/macros/s/.../exec`

2. **前端（Cloudflare Pages 範例）**
   - 把 `webapp/` 整個資料夾上傳到 Cloudflare Pages（或 Vercel / Netlify / GitHub Pages）
   - 在 DNS 新增 `bonus.drkentsui.com` 指向 Pages
   - 用手機開啟 `https://bonus.drkentsui.com`
   - 第一次：點右上角 ⚙️ → 貼上 Apps Script URL + 密碼 → 測試連線

3. **每天用法**
   - 班別交班時拍一張帳本照片
   - 開 Web App → 拍照 → 開始辨識 → 校對 → 儲存
   - 月底開「📊 統計」看誰拿多少

## 班別 → 試算表列對應

| 拍照標頭 | shift 值 | 分紅寫到 | 值日生寫到 |
|---|---|---|---|
| `(早)` | `morning` | `{name} A` | `{name} AM Duty` |
| `(午)` | `afternoon` | `{name} B` | `{name} PM Duty` |
| `(晚)` | `night` | `{name} C` | `{name} PM Duty` |

> 若想改規則，編輯 `Code.gs` 內 `SHIFT_TO_BONUS_ROW` / `SHIFT_TO_DUTY_ROW`。

## 「主欄 / 次欄 / 主+次」是什麼？

帳本上每位員工常有 2 欄數字：

- **主欄**：固定 200/150/80/0 等（這是基本分紅）
- **次欄**：變動 5/15/45/100/130/205 等（這是看診加給／看一次給一點）

對照你既有的 `bonus calculation` 試算表，可以發現裡面填的是**次欄**的數字。
所以前端預設「套用金額：次欄」，但你也可以改成「主欄」或「主+次」。

## 已知限制

- Gemini Vision 對手寫中文約 80–95% 命中率，**請務必在儲存前人工校對**
- Apps Script Web App 單次請求約 6 MB 上限 — 前端已壓縮到約 300–600 KB
- Apps Script 每天有執行時間配額（一般帳號 90 分鐘 / 天綽綽有餘）

## 安全性

- API Key 與 `WEBAPP_SECRET` 都只存在 Apps Script 後端（前端不會看到 Gemini Key）
- 前端的 URL + 密碼存在瀏覽器 `localStorage`
- 所有寫入都以**你**的身分執行（試算表權限不需另外開放）

## 既有的 Sheets 側邊欄版本仍可用

`index.html` 是綁在 Google Sheet 上的 Apps Script 側邊欄 UI（舊功能：打卡紙 + 帳本辨識 + 員工管理 + 月結），仍可繼續使用，與新版 Web App **並存**。
