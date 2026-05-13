/**
 * ============================================================
 * 徐嘉賢診所員工分紅自動化系統 - Apps Script
 * ============================================================
 * 版本：v1.1
 * 更新日期：2026-05-13
 * 變更：v1.1 — onDailyFormSubmit / onMetaFormSubmit 加上守衛,
 *              避免兩張 Form 共用同一份 Sheets 時互相誤觸發。
 *
 * 安裝步驟：
 * 1. 在 Google Sheets 開啟 擴充功能 → Apps Script
 * 2. 把整份檔案內容貼入 Code.gs(取代預設內容)
 * 3. 填入下方 FORM_ID_DAILY 與 FORM_ID_META(Form 編輯網址 /d/ 之後那串)
 * 4. 儲存後設定觸發器:
 *    - syncEmployeesToForms: 試算表編輯時
 *    - onDailyFormSubmit: 來自 Form 1 提交時
 *    - onMetaFormSubmit: 來自 Form 2 提交時
 * 5. 手動執行一次 syncEmployeesToForms() 完成首次同步
 *
 * 分紅規則(已對齊,不再修改):
 *   基礎分紅 = MAX(0, 看診人數 - 流感支數 - 60) * 5 + 流感支數 * 5
 *   值日生津貼 = 100 元 (僅上午/下午)
 *   代謝症候群收案 = 100 元/筆
 *   代謝症候群追蹤 = 20 元/筆
 * ============================================================
 */

// ============================================================
// 設定區塊
// ============================================================

const FORM_ID_DAILY = '在這裡填入Form1的ID';   // 每日當班結帳 Form
const FORM_ID_META  = '在這裡填入Form2的ID';   // 代謝症候群活動記錄 Form

const SHEET_EMPLOYEE = '員工總表';
const SHEET_DETAIL   = '分紅明細';

// 分紅參數(目前 hardcode,V1.5 可改為從「單價參數」工作表動態讀取)
const THRESHOLD_VISITS = 60;   // 看診人數門檻
const UNIT_PRICE       = 5;    // 超門檻每人單價
const FLU_PRICE        = 5;    // 自費流感每支單價
const DUTY_BONUS       = 100;  // 值日生津貼
const META_REGISTER    = 100;  // 代謝收案
const META_FOLLOWUP    = 20;   // 代謝追蹤


// ============================================================
// 函數 1:同步員工名單到兩張 Form 的下拉選項
// 觸發時機:員工總表編輯時
// ============================================================

function syncEmployeesToForms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEE);
  if (!sheet) {
    Logger.log('錯誤:找不到「' + SHEET_EMPLOYEE + '」工作表');
    return;
  }

  const data = sheet.getDataRange().getValues();
  // 第 1 列是標題,第 2 列起為員工資料
  // 欄位:A=姓名, B=職位, C=在職(✓), D=加入日期, E=離職日期
  const active = data
    .slice(1)
    .filter(row => row[2] === '✓' || row[2] === true)
    .map(row => String(row[0]).trim())
    .filter(name => name !== '');

  if (active.length === 0) {
    Logger.log('警告:目前沒有任何在職員工,Form 下拉選項將為空');
    return;
  }

  // 同步 Form 1
  if (FORM_ID_DAILY && FORM_ID_DAILY !== '在這裡填入Form1的ID') {
    updateFormChoices(FORM_ID_DAILY, {
      '本班當值人員': { type: 'checkbox', list: active },
      '本時段值日生': { type: 'list', list: ['無'].concat(active) },
      '填寫人': { type: 'list', list: active },
      '核對人': { type: 'list', list: active }
    });
  }

  // 同步 Form 2
  if (FORM_ID_META && FORM_ID_META !== '在這裡填入Form2的ID') {
    updateFormChoices(FORM_ID_META, {
      '執行員工': { type: 'list', list: active }
    });
  }

  Logger.log('員工名單同步完成,目前在職:' + active.join('、'));
}

function updateFormChoices(formId, mapping) {
  try {
    const form = FormApp.openById(formId);
    form.getItems().forEach(item => {
      const title = item.getTitle().trim();
      const conf = mapping[title];
      if (!conf) return;
      if (conf.type === 'checkbox') {
        item.asCheckboxItem().setChoiceValues(conf.list);
      } else if (conf.type === 'list') {
        item.asListItem().setChoiceValues(conf.list);
      }
    });
  } catch (err) {
    Logger.log('更新 Form ' + formId + ' 失敗:' + err.message);
  }
}

// 試算表編輯觸發器:當員工總表被編輯時自動同步
function onEdit(e) {
  if (!e || !e.source) return;
  const sheetName = e.source.getActiveSheet().getName();
  if (sheetName === SHEET_EMPLOYEE) {
    syncEmployeesToForms();
  }
}


// ============================================================
// 函數 2:Form 1 提交時自動拆人計算
// 觸發時機:Form 1(每日當班結帳)送出
// ============================================================

function onDailyFormSubmit(e) {
  try {
    const r = e.namedValues;
    // 守衛:兩張 Form 都寫入同一份 Sheets,每次提交兩個 onSubmit 函式都會被觸發。
    // 用 Form 1 獨有欄位「本班當值人員」判斷是否該由本函式處理。
    if (!r || !r['本班當值人員']) return;

    const date     = r['日期'][0];
    const period   = r['門診時段'][0];
    const employees = String(r['本班當值人員'][0]).split(', ').map(s => s.trim()).filter(s => s);
    const visits   = parseFloat(r['當時段有效看診人數'][0]) || 0;
    const flu      = parseFloat(r['自費流感疫苗支數'][0]) || 0;
    const duty     = (r['本時段值日生'] && r['本時段值日生'][0]) ? r['本時段值日生'][0].trim() : '';

    // 方案 Y:流感患者從總人數扣掉再算超門檻獎金
    const baseBonus = Math.max(0, visits - flu - THRESHOLD_VISITS) * UNIT_PRICE + flu * FLU_PRICE;

    const detail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DETAIL);
    if (!detail) {
      Logger.log('錯誤:找不到「' + SHEET_DETAIL + '」工作表');
      return;
    }

    const now = new Date();
    const rows = [];

    // 1. 每位當班員工的基礎分紅(僅當金額 > 0 時寫入,避免 0 元的空紀錄)
    if (baseBonus > 0) {
      employees.forEach(emp => {
        rows.push([now, date, period, emp, '基礎分紅', baseBonus, visits, flu]);
      });
    }

    // 2. 值日生津貼(僅上午/下午、且不是「無」)
    if (duty && duty !== '無' && (period === '上午' || period === '下午')) {
      rows.push([now, date, period, duty, '值日生津貼', DUTY_BONUS, '', '']);
    }

    if (rows.length > 0) {
      detail.getRange(detail.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
    }

    Logger.log('當班結帳處理完成:' + date + ' ' + period + ',基礎' + baseBonus + '元/人 × ' + employees.length + '人' + (duty && duty !== '無' ? ',值日生' + duty : ''));
  } catch (err) {
    Logger.log('onDailyFormSubmit 錯誤:' + err.message + '\n' + err.stack);
  }
}


// ============================================================
// 函數 3:Form 2 提交時寫入代謝症候群分紅
// 觸發時機:Form 2(代謝症候群活動記錄)送出
// ============================================================

function onMetaFormSubmit(e) {
  try {
    const r = e.namedValues;
    // 守衛:用 Form 2 獨有欄位「活動類型」+「執行員工」判斷是否該由本函式處理。
    if (!r || !r['活動類型'] || !r['執行員工']) return;

    const date = r['日期'][0];
    const period = r['時段'] && r['時段'][0] ? r['時段'][0] : '';
    const type = r['活動類型'][0];
    const emp  = r['執行員工'][0];

    let amount = 0;
    let label = '';
    if (type.includes('收案')) {
      amount = META_REGISTER;
      label = '代謝-收案';
    } else if (type.includes('追蹤')) {
      amount = META_FOLLOWUP;
      label = '代謝-追蹤';
    } else {
      Logger.log('警告:無法辨識的活動類型「' + type + '」,跳過');
      return;
    }

    const detail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DETAIL);
    detail.appendRow([new Date(), date, period, emp, label, amount, '', '']);

    Logger.log('代謝症候群處理完成:' + date + ' ' + emp + ' ' + label + ' ' + amount + '元');
  } catch (err) {
    Logger.log('onMetaFormSubmit 錯誤:' + err.message + '\n' + err.stack);
  }
}


// ============================================================
// 輔助函數:手動驗證計算邏輯(在 Apps Script 編輯器執行)
// ============================================================

function testCalculation() {
  const cases = [
    { visits: 50, flu: 0,  expected: 0   },
    { visits: 50, flu: 10, expected: 50  },
    { visits: 61, flu: 0,  expected: 5   },
    { visits: 80, flu: 0,  expected: 100 },
    { visits: 80, flu: 10, expected: 100 },
    { visits: 100, flu: 20, expected: 200 }
  ];

  Logger.log('=== 基礎分紅計算驗證 ===');
  cases.forEach(c => {
    const bonus = Math.max(0, c.visits - c.flu - THRESHOLD_VISITS) * UNIT_PRICE + c.flu * FLU_PRICE;
    const result = bonus === c.expected ? '✓' : '✗';
    Logger.log(result + ' 看診' + c.visits + '人/流感' + c.flu + '支 → ' + bonus + '元(預期 ' + c.expected + '元)');
  });
}
