/**
 * ============================================================
 * 徐嘉賢診所員工分紅自動化系統 - Apps Script
 * ============================================================
 * 版本：v1.6
 * 更新日期：2026-05-13
 * 變更紀錄：
 *   v1.6 — 新增「氣喘/濕疹評估筆數」(Peak flow/ACT/POEM,每筆 $5,
 *          計算方式鏡射自費流感:從門檻扣除+每筆獨立給價,進基礎分紅池)。
 *          Form 1 新增一欄,「分紅明細」新增 J 欄「評估筆數」。
 *          新增 migrateDailyFormAddAssessment() 給既有部署一鍵遷移。
 *   v1.5 — Form 2「病歷號」改成「可多筆」(逗號/換行/頓號分隔),
 *          腳本自動拆並產生多列分紅明細。「分紅明細」新增 I 欄「病歷號」
 *          追溯每筆分紅是哪位病人。分紅類別改集中設定 (BONUS_CATEGORIES),
 *          未來新增「衛教推廣」「SLIT」等類別只改一處。
 *   v1.4 — Form 2 欄位「患者代號」整個改名為「病歷號」;
 *          migrateMetaFormFieldNames() 同時處理此遷移。
 *   v1.3 — Form 2 欄位重新命名:「執行員工」→「執行人員」;
 *          「患者代號」說明文字改為「請填病歷號碼,避免重複收案」。
 *          新增 migrateMetaFormFieldNames() 給既有部署一鍵遷移。
 *          onMetaFormSubmit / syncEmployeesToForms 同時支援新舊名稱。
 *   v1.2 — 新增 createBothForms():一鍵自動建立兩張 Google Form,
 *          連結回應到當前 Sheets,Form ID 自動存入 ScriptProperties。
 *          syncEmployeesToForms() 改從 ScriptProperties 讀 Form ID,
 *          不再需要手動編輯程式碼。
 *   v1.1 — onDailyFormSubmit / onMetaFormSubmit 加上守衛,避免兩張
 *          Form 共用同一份 Sheets 時互相誤觸發。
 *
 * 安裝步驟(v1.2 簡化版):
 * 1. 上傳 clinic_bonus_template.xlsx 到 Google Drive,右鍵以 Sheets 開啟
 * 2. 擴充功能 → Apps Script,貼入本檔案完整內容,儲存
 * 3. 手動執行 createBothForms() 一次 → 兩張 Form 自動建立 + ID 自動儲存
 * 4. 設定 3 個觸發器(syncEmployeesToForms / onDailyFormSubmit / onMetaFormSubmit)
 * 5. 手動執行 syncEmployeesToForms() 一次完成首次員工名單同步
 * 6. 從執行記錄複製兩張 Form 的「填寫網址」給員工
 *
 * 分紅規則(已對齊,不再修改):
 *   基礎分紅 = MAX(0, 看診人數 - 流感支數 - 評估筆數 - 60) * 5
 *              + 流感支數 * 5 + 評估筆數 * 5
 *   (評估 = Peak flow/ACT/POEM,一位病人一次門診算 1 筆,進基礎分紅池)
 *   值日生津貼 = 100 元 (僅上午/下午)
 *   代謝症候群收案 = 100 元/筆
 *   代謝症候群追蹤 = 20 元/筆
 * ============================================================
 */

// ============================================================
// 設定區塊
// ============================================================

const SHEET_EMPLOYEE = '員工總表';
const SHEET_DETAIL   = '分紅明細';
const SHEET_FORM_DAILY = 'Form回應_當班';
const SHEET_FORM_META  = 'Form回應_代謝';

// 分紅參數(目前 hardcode,V1.5 可改為從「單價參數」工作表動態讀取)
const THRESHOLD_VISITS = 60;   // 看診人數門檻
const UNIT_PRICE       = 5;    // 超門檻每人單價
const FLU_PRICE        = 5;    // 自費流感每支單價
const ASSESS_PRICE     = 5;    // 氣喘/濕疹評估每筆單價 (Peak flow/ACT/POEM)
const DUTY_BONUS       = 100;  // 值日生津貼
const META_REGISTER    = 100;  // 代謝收案 (保留向下相容,實際金額讀 BONUS_CATEGORIES)
const META_FOLLOWUP    = 20;   // 代謝追蹤 (保留向下相容)

/**
 * 分紅類別中央設定 (v1.5 新增)
 * ----------------------------------------
 * 要新增類別?(例如未來加「衛教推廣」「SLIT 服務」)
 *   1. 在這裡加一筆 entry, key 必須與 Form 2「活動類型」下拉選項字串完全一致
 *   2. 到 Form 2 編輯頁手動把新選項加進「活動類型」下拉
 *   3. 到「單價參數」分頁加一列記錄該類別單價 (供日後對帳)
 *
 * 欄位說明:
 *   label       — 寫入「分紅明細」E 欄的類型名稱
 *   amount      — 每筆金額
 *   perPatient  — true: 病歷號填幾筆就產幾列分紅 (每筆獨立計算)
 *                false: 不論病歷號幾筆,只記 1 列 (用於非依病人計算的津貼)
 */
const BONUS_CATEGORIES = {
  '收案': { label: '代謝-收案', amount: 100, perPatient: true },
  '追蹤': { label: '代謝-追蹤', amount: 20,  perPatient: true }
  // 範例(將來新增):
  // '衛教推廣': { label: '衛教推廣', amount: 50, perPatient: false },
  // 'SLIT 服務': { label: 'SLIT-服務', amount: 30, perPatient: true },
};

// Form ID 從 ScriptProperties 動態讀取(由 createBothForms 自動寫入)
function getFormIds_() {
  const props = PropertiesService.getScriptProperties();
  return {
    daily: props.getProperty('FORM_ID_DAILY') || '',
    meta:  props.getProperty('FORM_ID_META')  || ''
  };
}


// ============================================================
// 函數 0:一鍵自動建立兩張 Google Form
// 觸發時機:首次部署時手動執行一次
// ============================================================

function createBothForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // 重複執行保護
  const existingDaily = props.getProperty('FORM_ID_DAILY');
  const existingMeta  = props.getProperty('FORM_ID_META');
  if (existingDaily || existingMeta) {
    const msg = [
      '⚠ Form 已建立過,跳過避免產生重複表單。',
      '  既有 Form 1 ID:' + existingDaily,
      '  既有 Form 2 ID:' + existingMeta,
      '',
      '若要重新建立,請先到 Apps Script → 專案設定 ⚙ → 指令碼屬性',
      '刪除 FORM_ID_DAILY 與 FORM_ID_META 後再執行本函式。'
    ].join('\n');
    Logger.log(msg);
    return msg;
  }

  // 預先刪掉 .xlsx 模板裡的空白佔位回應分頁(只有標題列才刪)
  [SHEET_FORM_DAILY, SHEET_FORM_META].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() <= 1) ss.deleteSheet(sh);
  });

  // === Form 1:每日當班結帳 ===
  const sheetsBefore1 = ss.getSheets().map(s => s.getName());
  const form1 = buildDailyForm_();
  form1.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  renameNewResponseSheet_(ss, sheetsBefore1, SHEET_FORM_DAILY);

  // === Form 2:代謝症候群活動記錄 ===
  const sheetsBefore2 = ss.getSheets().map(s => s.getName());
  const form2 = buildMetaForm_();
  form2.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  renameNewResponseSheet_(ss, sheetsBefore2, SHEET_FORM_META);

  // 儲存 Form ID 到 ScriptProperties
  props.setProperty('FORM_ID_DAILY', form1.getId());
  props.setProperty('FORM_ID_META',  form2.getId());

  const summary = [
    '===== Form 建立完成 =====',
    '',
    '【Form 1:每日當班結帳】',
    '  填寫網址(給員工):' + form1.getPublishedUrl(),
    '  編輯網址(管理用):' + form1.getEditUrl(),
    '  Form ID         :' + form1.getId(),
    '',
    '【Form 2:代謝症候群活動記錄】',
    '  填寫網址(給員工):' + form2.getPublishedUrl(),
    '  編輯網址(管理用):' + form2.getEditUrl(),
    '  Form ID         :' + form2.getId(),
    '',
    'Form ID 已自動儲存到 ScriptProperties。',
    '',
    '下一步:',
    '  1. 設定 3 個觸發器(syncEmployeesToForms / onDailyFormSubmit / onMetaFormSubmit)',
    '  2. 手動執行一次 syncEmployeesToForms() 完成員工名單首次同步',
    '  3. 把上方「填寫網址」用短網址服務縮短後,印 QR Code 給員工'
  ].join('\n');

  Logger.log(summary);
  return summary;
}

function buildDailyForm_() {
  const form = FormApp.create('每日當班結帳');
  form.setDescription('每個時段結帳時填一張。一天最多三張(上午/下午/晚上)。');

  form.addDateItem()
      .setTitle('日期')
      .setRequired(true);

  form.addMultipleChoiceItem()
      .setTitle('門診時段')
      .setChoiceValues(['上午', '下午', '晚上'])
      .setRequired(true);

  form.addListItem()
      .setTitle('本時段值日生')
      .setHelpText('晚上時段請選「無」(晚上沒有值日生津貼)')
      .setChoiceValues(['無'])
      .setRequired(true);

  form.addCheckboxItem()
      .setTitle('本班當值人員')
      .setHelpText('請勾選本時段在場的所有員工(含值日生本人)')
      .setChoiceValues(['(待同步)'])
      .setRequired(true);

  form.addListItem()
      .setTitle('填寫人')
      .setChoiceValues(['(待同步)'])
      .setRequired(true);

  form.addListItem()
      .setTitle('核對人')
      .setHelpText('主任或徐醫師(可後補)')
      .setChoiceValues(['(待同步)'])
      .setRequired(true);

  const visitItem = form.addTextItem()
      .setTitle('當時段有效看診人數')
      .setHelpText('本時段所有醫師合計看診人數')
      .setRequired(true);
  visitItem.setValidation(
    FormApp.createTextValidation()
      .setHelpText('請輸入大於或等於 0 的整數')
      .requireNumberGreaterThanOrEqualTo(0)
      .build()
  );

  const fluItem = form.addTextItem()
      .setTitle('自費流感疫苗支數')
      .setHelpText('沒有流感請填 0,不要留空')
      .setRequired(true);
  fluItem.setValidation(
    FormApp.createTextValidation()
      .setHelpText('請輸入大於或等於 0 的整數')
      .requireNumberGreaterThanOrEqualTo(0)
      .build()
  );

  const assessItem = form.addTextItem()
      .setTitle('氣喘/濕疹評估筆數')
      .setHelpText('Peak flow / ACT / POEM 執行人數。一位病人不管做幾項評估都算 1 筆。沒做請填 0')
      .setRequired(true);
  assessItem.setValidation(
    FormApp.createTextValidation()
      .setHelpText('請輸入大於或等於 0 的整數')
      .requireNumberGreaterThanOrEqualTo(0)
      .build()
  );

  form.addTextItem()
      .setTitle('備註')
      .setRequired(false);

  // 安全 & 顯示設定
  form.setAllowResponseEdits(false);
  form.setPublishingSummary(false);
  form.setShowLinkToRespondAgain(false);
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setConfirmationMessage('已收到,謝謝!');
  form.setProgressBar(true);
  return form;
}

function buildMetaForm_() {
  const form = FormApp.create('代謝症候群活動記錄');
  form.setDescription('同一人同一活動類型多筆病歷號可一張填完(病歷號用逗號分隔)。\n不同人或不同活動類型請分開填。');

  form.addDateItem()
      .setTitle('日期')
      .setRequired(true);

  form.addMultipleChoiceItem()
      .setTitle('時段')
      .setChoiceValues(['上午', '下午', '晚上'])
      .setRequired(true);

  form.addMultipleChoiceItem()
      .setTitle('活動類型')
      .setChoiceValues(['收案', '追蹤'])
      .setRequired(true);

  form.addListItem()
      .setTitle('執行人員')
      .setChoiceValues(['(待同步)'])
      .setRequired(true);

  form.addTextItem()
      .setTitle('病歷號')
      .setHelpText('多筆請用逗號分隔(例如:001,002,003),每筆獨立計算分紅')
      .setRequired(false);

  form.addTextItem()
      .setTitle('備註')
      .setRequired(false);

  form.setAllowResponseEdits(false);
  form.setPublishingSummary(false);
  form.setShowLinkToRespondAgain(false);
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setConfirmationMessage('已收到,謝謝!');
  form.setProgressBar(true);
  return form;
}

function renameNewResponseSheet_(ss, sheetsBefore, targetName) {
  const beforeSet = new Set(sheetsBefore);
  const newSheet = ss.getSheets().find(s => !beforeSet.has(s.getName()));
  if (newSheet) {
    newSheet.setName(targetName);
    Logger.log('已將自動產生的回應分頁改名為「' + targetName + '」');
  } else {
    Logger.log('警告:找不到新增的回應分頁,目標名稱「' + targetName + '」未套用');
  }
}


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

  const ids = getFormIds_();

  if (ids.daily) {
    updateFormChoices(ids.daily, {
      '本班當值人員': { type: 'checkbox', list: active },
      '本時段值日生': { type: 'list', list: ['無'].concat(active) },
      '填寫人': { type: 'list', list: active },
      '核對人': { type: 'list', list: active }
    });
  } else {
    Logger.log('警告:FORM_ID_DAILY 未設定,請先執行 createBothForms()');
  }

  if (ids.meta) {
    // 同時支援新舊欄位名稱(v1.3 改名「執行員工」→「執行人員」)
    updateFormChoices(ids.meta, {
      '執行人員': { type: 'list', list: active },
      '執行員工': { type: 'list', list: active }
    });
  } else {
    Logger.log('警告:FORM_ID_META 未設定,請先執行 createBothForms()');
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
    // v1.6:氣喘/濕疹評估 (Peak flow/ACT/POEM)。舊 Form 沒有此欄位時視為 0。
    const assess   = (r['氣喘/濕疹評估筆數'] && r['氣喘/濕疹評估筆數'][0])
                   ? (parseFloat(r['氣喘/濕疹評估筆數'][0]) || 0) : 0;
    const duty     = (r['本時段值日生'] && r['本時段值日生'][0]) ? r['本時段值日生'][0].trim() : '';

    // 方案 Y:非門檻類 (流感+評估) 從總人數扣掉再算超門檻獎金,
    // 非門檻類每筆固定給價,不受 60 人門檻限制
    const baseBonus = Math.max(0, visits - flu - assess - THRESHOLD_VISITS) * UNIT_PRICE
                    + flu * FLU_PRICE
                    + assess * ASSESS_PRICE;

    const detail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DETAIL);
    if (!detail) {
      Logger.log('錯誤:找不到「' + SHEET_DETAIL + '」工作表');
      return;
    }
    ensureDetailHeaders_(detail);

    const now = new Date();
    const rows = [];

    // 1. 每位當班員工的基礎分紅(僅當金額 > 0 時寫入,避免 0 元的空紀錄)
    if (baseBonus > 0) {
      employees.forEach(emp => {
        rows.push([now, date, period, emp, '基礎分紅', baseBonus, visits, flu, '', assess]);
      });
    }

    // 2. 值日生津貼(僅上午/下午、且不是「無」)
    if (duty && duty !== '無' && (period === '上午' || period === '下午')) {
      rows.push([now, date, period, duty, '值日生津貼', DUTY_BONUS, '', '', '', '']);
    }

    if (rows.length > 0) {
      detail.getRange(detail.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
    }

    Logger.log('當班結帳處理完成:' + date + ' ' + period + ',基礎' + baseBonus + '元/人 × ' + employees.length + '人' +
               '(看診' + visits + '/流感' + flu + '/評估' + assess + ')' +
               (duty && duty !== '無' ? ',值日生' + duty : ''));
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
    // 守衛:用 Form 2 獨有欄位「活動類型」+「執行人員/員工」判斷是否該由本函式處理。
    if (!r || !r['活動類型'] || !(r['執行人員'] || r['執行員工'])) return;

    const date = r['日期'][0];
    const period = r['時段'] && r['時段'][0] ? r['時段'][0] : '';
    const type = String(r['活動類型'][0]).trim();
    const emp  = (r['執行人員'] && r['執行人員'][0]) || r['執行員工'][0];
    const chartRaw = (r['病歷號'] && r['病歷號'][0])
                   || (r['患者代號'] && r['患者代號'][0])
                   || '';

    // 從中央設定找對應類別 (支援部分字串相符,例如「收案」匹配「收案」key)
    let category = null;
    for (const key in BONUS_CATEGORIES) {
      if (type === key || type.indexOf(key) !== -1) {
        category = BONUS_CATEGORIES[key];
        break;
      }
    }
    if (!category) {
      Logger.log('警告:無法辨識的活動類型「' + type + '」。' +
                 '請檢查 BONUS_CATEGORIES 是否需新增此類別。跳過此筆。');
      return;
    }

    // 拆病歷號 (支援半形逗號、全形逗號、頓號、換行)
    const chartNos = chartRaw.split(/[,\n，、]/).map(s => s.trim()).filter(s => s);

    // 決定要寫幾列:perPatient 類別依病歷號筆數;否則固定 1 列
    const count = category.perPatient ? Math.max(1, chartNos.length) : 1;

    const detail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DETAIL);
    if (!detail) {
      Logger.log('錯誤:找不到「' + SHEET_DETAIL + '」工作表');
      return;
    }

    // 確保 I 欄「病歷號」、J 欄「評估筆數」標題存在 (v1.6 schema)
    ensureDetailHeaders_(detail);

    const now = new Date();
    const rows = [];
    for (let i = 0; i < count; i++) {
      const chartNo = chartNos[i] || '';
      rows.push([now, date, period, emp, category.label, category.amount,
                 '', '', chartNo, '']);
    }
    detail.getRange(detail.getLastRow() + 1, 1, rows.length, 10).setValues(rows);

    const total = category.amount * count;
    Logger.log('代謝症候群處理完成:' + date + ' ' + emp + ' ' + category.label +
               ' x' + count + ' = ' + total + '元' +
               (chartNos.length ? '(病歷號 ' + chartNos.join(',') + ')' : ''));
  } catch (err) {
    Logger.log('onMetaFormSubmit 錯誤:' + err.message + '\n' + err.stack);
  }
}

function ensureDetailHeaders_(detail) {
  if (detail.getRange(1, 9).getValue() !== '病歷號') {
    detail.getRange(1, 9).setValue('病歷號');
  }
  if (detail.getRange(1, 10).getValue() !== '評估筆數') {
    detail.getRange(1, 10).setValue('評估筆數');
  }
}


// ============================================================
// 輔助函數:手動驗證計算邏輯(在 Apps Script 編輯器執行)
// ============================================================

function testCalculation() {
  const cases = [
    { visits: 50,  flu: 0,  assess: 0, expected: 0   },
    { visits: 50,  flu: 10, assess: 0, expected: 50  },
    { visits: 61,  flu: 0,  assess: 0, expected: 5   },
    { visits: 80,  flu: 0,  assess: 0, expected: 100 },
    { visits: 80,  flu: 10, assess: 0, expected: 100 },
    { visits: 100, flu: 20, assess: 0, expected: 200 },
    // v1.6:氣喘/濕疹評估 (行為鏡射流感)
    { visits: 50,  flu: 0,  assess: 3, expected: 15  },  // 未達門檻,3 筆評估仍給 $15
    { visits: 80,  flu: 10, assess: 3, expected: 100 },  // 過門檻,評估扣門檻+自身給價剛好抵消
    { visits: 60,  flu: 0,  assess: 5, expected: 25  },  // 卡門檻,5 筆評估帶來 $25
    { visits: 80,  flu: 5,  assess: 5, expected: 100 }   // 流感+評估混合
  ];

  Logger.log('=== 基礎分紅計算驗證 ===');
  cases.forEach(c => {
    const bonus = Math.max(0, c.visits - c.flu - c.assess - THRESHOLD_VISITS) * UNIT_PRICE
                + c.flu * FLU_PRICE
                + c.assess * ASSESS_PRICE;
    const result = bonus === c.expected ? '✓' : '✗';
    Logger.log(result + ' 看診' + c.visits + '/流感' + c.flu + '/評估' + c.assess +
               ' → ' + bonus + '元(預期 ' + c.expected + '元)');
  });
}


// ============================================================
// 輔助函數:遷移既有 Form 1 新增「氣喘/濕疹評估筆數」欄位(v1.6)
// 用法:在 Apps Script 編輯器手動執行一次即可
// ============================================================

function migrateDailyFormAddAssessment() {
  const ids = getFormIds_();
  if (!ids.daily) {
    Logger.log('錯誤:找不到 Form 1 ID,請先執行 createBothForms()');
    return;
  }

  const form = FormApp.openById(ids.daily);
  const items = form.getItems();

  // 已存在就跳過(可重複執行)
  const exists = items.some(item => item.getTitle().trim() === '氣喘/濕疹評估筆數');
  if (exists) {
    Logger.log('「氣喘/濕疹評估筆數」欄位已存在,跳過。');
    return;
  }

  // 新增欄位(先加在最後)
  const assessItem = form.addTextItem()
      .setTitle('氣喘/濕疹評估筆數')
      .setHelpText('Peak flow / ACT / POEM 執行人數。一位病人不管做幾項評估都算 1 筆。沒做請填 0')
      .setRequired(true);
  assessItem.setValidation(
    FormApp.createTextValidation()
      .setHelpText('請輸入大於或等於 0 的整數')
      .requireNumberGreaterThanOrEqualTo(0)
      .build()
  );

  // 移到「自費流感疫苗支數」正下方
  const fluIndex = form.getItems().findIndex(item => item.getTitle().trim() === '自費流感疫苗支數');
  if (fluIndex !== -1) {
    form.moveItem(assessItem.getIndex(), fluIndex + 1);
    Logger.log('✓ 已新增「氣喘/濕疹評估筆數」並移到「自費流感疫苗支數」下方');
  } else {
    Logger.log('✓ 已新增「氣喘/濕疹評估筆數」(找不到流感欄位,保持在表單最後)');
  }

  Logger.log('遷移完成。請開 Form 1 預覽確認欄位位置與說明文字。');
  Logger.log('注意:「Form回應_當班」分頁會自動在最右邊新增一欄收此值,不需手動處理。');
}


// ============================================================
// 輔助函數:遷移既有 Form 2 欄位名稱(v1.3 改名)
// 用法:在 Apps Script 編輯器手動執行一次即可
// ============================================================

function migrateMetaFormFieldNames() {
  const ids = getFormIds_();
  if (!ids.meta) {
    Logger.log('錯誤:找不到 Form 2 ID,請先執行 createBothForms()');
    return;
  }

  const form = FormApp.openById(ids.meta);
  let changed = 0;

  // 更新表單描述為 v1.5 多筆病歷號版
  const newDesc = '同一人同一活動類型多筆病歷號可一張填完(病歷號用逗號分隔)。\n不同人或不同活動類型請分開填。';
  if (form.getDescription() !== newDesc) {
    form.setDescription(newDesc);
    Logger.log('✓ Form 2 表單說明已更新為 v1.5 多筆版');
    changed++;
  }

  form.getItems().forEach(item => {
    const title = item.getTitle().trim();

    if (title === '執行員工') {
      item.setTitle('執行人員');
      Logger.log('✓ 欄位標題:「執行員工」→「執行人員」');
      changed++;
    }

    if (title === '患者代號' || title === '病歷號') {
      if (title === '患者代號') {
        item.setTitle('病歷號');
        Logger.log('✓ 欄位標題:「患者代號」→「病歷號」');
      }
      item.setHelpText('多筆請用逗號分隔(例如:001,002,003),每筆獨立計算分紅');
      Logger.log('✓ 「病歷號」說明文字已更新(v1.5 支援多筆)');
      changed++;
    }
  });

  if (changed === 0) {
    Logger.log('沒有發現需要遷移的欄位(可能已是 v1.3 新名)。');
  } else {
    Logger.log('遷移完成,共更新 ' + changed + ' 處。');
    Logger.log('建議接著手動執行一次 syncEmployeesToForms() 重新同步員工名單。');
  }
}


// ============================================================
// 輔助函數:印出兩張 Form 的填寫網址(忘記時用)
// ============================================================

function printFormUrls() {
  const ids = getFormIds_();
  if (!ids.daily && !ids.meta) {
    Logger.log('尚未建立 Form,請先執行 createBothForms()');
    return;
  }
  if (ids.daily) {
    const f = FormApp.openById(ids.daily);
    Logger.log('Form 1 填寫:' + f.getPublishedUrl());
    Logger.log('Form 1 編輯:' + f.getEditUrl());
  }
  if (ids.meta) {
    const f = FormApp.openById(ids.meta);
    Logger.log('Form 2 填寫:' + f.getPublishedUrl());
    Logger.log('Form 2 編輯:' + f.getEditUrl());
  }
}
