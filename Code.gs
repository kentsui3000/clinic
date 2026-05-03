/**
 * 診所薪資自動化系統 - Google Apps Script
 * 功能：透過 Gemini Vision API 辨識打卡紙與手寫帳本，自動計算薪資
 */

// ==================== 設定區 ====================
const GEMINI_API_KEY = 'YOUR_KEY_HERE'; // 請替換為您的 Gemini API Key
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent';

// 分頁名稱常數
const SHEET_NAMES = {
  CLOCK_IN: '打卡紀錄',
  BONUS_DUTY: '分紅與值班',
  HOLIDAYS: '國定假日',
  PAYROLL: '薪資計算結果',
  EMPLOYEES: '員工名單'
};

// ==================== 選單與初始化 ====================

/**
 * 當試算表開啟時建立自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 智慧薪資助理')
    .addItem('📤 開啟上傳介面', 'showSidebar')
    .addSeparator()
    .addItem('🔧 初始化系統 (建立分頁)', 'initializeSheets')
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 薪資計算')
      .addItem('執行當月結算', 'promptCalculatePayroll'))
    .addToUi();
}

/**
 * 顯示側邊欄
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('智慧薪資助理')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 取得員工名單供前端使用
 */
function getEmployeeList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);

  if (!sheet) {
    // 如果分頁不存在，建立並返回預設員工
    createSheetIfNotExists(ss, SHEET_NAMES.EMPLOYEES, ['員工姓名', '建立日期']);
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(row => row[0]).filter(name => name !== '');
}

/**
 * 新增員工
 * @param {string} employeeName - 員工姓名
 * @returns {object} - 處理結果
 */
function addEmployee(employeeName) {
  try {
    if (!employeeName || employeeName.trim() === '') {
      return { success: false, message: '員工姓名不能為空' };
    }

    const name = employeeName.trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 確保分頁存在
    createSheetIfNotExists(ss, SHEET_NAMES.EMPLOYEES, ['員工姓名', '建立日期']);
    const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);

    // 檢查是否已存在
    const existingEmployees = getEmployeeList();
    if (existingEmployees.includes(name)) {
      return { success: false, message: '員工已存在：' + name };
    }

    // 新增員工
    const lastRow = sheet.getLastRow();
    const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[name, today]]);

    return { success: true, message: '成功新增員工：' + name };
  } catch (error) {
    Logger.log('addEmployee Error: ' + error.toString());
    return { success: false, message: '新增員工時發生錯誤：' + error.toString() };
  }
}

/**
 * 更新員工姓名
 * @param {string} oldName - 舊姓名
 * @param {string} newName - 新姓名
 * @returns {object} - 處理結果
 */
function updateEmployee(oldName, newName) {
  try {
    if (!oldName || !newName || newName.trim() === '') {
      return { success: false, message: '員工姓名不能為空' };
    }

    const newNameTrimmed = newName.trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);

    if (!sheet) {
      return { success: false, message: '員工名單分頁不存在' };
    }

    // 檢查新名稱是否已存在
    const existingEmployees = getEmployeeList();
    if (existingEmployees.includes(newNameTrimmed) && newNameTrimmed !== oldName) {
      return { success: false, message: '員工姓名已存在：' + newNameTrimmed };
    }

    // 找到並更新
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: '找不到員工：' + oldName };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let found = false;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === oldName) {
        sheet.getRange(i + 2, 1).setValue(newNameTrimmed);
        found = true;
        break;
      }
    }

    if (!found) {
      return { success: false, message: '找不到員工：' + oldName };
    }

    // 同步更新其他分頁中的員工姓名
    updateEmployeeNameInAllSheets(oldName, newNameTrimmed);

    return { success: true, message: '成功更新員工姓名：' + oldName + ' → ' + newNameTrimmed };
  } catch (error) {
    Logger.log('updateEmployee Error: ' + error.toString());
    return { success: false, message: '更新員工時發生錯誤：' + error.toString() };
  }
}

/**
 * 刪除員工
 * @param {string} employeeName - 員工姓名
 * @returns {object} - 處理結果
 */
function deleteEmployee(employeeName) {
  try {
    if (!employeeName) {
      return { success: false, message: '員工姓名不能為空' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);

    if (!sheet) {
      return { success: false, message: '員工名單分頁不存在' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: '找不到員工：' + employeeName };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowToDelete = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === employeeName) {
        rowToDelete = i + 2; // +2 因為標題列和 0-index
        break;
      }
    }

    if (rowToDelete === -1) {
      return { success: false, message: '找不到員工：' + employeeName };
    }

    sheet.deleteRow(rowToDelete);

    return { success: true, message: '成功刪除員工：' + employeeName };
  } catch (error) {
    Logger.log('deleteEmployee Error: ' + error.toString());
    return { success: false, message: '刪除員工時發生錯誤：' + error.toString() };
  }
}

/**
 * 同步更新所有分頁中的員工姓名
 */
function updateEmployeeNameInAllSheets(oldName, newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 更新打卡紀錄
  const clockInSheet = ss.getSheetByName(SHEET_NAMES.CLOCK_IN);
  if (clockInSheet) {
    updateNameInColumn(clockInSheet, 1, oldName, newName);
  }

  // 更新分紅與值班
  const bonusSheet = ss.getSheetByName(SHEET_NAMES.BONUS_DUTY);
  if (bonusSheet) {
    updateNameInColumn(bonusSheet, 2, oldName, newName);
  }

  // 更新薪資計算結果
  const payrollSheet = ss.getSheetByName(SHEET_NAMES.PAYROLL);
  if (payrollSheet) {
    updateNameInColumn(payrollSheet, 3, oldName, newName);
  }
}

/**
 * 更新指定欄位中的姓名
 */
function updateNameInColumn(sheet, column, oldName, newName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, column, lastRow - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === oldName) {
      sheet.getRange(i + 2, column).setValue(newName);
    }
  }
}

// ==================== 分頁初始化 ====================

/**
 * 初始化所有必要的分頁結構
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 打卡紀錄
  createSheetIfNotExists(ss, SHEET_NAMES.CLOCK_IN,
    ['員工姓名', '日期', '上班時間', '下班時間', '當日工時(小時)']);

  // 2. 分紅與值班
  createSheetIfNotExists(ss, SHEET_NAMES.BONUS_DUTY,
    ['日期', '員工姓名', '分紅金額', '值日費', '備註']);

  // 3. 國定假日
  createSheetIfNotExists(ss, SHEET_NAMES.HOLIDAYS,
    ['日期', '假日名稱']);

  // 4. 薪資計算結果
  createSheetIfNotExists(ss, SHEET_NAMES.PAYROLL,
    ['年份', '月份', '員工姓名', '公司應上工時', '實際上班工時', '時數差異(抵扣額度)', '總分紅', '總值日費']);

  // 5. 員工名單
  createSheetIfNotExists(ss, SHEET_NAMES.EMPLOYEES,
    ['員工姓名', '建立日期']);

  SpreadsheetApp.getUi().alert('✅ 系統初始化完成！\n已建立/確認所有必要分頁。');
}

/**
 * 建立分頁（若不存在）
 */
function createSheetIfNotExists(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log(`已建立分頁: ${sheetName}`);
  } else {
    // 確認標題列存在
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (existingHeaders[0] !== headers[0]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    Logger.log(`分頁已存在: ${sheetName}`);
  }

  return sheet;
}

// ==================== Function A: 處理打卡紙 ====================

/**
 * 處理打卡紙圖片辨識
 * @param {string} base64Image - 圖片的 Base64 編碼
 * @param {string} employeeName - 員工姓名
 * @returns {object} - 處理結果
 */
function processTimeCardImage(base64Image, employeeName) {
  try {
    // 確保分頁存在
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    createSheetIfNotExists(ss, SHEET_NAMES.CLOCK_IN,
      ['員工姓名', '日期', '上班時間', '下班時間', '當日工時(小時)']);

    const prompt = `你是一個專門辨識傳統打卡紙的 AI 助手。請仔細分析這張打卡紙圖片，執行以下任務：

1. 辨識每一行的打卡記錄
2. 提取以下資訊：
   - 日期：轉換為 YYYY-MM-DD 格式
   - 最早打卡時間 (In)：格式為 HH:MM
   - 最晚打卡時間 (Out)：格式為 HH:MM

規則：
- 忽略週六、週日的記錄
- 忽略沒有任何打卡記錄的行
- 如果某天只有一個打卡記錄，In 和 Out 設為相同時間
- 忽略任何無法辨識或不清楚的行

請以純 JSON Array 格式回傳，不要包含任何其他文字或 markdown 標記：
[
  {"date": "2024-01-02", "timeIn": "08:30", "timeOut": "17:30"},
  {"date": "2024-01-03", "timeIn": "08:25", "timeOut": "18:00"}
]

如果無法辨識任何有效記錄，請回傳空陣列：[]`;

    const result = callGeminiVisionAPI(base64Image, prompt);

    if (!result.success) {
      return { success: false, message: result.message };
    }

    const records = parseJSONResponse(result.data);

    if (!records || records.length === 0) {
      return { success: false, message: '無法從圖片中辨識出有效的打卡記錄' };
    }

    // 寫入試算表
    const sheet = ss.getSheetByName(SHEET_NAMES.CLOCK_IN);
    const rowsToAdd = [];

    for (const record of records) {
      const hours = calculateHours(record.timeIn, record.timeOut);
      rowsToAdd.push([
        employeeName,
        record.date,
        record.timeIn,
        record.timeOut,
        hours
      ]);
    }

    if (rowsToAdd.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    }

    return {
      success: true,
      message: `成功辨識並寫入 ${rowsToAdd.length} 筆打卡記錄`,
      recordCount: rowsToAdd.length
    };

  } catch (error) {
    Logger.log('processTimeCardImage Error: ' + error.toString());
    return { success: false, message: '處理打卡紙時發生錯誤: ' + error.toString() };
  }
}

/**
 * 計算工時（小時）
 */
function calculateHours(timeIn, timeOut) {
  try {
    const [inHour, inMin] = timeIn.split(':').map(Number);
    const [outHour, outMin] = timeOut.split(':').map(Number);

    const inMinutes = inHour * 60 + inMin;
    const outMinutes = outHour * 60 + outMin;

    const diffMinutes = outMinutes - inMinutes;
    const hours = Math.round((diffMinutes / 60) * 100) / 100;

    return hours > 0 ? hours : 0;
  } catch (e) {
    return 0;
  }
}

// ==================== Function B: 處理手寫帳本 ====================

/**
 * 處理手寫帳本圖片辨識
 * @param {string} base64Image - 圖片的 Base64 編碼
 * @returns {object} - 處理結果
 */
function processLedgerImage(base64Image) {
  try {
    // 確保分頁存在
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    createSheetIfNotExists(ss, SHEET_NAMES.BONUS_DUTY,
      ['日期', '員工姓名', '分紅金額', '值日費', '備註']);

    const prompt = `你是一個專門辨識診所手寫帳本的 AI 助手。請仔細分析這張手寫帳本圖片，執行以下任務：

1. 找出帳本記錄的日期，轉換為 YYYY-MM-DD 格式
2. 找出所有員工姓名與對應的分紅金額 (Bonus)
3. 判斷誰是當日值日生：
   - 值日生的標記可能是：名字旁有 '$100'、'值'、'Day'、被圈起來、或有特殊符號
   - 值日生的 dutyFee = 100
   - 非值日生的 dutyFee = 0

請以純 JSON Array 格式回傳，不要包含任何其他文字或 markdown 標記：
[
  {"date": "2024-01-15", "name": "王小明", "bonus": 500, "dutyFee": 100},
  {"date": "2024-01-15", "name": "李小華", "bonus": 450, "dutyFee": 0}
]

注意事項：
- bonus 必須是數字（整數）
- dutyFee 只能是 100 或 0
- 如果無法確定分紅金額，設為 0
- 如果無法辨識任何有效記錄，請回傳空陣列：[]`;

    const result = callGeminiVisionAPI(base64Image, prompt);

    if (!result.success) {
      return { success: false, message: result.message };
    }

    const records = parseJSONResponse(result.data);

    if (!records || records.length === 0) {
      return { success: false, message: '無法從圖片中辨識出有效的帳本記錄' };
    }

    // 寫入試算表
    const sheet = ss.getSheetByName(SHEET_NAMES.BONUS_DUTY);
    const rowsToAdd = [];

    for (const record of records) {
      rowsToAdd.push([
        record.date,
        record.name,
        record.bonus || 0,
        record.dutyFee || 0,
        '' // 備註欄位
      ]);
    }

    if (rowsToAdd.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    }

    return {
      success: true,
      message: `成功辨識並寫入 ${rowsToAdd.length} 筆分紅/值班記錄`,
      recordCount: rowsToAdd.length
    };

  } catch (error) {
    Logger.log('processLedgerImage Error: ' + error.toString());
    return { success: false, message: '處理帳本時發生錯誤: ' + error.toString() };
  }
}

// ==================== Gemini API 呼叫 ====================

/**
 * 呼叫 Gemini Vision API
 * @param {string} base64Image - 圖片的 Base64 編碼
 * @param {string} prompt - 提示詞
 * @returns {object} - API 回應結果
 */
function callGeminiVisionAPI(base64Image, prompt) {
  try {
    // 移除 data URL 前綴（如果有的話）
    const base64Data = base64Image.replace(/^data:image\/\w+;base64,/, '');

    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: 'image/jpeg',
              data: base64Data
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        topK: 1,
        topP: 1,
        maxOutputTokens: 4096
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const url = `${GEMINI_API_URL}?key=${GEMINI_API_KEY}`;
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log('Gemini API Error: ' + responseText);
      return {
        success: false,
        message: `API 錯誤 (${responseCode}): ${responseText}`
      };
    }

    const jsonResponse = JSON.parse(responseText);

    if (jsonResponse.candidates && jsonResponse.candidates[0] &&
        jsonResponse.candidates[0].content &&
        jsonResponse.candidates[0].content.parts) {
      const text = jsonResponse.candidates[0].content.parts[0].text;
      return { success: true, data: text };
    }

    return { success: false, message: 'API 回應格式錯誤' };

  } catch (error) {
    Logger.log('callGeminiVisionAPI Error: ' + error.toString());
    return { success: false, message: 'API 呼叫失敗: ' + error.toString() };
  }
}

/**
 * 解析 JSON 回應（處理可能的格式問題）
 */
function parseJSONResponse(text) {
  try {
    // 嘗試直接解析
    return JSON.parse(text);
  } catch (e) {
    // 嘗試提取 JSON 陣列
    const match = text.match(/\[[\s\S]*\]/);
    if (match) {
      try {
        return JSON.parse(match[0]);
      } catch (e2) {
        Logger.log('JSON Parse Error: ' + e2.toString());
        return null;
      }
    }
    Logger.log('No JSON array found in response');
    return null;
  }
}

// ==================== Function C: 薪資計算 ====================

/**
 * 彈出視窗詢問年月並執行薪資計算
 */
function promptCalculatePayroll() {
  const ui = SpreadsheetApp.getUi();

  // 詢問年份
  const yearResponse = ui.prompt(
    '💰 薪資計算',
    '請輸入年份 (例如: 2024):',
    ui.ButtonSet.OK_CANCEL
  );

  if (yearResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const year = parseInt(yearResponse.getResponseText().trim());
  if (isNaN(year) || year < 2000 || year > 2100) {
    ui.alert('❌ 錯誤', '請輸入有效的年份 (2000-2100)', ui.ButtonSet.OK);
    return;
  }

  // 詢問月份
  const monthResponse = ui.prompt(
    '💰 薪資計算',
    '請輸入月份 (1-12):',
    ui.ButtonSet.OK_CANCEL
  );

  if (monthResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const month = parseInt(monthResponse.getResponseText().trim());
  if (isNaN(month) || month < 1 || month > 12) {
    ui.alert('❌ 錯誤', '請輸入有效的月份 (1-12)', ui.ButtonSet.OK);
    return;
  }

  // 執行計算
  const result = calculateMonthlyPayroll(year, month);

  if (result.success) {
    ui.alert('✅ 計算完成', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('❌ 計算失敗', result.message, ui.ButtonSet.OK);
  }
}

/**
 * 計算指定月份的薪資
 * @param {number} year - 年份
 * @param {number} month - 月份
 * @returns {object} - 計算結果
 */
function calculateMonthlyPayroll(year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 確保分頁存在
    createSheetIfNotExists(ss, SHEET_NAMES.PAYROLL,
      ['年份', '月份', '員工姓名', '公司應上工時', '實際上班工時', '時數差異(抵扣額度)', '總分紅', '總值日費']);

    // 1. 計算標準工時
    const standardHours = calculateStandardHours(year, month);

    // 2. 取得打卡記錄
    const clockInRecords = getClockInRecords(year, month);

    // 3. 取得分紅與值班記錄
    const bonusDutyRecords = getBonusDutyRecords(year, month);

    // 4. 取得員工名單
    const employees = getEmployeeList();

    if (employees.length === 0) {
      return { success: false, message: '尚未設定任何員工，請先在「員工名單」分頁新增員工' };
    }

    // 5. 計算每位員工的薪資
    const payrollData = [];

    for (const employee of employees) {
      // 計算實際工時
      const actualHours = clockInRecords
        .filter(r => r.name === employee)
        .reduce((sum, r) => sum + r.hours, 0);

      // 計算時數差異
      const hoursDiff = standardHours - actualHours;

      // 計算分紅與值日費總和
      const employeeBonusDuty = bonusDutyRecords.filter(r => r.name === employee);
      const totalBonus = employeeBonusDuty.reduce((sum, r) => sum + r.bonus, 0);
      const totalDutyFee = employeeBonusDuty.reduce((sum, r) => sum + r.dutyFee, 0);

      payrollData.push([
        year,
        month,
        employee,
        standardHours,
        Math.round(actualHours * 100) / 100,
        Math.round(hoursDiff * 100) / 100,
        totalBonus,
        totalDutyFee
      ]);
    }

    // 5. 清空並寫入薪資計算結果
    const payrollSheet = ss.getSheetByName(SHEET_NAMES.PAYROLL);

    // 清除舊資料（保留標題列）
    const lastRow = payrollSheet.getLastRow();
    if (lastRow > 1) {
      payrollSheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }

    // 寫入新資料
    if (payrollData.length > 0) {
      payrollSheet.getRange(2, 1, payrollData.length, 8).setValues(payrollData);
    }

    return {
      success: true,
      message: `${year}年${month}月薪資計算完成！\n\n` +
               `📊 標準工時: ${standardHours} 小時\n` +
               `👥 已計算 ${employees.length} 位員工的薪資資料`
    };

  } catch (error) {
    Logger.log('calculateMonthlyPayroll Error: ' + error.toString());
    return { success: false, message: '計算薪資時發生錯誤: ' + error.toString() };
  }
}

/**
 * 計算標準工時（Dr. Xu's Rule）
 * 公式：標準工時 = (總天數 - 週末天數 - 國定假日數) * 8
 */
function calculateStandardHours(year, month) {
  // 取得該月總天數
  const daysInMonth = new Date(year, month, 0).getDate();

  // 計算週末天數
  let weekendDays = 0;
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) { // 週日 = 0, 週六 = 6
      weekendDays++;
    }
  }

  // 取得國定假日數
  const holidays = getHolidays(year, month);
  const holidayCount = holidays.length;

  // 計算標準工時
  const workDays = daysInMonth - weekendDays - holidayCount;
  const standardHours = workDays * 8;

  Logger.log(`${year}/${month}: 總天數=${daysInMonth}, 週末=${weekendDays}, 國定假日=${holidayCount}, 工作日=${workDays}, 標準工時=${standardHours}`);

  return standardHours;
}

/**
 * 取得指定月份的國定假日
 */
function getHolidays(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HOLIDAYS);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const holidays = [];

  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);

  for (const row of data) {
    if (!row[0]) continue;

    const holidayDate = new Date(row[0]);
    if (holidayDate >= monthStart && holidayDate <= monthEnd) {
      // 確認該假日不是週末（避免重複扣除）
      const dayOfWeek = holidayDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        holidays.push({
          date: holidayDate,
          name: row[1]
        });
      }
    }
  }

  return holidays;
}

/**
 * 取得指定月份的打卡記錄
 */
function getClockInRecords(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CLOCK_IN);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const records = [];

  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);

  for (const row of data) {
    if (!row[1]) continue;

    const recordDate = new Date(row[1]);
    if (recordDate >= monthStart && recordDate <= monthEnd) {
      records.push({
        name: row[0],
        date: recordDate,
        timeIn: row[2],
        timeOut: row[3],
        hours: parseFloat(row[4]) || 0
      });
    }
  }

  return records;
}

/**
 * 取得指定月份的分紅與值班記錄
 */
function getBonusDutyRecords(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.BONUS_DUTY);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const records = [];

  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);

  for (const row of data) {
    if (!row[0]) continue;

    const recordDate = new Date(row[0]);
    if (recordDate >= monthStart && recordDate <= monthEnd) {
      records.push({
        date: recordDate,
        name: row[1],
        bonus: parseFloat(row[2]) || 0,
        dutyFee: parseFloat(row[3]) || 0,
        note: row[4]
      });
    }
  }

  return records;
}

// ====================================================================
// 行動版 Web App  (給 webapp/ 靜態頁面呼叫的 REST 風格端點)
// 部署：發行 → 新增部署作業 → 類型「網頁應用程式」
//   - 執行身分：我（這樣會以你的權限寫入試算表）
//   - 誰可以存取：「任何人」
// 部署後得到一組 https://script.google.com/macros/s/.../exec URL
// 把這個 URL + WEBAPP_SECRET 填到 webapp 的「設定」頁。
// ====================================================================

const WEBAPP_SECRET = 'change-me-before-deploy';   // ⚠️ 請改成自己的密碼

// 班別 → 試算表上的列尾標籤
const SHIFT_TO_BONUS_ROW = { morning: 'A', afternoon: 'B', night: 'C' };
const SHIFT_TO_DUTY_ROW  = { morning: 'AM Duty', afternoon: 'PM Duty', night: 'PM Duty' };
const EMPLOYEE_ROW_SUFFIXES = ['AM Duty', 'PM Duty', 'A', 'B', 'C'];
const DUTY_FEE = 100;

function doGet(e) {
  return jsonOut({
    status: 'ok',
    service: 'clinic-bonus-ocr',
    version: '1.0',
    actions: ['ping', 'ocr', 'save', 'summary', 'employees']
  });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ success: false, message: '請求 body 不是有效 JSON' });
  }

  if (body.secret !== WEBAPP_SECRET) {
    return jsonOut({ success: false, message: '密碼錯誤 (secret mismatch)' });
  }

  let result;
  try {
    switch (body.action) {
      case 'ping':
        result = { success: true, message: 'pong', time: new Date().toISOString() };
        break;
      case 'ocr':
        result = ocrBonusDutyShift(body.image);
        break;
      case 'save':
        result = saveBonusDutyShiftRecord(body.record, body.spreadsheetId);
        break;
      case 'summary':
        result = getBonusMonthSummary(body.year, body.month, body.spreadsheetId);
        break;
      case 'employees':
        result = { success: true, employees: getKnownEmployees(body.spreadsheetId) };
        break;
      default:
        result = { success: false, message: '未知 action: ' + body.action };
    }
  } catch (err) {
    Logger.log('doPost Error: ' + err.toString());
    result = { success: false, message: '伺服器錯誤: ' + err.toString() };
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== Bonus + Shift 專用 OCR ====================

function ocrBonusDutyShift(base64Image) {
  if (!base64Image) {
    return { success: false, message: '缺少圖片資料' };
  }

  const prompt = [
    '你是一位有 20 年經驗的中文手寫文字辨識專家，特別擅長辨識診所員工手寫帳本。',
    '請仔細分析這張圖片，提取以下資訊：',
    '',
    '【步驟 1：找出標頭資訊（重點看圖片最上方）】',
    '- 日期：通常在頂部中央，格式如「2026.04.06 W一(早)」、「2026.04.01 W三(午)」、「2026.04.06 W一(晚)」',
    '  → 抽取年/月/日，輸出 YYYY-MM-DD 格式',
    '  → 抽取班別括號內：「早」=morning、「午」=afternoon、「晚」=night',
    '- 值日生：通常在「左上角」寫著「值日生：XXX」，可能是手寫名字或紅色印章',
    '  → 也可能在「右上角」（最多 3 位）',
    '  → 印章字尾若有「RN」/「醫」/「藥」等職稱請去除（例：「江芯儀RN」→「江芯儀」）',
    '',
    '【步驟 2：辨識表格】',
    '表格欄位通常為：序號 | 姓名 | 主分紅金額 | 次分紅金額 | 備註',
    '每一行請提取：',
    '- name: 員工姓名（中文 2~4 字）',
    '- bonus_main: 主欄分紅（常見值 0/80/150/200/300）',
    '- bonus_extra: 次欄分紅（常見值 5/15/20/45/60/100/130/205 等小整數，無則 0）',
    '- note: 備註文字（如「備藥」「榮民」「慢箋」「殘」），無則空字串',
    '',
    '【輸出格式】',
    '只輸出純 JSON，不要任何 ``` 標記或說明文字：',
    '{',
    '  "date": "2026-04-06",',
    '  "shift": "morning",',
    '  "duty_persons": ["郭雅如"],',
    '  "entries": [',
    '    {"name": "莊榆植", "bonus_main": 0,   "bonus_extra": 0,   "note": "備藥"},',
    '    {"name": "莊統憲", "bonus_main": 200, "bonus_extra": 0,   "note": ""},',
    '    {"name": "林郁言", "bonus_main": 200, "bonus_extra": 60,  "note": ""}',
    '  ]',
    '}',
    '',
    '【注意】',
    '- 所有金額必須是整數，無法辨識則填 0',
    '- 多位值日生請全部列出',
    '- 跳過空白列、模糊不清難以辨識的列',
    '- 不要編造任何資訊',
    '- 若完全無法辨識，回傳 {"date": null, "shift": null, "duty_persons": [], "entries": []}'
  ].join('\n');

  const apiResult = callGeminiVisionAPI(base64Image, prompt);
  if (!apiResult.success) return apiResult;

  let text = (apiResult.data || '').trim();
  // 去除 markdown code fence
  text = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');

  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    const m = text.match(/\{[\s\S]*\}/);
    if (m) {
      try { parsed = JSON.parse(m[0]); } catch (e2) {}
    }
  }

  if (!parsed) {
    return { success: false, message: '解析 OCR 結果失敗', raw: apiResult.data };
  }

  // 正規化
  parsed.duty_persons = (parsed.duty_persons || []).filter(Boolean);
  parsed.entries = (parsed.entries || []).map(function (en) {
    return {
      name: (en.name || '').trim(),
      bonus_main: parseInt(en.bonus_main, 10) || 0,
      bonus_extra: parseInt(en.bonus_extra, 10) || 0,
      note: (en.note || '').trim()
    };
  }).filter(function (en) { return en.name; });

  return { success: true, data: parsed };
}

// ==================== Save：寫入矩陣式月份分頁 ====================

function saveBonusDutyShiftRecord(record, spreadsheetId) {
  if (!record || !record.date || !record.shift) {
    return { success: false, message: '資料不完整：缺少 date 或 shift' };
  }

  const bonusSuffix = SHIFT_TO_BONUS_ROW[record.shift];
  const dutySuffix  = SHIFT_TO_DUTY_ROW[record.shift];
  if (!bonusSuffix) {
    return { success: false, message: '無效班別：' + record.shift };
  }

  const m = String(record.date).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) {
    return { success: false, message: '日期格式錯誤，應為 YYYY-MM-DD：' + record.date };
  }
  const year  = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  const day   = parseInt(m[3], 10);
  if (day < 1 || day > 31) {
    return { success: false, message: '日期超出範圍' };
  }

  const ss = openSpreadsheet(spreadsheetId);
  const sheetName = year + '-' + month;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = createMonthSheet(ss, year, month);

  const dayCol = day + 1; // A=1 為標籤欄，第 1 日 = B 欄 = 2
  const writes = [];

  // 寫分紅
  (record.entries || []).forEach(function (entry) {
    if (!entry.name) return;
    const amount = parseInt(entry.bonus, 10);
    if (!amount || amount === 0) return; // 0 不寫
    const rowIndex = ensureEmployeeRow(sheet, entry.name, bonusSuffix);
    sheet.getRange(rowIndex, dayCol).setValue(amount);
    writes.push({ row: entry.name + ' ' + bonusSuffix, day: day, value: amount });
  });

  // 寫值日費
  (record.duty_persons || []).forEach(function (name) {
    if (!name) return;
    const rowIndex = ensureEmployeeRow(sheet, name, dutySuffix);
    sheet.getRange(rowIndex, dayCol).setValue(DUTY_FEE);
    writes.push({ row: name + ' ' + dutySuffix, day: day, value: DUTY_FEE });
  });

  refreshTotalsColumn(sheet);

  return {
    success: true,
    sheetName: sheetName,
    day: day,
    shift: record.shift,
    writeCount: writes.length,
    writes: writes,
    message: '已寫入 ' + sheetName + ' (' + day + '日 ' + record.shift + ')，共 ' + writes.length + ' 筆'
  };
}

function openSpreadsheet(spreadsheetId) {
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('未提供 spreadsheetId 且 Apps Script 未綁定試算表');
}

function createMonthSheet(ss, year, month) {
  const sheetName = year + '-' + month;
  const sheet = ss.insertSheet(sheetName);

  // A1 月份標籤
  sheet.getRange(1, 1).setValue(year + '-' + month);

  // Row 2: 1-31
  const days = [];
  for (let d = 1; d <= 31; d++) days.push(d);
  sheet.getRange(2, 2, 1, 31).setValues([days]);
  sheet.getRange(2, 33).setValue('total');

  // Row 3: 星期
  const wkLetters = ['日', '一', '二', '三', '四', '五', '六'];
  const wks = [];
  for (let d = 1; d <= 31; d++) {
    const dt = new Date(year, month - 1, d);
    wks.push(dt.getMonth() === month - 1 ? wkLetters[dt.getDay()] : '');
  }
  sheet.getRange(3, 2, 1, 31).setValues([wks]);

  sheet.getRange(1, 1, 3, 33)
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 130);
  for (let c = 2; c <= 32; c++) sheet.setColumnWidth(c, 38);
  sheet.setColumnWidth(33, 70);

  return sheet;
}

/**
 * 確保某員工的 5 列（AM Duty / PM Duty / A / B / C）都存在
 * 並回傳指定 suffix 的列號
 */
function ensureEmployeeRow(sheet, employeeName, suffix) {
  const lastRow = sheet.getLastRow();
  const startRow = 4;
  const targetLabel = employeeName + ' ' + suffix;

  let labels = [];
  if (lastRow >= startRow) {
    labels = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
    for (let i = 0; i < labels.length; i++) {
      if (labels[i][0] === targetLabel) return startRow + i;
    }
  }

  // 找出該員工現有的列（任何 suffix）
  const existing = {};
  for (let i = 0; i < labels.length; i++) {
    const lbl = labels[i][0];
    if (typeof lbl !== 'string') continue;
    EMPLOYEE_ROW_SUFFIXES.forEach(function (sfx) {
      if (lbl === employeeName + ' ' + sfx) {
        existing[sfx] = startRow + i;
      }
    });
  }

  // 起始插入位置：若已有任何列 → 緊接著最後一列；否則接到表尾
  let insertAt = (Object.keys(existing).length > 0)
    ? Math.max.apply(null, Object.values(existing)) + 1
    : Math.max(lastRow + 1, startRow);

  EMPLOYEE_ROW_SUFFIXES.forEach(function (sfx) {
    if (!existing[sfx]) {
      sheet.getRange(insertAt, 1).setValue(employeeName + ' ' + sfx);
      existing[sfx] = insertAt;
      insertAt++;
    }
  });

  return existing[suffix];
}

function refreshTotalsColumn(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;
  const formulas = [];
  for (let r = 4; r <= lastRow; r++) {
    formulas.push(['=SUM(B' + r + ':AF' + r + ')']);
  }
  sheet.getRange(4, 33, formulas.length, 1).setFormulas(formulas);
}

// ==================== Summary：本月每位員工的總額 ====================

function getBonusMonthSummary(year, month, spreadsheetId) {
  const ss = openSpreadsheet(spreadsheetId);
  const sheetName = year + '-' + month;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: true, sheetName: sheetName, employees: [], grandTotal: 0 };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) {
    return { success: true, sheetName: sheetName, employees: [], grandTotal: 0 };
  }

  const labels = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
  const totals = sheet.getRange(4, 33, lastRow - 3, 1).getValues();

  const map = {};
  for (let i = 0; i < labels.length; i++) {
    const lbl = labels[i][0];
    if (typeof lbl !== 'string' || !lbl) continue;
    const m = lbl.match(/^(.+?)\s+(AM Duty|PM Duty|A|B|C)$/);
    if (!m) continue;
    const name = m[1];
    const suffix = m[2];
    const total = parseFloat(totals[i][0]) || 0;
    if (!map[name]) {
      map[name] = { name: name, am_duty: 0, pm_duty: 0, a: 0, b: 0, c: 0, bonus: 0, duty: 0, total: 0 };
    }
    if (suffix === 'AM Duty')      { map[name].am_duty = total; map[name].duty += total; }
    else if (suffix === 'PM Duty') { map[name].pm_duty = total; map[name].duty += total; }
    else if (suffix === 'A')       { map[name].a = total; map[name].bonus += total; }
    else if (suffix === 'B')       { map[name].b = total; map[name].bonus += total; }
    else if (suffix === 'C')       { map[name].c = total; map[name].bonus += total; }
    map[name].total += total;
  }

  const employees = Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return b.total - a.total; });
  const grandTotal = employees.reduce(function (s, e) { return s + e.total; }, 0);

  return { success: true, sheetName: sheetName, employees: employees, grandTotal: grandTotal };
}

function getKnownEmployees(spreadsheetId) {
  const ss = openSpreadsheet(spreadsheetId);
  const sheets = ss.getSheets();
  const set = {};

  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    if (!/^\d{4}-\d{1,2}$/.test(sh.getName())) continue;
    const lastRow = sh.getLastRow();
    if (lastRow < 4) continue;
    const labels = sh.getRange(4, 1, lastRow - 3, 1).getValues();
    for (let j = 0; j < labels.length; j++) {
      const lbl = labels[j][0];
      if (typeof lbl !== 'string') continue;
      const m = lbl.match(/^(.+?)\s+(AM Duty|PM Duty|A|B|C)$/);
      if (m) set[m[1]] = true;
    }
  }
  return Object.keys(set).sort();
}
