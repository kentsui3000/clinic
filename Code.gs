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
