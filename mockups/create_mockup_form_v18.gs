/**
 * ============================================================
 * 【MOCKUP 產生器】每日當班結帳 v1.8 合併版 — 真實 Google Form 試填版
 * ============================================================
 * 用途:在正式改版前,產生一張「可實際試填」的 Google Form,
 *       體驗 v1.8 合併後的完整流程(含分段跳轉)。
 *
 * 安全性:
 *   - 不連結任何試算表、不掛觸發器、不寫 ScriptProperties
 *   - 填了也不會計算任何分紅
 *   - 與正式系統完全無關,審完直接把表單丟垃圾桶即可
 *
 * 用法:
 *   1. 把本檔內容「附加」到現有 Apps Script 專案最下方(不要覆蓋既有程式碼)
 *   2. 函式下拉選 createMockupFormV18 → ▶ 執行
 *   3. 從執行記錄複製「試填網址」,用手機開啟實際走一遍
 *   4. 審完:刪除 Drive 裡的【MOCKUP】表單 + 刪掉本段程式碼
 * ============================================================
 */

function createMockupFormV18() {
  const STAFF = ['芯儀', '郁芹', '佩芷', '嘉惠', '珈宜', '雅玲'];
  const nonNeg = () => FormApp.createTextValidation()
      .setHelpText('請輸入大於或等於 0 的整數')
      .requireNumberGreaterThanOrEqualTo(0)
      .build();

  const form = FormApp.create('【MOCKUP】每日當班結帳 v1.8');
  form.setDescription(
    '⚠ 這是 v1.8 合併版的「試填用」mockup — 沒有連結試算表、不會計算分紅。\n' +
    '請實際填一次體驗流程(特別試「無代謝活動」和「有代謝活動」兩條路),審完請刪除此表單。'
  );

  // ===== 第 1 頁:本班資訊(4 欄,全點選) =====
  form.addDateItem()
      .setTitle('日期')
      .setRequired(true);

  form.addMultipleChoiceItem()
      .setTitle('門診時段')
      .setChoiceValues(['上午', '下午', '晚上'])
      .setRequired(true);

  form.addCheckboxItem()
      .setTitle('本班當值人員')
      .setHelpText('請勾選本時段在場的所有員工(含值日生本人)')
      .setChoiceValues(STAFF)
      .setRequired(true);

  form.addListItem()
      .setTitle('本時段值日生')
      .setHelpText('晚上時段請選「無」')
      .setChoiceValues(['無'].concat(STAFF))
      .setRequired(true);

  // ===== 第 2 頁:結帳數字(3 個數字 + 路由題) =====
  form.addPageBreakItem()
      .setTitle('結帳數字')
      .setHelpText('只填 3 個數字 — 這是唯一要打字的地方');

  form.addTextItem()
      .setTitle('當時段有效看診人數')
      .setHelpText('本時段所有醫師合計')
      .setRequired(true)
      .setValidation(nonNeg());

  form.addTextItem()
      .setTitle('自費流感疫苗支數')
      .setHelpText('沒有請填 0,不要留空')
      .setRequired(true)
      .setValidation(nonNeg());

  form.addTextItem()
      .setTitle('氣喘/濕疹評估筆數')
      .setHelpText('Peak flow / ACT / POEM 執行人數,一位病人算 1 筆。沒做請填 0')
      .setRequired(true)
      .setValidation(nonNeg());

  const routing = form.addMultipleChoiceItem()
      .setTitle('本時段有無代謝症候群活動?')
      .setHelpText('選「無」會直接跳到簽核頁')
      .setRequired(true);

  // ===== 第 3 頁:代謝症候群活動(選「有」才會看到) =====
  const sec3 = form.addPageBreakItem()
      .setTitle('代謝症候群活動')
      .setHelpText('活動 1 必填。同時段有第二位執行者或第二種類型才填活動 2,沒有就整組留空');

  form.addListItem()
      .setTitle('活動1-活動類型')
      .setHelpText('照健保申報代碼選')
      .setChoiceValues(['收案 (P7501)', '追蹤 (P7502)', '年度評估 (P7503)'])
      .setRequired(true);

  form.addListItem()
      .setTitle('活動1-執行人員')
      .setChoiceValues(STAFF)
      .setRequired(true);

  form.addTextItem()
      .setTitle('活動1-病歷號')
      .setHelpText('多筆用逗號分隔(例:0031234, 0035678),每筆獨立計算')
      .setRequired(true);

  form.addListItem()
      .setTitle('活動2-活動類型(沒有就留空)')
      .setChoiceValues(['收案 (P7501)', '追蹤 (P7502)', '年度評估 (P7503)'])
      .setRequired(false);

  form.addListItem()
      .setTitle('活動2-執行人員')
      .setChoiceValues(STAFF)
      .setRequired(false);

  form.addTextItem()
      .setTitle('活動2-病歷號')
      .setHelpText('多筆用逗號分隔')
      .setRequired(false);

  // ===== 第 4 頁:簽核送出 =====
  const sec4 = form.addPageBreakItem()
      .setTitle('簽核送出');

  form.addListItem()
      .setTitle('填寫人')
      .setChoiceValues(STAFF)
      .setRequired(true);

  form.addListItem()
      .setTitle('核對人')
      .setHelpText('主任或徐醫師,可後補')
      .setChoiceValues(STAFF)
      .setRequired(true);

  form.addTextItem()
      .setTitle('備註')
      .setRequired(false);

  // 路由:無 → 跳過代謝頁直達簽核;有 → 進代謝頁
  routing.setChoices([
    routing.createChoice('無', sec4),
    routing.createChoice('有,展開登錄', sec3)
  ]);

  // 顯示設定(與正式版相同)
  form.setProgressBar(true);
  form.setConfirmationMessage('MOCKUP 試填完成 — 正式版送出後才會自動計算分紅');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setShowLinkToRespondAgain(true); // mockup 方便重複試填,正式版為 false

  const msg = [
    '===== v1.8 MOCKUP 表單已建立 =====',
    '',
    '試填網址(手機開這個):' + form.getPublishedUrl(),
    '編輯網址:' + form.getEditUrl(),
    '',
    '建議試兩條路:',
    '  1. 「無」代謝活動 → 應該從結帳數字直接跳到簽核頁',
    '  2. 「有」代謝活動 → 應該進入代謝頁,活動2留空可直接過',
    '',
    '審完:到 Drive 把【MOCKUP】表單丟垃圾桶,並刪除本段程式碼。',
    '正式 v1.8 不會用這個函式 — 會由改版後的 buildDailyForm_() 與遷移函式處理。'
  ].join('\n');
  Logger.log(msg);
  return msg;
}
