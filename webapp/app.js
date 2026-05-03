/* 診所分紅 OCR 助手 - 前端 (純靜態，呼叫 Apps Script Web App) */
(function () {
  'use strict';

  // ==================== Storage helpers ====================
  const KEYS = {
    URL: 'clinic.ocr.url',
    SECRET: 'clinic.ocr.secret',
    SHEET_ID: 'clinic.ocr.sheetId',
    BONUS_SRC: 'clinic.ocr.bonusSrc',
    RECENT: 'clinic.ocr.recent'
  };
  const get = k => localStorage.getItem(k) || '';
  const set = (k, v) => localStorage.setItem(k, v);

  // ==================== View routing ====================
  const views = ['viewCapture', 'viewLoading', 'viewReview', 'viewSuccess', 'viewSettings', 'viewSummary'];
  function showView(id) {
    views.forEach(v => {
      const el = document.getElementById(v);
      if (el) el.classList.toggle('active', v === id);
    });
    window.scrollTo(0, 0);
  }

  // ==================== Toast ====================
  let toastTimer;
  function toast(msg, type) {
    const t = document.getElementById('toast');
    t.className = 'toast show ' + (type || '');
    t.textContent = msg;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
  }

  // ==================== App state ====================
  const state = {
    imageBase64: null,    // raw base64 (no prefix) - used for OCR
    imageDataUrl: null,   // data URL - for preview
    ocrData: null,        // last OCR result {date, shift, duty_persons, entries}
    bonusSrc: get(KEYS.BONUS_SRC) || 'extra'
  };

  // ==================== API call helper ====================
  async function callApi(action, body) {
    const url = get(KEYS.URL);
    const secret = get(KEYS.SECRET);
    if (!url || !secret) {
      throw new Error('尚未設定 URL 或密碼，請到設定頁。');
    }
    const payload = Object.assign({ action: action, secret: secret }, body || {});
    const sheetId = get(KEYS.SHEET_ID);
    if (sheetId && !('spreadsheetId' in payload)) payload.spreadsheetId = sheetId;

    // Use text/plain to avoid CORS preflight on Apps Script
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow'
    });
    const txt = await res.text();
    let json;
    try { json = JSON.parse(txt); }
    catch (e) { throw new Error('伺服器回傳非 JSON：' + txt.slice(0, 200)); }
    if (json.success === false) throw new Error(json.message || '未知錯誤');
    return json;
  }

  // ==================== Image processing ====================
  // Resize to max 1600px, then convert to JPEG base64 (compressed)
  async function fileToCompressedBase64(file) {
    const dataUrl = await new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = reject;
      r.readAsDataURL(file);
    });
    const img = await new Promise((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = reject;
      i.src = dataUrl;
    });
    const MAX = 1800;
    let w = img.width, h = img.height;
    if (Math.max(w, h) > MAX) {
      const ratio = MAX / Math.max(w, h);
      w = Math.round(w * ratio);
      h = Math.round(h * ratio);
    }
    const canvas = document.createElement('canvas');
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, w, h);
    const compressedUrl = canvas.toDataURL('image/jpeg', 0.85);
    return {
      dataUrl: compressedUrl,
      base64: compressedUrl.replace(/^data:image\/[^;]+;base64,/, ''),
      width: w,
      height: h,
      sizeKB: Math.round(compressedUrl.length * 0.75 / 1024)
    };
  }

  // ==================== Capture / preview ====================
  const fileInput = document.getElementById('fileInput');
  const previewWrap = document.getElementById('previewWrap');
  const previewImg = document.getElementById('preview');
  const imageInfo = document.getElementById('imageInfo');

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      toast('壓縮圖片中…');
      const out = await fileToCompressedBase64(file);
      state.imageBase64 = out.base64;
      state.imageDataUrl = out.dataUrl;
      previewImg.src = out.dataUrl;
      imageInfo.textContent = `${out.width}×${out.height}，約 ${out.sizeKB} KB`;
      previewWrap.hidden = false;
      toast('');
    } catch (err) {
      toast('讀取失敗：' + err.message, 'error');
    }
  });

  document.getElementById('btnReplace').addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });

  // ==================== OCR ====================
  document.getElementById('btnRecognize').addEventListener('click', async () => {
    if (!state.imageBase64) { toast('請先選擇圖片', 'error'); return; }
    showView('viewLoading');
    document.getElementById('loadingText').textContent = '辨識中… (Gemini Vision)';
    try {
      const resp = await callApi('ocr', { image: state.imageBase64 });
      state.ocrData = resp.data || {};
      renderReview(state.ocrData);
      showView('viewReview');
    } catch (err) {
      toast('辨識失敗：' + err.message, 'error');
      showView('viewCapture');
    }
  });

  // ==================== Review form ====================
  function renderReview(data) {
    document.getElementById('reviewDate').value = data.date || todayISO();
    document.getElementById('reviewShift').value = data.shift || guessShift();
    document.getElementById('reviewDuty').value = (data.duty_persons || []).join(', ');

    const list = document.getElementById('entriesList');
    list.innerHTML = '';

    // Header row
    const head = document.createElement('div');
    head.className = 'entry-head';
    head.innerHTML = '<div>姓名</div><div>主</div><div>次</div><div>備註</div><div></div>';
    list.appendChild(head);

    (data.entries || []).forEach(en => list.appendChild(makeEntryRow(en)));
    if (!data.entries || data.entries.length === 0) {
      list.appendChild(makeEntryRow({}));
    }
  }

  function makeEntryRow(en) {
    const div = document.createElement('div');
    div.className = 'entry-row';
    div.innerHTML = `
      <input class="name" type="text" placeholder="姓名" value="${esc(en.name || '')}">
      <input class="num main" type="number" inputmode="numeric" placeholder="0" value="${en.bonus_main != null ? en.bonus_main : ''}">
      <input class="num extra" type="number" inputmode="numeric" placeholder="0" value="${en.bonus_extra != null ? en.bonus_extra : ''}">
      <input class="note" type="text" placeholder="備註" value="${esc(en.note || '')}">
      <button class="del" aria-label="刪除">×</button>
    `;
    div.querySelector('.del').addEventListener('click', () => div.remove());
    return div;
  }

  function esc(s) {
    return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  document.getElementById('btnAddEntry').addEventListener('click', () => {
    document.getElementById('entriesList').appendChild(makeEntryRow({}));
  });

  document.getElementById('btnBack').addEventListener('click', () => showView('viewCapture'));

  // Bonus source toggle
  document.querySelectorAll('input[name="bonusSrc"]').forEach(r => {
    if (r.value === state.bonusSrc) r.checked = true;
    r.addEventListener('change', () => {
      state.bonusSrc = r.value;
      set(KEYS.BONUS_SRC, r.value);
    });
  });

  function collectReview() {
    const date = document.getElementById('reviewDate').value;
    const shift = document.getElementById('reviewShift').value;
    const dutyRaw = document.getElementById('reviewDuty').value || '';
    const duty_persons = dutyRaw.split(/[,，、；;\s]+/).map(s => s.trim()).filter(Boolean);

    const src = (document.querySelector('input[name="bonusSrc"]:checked') || {}).value || 'extra';
    const rows = document.querySelectorAll('#entriesList .entry-row');
    const entries = [];
    rows.forEach(r => {
      const name = r.querySelector('.name').value.trim();
      const main = parseInt(r.querySelector('.main').value, 10) || 0;
      const extra = parseInt(r.querySelector('.extra').value, 10) || 0;
      const note = r.querySelector('.note').value.trim();
      let bonus = 0;
      if (src === 'main') bonus = main;
      else if (src === 'extra') bonus = extra;
      else bonus = main + extra;
      if (name) entries.push({ name, bonus, bonus_main: main, bonus_extra: extra, note });
    });
    return { date, shift, duty_persons, entries };
  }

  // ==================== Save ====================
  document.getElementById('btnSave').addEventListener('click', async () => {
    const record = collectReview();
    if (!record.date || !record.shift) { toast('請填寫日期與班別', 'error'); return; }
    if (record.entries.length === 0 && record.duty_persons.length === 0) {
      toast('沒有任何資料可儲存', 'error');
      return;
    }

    const btn = document.getElementById('btnSave');
    btn.disabled = true;
    btn.textContent = '儲存中…';
    try {
      const resp = await callApi('save', { record: record });
      const detail = `分頁：${resp.sheetName}　日：${resp.day}　班別：${shiftLabel(resp.shift)}<br>共寫入 ${resp.writeCount} 格`;
      document.getElementById('successDetail').innerHTML = detail;
      pushRecent({ at: Date.now(), date: record.date, shift: record.shift,
                   count: record.entries.length, duty: record.duty_persons.length });
      showView('viewSuccess');
      toast('已儲存', 'success');
    } catch (err) {
      toast('儲存失敗：' + err.message, 'error');
    } finally {
      btn.disabled = false;
      btn.textContent = '💾 儲存到試算表';
    }
  });

  function shiftLabel(s) {
    return ({ morning: '早 (A)', afternoon: '午 (B)', night: '晚 (C)' })[s] || s;
  }

  document.getElementById('btnNew').addEventListener('click', () => {
    state.imageBase64 = null;
    state.imageDataUrl = null;
    state.ocrData = null;
    fileInput.value = '';
    previewWrap.hidden = true;
    showView('viewCapture');
    refreshRecent();
  });

  document.getElementById('btnViewSummary').addEventListener('click', () => {
    showView('viewSummary');
    document.getElementById('btnRefreshSummary').click();
  });

  // ==================== Settings ====================
  function loadSettings() {
    document.getElementById('settingsUrl').value = get(KEYS.URL);
    document.getElementById('settingsSecret').value = get(KEYS.SECRET);
    document.getElementById('settingsSheetId').value = get(KEYS.SHEET_ID);
  }
  loadSettings();

  document.getElementById('btnSaveSettings').addEventListener('click', () => {
    set(KEYS.URL, document.getElementById('settingsUrl').value.trim());
    set(KEYS.SECRET, document.getElementById('settingsSecret').value.trim());
    set(KEYS.SHEET_ID, document.getElementById('settingsSheetId').value.trim());
    toast('已儲存設定', 'success');
    showView('viewCapture');
  });

  document.getElementById('btnTestConn').addEventListener('click', async () => {
    // Save first so callApi sees the latest values
    set(KEYS.URL, document.getElementById('settingsUrl').value.trim());
    set(KEYS.SECRET, document.getElementById('settingsSecret').value.trim());
    set(KEYS.SHEET_ID, document.getElementById('settingsSheetId').value.trim());
    try {
      const resp = await callApi('ping');
      toast('連線成功：' + resp.message, 'success');
    } catch (err) {
      toast('連線失敗：' + err.message, 'error');
    }
  });

  document.getElementById('btnClearRecent').addEventListener('click', () => {
    localStorage.removeItem(KEYS.RECENT);
    refreshRecent();
    toast('已清空');
  });

  document.getElementById('btnExport').addEventListener('click', () => {
    const data = {};
    Object.values(KEYS).forEach(k => { data[k] = localStorage.getItem(k); });
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'clinic-ocr-settings.json';
    a.click();
  });

  // ==================== Summary ====================
  function setDefaultSummaryMonth() {
    const inp = document.getElementById('summaryMonth');
    if (inp.value) return;
    const d = new Date();
    inp.value = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }

  document.getElementById('btnRefreshSummary').addEventListener('click', async () => {
    setDefaultSummaryMonth();
    const monthVal = document.getElementById('summaryMonth').value;
    if (!monthVal) { toast('請選擇月份', 'error'); return; }
    const [y, m] = monthVal.split('-').map(Number);
    const tableEl = document.getElementById('summaryTable');
    const grandEl = document.getElementById('summaryGrand');
    tableEl.innerHTML = '<div class="hint">載入中…</div>';
    grandEl.textContent = '';
    try {
      const resp = await callApi('summary', { year: y, month: m });
      renderSummary(resp);
    } catch (err) {
      tableEl.innerHTML = '<div class="hint" style="color:#d93025">' + err.message + '</div>';
    }
  });

  function renderSummary(resp) {
    const grandEl = document.getElementById('summaryGrand');
    const tableEl = document.getElementById('summaryTable');
    if (!resp.employees || resp.employees.length === 0) {
      tableEl.innerHTML = '<div class="hint">本月份分頁尚無資料</div>';
      grandEl.textContent = '分頁：' + resp.sheetName + '　無資料';
      return;
    }
    grandEl.textContent = `分頁：${resp.sheetName}　全店總額：$${resp.grandTotal.toLocaleString()}`;
    let html = '<table><thead><tr><th>員工</th><th>早(A)</th><th>午(B)</th><th>晚(C)</th><th>分紅</th><th>值日</th><th>總計</th></tr></thead><tbody>';
    resp.employees.forEach(e => {
      html += `<tr>
        <td>${esc(e.name)}</td>
        <td>${e.a || ''}</td>
        <td>${e.b || ''}</td>
        <td>${e.c || ''}</td>
        <td>${e.bonus || ''}</td>
        <td>${e.duty || ''}</td>
        <td><strong>${e.total ? '$' + e.total.toLocaleString() : ''}</strong></td>
      </tr>`;
    });
    html += '</tbody></table>';
    tableEl.innerHTML = html;
  }

  // ==================== Recent records (local only) ====================
  function pushRecent(entry) {
    const arr = JSON.parse(get(KEYS.RECENT) || '[]');
    arr.unshift(entry);
    if (arr.length > 10) arr.length = 10;
    set(KEYS.RECENT, JSON.stringify(arr));
    refreshRecent();
  }
  function refreshRecent() {
    const arr = JSON.parse(get(KEYS.RECENT) || '[]');
    const ul = document.getElementById('recentList');
    const card = document.getElementById('recentCard');
    if (!arr.length) { card.hidden = true; return; }
    card.hidden = false;
    ul.innerHTML = arr.map(e => `
      <li>
        <div>${e.date} ${shiftLabel(e.shift)}　共 ${e.count} 人 / ${e.duty} 值日</div>
        <div class="ts">${new Date(e.at).toLocaleString('zh-TW')}</div>
      </li>
    `).join('');
  }
  refreshRecent();

  // ==================== Top bar nav ====================
  document.querySelectorAll('[data-go]').forEach(b => {
    b.addEventListener('click', () => {
      const v = b.getAttribute('data-go');
      if (v === 'viewSummary') {
        showView('viewSummary');
        setDefaultSummaryMonth();
      } else {
        showView(v);
      }
    });
  });

  // ==================== Helpers ====================
  function todayISO() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function guessShift() {
    const h = new Date().getHours();
    if (h < 12) return 'morning';
    if (h < 18) return 'afternoon';
    return 'night';
  }

  // First-run hint
  if (!get(KEYS.URL)) {
    setTimeout(() => {
      toast('首次使用請先到 ⚙️ 設定填入 Apps Script URL', '');
      showView('viewSettings');
    }, 400);
  }
})();
