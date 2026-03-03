const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const https = require('https');

const app = express();
const PORT = 3001;

const GOOGLE_APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL ||
  'https://script.google.com/macros/s/AKfycbzRQEjc8Cq4vSxpxVk1hwEcbFWLf9XUuNO8TDV58OJwAOfDW52PoS6Iunst35h-DH2_Iw/exec';

const EXCEL_FILE_PATH = path.join(__dirname, 'bookings.xlsx');
const SHEET_NAME = '預約紀錄';

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// 讀取或建立 Excel 工作簿
function getOrCreateWorkbook() {
  if (fs.existsSync(EXCEL_FILE_PATH)) {
    return XLSX.readFile(EXCEL_FILE_PATH);
  }
  const wb = XLSX.utils.book_new();
  const headers = [['時段', '姓名', '預約時間', '時段 ID']];
  const ws = XLSX.utils.aoa_to_sheet(headers);

  // 設定欄寬
  ws['!cols'] = [
    { wch: 10 },
    { wch: 16 },
    { wch: 22 },
    { wch: 16 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
  XLSX.writeFile(wb, EXCEL_FILE_PATH);
  return wb;
}

// 取得工作表內現有資料（以 slotId 為 key）
function getExistingBookings(wb) {
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) return {};
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const existing = {};
  // rows[0] 為標題列，從 index 1 開始
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row && row[3]) {
      existing[row[3]] = { time: row[0], name: row[1], bookedAt: row[2], rowIndex: i };
    }
  }
  return existing;
}

// 將所有預約資料同步寫入 Excel
function syncBookingsToExcel(bookings) {
  const wb = getOrCreateWorkbook();
  const rows = [['時段', '姓名', '預約時間', '時段 ID']];

  const sorted = Object.entries(bookings).sort(([idA], [idB]) => idA.localeCompare(idB));

  for (const [slotId, entry] of sorted) {
    const time = typeof entry === 'object' ? entry.time : '—';
    const name = typeof entry === 'object' ? entry.name : entry;
    const bookedAt = typeof entry === 'object' ? entry.bookedAt : new Date().toISOString();
    rows.push([time, name, bookedAt, slotId]);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 22 }, { wch: 16 }];
  wb.Sheets[SHEET_NAME] = ws;
  XLSX.writeFile(wb, EXCEL_FILE_PATH);
}

// 轉送至 Google Apps Script（自動跟隨 302 重新導向）
function forwardToAppsScript(payload, redirectCount = 0) {
  if (!GOOGLE_APPS_SCRIPT_URL) return Promise.resolve({ skipped: true });
  if (redirectCount > 5) return Promise.reject(new Error('重新導向次數過多'));

  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload);
    const url = new URL(
      redirectCount === 0 ? GOOGLE_APPS_SCRIPT_URL : payload._redirectUrl
    );
    const isHttps = url.protocol === 'https:';
    const transport = isHttps ? https : require('http');

    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data)
      }
    };

    const req = transport.request(options, (res) => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        const location = res.headers['location'];
        if (location) {
          // 跟隨重新導向，改用 GET（Apps Script 302 後通常接受 GET）
          return followRedirect(location).then(resolve).catch(reject);
        }
      }
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body }));
    });

    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

function followRedirect(location) {
  return new Promise((resolve, reject) => {
    const url = new URL(location);
    const transport = url.protocol === 'https:' ? https : require('http');
    transport.get(location, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body }));
    }).on('error', reject);
  });
}

// POST /api/bookings - 儲存單筆預約
app.post('/api/bookings', async (req, res) => {
  const { slotId, time, name } = req.body;

  if (!slotId || !name || !time) {
    return res.status(400).json({ error: '缺少必要欄位：slotId、time、name' });
  }

  const bookedAt = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  const entry = { time, name, bookedAt };

  try {
    // 讀取現有 Excel，新增這筆
    const wb = getOrCreateWorkbook();
    const existing = getExistingBookings(wb);

    if (existing[slotId]) {
      return res.status(409).json({ error: '此時段已被預約' });
    }

    const ws = wb.Sheets[SHEET_NAME];
    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const newRow = [time, name, bookedAt, slotId];
    allRows.push(newRow);
    const newWs = XLSX.utils.aoa_to_sheet(allRows);
    newWs['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 22 }, { wch: 16 }];
    wb.Sheets[SHEET_NAME] = newWs;
    XLSX.writeFile(wb, EXCEL_FILE_PATH);

    // 同步至 Google Sheets（非阻塞）
    forwardToAppsScript({ action: 'add', slotId, time, name, bookedAt })
      .then(r => { if (!r.skipped) console.log('[Sheets 同步]', r.status); })
      .catch(e => console.error('[Sheets 同步失敗]', e.message));

    console.log(`[新增預約] ${time} - ${name} (${slotId})`);
    res.json({ success: true, entry });
  } catch (err) {
    console.error('[寫入 Excel 失敗]', err);
    res.status(500).json({ error: '伺服器錯誤，請稍後再試' });
  }
});

// DELETE /api/bookings/:slotId - 取消預約
app.delete('/api/bookings/:slotId', async (req, res) => {
  const { slotId } = req.params;

  try {
    const wb = getOrCreateWorkbook();
    const ws = wb.Sheets[SHEET_NAME];
    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const filtered = allRows.filter((row, i) => i === 0 || row[3] !== slotId);

    if (filtered.length === allRows.length) {
      return res.status(404).json({ error: '找不到此預約' });
    }

    const newWs = XLSX.utils.aoa_to_sheet(filtered);
    newWs['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 22 }, { wch: 16 }];
    wb.Sheets[SHEET_NAME] = newWs;
    XLSX.writeFile(wb, EXCEL_FILE_PATH);

    forwardToAppsScript({ action: 'delete', slotId })
      .then(r => { if (!r.skipped) console.log('[Sheets 同步刪除]', r.status); })
      .catch(e => console.error('[Sheets 同步失敗]', e.message));

    console.log(`[取消預約] ${slotId}`);
    res.json({ success: true });
  } catch (err) {
    console.error('[刪除 Excel 失敗]', err);
    res.status(500).json({ error: '伺服器錯誤' });
  }
});

// GET /api/bookings - 讀取所有預約（供頁面初始化）
app.get('/api/bookings', (req, res) => {
  try {
    const wb = getOrCreateWorkbook();
    const ws = wb.Sheets[SHEET_NAME];
    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const bookings = {};

    for (let i = 1; i < allRows.length; i++) {
      const row = allRows[i];
      if (row && row[3] && row[1]) {
        bookings[row[3]] = { time: row[0], name: row[1], bookedAt: row[2] };
      }
    }

    res.json({ bookings });
  } catch (err) {
    console.error('[讀取 Excel 失敗]', err);
    res.status(500).json({ error: '伺服器錯誤' });
  }
});

// GET /api/export - 下載 Excel 檔案
app.get('/api/export', (req, res) => {
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    return res.status(404).json({ error: '尚無預約紀錄' });
  }
  res.download(EXCEL_FILE_PATH, '面談預約紀錄.xlsx');
});

app.listen(PORT, () => {
  console.log(`✓ 面談預約後端已啟動：http://localhost:${PORT}`);
  console.log(`  Excel 儲存路徑：${EXCEL_FILE_PATH}`);
  if (!GOOGLE_APPS_SCRIPT_URL) {
    console.log('  ⚠ 尚未設定 APPS_SCRIPT_URL，Google Sheets 同步已停用');
  }
});
