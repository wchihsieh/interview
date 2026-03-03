/**
 * Google Apps Script - 面談預約 Google Sheets 後端
 *
 * 部署步驟：
 * 1. 開啟目標 Google Sheets
 * 2. 點選「擴充功能」>「Apps Script」
 * 3. 貼上此檔案內容（取代原有內容）
 * 4. 點選「部署」>「新增部署作業」
 *    - 類型選「網路應用程式」
 *    - 執行身分：「我自己」
 *    - 存取權：「所有人」（或「所有已登入 Google 的使用者」）
 * 5. 複製部署 URL，貼入 server.js 的 GOOGLE_APPS_SCRIPT_URL 變數
 *    或設為環境變數：APPS_SCRIPT_URL=https://script.google.com/macros/s/...
 */

const SHEET_NAME = '預約紀錄';
const HEADERS = ['時段', '姓名', '預約時間', '時段 ID'];

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);

    // 設定標題列樣式
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#b85c38');
    headerRange.setFontColor('#ffffff');

    // 設定欄寬
    sheet.setColumnWidth(1, 90);   // 時段
    sheet.setColumnWidth(2, 130);  // 姓名
    sheet.setColumnWidth(3, 200);  // 預約時間
    sheet.setColumnWidth(4, 140);  // 時段 ID

    sheet.setFrozenRows(1);
  }

  return sheet;
}

function findRowBySlotId(sheet, slotId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === slotId) return i + 1; // 1-based row number
  }
  return -1;
}

function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    const payload = JSON.parse(e.postData.contents);
    const { action, slotId, time, name, bookedAt } = payload;
    const sheet = getOrCreateSheet();

    if (action === 'add') {
      // 檢查是否已存在
      const existingRow = findRowBySlotId(sheet, slotId);
      if (existingRow > 0) {
        return ContentService
          .createTextOutput(JSON.stringify({ error: '此時段已被預約' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      sheet.appendRow([time, name, bookedAt, slotId]);

      // 新增列的樣式（交替背景色）
      const lastRow = sheet.getLastRow();
      const rowRange = sheet.getRange(lastRow, 1, 1, 4);
      rowRange.setBackground(lastRow % 2 === 0 ? '#f5f0e8' : '#ffffff');

      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'delete') {
      const rowIndex = findRowBySlotId(sheet, slotId);
      if (rowIndex < 0) {
        return ContentService
          .createTextOutput(JSON.stringify({ error: '找不到此預約' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet.deleteRow(rowIndex);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ error: '未知的 action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 支援 CORS 預檢請求
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '面談預約 API 運作中' }))
    .setMimeType(ContentService.MimeType.JSON);
}
