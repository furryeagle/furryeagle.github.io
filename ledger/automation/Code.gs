/**
 * 我的帳本 — 通知收件匣（Google Apps Script）
 *
 * 用法：在 Google 試算表開啟「擴充功能 → Apps Script」，貼上這份程式，
 * 把下面的 KEY 改成你自己的一串亂碼，再部署成「網頁應用程式」。
 *
 * POST  …/exec?key=KEY   內文（純文字）第 1 行＝App 名稱、第 2 行＝通知標題、其餘＝通知內文
 * GET   …/exec?key=KEY&after=N   回傳第 N 列之後的通知（JSON）
 */
const KEY = '請改成你自己的一串亂碼';   // 至少 20 個字元，不要告訴別人
const SHEET_NAME = '通知';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['時間', 'App', '標題', '內文']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function authorized_(e) {
  return KEY !== '請改成你自己的一串亂碼' && e && e.parameter && e.parameter.key === KEY;
}

function doPost(e) {
  if (!authorized_(e)) return json_({ok: false, error: 'bad key'});
  const body = (e.postData && e.postData.contents) ? String(e.postData.contents) : '';
  const lines = body.replace(/\r\n/g, '\n').split('\n');
  const app = (lines.shift() || '').trim();
  const title = (lines.shift() || '').trim();
  const text = lines.join('\n').trim();
  if (!title && !text) return json_({ok: false, error: 'empty'});

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const row = sh.getLastRow() + 1;
    // 先把儲存格設成純文字，避免通知內容被試算表當成公式或數字
    sh.getRange(row, 2, 1, 3).setNumberFormat('@');
    sh.getRange(row, 1, 1, 4).setValues([[new Date(), app, title, text]]);
    return json_({ok: true, row: row});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (!authorized_(e)) return json_({ok: false, error: 'bad key'});
  const sh = getSheet_();
  const last = sh.getLastRow();
  const after = Math.max(1, Math.min(Number(e.parameter.after) || 1, last));
  const rows = [];
  if (last > after) {
    sh.getRange(after + 1, 1, last - after, 4).getValues().forEach(function (r, i) {
      rows.push({
        row: after + 1 + i,
        ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        app: String(r[1]), title: String(r[2]), text: String(r[3]),
      });
    });
  }
  return json_({ok: true, last: last, rows: rows});
}
