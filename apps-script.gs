/**
 * Video Ads — form backend (Google Apps Script, free).
 * Saves every request as a row in this Google Sheet and emails you a copy.
 * Setup (one time, ~3 minutes, on a computer):
 *  1. Open https://sheets.new  → name it "Video Ads Requests".
 *  2. Extensions → Apps Script → delete the sample code → paste this file → Save.
 *  3. Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone → Deploy → allow access.
 *  4. Copy the Web app URL (ends with /exec) and send it to Claude (or paste it into ENDPOINT in index.html).
 */
const SHEET_NAME = 'Requests';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.website) return out({ ok: true });            // honeypot
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    let headers = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (!headers.length) { headers = ['received_at']; sh.appendRow(headers); sh.setFrozenRows(1); }
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    });
    const row = headers.map(function (h) {
      if (h === 'received_at') return new Date();
      const v = data[h];
      return Array.isArray(v) ? v.join('، ') : (v == null ? '' : String(v));
    });
    sh.appendRow(row);
    try {
      const body = Object.keys(data).map(function (k) {
        const v = data[k]; return k + ': ' + (Array.isArray(v) ? v.join('، ') : v);
      }).join('\n');
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        'طلب فيديو جديد — ' + (data.business_name || data.name || ''), body);
    } catch (err) {}
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out({ ok: true, service: 'Video Ads form' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
