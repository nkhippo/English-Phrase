const SHEET_NAME = 'english-phrase';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAll') return handleGetAll();
  return jsonResponse({ error: 'unknown action' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === 'save') return handleSave(data.entry);
  return jsonResponse({ error: 'unknown action' });
}

function handleGetAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return jsonResponse({ entries: [] });

  const entries = rows.slice(1).map(r => ({
    id:       r[0],
    input:    r[1],
    keyword:  r[2],
    examples: JSON.parse(r[3]),
    date:     r[4],
    note:     r[5] || '',
    idx:      0,
    show:     false
  })).reverse();

  return jsonResponse({ entries });
}

function handleSave(entry) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([
    entry.id,
    entry.input,
    entry.keyword,
    JSON.stringify(entry.examples),
    entry.date,
    entry.note || ''
  ]);
  return jsonResponse({ ok: true });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
