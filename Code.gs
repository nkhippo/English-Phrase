const SHEET_NAME = 'english-phrase';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAll') return handleGetAll();
  return jsonResponse({ error: 'unknown action' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === 'save') return handleSave(data.entry);
  if (data.action === 'generateAndSave') return handleGenerateAndSave(data.input, data.note, data.apiKey);
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
  appendEntry(entry);
  return jsonResponse({ ok: true });
}

function handleGenerateAndSave(input, note, apiKey) {
  try {
    const parsed = generateExamples(input, note || '', apiKey);
    const entry = {
      id: Date.now(),
      input: input,
      note: note || '',
      keyword: parsed.keyword,
      examples: parsed.examples,
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'MM/dd')
    };
    appendEntry(entry);
    return jsonResponse({ ok: true, entry: entry });
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) });
  }
}

function appendEntry(entry) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([
    entry.id,
    entry.input,
    entry.keyword,
    JSON.stringify(entry.examples),
    entry.date,
    entry.note || ''
  ]);
}

function generateExamples(input, note, apiKey) {
  if (!apiKey) throw new Error('APIキーが未設定です。アプリ上部の「APIキー設定」から登録してください。');

  const notePart = note ? `補足メモ：${note}。この補足も例文生成に活かしてください。` : '';
  const prompt =
    `ユーザーが「${input}」を英語で学びたいと入力しました。${notePart}` +
    'この意図・単語・フレーズに合った英文例を5つ作成してください。日常会話・ビジネス・留学生活に自然なものを選んでください。必ずJSON形式のみで回答（前置き・説明・コードブロック不要）:\n' +
    '{"keyword":"核となる英単語またはフレーズ","examples":[{"en":"英文1","ja":"日本語訳1"},{"en":"英文2","ja":"日本語訳2"},{"en":"英文3","ja":"日本語訳3"},{"en":"英文4","ja":"日本語訳4"},{"en":"英文5","ja":"日本語訳5"}]}';

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (status !== 200) {
    throw new Error(body.error?.message || `Claude API error (${status})`);
  }

  const raw = body.content.map(b => b.text || '').join('');
  return JSON.parse(raw.replace(/```json|```/g, '').trim());
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
