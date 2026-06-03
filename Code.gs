const SHEET_NAME = 'english-phrase';
const CLAUDE_MODEL = 'claude-sonnet-4-6';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAll') return handleGetAll();
  if (action === 'health') return handleHealth();
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
    date:     formatSheetDate(r[4]),
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
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'M/d')
    };
    appendEntry(entry);
    return jsonResponse({ ok: true, entry: entry });
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) });
  }
}

function formatSheetDate(cell) {
  if (cell === '' || cell == null) return '';
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone() || 'Asia/Tokyo', 'M/d');
  }
  const s = String(cell).trim();
  if (/^\d{1,2}\/\d{1,2}$/.test(s)) return s;
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone() || 'Asia/Tokyo', 'M/d');
  }
  return s;
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
      model: CLAUDE_MODEL,
      max_tokens: 1000,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (status !== 200) {
    const msg = body.error?.message || `Claude API error (${status})`;
    if (/model/i.test(msg)) {
      throw new Error('Claude のモデルが利用できません。GAS の Code.gs を最新版に更新して再デプロイしてください。');
    }
    throw new Error(msg);
  }

  const raw = body.content.map(b => b.text || '').join('');
  return JSON.parse(raw.replace(/```json|```/g, '').trim());
}

function handleHealth() {
  try {
    UrlFetchApp.fetch('https://httpbin.org/get', { muteHttpExceptions: true });
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message || String(err) });
  }
}

/**
 * 権限エラー時: Apps Script エディタでこの関数を1回実行 → 承認 → Webアプリを再デプロイ
 * デプロイ設定: 実行者=自分 / アクセス=全員
 */
function authorizeExternalRequest() {
  UrlFetchApp.fetch('https://httpbin.org/get', { muteHttpExceptions: true });
  PropertiesService.getScriptProperties().setProperty('EXTERNAL_AUTH_OK', '1');
}

/** スプレッドシートを開いたとき、所有者に権限付与を促す（1回） */
function onOpen() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('EXTERNAL_AUTH_OK') === '1') return;
  const ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  if (Session.getEffectiveUser().getEmail() !== ownerEmail) return;
  try {
    authorizeExternalRequest();
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Phrase: Claude API を使う権限が必要です。\n\n' +
      '1. 拡張機能 → Apps Script\n' +
      '2. authorizeExternalRequest を実行して「許可」\n' +
      '3. デプロイを更新（実行者: 自分 / アクセス: 全員）'
    );
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
