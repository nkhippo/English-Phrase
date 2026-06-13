const SHEET_NAME = 'english-phrase';
const CLAUDE_MODEL = 'claude-sonnet-4-6';
const DRIVE_FOLDER_NAME = 'EnglishPhrase_Audio';
const EXAMPLES_PER_ENTRY = 5;
const OPENAI_TTS_MODEL = 'gpt-4o-mini-tts';
const OPENAI_TTS_VOICE = 'nova';
const OPENAI_TTS_INSTRUCTIONS = 'Speak clearly and at a moderate pace, suitable for English language learning. Pronounce each word distinctly.';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAll') return handleGetAll();
  if (action === 'health') return handleHealth();
  if (action === 'getAudio') return handleGetAudio(e.parameter.id, Number(e.parameter.idx || 0));
  return jsonResponse({ error: 'unknown action' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === 'save') return handleSave(data.entry);
  if (data.action === 'generateAndSave') return handleGenerateAndSave(data.input, data.note, data.apiKey);
  if (data.action === 'confirmCard') return handleConfirmCard(data.id, data.confirmedAt);
  if (data.action === 'deleteEntry') return handleDeleteEntry(data.id);
  if (data.action === 'generateAudioBatch') return handleGenerateAudioBatch();
  return jsonResponse({ error: 'unknown action' });
}

function handleGetAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return jsonResponse({ entries: [] });

  const entries = rows.slice(1).map(r => ({
    id:           r[0],
    input:        r[1],
    keyword:      r[2],
    examples:     JSON.parse(r[3]),
    date:         formatSheetDate(r[4]),
    note:         r[5] || '',
    ipa:          r[6] || '',
    pos:          normalizePos(r[7]),
    confirmedAt:  parseConfirmedAt(r[8]),
    audioUrls:    parseAudioUrls(r[9]),
    idx:          0,
    show:         false
  })).reverse();

  return jsonResponse({ entries });
}

function handleSave(entry) {
  appendEntry(entry);
  return jsonResponse({ ok: true });
}

function toLowerEntryText(s) {
  return String(s || '').toLowerCase();
}

function handleGenerateAndSave(input, note, apiKey) {
  try {
    const normalizedInput = toLowerEntryText(input);
    const generated = generateExamples(normalizedInput, note || '', apiKey);
    const parsed = generated.parsed;
    const keyword = toLowerEntryText(parsed.keyword);
    const ipa = ensureIpa(keyword, apiKey, parsed, generated.raw);
    const entryId = Date.now();
    const entry = {
      id: entryId,
      input: normalizedInput,
      note: note || '',
      keyword: keyword,
      ipa: ipa,
      pos: normalizePos(parsed.partOfSpeech || parsed.pos || parsed.part_of_speech),
      examples: parsed.examples,
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'M/d'),
      confirmedAt: '',
      audioUrls: {}
    };

    try {
      entry.audioUrls = generateAudioForAllExamples(entry);
    } catch (audioErr) {
      console.error('Audio generation failed: ' + audioErr.message);
    }

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
    entry.note || '',
    entry.ipa || '',
    entry.pos || '',
    entry.confirmedAt || '',
    JSON.stringify(entry.audioUrls || {})
  ]);
}

function handleConfirmCard(id, confirmedAt) {
  if (id == null || id === '') return jsonResponse({ error: 'id required' });
  const ok = updateConfirmedAt(id, confirmedAt != null ? confirmedAt : Date.now());
  if (!ok) return jsonResponse({ error: 'entry not found' });
  return jsonResponse({ ok: true });
}

function updateConfirmedAt(id, confirmedAt) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 9).setValue(Number(confirmedAt) || Date.now());
      return true;
    }
  }
  return false;
}

function handleDeleteEntry(id) {
  if (id == null || id === '') return jsonResponse({ error: 'id required' });
  const ok = deleteEntryById(id);
  if (!ok) return jsonResponse({ error: 'entry not found' });
  return jsonResponse({ ok: true });
}

function deleteEntryById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function parseAudioUrls(cell) {
  if (!cell || cell === '') return {};
  try { return JSON.parse(cell); } catch (e) { return {}; }
}

function parseConfirmedAt(cell) {
  if (cell === '' || cell == null) return '';
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    return cell.getTime();
  }
  const n = Number(cell);
  return isNaN(n) || n <= 0 ? '' : n;
}

function normalizePos(pos) {
  const raw = String(pos || '').trim().toLowerCase();
  if (!raw) return '';
  const aliases = {
    'n': 'noun', 'n.': 'noun', 'noun': 'noun',
    'v': 'verb', 'v.': 'verb', 'verb': 'verb',
    'adj': 'adjective', 'adj.': 'adjective', 'adjective': 'adjective',
    'adv': 'adverb', 'adv.': 'adverb', 'adverb': 'adverb',
    'prep': 'preposition', 'prep.': 'preposition', 'preposition': 'preposition',
    'conj': 'conjunction', 'conj.': 'conjunction', 'conjunction': 'conjunction',
    'pron': 'pronoun', 'pron.': 'pronoun', 'pronoun': 'pronoun',
    'interj': 'interjection', 'interjection': 'interjection',
    'idiom': 'idiom', 'phrase': 'phrase', 'expression': 'phrase'
  };
  return aliases[raw] || raw.replace(/\s+/g, ' ');
}

function generateExamples(input, note, apiKey) {
  if (!apiKey) throw new Error('APIキーが未設定です。アプリ上部の「APIキー設定」から登録してください。');

  const notePart = note ? `補足メモ：${note}。この補足も例文生成に活かしてください。` : '';
  const prompt =
    `ユーザーが「${input}」を英語で学びたいと入力しました。${notePart}` +
    'この意図・単語・フレーズに合った英文例を5つ作成してください。日常会話・ビジネス・留学生活に自然なものを選んでください。\n' +
    '例文の形式ルール:\n' +
    '- 各例文は1つの完結した文にすること（A: / B: などの対話形式は使わない）\n' +
    '- 複数話者の会話を1つの en フィールドにまとめないこと\n' +
    '- 会話調にしたい場合も、登録ワードを含む話者の発話だけを1文として書くこと\n' +
    '必ず次のJSONのみを返してください（前置き・説明・コードブロック禁止）。ipa は必須で空文字不可:\n' +
    '{"keyword":"核となる英単語またはフレーズ（小文字）","ipa":"/IPA発音記号/","partOfSpeech":"品詞（noun, verb, adjective, adverb, preposition, conjunction, pronoun, interjection, idiom, phrase のいずれか1つ）","examples":[{"en":"英文1","ja":"日本語訳1"},{"en":"英文2","ja":"日本語訳2"},{"en":"英文3","ja":"日本語訳3"},{"en":"英文4","ja":"日本語訳4"},{"en":"英文5","ja":"日本語訳5"}]}';

  const raw = callClaude(apiKey, prompt, 1500);
  const parsed = parseGeneratedJson(raw);
  if (!parsed.examples || !parsed.examples.length) {
    throw new Error('例文の生成に失敗しました。もう一度お試しください。');
  }
  return { parsed: parsed, raw: raw };
}

function callClaude(apiKey, prompt, maxTokens) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: maxTokens,
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

  return body.content.map(b => b.text || '').join('');
}

function parseGeneratedJson(raw) {
  const cleaned = raw.replace(/```json|```/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (err) {
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    throw err;
  }
}

function extractIpaFromParsed(parsed, raw) {
  const candidate = parsed.ipa || parsed.IPA || parsed.pronunciation || parsed.phonetic;
  if (candidate) return normalizeIpa(candidate);

  const quoted = raw.match(/"ipa"\s*:\s*"((?:\\.|[^"\\])*)"/i);
  if (quoted) return normalizeIpa(quoted[1].replace(/\\"/g, '"'));

  const slash = raw.match(/\/[^/\n]{1,80}\//);
  if (slash) return normalizeIpa(slash[0]);

  return '';
}

function normalizeIpa(ipa) {
  let s = String(ipa || '').trim();
  if (!s) return '';
  s = s.replace(/^['"]|['"]$/g, '');
  if (!s.startsWith('/')) s = '/' + s;
  if (!s.endsWith('/')) s = s + '/';
  return s;
}

function ensureIpa(keyword, apiKey, parsed, raw) {
  let ipa = extractIpaFromParsed(parsed, raw);
  if (ipa) return ipa;

  const fallbackPrompt =
    `英語の単語またはフレーズ「${keyword}」のIPA発音記号だけを返してください。` +
    'JSONのみ: {"ipa":"/ここにIPA/"}';
  const fallbackRaw = callClaude(apiKey, fallbackPrompt, 200);
  const fallbackParsed = parseGeneratedJson(fallbackRaw);
  ipa = extractIpaFromParsed(fallbackParsed, fallbackRaw);
  if (ipa) return ipa;

  throw new Error('IPA発音記号の生成に失敗しました。もう一度お試しください。');
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

function callOpenAiTts(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY がスクリプトプロパティに未設定です。');

  const payload = {
    model: OPENAI_TTS_MODEL,
    voice: OPENAI_TTS_VOICE,
    input: text,
    instructions: OPENAI_TTS_INSTRUCTIONS,
    response_format: 'mp3'
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/audio/speech', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  if (status !== 200) {
    const body = res.getContentText();
    throw new Error('OpenAI TTS API error (' + status + '): ' + body);
  }

  return res.getContent();
}

function saveAudioToDrive(audioBytes, entryId, idx) {
  const folder = getOrCreateAudioFolder();
  const fileName = entryId + '_' + idx + '.mp3';

  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const blob = Utilities.newBlob(audioBytes, 'audio/mpeg', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return 'https://drive.google.com/uc?export=download&id=' + file.getId();
}

function getOrCreateAudioFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  return folder;
}

function generateAudioForAllExamples(entry, existingUrls) {
  const audioUrls = Object.assign({}, existingUrls || {});
  const examples = entry.examples || [];
  const count = Math.min(examples.length, EXAMPLES_PER_ENTRY);

  for (let i = 0; i < count; i++) {
    const key = String(i);
    if (audioUrls[key]) continue;

    const text = examples[i] && examples[i].en;
    if (!text) continue;
    try {
      const audioBytes = callOpenAiTts(text);
      const url = saveAudioToDrive(audioBytes, entry.id, i);
      audioUrls[key] = url;
      Utilities.sleep(300);
    } catch (err) {
      console.error('Audio failed for idx ' + i + ': ' + err.message);
    }
  }
  return audioUrls;
}

function isAllAudioCached(examples, audioUrls) {
  const count = Math.min((examples || []).length, EXAMPLES_PER_ENTRY);
  if (count === 0) return true;
  for (let i = 0; i < count; i++) {
    if (!audioUrls[String(i)]) return false;
  }
  return true;
}

function handleGetAudio(id, idx) {
  if (!id) return jsonResponse({ error: 'id required' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const audioUrls = parseAudioUrls(rows[i][9]);
    const key = String(idx);

    if (audioUrls[key]) {
      return jsonResponse({ audioUrl: audioUrls[key] });
    }

    const examples = JSON.parse(rows[i][3] || '[]');
    const text = examples[idx] && examples[idx].en;
    if (!text) return jsonResponse({ error: 'example not found' });

    try {
      const audioBytes = callOpenAiTts(text);
      const url = saveAudioToDrive(audioBytes, id, idx);
      audioUrls[key] = url;
      sheet.getRange(i + 1, 10).setValue(JSON.stringify(audioUrls));
      return jsonResponse({ audioUrl: url });
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }

  return jsonResponse({ error: 'entry not found' });
}

function handleGenerateAudioBatch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  let processed = 0;
  let errors = 0;

  for (let i = 1; i < rows.length; i++) {
    const id = rows[i][0];
    const existingUrls = parseAudioUrls(rows[i][9]);
    const examples = JSON.parse(rows[i][3] || '[]');
    if (isAllAudioCached(examples, existingUrls)) continue;

    try {
      const merged = generateAudioForAllExamples({ id: id, examples: examples }, existingUrls);
      sheet.getRange(i + 1, 10).setValue(JSON.stringify(merged));
      processed++;
      Utilities.sleep(500);
    } catch (err) {
      console.error('Batch audio error for id=' + id + ': ' + err.message);
      errors++;
    }
  }

  return jsonResponse({ ok: true, processed: processed, errors: errors });
}

function runAudioBatchGeneration() {
  const result = handleGenerateAudioBatch();
  console.log(result.getContent());
}

function generateAudioBatchFrom(startRow, n) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  let processed = 0;

  for (let i = startRow; i < Math.min(startRow + n, rows.length); i++) {
    const id = rows[i][0];
    const existingUrls = parseAudioUrls(rows[i][9]);
    const examples = JSON.parse(rows[i][3] || '[]');
    if (isAllAudioCached(examples, existingUrls)) continue;

    try {
      const merged = generateAudioForAllExamples({ id: id, examples: examples }, existingUrls);
      sheet.getRange(i + 1, 10).setValue(JSON.stringify(merged));
      processed++;
      Utilities.sleep(500);
    } catch (err) {
      console.error('Row ' + i + ' failed: ' + err.message);
    }
  }
  console.log('Done. processed=' + processed);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
