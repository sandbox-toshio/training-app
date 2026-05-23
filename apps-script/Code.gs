/**
 * ============================================================
 *  VOCAB APP — Google Apps Script バックエンド
 *  単語マスター読み込み + 進捗の読み書きを提供する Web App。
 *  デプロイ方法は README.md を参照。
 * ============================================================
 */

// ── 設定（必要に応じて変更） ───────────────────────────────
const SHEET_WORDS    = '単語';   // 単語マスタータブ名
const SHEET_PROGRESS = '進捗';   // 進捗タブ名（無ければ自動生成）

// 進捗タブのカラム（順序固定）
const PROGRESS_HEADERS = [
  'en', 'ja',
  'total_EJ', 'correct_EJ', 'streak_EJ',
  'total_JE', 'correct_JE', 'streak_JE',
  'last_wrong', 'updated_at'
];

// ────────────────────────────────────────────────────────────
// HTTP エンドポイント
// ────────────────────────────────────────────────────────────

/**
 * GET: 単語マスター + 進捗を一括取得
 *   URL末尾に ?action=all（省略可）
 */
function doGet(e) {
  try {
    return jsonOk({
      words:    readWords(),
      progress: readProgress(),
      ts:       new Date().toISOString(),
    });
  } catch (err) {
    return jsonErr(err);
  }
}

/**
 * POST: 進捗をまとめて upsert
 *   body: { action: 'updateProgress', updates: [{...}, ...] }
 *   ※ Content-Type は text/plain にすること（CORS preflight 回避）
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'updateProgress') {
      const n = writeProgress(body.updates || []);
      return jsonOk({ ok: true, written: n });
    }
    return jsonErr(new Error('unknown action: ' + body.action));
  } catch (err) {
    return jsonErr(err);
  }
}

// ────────────────────────────────────────────────────────────
// READ: 単語マスター
// 列: A:単元 / B:英単語 / C:意味 / D:出題ON（チェックボックス）
// ────────────────────────────────────────────────────────────
function readWords() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_WORDS);
  if (!sh) throw new Error('シートが見つかりません: ' + SHEET_WORDS);

  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, 4).getValues();

  return values
    .filter(r => String(r[1] || '').trim())  // 英単語が空の行は除外
    .map(r => ({
      unit: String(r[0] || '').trim(),
      en:   String(r[1] || '').trim(),
      ja:   String(r[2] || '').trim(),
      on:   r[3] === true || r[3] === 'TRUE' || r[3] === 1,
    }));
}

// ────────────────────────────────────────────────────────────
// READ: 進捗
// ────────────────────────────────────────────────────────────
function readProgress() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PROGRESS);
  if (!sh) return [];  // 未作成なら空
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(1, 1, last, PROGRESS_HEADERS.length).getValues();
  const headers = data[0].map(String);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ────────────────────────────────────────────────────────────
// WRITE: 進捗（en+ja でキーマッチして upsert）
// ────────────────────────────────────────────────────────────
function writeProgress(updates) {
  if (!Array.isArray(updates) || updates.length === 0) return 0;

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PROGRESS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROGRESS);
    sh.getRange(1, 1, 1, PROGRESS_HEADERS.length).setValues([PROGRESS_HEADERS]);
    sh.setFrozenRows(1);
  }

  const last = sh.getLastRow();
  const data = last > 0 ? sh.getRange(1, 1, last, PROGRESS_HEADERS.length).getValues() : [PROGRESS_HEADERS];
  const headers = data[0].map(String);
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  // 既存行を key で索引化
  const keyOf = (en, ja) => String(en) + '||' + String(ja);
  const rowMap = new Map();
  for (let r = 1; r < data.length; r++) {
    rowMap.set(keyOf(data[r][idx.en], data[r][idx.ja]), r + 1); // sheet rowは1始まり
  }

  const now = new Date().toISOString();
  const toAppend = [];
  let updatedCount = 0;

  updates.forEach(u => {
    const row = [
      u.en, u.ja,
      Number(u.total_EJ)   || 0,
      Number(u.correct_EJ) || 0,
      Number(u.streak_EJ)  || 0,
      Number(u.total_JE)   || 0,
      Number(u.correct_JE) || 0,
      Number(u.streak_JE)  || 0,
      u.last_wrong || '',
      now,
    ];
    const k = keyOf(u.en, u.ja);
    if (rowMap.has(k)) {
      sh.getRange(rowMap.get(k), 1, 1, PROGRESS_HEADERS.length).setValues([row]);
      updatedCount++;
    } else {
      toAppend.push(row);
    }
  });

  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, PROGRESS_HEADERS.length).setValues(toAppend);
  }
  return updatedCount + toAppend.length;
}

// ────────────────────────────────────────────────────────────
// JSON レスポンス
// ────────────────────────────────────────────────────────────
function jsonOk(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonErr(err) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err) }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────
// (開発用) ローカルテスト: GAS エディタで直接実行できる
// ────────────────────────────────────────────────────────────
function _devTestRead() {
  Logger.log(JSON.stringify({ words: readWords().slice(0, 3), progress: readProgress().slice(0, 3) }, null, 2));
}
function _devTestWrite() {
  const n = writeProgress([{
    en: 'I', ja: 'わたしは',
    total_EJ: 1, correct_EJ: 1, streak_EJ: 1,
    total_JE: 0, correct_JE: 0, streak_JE: 0,
    last_wrong: '',
  }]);
  Logger.log('written: ' + n);
}
