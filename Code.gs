/**
 * ============================================================
 *  Vocab App — Google Apps Script (Web App)
 *  Sheets を読み書きして、アプリと同期させるためのバックエンド
 * ============================================================
 *
 *  使い方:
 *   1. スプレッドシートを開く
 *   2. 「拡張機能」→「Apps Script」
 *   3. デフォルトの Code.gs の中身を全部消して、このファイルを丸ごと貼り付け
 *   4. 上のメニューから「デプロイ」→「新しいデプロイ」
 *   5. 種類: ウェブアプリ
 *      - 実行ユーザー: 自分
 *      - アクセスできるユーザー: 全員
 *   6. 表示された Web App URL をコピーして、アプリの設定画面に貼る
 *
 *  ※ 初回デプロイ時に Google から認証承認を求められます。
 *     「詳細」→「(プロジェクト名)に移動」→ 許可、で進めてください。
 *
 *  タブ名を変えている場合は、下の WORDS_TAB / PROGRESS_TAB を書き換えてください。
 */

const WORDS_TAB    = '単語';   // 単語マスタータブ（A:単元 B:英単語 C:意味 D:出題ON）
const PROGRESS_TAB = '進捗';   // 進捗タブ（無ければ自動作成）

// 進捗タブの列順（このまま変えないでください。アプリと対応しています）
const PROGRESS_HEADERS = [
  'en', 'ja',
  'total_ej', 'correct_ej', 'streak_ej',
  'total_je', 'correct_je', 'streak_je',
  'last_wrong', 'updated_at'
];

/* ============================================================
 *  TOKEN — スクリプトプロパティ SECRET_TOKEN を必須化
 *  設定方法:
 *    Apps Script エディタ → 左メニューの歯車（プロジェクトの設定）
 *    → スクリプトプロパティ → 追加
 *    キー: SECRET_TOKEN  /  値: 任意のランダム文字列
 *  値を設定したら、必ず「デプロイ → デプロイを管理 → 編集 → バージョン:新しいバージョン → デプロイ」で再デプロイすること
 * ============================================================ */
function requireToken_(e, body) {
  const expected = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN');
  if (!expected) {
    throw new Error('SECRET_TOKEN is not set on the server. Set it in Script Properties.');
  }
  const provided =
    (e && e.parameter && e.parameter.token) ||
    (body && body.token) || '';
  if (provided !== expected) {
    throw new Error('unauthorized');
  }
}

/* ============================================================
 *  GET — アプリ起動時に呼ばれる。単語マスタと進捗を返す
 *    クエリパラメータ: ?token=xxxx
 * ============================================================ */
function doGet(e) {
  try {
    requireToken_(e, null);
    const ss = SpreadsheetApp.getActive();
    const words    = readWordsTab(ss);
    const progress = readProgressTab(ss);
    return jsonResponse({
      ok: true,
      server_time: new Date().toISOString(),
      words,
      progress
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/* ============================================================
 *  POST — セッション終了時にアプリから呼ばれる。進捗を upsert
 *  body 例:
 *    { "token":"xxxx",
 *      "updates": [
 *        { "en":"teacher", "ja":"先生/教師",
 *          "total_ej":12, "correct_ej":10, "streak_ej":3,
 *          "total_je":8,  "correct_je":7,  "streak_je":5,
 *          "last_wrong":"2026-05-20T12:34:56Z" },
 *        ...
 *      ] }
 * ============================================================ */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    requireToken_(e, body);
    const updates = Array.isArray(body.updates) ? body.updates : [];
    const ss = SpreadsheetApp.getActive();
    const result = upsertProgress(ss, updates);
    return jsonResponse({ ok: true, ...result });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/* ============================================================
 *  単語マスタを読む（A:単元 B:英単語 C:意味 D:出題ON）
 * ============================================================ */
function readWordsTab(ss) {
  const sheet = ss.getSheetByName(WORDS_TAB);
  if (!sheet) throw new Error('「' + WORDS_TAB + '」タブが見つかりません');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  // 1行目はヘッダ想定。データは2行目以降
  return values.slice(1)
    .map(row => ({
      unit: String(row[0] || '').trim(),
      en:   String(row[1] || '').trim(),
      ja:   String(row[2] || '').trim(),
      on:   toBool(row[3]),
    }))
    .filter(w => w.en && w.ja);
}

function toBool(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'ON' || s === '✓' || s === 'YES';
}

/* ============================================================
 *  進捗タブを読む（無ければ作る）
 * ============================================================ */
function readProgressTab(ss) {
  const sheet = ensureProgressSheet(ss);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(o => o.en);
}

/* ============================================================
 *  進捗タブの存在確認 + ヘッダ自動作成
 * ============================================================ */
function ensureProgressSheet(ss) {
  let sheet = ss.getSheetByName(PROGRESS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRESS_TAB);
  }
  // ヘッダ行が空 or 不一致なら書く
  const firstRow = sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length).getValues()[0];
  const isEmpty = firstRow.every(v => v === '' || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length).setValues([PROGRESS_HEADERS]);
    sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ============================================================
 *  進捗を upsert（en+ja の組み合わせをキーに、既存なら上書き、無ければ追加）
 * ============================================================ */
function upsertProgress(ss, updates) {
  const sheet = ensureProgressSheet(ss);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const enIdx = headers.indexOf('en');
  const jaIdx = headers.indexOf('ja');
  if (enIdx < 0 || jaIdx < 0) throw new Error('進捗タブのヘッダに en/ja がありません');

  // 既存行のキー → 行番号(1-indexed)
  const rowByKey = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][enIdx]) + '' + String(values[i][jaIdx]);
    if (values[i][enIdx]) rowByKey[key] = i + 1;
  }

  const now = new Date().toISOString();
  let updated = 0;
  const appends = [];

  updates.forEach(u => {
    const key = String(u.en) + '' + String(u.ja);
    const rowData = PROGRESS_HEADERS.map(h => {
      if (h === 'updated_at') return now;
      const v = u[h];
      return (v === undefined || v === null) ? '' : v;
    });
    if (rowByKey[key]) {
      sheet.getRange(rowByKey[key], 1, 1, PROGRESS_HEADERS.length).setValues([rowData]);
      updated++;
    } else {
      appends.push(rowData);
    }
  });

  if (appends.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, appends.length, PROGRESS_HEADERS.length).setValues(appends);
  }

  return { updated, inserted: appends.length };
}

/* ============================================================
 *  JSONレスポンス共通
 * ============================================================ */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 *  自己テスト — Apps Script エディタで実行ボタンから動作確認できる
 *  関数名 testRead を選んで実行すると、ログに件数が出ます
 *  （エディタから直接実行する場合はトークン不要）
 * ============================================================ */
function testRead() {
  const ss = SpreadsheetApp.getActive();
  const words = readWordsTab(ss);
  const progress = readProgressTab(ss);
  Logger.log('単語: ' + words.length + '件');
  Logger.log('  最初の3件: ' + JSON.stringify(words.slice(0, 3)));
  Logger.log('進捗: ' + progress.length + '件');
  const token = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN');
  Logger.log('SECRET_TOKEN: ' + (token ? '設定済み（' + token.length + '文字）' : '⚠️ 未設定！'));
}
