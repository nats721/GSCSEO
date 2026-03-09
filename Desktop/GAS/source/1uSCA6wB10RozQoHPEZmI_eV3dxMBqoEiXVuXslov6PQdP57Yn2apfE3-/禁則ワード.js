/**
 * スタンドアロンGAS用：禁則ワード検出（件数）＋月間クリック数（GSC最新）を挿入
 * + C列とD列の間に「プロンプト」列を追加（ChatGPT投入用）
 *
 * 参照元（別スプレッドシート）
 * 1) HTMLキャッシュ：A列=URL / H列=本文
 * 2) 禁則ワード：A列
 * 3) GSC最新：A列=URL / B列=直近30日のクリック数
 *
 * 出力先（禁則ワード側ブック）
 * - 「チェック結果」シート
 *   A列：URL
 *   B列：月間クリック数（直近30日）
 *   C列：合計（各禁則ワードの出現回数の合計）
 *   D列：プロンプト（ChatGPTに流す用）
 *   E列〜：ヘッダは「実際の禁則ワード」／セルは出現回数
 */

// ====== 設定 ======
const CACHE_SPREADSHEET_ID = "1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o";
const WORDS_SPREADSHEET_ID = "1-LnhzG2eE0C3XR0RK6G7-2AQIzT1cBeGMHMiZkpDNls";
const GSC_SPREADSHEET_ID   = "1ZHMy3vVUd_0MrSduXTe1etBb1HNpAlGwQr006ljwMiM";

const WORDS_SHEET_NAME  = "禁則ワード";
const RESULT_SHEET_NAME = "チェック結果";
const GSC_SHEET_NAME    = "GSC最新";

// 先頭行が見出しなら2、見出し無しなら1
const CACHE_START_ROW = 2;
const WORDS_START_ROW = 1;
const GSC_START_ROW   = 2; // 見出しがある前提。無ければ1に

// ====== 実行 ======
function runForbiddenWordCheck() {
  const cacheSS = SpreadsheetApp.openById(CACHE_SPREADSHEET_ID);
  const cacheSh = getSheetOrThrow_(cacheSS, CACHE_SHEET_NAME);

  const wordsSS = SpreadsheetApp.openById(WORDS_SPREADSHEET_ID);
  const wordsSh = getSheetOrThrow_(wordsSS, WORDS_SHEET_NAME);

  const gscSS = SpreadsheetApp.openById(GSC_SPREADSHEET_ID);
  const gscSh = getSheetOrThrow_(gscSS, GSC_SHEET_NAME);

  const words  = loadWords_(wordsSh, WORDS_START_ROW);           // ["ワード1", ...]
  const rows   = loadCacheRows_(cacheSh, CACHE_START_ROW);       // [{url, body}, ...]
  const gscMap = loadGscClicksMap_(gscSh, GSC_START_ROW);        // {url: clicks}

  writeResultCounts_(wordsSS, rows, words, gscMap);
}

// ====== データ取得 ======
function loadWords_(sh, startRow) {
  const last = sh.getLastRow();
  if (last < startRow) return [];

  // A:B（2列） A=禁則ワード / B=除外語（任意）
  const vals = sh.getRange(startRow, 1, last - startRow + 1, 2).getValues();

  const seen = new Set();
  const out = [];
  for (const row of vals) {
    const wRaw = row[0] == null ? "" : String(row[0]).trim();
    if (!wRaw) continue;
    if (seen.has(wRaw)) continue;
    seen.add(wRaw);

    const exRaw = row[1] == null ? "" : String(row[1]).trim();
    const excludes = exRaw
      ? exRaw.split(/\r?\n|\|/).map(s => s.trim()).filter(Boolean)
      : [];

    out.push({ word: wRaw, excludes });
  }
  return out; // [{word:"型", excludes:["文型"]}, ...]
}


function loadCacheRows_(sh, startRow) {
  const last = sh.getLastRow();
  if (last < startRow) return [];

  // A〜H（8列）を読む：A=URL(1) / H=本文(8)
  const vals = sh.getRange(startRow, 1, last - startRow + 1, 8).getValues();

  const out = [];
  for (const row of vals) {
    const url  = row[0] == null ? "" : String(row[0]).trim();
    const body = row[7] == null ? "" : String(row[7]);
    if (!url) continue;
    out.push({ url, body });
  }
  return out;
}

function loadGscClicksMap_(sh, startRow) {
  const last = sh.getLastRow();
  if (last < startRow) return {};

  // A:B（2列） A=URL / B=clicks
  const vals = sh.getRange(startRow, 1, last - startRow + 1, 2).getValues();

  const map = Object.create(null);
  for (const row of vals) {
    const url = row[0] == null ? "" : String(row[0]).trim();
    if (!url) continue;

    // clicks が空なら 0、数値/文字列の両対応
    let clicks = row[1];
    if (clicks == null || clicks === "") clicks = 0;
    const n = Number(clicks);
    map[url] = Number.isFinite(n) ? n : 0;
  }
  return map;
}

// ====== 出力（件数 + 月間クリック数 + プロンプト） ======
function writeResultCounts_(ss, cacheRows, words, gscMap) {
  let sh = ss.getSheetByName(RESULT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(RESULT_SHEET_NAME);

  sh.clearContents();

  // ヘッダ
  const headerWords = words.map(o => toSafeHeader_(o.word));
  const header = ["URL", "月間クリック数(30日)", "合計", "プロンプト", ...headerWords];

  const values = [header];

  for (const r of cacheRows) {
    const text = r.body || "";

    const counts = [];
    let sum = 0;

    for (const o of words) {
      const cnt = countOccurrencesWithExcludes_(text, o.word, o.excludes);
      counts.push(cnt);
      sum += cnt;
    }

    const clicks = (gscMap && Object.prototype.hasOwnProperty.call(gscMap, r.url)) ? gscMap[r.url] : 0;

    const usedWords = [];
    for (let i = 0; i < words.length; i++) {
      if (counts[i] > 0) usedWords.push(words[i].word);
    }

    const prompt = buildPrompt_(r.url, text, usedWords);

    values.push([r.url, clicks, sum, prompt, ...counts]);
  }

  sh.getRange(1, 1, values.length, header.length).setValues(values);
  sh.setFrozenRows(1);
}

function countOccurrencesWithExcludes_(text, word, excludes) {
  if (!text || !word) return 0;

  let t = String(text);

  // 例外語を「□」でマスク（wordが中にあってもカウントされない）
  if (excludes && excludes.length) {
    for (const ex of excludes) {
      if (!ex) continue;
      const placeholder = "□".repeat(ex.length);
      t = t.split(ex).join(placeholder); // 文字列置換（正規表現不要で安全）
    }
  }

  // 元の countOccurrences_ と同じ数え方
  let count = 0;
  let pos = 0;
  while (true) {
    const found = t.indexOf(word, pos);
    if (found === -1) break;
    count++;
    pos = found + word.length;
  }
  return count;
}


// ====== ChatGPT投入用プロンプト ======
function buildPrompt_(url, body, usedWords) {
  const used = (usedWords && usedWords.length)
    ? usedWords.map(w => `- ${w}`).join("\n")
    : "（検出なし）";

  // セル上限対策（ざっくり）
  const safeBody = truncateForCell_(body, 45000);

  const p =
`以下のページは不自然なワードを使用しています。
これらを適切に自然な言葉に置き換えて下さい。
なお、本文の装飾や構造は絶対に変えないでください。

【対象ページURL】
${url}

【検出された不自然ワード】
${used}

【本文】
------
${safeBody}
------`;

  return p;
}

function truncateForCell_(s, maxLen) {
  const str = (s == null) ? "" : String(s);
  if (str.length <= maxLen) return str;
  return str.slice(0, maxLen) + "\n\n（以下略：文字数上限のため途中で省略）";
}

// ====== 文字列の出現回数（重なりは数えない） ======
function countOccurrences_(text, word) {
  if (!text || !word) return 0;

  let count = 0;
  let pos = 0;

  while (true) {
    const found = text.indexOf(word, pos);
    if (found === -1) break;
    count++;
    pos = found + word.length;
  }
  return count;
}

// ====== ユーティリティ ======
function getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("シートが見つかりません: " + name);
  return sh;
}

function toSafeHeader_(s) {
  const t = String(s);
  if (/^[=+\-@]/.test(t)) return "'" + t;
  return t;
}

/**
 * （任意）時間主導トリガーを作る：毎日1回など
 */
function installDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "runForbiddenWordCheck") ScriptApp.deleteTrigger(t);
  }

  ScriptApp.newTrigger("runForbiddenWordCheck")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}
