/**
 * スタンドアロンGAS用：
 * 指定スプレッドシートの CacheDiffLog から
 * 「fetch日時が7日より前」の行をブランド（ドメイン）別 DiffLog_<domain> に退避し、
 * 元から削除します。
 *
 * 対象ブック：
 * https://docs.google.com/spreadsheets/d/1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o/
 */

const SPREADSHEET_ID = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o';
const SOURCE_SHEET_NAME = 'CacheDiffLog';
const DAYS_TO_KEEP = 2;
const HEADER_ROW = 1;
const DEST_PREFIX = 'DiffLog_';

const ALLOWED_DOMAINS = [
  'natsui-sansu-juku.com',
  'dokkai-labo.tokyo',
  'juku-escot.com',
  'jukureiwa.com',
  'koten-kakitsubata.jp',
  'eigo-rewrite.com',
  'eng-support.com',
  'kokugo-trigger.com',
  'rika-quest.com',
];

// 許可リスト外をどうするか：true=元に残す（安全） / false=unknownへ退避
const KEEP_UNKNOWN_DOMAIN_ROWS = true;
const UNKNOWN_DOMAIN_SHEET = 'DiffLog_unknown';

function archiveOldCacheDiffLogByBrand() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) throw new Error(`シートが見つかりません: ${SOURCE_SHEET_NAME}`);

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow <= HEADER_ROW) return;

  const header = src.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const idxOf = (name) => {
    const i = header.indexOf(name);
    if (i === -1) throw new Error(`ヘッダに「${name}」列が見つかりません`);
    return i;
  };

  const urlIdx = idxOf('URL');
  const fetchIdx = idxOf('fetch日時');

  const data = src.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();

  const now = new Date();
  const threshold = new Date(now.getTime() - DAYS_TO_KEEP * 24 * 60 * 60 * 1000);

  const allowedSet = new Set(ALLOWED_DOMAINS);

  // destSheetName -> rows
  const bucket = new Map();
  const rowNumsToDelete = [];

  data.forEach((row, i) => {
    const sheetRowNum = HEADER_ROW + 1 + i;

    const fetchDate = parseFetchDate_(row[fetchIdx]);
    if (!fetchDate) return; // 解釈不能は触らない（安全）

    if (fetchDate < threshold) {
      const url = String(row[urlIdx] || '');
      const domain = normalizeDomain_(extractDomain_(url));
      if (!domain) return;

      if (!allowedSet.has(domain)) {
        if (KEEP_UNKNOWN_DOMAIN_ROWS) return;

        const destName = sanitizeSheetName_(UNKNOWN_DOMAIN_SHEET);
        if (!bucket.has(destName)) bucket.set(destName, []);
        bucket.get(destName).push(row);
        rowNumsToDelete.push(sheetRowNum);
        return;
      }

      const destName = sanitizeSheetName_(`${DEST_PREFIX}${domain}`);
      if (!bucket.has(destName)) bucket.set(destName, []);
      bucket.get(destName).push(row);
      rowNumsToDelete.push(sheetRowNum);
    }
  });

  if (bucket.size === 0) return;

  // 退避先へappend（必要なら作成＋ヘッダ同期）
  bucket.forEach((rows, destSheetName) => {
    const dest = ensureSheetWithHeader_(ss, destSheetName, header);
    const writeStartRow = Math.max(dest.getLastRow() + 1, 2);
    dest.getRange(writeStartRow, 1, rows.length, lastCol).setValues(rows);
  });

  // 元から削除（下から）
  deleteRowsInRuns_(src, rowNumsToDelete);
}

/** fetch日時 を Date に寄せる（2桁年/4桁年、秒省略にも対応） */
function parseFetchDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;

  const s = String(val).trim();
  const m = s.match(/^(\d{2,4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (!m) return null;

  let year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  const hh = Number(m[4]);
  const mm = Number(m[5]);
  const ss = Number(m[6] || 0);

  if (year < 100) year += 2000;

  const d = new Date(year, month - 1, day, hh, mm, ss);
  return isNaN(d.getTime()) ? null : d;
}

function extractDomain_(url) {
  try {
    return new URL(url).hostname;
  } catch (e) {
    const m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
    return m ? m[1] : null;
  }
}

function normalizeDomain_(hostname) {
  let h = String(hostname || '').toLowerCase().trim();
  if (!h) return null;
  if (h.startsWith('www.')) h = h.slice(4);
  return h;
}

function sanitizeSheetName_(name) {
  let s = String(name).replace(/[\[\]\:\*\?\/\\]/g, '_');
  if (s.length > 100) s = s.slice(0, 100);
  return s;
}

function ensureSheetWithHeader_(ss, sheetName, headerRowValues) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, headerRowValues.length).setValues([headerRowValues]);
    sh.setFrozenRows(1);
  } else {
    // ヘッダが空なら補完（既存ヘッダは壊さない）
    const lastCol = headerRowValues.length;
    const current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const isEmpty = current.every(v => v === '' || v === null);
    if (isEmpty) {
      sh.getRange(1, 1, 1, lastCol).setValues([headerRowValues]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function deleteRowsInRuns_(sheet, rowNums) {
  const nums = Array.from(new Set(rowNums)).sort((a, b) => a - b);
  if (nums.length === 0) return;

  const runs = [];
  let start = nums[0];
  let prev = nums[0];

  for (let i = 1; i < nums.length; i++) {
    const cur = nums[i];
    if (cur === prev + 1) {
      prev = cur;
      continue;
    }
    runs.push({ start, count: prev - start + 1 });
    start = cur;
    prev = cur;
  }
  runs.push({ start, count: prev - start + 1 });

  for (let i = runs.length - 1; i >= 0; i--) {
    sheet.deleteRows(runs[i].start, runs[i].count);
  }
}