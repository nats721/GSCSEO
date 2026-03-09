/************************************************************
 * UpdateImpactReport（本文更新→次更新 区間 × GSC日別集計）
 *
 * 入力：
 * - GSC日別（SS: 1ZHMy... / gid=1223550253）
 *   列: 取得日時 / データ基準日 / ドメイン / URL / clicks / impressions / ctr / position
 *
 * - CacheDiffLog（SS: 1Zyg... / gid=1003771447）
 *   列: URL / 分類 / lastmod / fetch日時 / title / metaDesc / H1 / entryBodyHtml / ...
 *
 * 参考（今回は必須ではない）：
 * - HTMLキャッシュ（SS: 1Zyg... / gid=1666145974）
 *
 * 出力：
 * - GSC SS 内に "UpdateImpactReport" シートを作成/上書き
 ************************************************************/

// ====== スプレッドシート（衝突回避のため UIR_ プレフィックス） ======
const UIR_GSC_SS_ID = '1ZHMy3vVUd_0MrSduXTe1etBb1HNpAlGwQr006ljwMiM';
const UIR_GSC_DAILY_SHEET_GID = 1223550253;

const UIR_CACHE_SS_ID = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o';
const UIR_HTML_CACHE_SHEET_GID = 1666145974;      // 今回は必須ではない（存在チェックもしない）
const UIR_CACHEDIFFLOG_SHEET_GID = 1003771447;    // ここが履歴の本体

const REPORT_SHEET_NAME = 'UpdateImpactReport';

// ====== 列定義 ======
// GSC日別: 取得日時 / データ基準日 / ドメイン / URL / clicks / impressions / ctr / position
const UIR_DAILY_COL = { runAt: 1, dataDate: 2, domain: 3, url: 4, clicks: 5, impr: 6, ctr: 7, pos: 8 };

// CacheDiffLog/HTMLキャッシュ共通: URL / 分類 / lastmod / fetch日時 / title / metaDesc / H1 / entryBodyHtml / ...
const UIR_CACHE_COL = { url: 1, fetchedAt: 4, title: 5, meta: 6, body: 8 };

// ====== パラメータ ======
const UIR_LAG_DAYS = 3;             // 区間端の削り（GSC遅延/揺れ）
const UIR_SEG_DAYS_DEFAULT = 7;     // 前半/後半の比較幅
const UIR_MIN_INTERVAL_DAYS = 7;    // 短すぎる区間は「短区間モード」へ（除外しない）
const UIR_SAMPLE_CLIP = 500;        // Added/Removedサンプルの最大文字数

// 短い更新間隔（週次など）でも比較できるようにする
const UIR_SHORT_WINDOW_DAYS = 3;      // 短区間用の比較窓（日数）
const UIR_SHORT_SKIP_UPDATE_DAY = 1;  // 更新当日は除外（翌日から）

/***********************
 * ENTRY
 ***********************/
function run_UpdateImpactReport() {
  const gscSS   = SpreadsheetApp.openById(UIR_GSC_SS_ID);
  const cacheSS = SpreadsheetApp.openById(UIR_CACHE_SS_ID);

  const dailySheet   = UIR_getSheetByGid_(gscSS, UIR_GSC_DAILY_SHEET_GID, 'GSC日別');
  const diffLogSheet = UIR_getSheetByGid_(cacheSS, UIR_CACHEDIFFLOG_SHEET_GID, 'CacheDiffLog');

  // 参考（今回は使わないが、取れるようにだけしておく）
  const htmlCacheSheet = UIR_getSheetByGid_(cacheSS, UIR_HTML_CACHE_SHEET_GID, 'HTMLキャッシュ');

  if (!dailySheet)   throw new Error('GSC日別シートが見つかりません（gid/nameを確認）');
  if (!diffLogSheet) throw new Error('CacheDiffLogシートが見つかりません（gid/nameを確認）');

  // 取り込み（必要列だけ）
  const dailyRows   = UIR_readSheetRows_(dailySheet, 8);     // A-H
  const diffLogRows = UIR_readSheetRows_(diffLogSheet, 8);   // A-H（entryBodyHtmlまで）

  // インデックス作成（正規化込み）
  const dailyIndex = UIR_buildDailyIndex_(dailyRows);              // domain|||url|||date -> metrics
  const dailyPairs = UIR_buildDailyPairs_(dailyRows);              // Set(domain|||url)
  const updateEventsByUrl = UIR_buildUpdateEventsByUrl_(diffLogRows); // url -> events[]

  const report = [];
  report.push([
    'Domain',
    'URL',
    'UpdateAt',
    'NextUpdateAt',
    'IntervalStart',
    'IntervalEnd',
    'IntervalDays',
    'TitleChanged',
    'MetaChanged',
    'AddedLines',
    'RemovedLines',
    'AddedSample',
    'RemovedSample',
    'FoundDays(FirstSeg)',
    'FoundDays(LastSeg)',
    'Clicks(FirstSeg)',
    'Clicks(LastSeg)',
    'ΔClicks',
    'Impr(FirstSeg)',
    'Impr(LastSeg)',
    'ΔImpr',
    'CTR(FirstSeg)',
    'CTR(LastSeg)',
    'ΔCTR',
    'Pos(FirstSeg)',
    'Pos(LastSeg)',
    'ΔPos',
    'Judgement',
  ]);

  const urls = Object.keys(updateEventsByUrl).sort();
  for (const urlKey of urls) {
    const events = updateEventsByUrl[urlKey];
    if (!events || events.length === 0) continue;

    // 日別に存在する domain|||url のうち、URL一致のドメインを拾う
    const domainsForUrl = [];
    for (const key of dailyPairs) {
      const [dom, u] = key.split('|||');
      if (u === urlKey) domainsForUrl.push(dom);
    }
    if (domainsForUrl.length === 0) continue;

    for (const domainKey of domainsForUrl) {
      const maxDaily = UIR_maxDailyDateFor_(dailyRows, domainKey, urlKey);
      if (!maxDaily) continue;

      for (let i = 0; i < events.length; i++) {
        const ev = events[i];
        const nextEv = events[i + 1] || null;

        const updateAt = ev.updateAt;
        const nextUpdateAt = nextEv ? nextEv.updateAt : null;

        const w = UIR_computeWindows_(updateAt, nextUpdateAt, maxDaily);
        if (!w) continue;

        const intervalStart = w.intervalStart;
        const intervalEnd   = w.intervalEnd;
        const intervalDays  = w.intervalDays;

        const firstStart = w.firstStart;
        const firstEnd   = w.firstEnd;

        const lastStart  = w.lastStart;
        const lastEnd    = w.lastEnd;

        const firstAgg = UIR_aggDaily_(dailyIndex, domainKey, urlKey, firstStart, firstEnd);
        const lastAgg  = UIR_aggDaily_(dailyIndex, domainKey, urlKey, lastStart, lastEnd);

        // ====== ここが今回の修正ポイント ======
        // GSCは imp=0 の日は行が出ないので、FoundDays=0 は「その窓は実質 imp=0」とみなす。
        // ただし「片側だけ0」の場合は、その側を0として比較しつつ、Judgementに印を付ける。

        const missingFirst = (firstAgg.foundDays === 0);
        const missingLast  = (lastAgg.foundDays === 0);

        const dClicks = lastAgg.clicks - firstAgg.clicks;
        const dImpr   = lastAgg.impr   - firstAgg.impr;
        const dCtr    = lastAgg.ctr    - firstAgg.ctr;

        let dPos = '';
        if (isFinite(firstAgg.pos) && isFinite(lastAgg.pos)) {
          dPos = (lastAgg.pos - firstAgg.pos);
        }

        let judgement = '';
        if (missingFirst && missingLast) {
          // 両窓とも行がない＝露出ゼロ（ほぼ）
          judgement = 'ZeroImpr';
        } else {
          const base = UIR_judge_(dClicks, dImpr, dCtr, dPos);

          if (missingFirst && !missingLast) judgement = base + '_AssumedZeroFirst';
          else if (!missingFirst && missingLast) judgement = base + '_AssumedZeroLast';
          else judgement = base;
        }

        report.push([
          domainKey,
          urlKey,
          UIR_formatDateTime_(updateAt),
          nextUpdateAt ? UIR_formatDateTime_(nextUpdateAt) : '',
          UIR_formatYmd_(intervalStart),
          UIR_formatYmd_(intervalEnd),
          intervalDays,
          ev.titleChanged ? '1' : '0',
          ev.metaChanged ? '1' : '0',
          ev.diff.addedCount,
          ev.diff.removedCount,
          ev.diff.addedSample,
          ev.diff.removedSample,
          firstAgg.foundDays,
          lastAgg.foundDays,
          firstAgg.clicks,
          lastAgg.clicks,
          dClicks,
          firstAgg.impr,
          lastAgg.impr,
          dImpr,
          UIR_round_(firstAgg.ctr, 6),
          UIR_round_(lastAgg.ctr, 6),
          UIR_round_(dCtr, 6),
          (isFinite(firstAgg.pos) ? UIR_round_(firstAgg.pos, 3) : ''),
          (isFinite(lastAgg.pos)  ? UIR_round_(lastAgg.pos, 3)  : ''),
          (dPos === '' ? '' : UIR_round_(dPos, 3)),
          judgement,
        ]);
      }
    }
  }

  UIR_writeSheet_(gscSS, REPORT_SHEET_NAME, report);
}

/***********************
 * 更新イベント（本文ハッシュ変化）: CacheDiffLog から生成
 ***********************/
function UIR_buildUpdateEventsByUrl_(rows) {
  const byUrl = {};

  for (const r of rows) {
    const rawUrl = (r[UIR_CACHE_COL.url - 1] || '').toString();
    const urlKey = UIR_normalizeUrl_(rawUrl);
    if (!urlKey) continue;

    const fetchedAt = UIR_parseDate_(r[UIR_CACHE_COL.fetchedAt - 1]);
    if (!fetchedAt) continue;

    const title = (r[UIR_CACHE_COL.title - 1] || '').toString();
    const meta  = (r[UIR_CACHE_COL.meta - 1]  || '').toString();
    const body  = UIR_normalizeBody_(r[UIR_CACHE_COL.body - 1]);

    if (!byUrl[urlKey]) byUrl[urlKey] = [];
    byUrl[urlKey].push({ fetchedAt, title, meta, body });
  }

  const eventsByUrl = {};

  for (const urlKey of Object.keys(byUrl)) {
    const snaps = byUrl[urlKey].sort((a, b) => a.fetchedAt - b.fetchedAt);
    if (snaps.length < 2) continue;

    let prev = snaps[0];
    let prevHash = UIR_hashMd5Hex_(prev.body);

    const events = [];

    for (let i = 1; i < snaps.length; i++) {
      const cur = snaps[i];
      const curHash = UIR_hashMd5Hex_(cur.body);

      // body が同一なら更新イベントにしない
      if (prevHash === curHash) {
        prev = cur;
        continue;
      }

      const titleChanged = (prev.title || '') !== (cur.title || '');
      const metaChanged  = (prev.meta  || '') !== (cur.meta  || '');

      const diff = UIR_diffLinesSummary_(prev.body, cur.body);

      events.push({
        updateAt: cur.fetchedAt,
        titleChanged,
        metaChanged,
        diff,
      });

      prev = cur;
      prevHash = curHash;
    }

    if (events.length) eventsByUrl[urlKey] = events;
  }

  return eventsByUrl;
}

/***********************
 * GSC日別 index（正規化込み）
 ***********************/
function UIR_buildDailyIndex_(rows) {
  const idx = {};
  for (const r of rows) {
    const domainKey = UIR_normalizeDomain_(r[UIR_DAILY_COL.domain - 1]);
    const urlKey    = UIR_normalizeUrl_(r[UIR_DAILY_COL.url - 1]);
    const dateKey   = UIR_toDateKey_(r[UIR_DAILY_COL.dataDate - 1]);
    if (!domainKey || !urlKey || !dateKey) continue;

    idx[`${domainKey}|||${urlKey}|||${dateKey}`] = {
      clicks: UIR_num_(r[UIR_DAILY_COL.clicks - 1]),
      impr:   UIR_num_(r[UIR_DAILY_COL.impr - 1]),
      pos:    UIR_toNumOrNaN_(r[UIR_DAILY_COL.pos - 1]),
    };
  }
  return idx;
}

function UIR_buildDailyPairs_(rows) {
  const s = new Set();
  for (const r of rows) {
    const domainKey = UIR_normalizeDomain_(r[UIR_DAILY_COL.domain - 1]);
    const urlKey    = UIR_normalizeUrl_(r[UIR_DAILY_COL.url - 1]);
    if (domainKey && urlKey) s.add(`${domainKey}|||${urlKey}`);
  }
  return s;
}

/***********************
 * 区間/窓の決定（長区間=端削って前後比較／短区間=固定窓比較）
 ***********************/
function UIR_computeWindows_(updateAt, nextUpdateAt, maxDaily) {
  const updDay  = UIR_floorDate_(updateAt);
  const nextDay = nextUpdateAt ? UIR_floorDate_(nextUpdateAt) : null;

  // raw区間：更新日を除外して翌日から／次更新日は除外して前日まで
  const rawStart = UIR_addDays_(updDay, UIR_SHORT_SKIP_UPDATE_DAY);
  const rawEnd   = nextDay ? UIR_addDays_(nextDay, -1) : maxDaily;

  if (!rawEnd || rawEnd < rawStart) return null;

  const rawDays = UIR_dayDiffInclusive_(rawStart, rawEnd);

  // 十分長い区間 → 端を削って前半/後半
  if (rawDays >= UIR_MIN_INTERVAL_DAYS) {
    const intervalStart = UIR_addDays_(updDay, UIR_LAG_DAYS);
    const intervalEnd   = UIR_addDays_(rawEnd, -UIR_LAG_DAYS);

    if (intervalEnd < intervalStart) {
      return UIR_computeWindowsShort_(rawStart, rawEnd);
    }

    const intervalDays = UIR_dayDiffInclusive_(intervalStart, intervalEnd);
    if (intervalDays < 1) return null;

    let segDays = Math.min(UIR_SEG_DAYS_DEFAULT, Math.floor(intervalDays / 2));
    segDays = Math.max(1, segDays);
    segDays = Math.min(segDays, intervalDays);

    const firstStart = intervalStart;
    const firstEnd   = UIR_addDays_(firstStart, segDays - 1);

    const lastEnd    = intervalEnd;
    const lastStart  = UIR_addDays_(lastEnd, -(segDays - 1));

    return { intervalStart, intervalEnd, intervalDays, firstStart, firstEnd, lastStart, lastEnd };
  }

  // 短い区間 → 固定窓で「区間前半 vs 区間後半」
  return UIR_computeWindowsShort_(rawStart, rawEnd);
}

function UIR_computeWindowsShort_(rawStart, rawEnd) {
  const rawDays = UIR_dayDiffInclusive_(rawStart, rawEnd);
  if (rawDays < 1) return null;

  let w = Math.min(UIR_SHORT_WINDOW_DAYS, rawDays);
  if (rawDays >= 2 && rawDays < 2 * w) w = Math.max(1, Math.floor(rawDays / 2));

  const intervalStart = rawStart;
  const intervalEnd   = rawEnd;
  const intervalDays  = rawDays;

  const firstStart = intervalStart;
  const firstEnd   = UIR_addDays_(firstStart, w - 1);

  const lastEnd    = intervalEnd;
  const lastStart  = UIR_addDays_(lastEnd, -(w - 1));

  return { intervalStart, intervalEnd, intervalDays, firstStart, firstEnd, lastStart, lastEnd };
}

/***********************
 * maxDaily（URLごとの日別の最新日）
 ***********************/
function UIR_maxDailyDateFor_(rows, domainKey, urlKey) {
  let max = null;
  for (const r of rows) {
    const d = UIR_normalizeDomain_(r[UIR_DAILY_COL.domain - 1]);
    const u = UIR_normalizeUrl_(r[UIR_DAILY_COL.url - 1]);
    if (d !== domainKey || u !== urlKey) continue;

    const dt = UIR_parseDate_(r[UIR_DAILY_COL.dataDate - 1]);
    if (!dt) continue;
    if (!max || dt > max) max = dt;
  }
  return max ? UIR_floorDate_(max) : null;
}

/***********************
 * 日別集計
 ***********************/
function UIR_aggDaily_(idx, domainKey, urlKey, startDate, endDate) {
  let clicks = 0, impr = 0;
  let posWeighted = 0, posWeight = 0;
  let foundDays = 0;

  const cur = new Date(startDate.getTime());
  while (cur <= endDate) {
    const ymd = UIR_formatYmd_(cur);
    const v = idx[`${domainKey}|||${urlKey}|||${ymd}`];
    if (v) {
      foundDays++;
      clicks += v.clicks;
      impr   += v.impr;
      if (isFinite(v.pos) && v.impr > 0) {
        posWeighted += v.pos * v.impr;
        posWeight   += v.impr;
      }
    }
    cur.setDate(cur.getDate() + 1);
  }

  const ctr = impr > 0 ? clicks / impr : 0;
  const pos = posWeight > 0 ? (posWeighted / posWeight) : NaN;

  return { clicks, impr, ctr, pos, foundDays };
}

function UIR_judge_(dClicks, dImpr, dCtr, dPos) {
  const primary = (dClicks > 0) || (dImpr > 0);
  const aux = (dCtr > 0) || (dPos !== '' && isFinite(dPos) ? (dPos < 0) : false);
  if (primary && aux) return 'Improved';
  if (primary) return 'MostlyImproved';
  if (aux) return 'SlightlyImproved';
  return 'NotImproved';
}

/***********************
 * 簡易DIFF（タグ境界で改行 → 行差分）
 ***********************/
function UIR_diffLinesSummary_(oldBody, newBody) {
  const oldText = UIR_prepareForLineDiff_(oldBody);
  const newText = UIR_prepareForLineDiff_(newBody);

  const oldLines = oldText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const newLines = newText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

  const oldSet = new Set(oldLines);
  const newSet = new Set(newLines);

  const added = newLines.filter(l => !oldSet.has(l));
  const removed = oldLines.filter(l => !newSet.has(l));

  const addedSample = UIR_clip_(added.slice(0, 8).join('\n'), UIR_SAMPLE_CLIP);
  const removedSample = UIR_clip_(removed.slice(0, 8).join('\n'), UIR_SAMPLE_CLIP);

  return {
    addedCount: added.length,
    removedCount: removed.length,
    addedSample,
    removedSample,
  };
}

function UIR_prepareForLineDiff_(html) {
  const s = (html || '').toString();
  return s.replace(/></g, '>\n<');
}

function UIR_normalizeBody_(v) {
  return (v || '')
    .toString()
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function UIR_hashMd5Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(text || ''), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/***********************
 * 正規化
 ***********************/
function UIR_normalizeDomain_(d) {
  if (d == null) return '';
  return String(d).trim().toLowerCase().replace(/^sc-domain:/, '');
}

function UIR_normalizeUrl_(u) {
  if (u == null) return '';
  let s = String(u).trim();
  if (!s) return '';

  s = UIR_safeDecode_(s);
  s = s.replace(/^http:\/\//i, 'https://');
  s = s.replace(/#.*$/, '');
  s = s.replace(/\?.*$/, '');
  if (s.length > 8 && s.endsWith('/')) s = s.slice(0, -1);

  return s;
}

function UIR_safeDecode_(s) {
  try { return decodeURIComponent(s); } catch (e) { return s; }
}

function UIR_toDateKey_(v) {
  const d = UIR_parseDate_(v);
  if (!d) return (v == null ? '' : String(v).trim());
  return UIR_formatYmd_(d);
}

/***********************
 * Sheet I/O（必要列だけ読む）
 ***********************/
function UIR_getSheetByGid_(ss, gid, fallbackName) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (sh.getSheetId && sh.getSheetId() === gid) return sh;
  }
  if (fallbackName) return ss.getSheetByName(fallbackName);
  return null;
}

function UIR_readSheetRows_(sheet, colCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const cols = Math.min(colCount || sheet.getLastColumn(), sheet.getLastColumn());
  return sheet.getRange(2, 1, lastRow - 1, cols).getValues();
}

function UIR_writeSheet_(ss, name, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  //sh.autoResizeColumns(1, rows[0].length);
}

/***********************
 * util
 ***********************/
function UIR_parseDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  const s = (v == null ? '' : String(v)).trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) {
    const yy = Number(m[1]), mo = Number(m[2]) - 1, dd = Number(m[3]);
    const hh = Number(m[4] || 0), mi = Number(m[5] || 0), ss = Number(m[6] || 0);
    return new Date(yy, mo, dd, hh, mi, ss);
  }

  m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) {
    const yy = Number(m[1]), mo = Number(m[2]) - 1, dd = Number(m[3]);
    const hh = Number(m[4] || 0), mi = Number(m[5] || 0), ss = Number(m[6] || 0);
    return new Date(yy, mo, dd, hh, mi, ss);
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function UIR_floorDate_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function UIR_addDays_(d, days) { const x = new Date(d.getTime()); x.setDate(x.getDate() + days); return x; }

function UIR_dayDiffInclusive_(a, b) {
  const ms = 86400 * 1000;
  const diff = Math.floor((UIR_floorDate_(b) - UIR_floorDate_(a)) / ms);
  return diff + 1;
}

function UIR_formatYmd_(d) {
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return `${y}-${m}-${day}`;
}

function UIR_formatDateTime_(d) {
  const ymd = UIR_formatYmd_(d);
  const hh = ('0' + d.getHours()).slice(-2);
  const mm = ('0' + d.getMinutes()).slice(-2);
  const ss = ('0' + d.getSeconds()).slice(-2);
  return `${ymd} ${hh}:${mm}:${ss}`;
}

function UIR_num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function UIR_toNumOrNaN_(v) { const n = Number(v); return isFinite(n) ? n : NaN; }

function UIR_round_(v, digits) {
  if (!isFinite(v)) return '';
  const p = Math.pow(10, digits);
  return Math.round(v * p) / p;
}

function UIR_clip_(s, n) {
  s = (s == null ? '' : String(s));
  return s.length <= n ? s : s.slice(0, n) + '...';
}
