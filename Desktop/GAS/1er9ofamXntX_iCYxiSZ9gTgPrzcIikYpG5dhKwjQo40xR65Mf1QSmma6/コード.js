  /************************************************************
 * GSCデータを“別スプレッドシート”に書き込む最適化版（拡張）
 *
 * 【方針】
 * - 既存列は一切変更しない
 * - 右側に impressions / ctr / position を追加するだけ
 * - ドメインごとにページ次元で一括取得
 ************************************************************/

const spreadsheetId   = '1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs';
const GSC_SPREADSHEET_ID = '1ZHMy3vVUd_0MrSduXTe1etBb1HNpAlGwQr006ljwMiM';

const GSC_LATEST_SHEET_NAME  = 'GSC最新';
const GSC_HISTORY_SHEET_NAME = 'GSC履歴';

const GSC_DATA_LAG_DAYS = 3;

/*******************************************************************
 * ENTRY POINT
 *******************************************************************/
function updateGSC_AllDomains() {
  const mainSS = SpreadsheetApp.openById(spreadsheetId);
  const gscSS  = SpreadsheetApp.openById(GSC_SPREADSHEET_ID);

  const today = new Date();
  const latestDataDate = new Date(today.getTime() - GSC_DATA_LAG_DAYS * 86400 * 1000);
  const ranges = buildGscDateRanges(latestDataDate);

  let allResults = [];

  mainSS.getSheets().forEach(sheet => {
    const name = sheet.getName();

    // 有効ドメイン以外をスキップ
    if (name === '設定') return;
    if (!name.includes('.')) return;

    const propertyUrl = `sc-domain:${name}`;

    // ページ次元で一括取得
    const gsc30   = fetchAllPagesMetrics(propertyUrl, ranges.cur30);
    const gsc90   = fetchAllPagesMetrics(propertyUrl, ranges.cur90);
    const gprev30 = fetchAllPagesMetrics(propertyUrl, ranges.prev30);
    const gprev90 = fetchAllPagesMetrics(propertyUrl, ranges.prev90);

    const urls = getUrlsFromSheet(sheet);

    urls.forEach(url => {
      const a30 = gsc30[url]   || {};
      const a90 = gsc90[url]   || {};
      const b30 = gprev30[url] || {};
      const b90 = gprev90[url] || {};

      allResults.push({
        url,

        // ===== 既存列（変更しない） =====
        click30 : a30.clicks || 0,
        click90 : a90.clicks || 0,
        prev30  : b30.clicks || 0,
        prev90  : b90.clicks || 0,

        // ===== 追加列 =====
        imp30 : a30.impressions || 0,
        ctr30 : a30.ctr || 0,
        pos30 : a30.position ?? null,

        imp90 : a90.impressions || 0,
        ctr90 : a90.ctr || 0,
        pos90 : a90.position ?? null,

        pimp30 : b30.impressions || 0,
        pctr30 : b30.ctr || 0,
        ppos30 : b30.position ?? null,

        pimp90 : b90.impressions || 0,
        pctr90 : b90.ctr || 0,
        ppos90 : b90.position ?? null
      });
    });
  });

  writeToGscLatest(gscSS, allResults, latestDataDate);
  writeToGscHistory(gscSS, allResults, latestDataDate);
}

/*******************************************************************
 * 本体のURL一覧（サイトマップ由来）
 *******************************************************************/
function getUrlsFromSheet(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2,1,last-1,1)
    .getValues()
    .map(r => r[0])
    .filter(Boolean);
}

/*******************************************************************
 * 日付レンジ作成
 *******************************************************************/
function buildGscDateRanges(latestDate) {
  const DAY = 86400 * 1000;

  const cur30 = {
    start: new Date(latestDate.getTime() - 29*DAY),
    end: latestDate
  };
  const cur90 = {
    start: new Date(latestDate.getTime() - 89*DAY),
    end: latestDate
  };

  const prev30 = {
    start: new Date(cur30.start),
    end: new Date(cur30.end)
  };
  prev30.start.setFullYear(prev30.start.getFullYear() - 1);
  prev30.end.setFullYear(prev30.end.getFullYear() - 1);

  const prev90 = {
    start: new Date(cur90.start),
    end: new Date(cur90.end)
  };
  prev90.start.setFullYear(prev90.start.getFullYear() - 1);
  prev90.end.setFullYear(prev90.end.getFullYear() - 1);

  return { cur30, cur90, prev30, prev90 };
}

/*******************************************************************
 * GSCページ次元データ取得（clicks + imp + ctr + pos）
 *******************************************************************/
function fetchAllPagesMetrics(siteUrl, range) {
  const token = ScriptApp.getOAuthToken();

  const url =
    'https://searchconsole.googleapis.com/webmasters/v3/sites/' +
    encodeURIComponent(siteUrl) +
    '/searchAnalytics/query';

  const payload = {
    startDate: formatDate(range.start),
    endDate:   formatDate(range.end),
    dimensions: ['page'],
    rowLimit: 25000
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('GSC error for ' + siteUrl + ' : ' + response.getContentText());
    return {};
  }

  const json = JSON.parse(response.getContentText());
  const map = {};

  if (!json.rows) return map;

  json.rows.forEach(r => {
    const pageUrl = safeDecode_(r.keys[0]);

    map[pageUrl] = {
      clicks: r.clicks || 0,
      impressions: r.impressions || 0,
      ctr: r.ctr || 0,
      position: r.position ?? null
    };
  });

  return map;
}

function formatDate(dt) {
  return Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/*******************************************************************
 * GSC最新シート（全上書き）
 * ※ 既存列を維持し、右に追加
 *******************************************************************/
function writeToGscLatest(gscSS, rows, latestDate) {
  let sheet = gscSS.getSheetByName(GSC_LATEST_SHEET_NAME);
  if (!sheet) sheet = gscSS.insertSheet(GSC_LATEST_SHEET_NAME);

  sheet.clearContents();

  const header = [
    'URL',
    '直近30日','直近90日','前年同期30日','前年同期90日','データ基準日',

    // 追加
    '30d_imp','30d_ctr','30d_pos',
    '90d_imp','90d_ctr','90d_pos',
    'py30_imp','py30_ctr','py30_pos',
    'py90_imp','py90_ctr','py90_pos'
  ];
  sheet.getRange(1,1,1,header.length).setValues([header]);

  const dateStr = formatDate(latestDate);

  const values = rows.map(r => [
    r.url,
    r.click30, r.click90, r.prev30, r.prev90, dateStr,

    r.imp30, r.ctr30, r.pos30,
    r.imp90, r.ctr90, r.pos90,
    r.pimp30, r.pctr30, r.ppos30,
    r.pimp90, r.pctr90, r.ppos90
  ]);

  if (values.length) {
    sheet.getRange(2,1,values.length,values[0].length).setValues(values);
  }
}

/*******************************************************************
 * GSC履歴シート（append）
 * ※ 既存列を維持し、右に追加
 *******************************************************************/
function writeToGscHistory(gscSS, rows, latestDate) {
  let sheet = gscSS.getSheetByName(GSC_HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = gscSS.insertSheet(GSC_HISTORY_SHEET_NAME);
    sheet.appendRow([
      '取得日時','データ基準日','URL',
      '直近30','直近90','前年同期30','前年同期90',

      '30d_imp','30d_ctr','30d_pos',
      '90d_imp','90d_ctr','90d_pos'
    ]);
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const dateStr = formatDate(latestDate);

  const values = rows.map(r => [
    now, dateStr, r.url,
    r.click30, r.click90, r.prev30, r.prev90,

    r.imp30, r.ctr30, r.pos30,
    r.imp90, r.ctr90, r.pos90,
    r.pimp30, r.pctr30, r.ppos30,
    r.pimp90, r.pctr90, r.ppos90
  ]);

  sheet.getRange(sheet.getLastRow()+1,1,values.length,values[0].length)
       .setValues(values);
}


/************************************************************
 * GSC日別 backfill（date x page）
 *
 * 出力シート: GSC日別
 * 列: 取得日時 / データ基準日 / ドメイン / URL / clicks / impressions / ctr / position
 *
 * - dimensions: ['date','page']
 * - rowLimit: 25000 + startRow でページング（公式仕様）:contentReference[oaicite:1]{index=1}
 * - まず指定期間の既存日別データを削除してから再投入（重複防止）
 ************************************************************/

const GSC_DAILY_SHEET_NAME = 'GSC日別';

// backfill期間（とりあえず「2025-12-01〜最新(ラグ考慮)」で作る）
function backfillGSC_Daily_AllDomains() {
  const mainSS = SpreadsheetApp.openById(spreadsheetId);
  const gscSS  = SpreadsheetApp.openById(GSC_SPREADSHEET_ID);

  const today = new Date();
  const latestDataDate = new Date(today.getTime() - GSC_DATA_LAG_DAYS * 86400 * 1000);

  // ★必要ならここだけ変えてください
  const start = new Date(2025, 11, 1);            // 2025-12-01
  const end   = floorDate_(latestDataDate);       // ラグ考慮した最新日

  const chunkDays = 14; // 14日刻み（安全側）

  backfillGSC_Daily_(mainSS, gscSS, start, end, chunkDays);
}

function backfillGSC_Daily_(mainSS, gscSS, startDate, endDate, chunkDays) {
  const dailySheet = getOrCreateDailySheet_(gscSS);

  // 指定期間の既存データを削除（重複防止）
  removeDailyRowsInRange_(dailySheet, startDate, endDate);

  const startStr = formatDate(startDate);
  const endStr   = formatDate(endDate);
  Logger.log(`Backfill start: ${startStr} -> ${endStr} (chunk=${chunkDays}d)`);

  const runAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  mainSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === '設定') return;
    if (!name.includes('.')) return; // ドメイン以外スキップ

    const domain = name;
    const propertyUrl = `sc-domain:${domain}`;

    // URLリスト（サイトマップ由来）に限定したい場合はこれを使う
    const urlSet = new Set(getUrlsFromSheet(sheet).map(u => u.toString().trim()));

    Logger.log(`Domain: ${domain} urls=${urlSet.size}`);

    let cur = floorDate_(startDate);
    while (cur <= endDate) {
      const chunkStart = new Date(cur.getTime());
      const chunkEnd = minDate_(endDate, addDays_(chunkStart, chunkDays - 1));

      const cs = formatDate(chunkStart);
      const ce = formatDate(chunkEnd);
      Logger.log(`  Fetch: ${cs} -> ${ce}`);

      // date×page をページングしながら全部取る
      const rows = fetchDatePageMetricsAll_(propertyUrl, chunkStart, chunkEnd);

      // 取得したものをシートに追記（URLリストにあるものだけ保存）
      const values = [];
      for (const r of rows) {
        const pageUrl = r.page;
        if (urlSet.size && !urlSet.has(pageUrl)) continue; // 限定しないならこの行をコメントアウト

        values.push([
          runAt,          // 取得日時
          r.date,         // データ基準日（YYYY-MM-DD）
          domain,         // ドメイン
          pageUrl,        // URL
          r.clicks,
          r.impressions,
          r.ctr,
          r.position
        ]);
      }

      appendValuesChunked_(dailySheet, values, 5000);

      cur = addDays_(chunkEnd, 1);
    }
  });

  Logger.log('Backfill done.');
}

/***********************
 * Search Analytics API
 ***********************/
function fetchDatePageMetricsAll_(siteUrl, startDate, endDate) {
  const token = ScriptApp.getOAuthToken();
  const api =
    'https://searchconsole.googleapis.com/webmasters/v3/sites/' +
    encodeURIComponent(siteUrl) +
    '/searchAnalytics/query';

  const out = [];
  const rowLimit = 25000; // 最大25,000 :contentReference[oaicite:2]{index=2}
  let startRow = 0;

  while (true) {
    const payload = {
      startDate: formatDate(startDate),
      endDate:   formatDate(endDate),
      dimensions: ['date', 'page'],
      rowLimit: rowLimit,
      startRow: startRow
    };

    const res = UrlFetchApp.fetch(api, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`GSC error ${res.getResponseCode()} for ${siteUrl}: ${res.getContentText()}`);
      break;
    }

    const json = JSON.parse(res.getContentText());
    const rows = json.rows || [];
    if (!rows.length) break;

    for (const r of rows) {
      const keys = r.keys || [];
      const rawDate = keys[0];
      const rawPage = keys[1];

      const dateStr = normalizeGscDateKey_(rawDate);
      const pageUrl = safeDecode_(rawPage);

      out.push({
        date: dateStr,
        page: pageUrl,
        clicks: r.clicks || 0,
        impressions: r.impressions || 0,
        ctr: r.ctr || 0,
        position: (r.position ?? '')
      });
    }

    // ページング（startRow） :contentReference[oaicite:3]{index=3}
    if (rows.length < rowLimit) break;
    startRow += rowLimit;
  }

  return out;
}

function normalizeGscDateKey_(k) {
  // 返却が YYYY-MM-DD の想定だが、念のため YYYYMMDD も吸収
  const s = (k || '').toString().trim();
  if (/^\d{8}$/.test(s)) {
    return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
  }
  return s;
}

function safeDecode_(s) {
  const str = (s || '').toString();
  try {
    return decodeURIComponent(str);
  } catch (e) {
    return str;
  }
}

/***********************
 * Sheet I/O
 ***********************/
function getOrCreateDailySheet_(gscSS) {
  let sh = gscSS.getSheetByName(GSC_DAILY_SHEET_NAME);
  if (!sh) {
    sh = gscSS.insertSheet(GSC_DAILY_SHEET_NAME);
    sh.appendRow(['取得日時','データ基準日','ドメイン','URL','clicks','impressions','ctr','position']);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(['取得日時','データ基準日','ドメイン','URL','clicks','impressions','ctr','position']);
  }
  return sh;
}

function removeDailyRowsInRange_(sheet, startDate, endDate) {
  const last = sheet.getLastRow();
  if (last < 2) return;

  const startStr = formatDate(startDate);
  const endStr   = formatDate(endDate);

  // 全行読み→フィルタ→全書き戻し（deleteRow連発より速い）
  const data = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  const header = data[0];
  const kept = [header];

  for (let i = 1; i < data.length; i++) {
    const d = (data[i][1] || '').toString(); // B列: データ基準日
    // 範囲内は捨てる
    if (d && d >= startStr && d <= endStr) continue;
    kept.push(data[i]);
  }

  sheet.clearContents();
  sheet.getRange(1,1,kept.length,kept[0].length).setValues(kept);
}

function appendValuesChunked_(sheet, values, chunkSize) {
  if (!values || !values.length) return;
  const cols = values[0].length;
  for (let i = 0; i < values.length; i += chunkSize) {
    const chunk = values.slice(i, i + chunkSize);
    sheet.getRange(sheet.getLastRow() + 1, 1, chunk.length, cols).setValues(chunk);
  }
}

/***********************
 * Date utils
 ***********************/
function floorDate_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function addDays_(d, days) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + days);
  return x;
}
function minDate_(a, b) {
  return (a.getTime() <= b.getTime()) ? a : b;
}

