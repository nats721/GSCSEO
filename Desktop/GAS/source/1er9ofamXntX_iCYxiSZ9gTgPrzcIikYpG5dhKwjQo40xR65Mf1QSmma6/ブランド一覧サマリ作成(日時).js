/*************************************************************
 * 概観ダッシュボード生成（ブランド横断）＋履歴保存（日次運用向け）
 * - スタンドアロンGASで実行（スプレッドシートに付属しない）
 * - 対象スプレッドシートIDを固定参照
 *
 * 生成シート：
 *   1) 概観
 *   2) 概観_未分類
 *   3) 概観_履歴（同日分は上書き）
 *************************************************************/

// ===== 対象スプレッドシート（固定）=====
const TARGET_SPREADSHEET_ID = '1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs';

// ===== シート名 =====
const OVERVIEW_SHEET         = '概観';
const OVERVIEW_UNCLASS_SHEET = '概観_未分類';
const HISTORY_SHEET          = '概観_履歴';
const SETTINGS_SHEET         = '設定';

// ===== 分類定義 =====
const EXCLUDE_CLASSES = new Set(['NW', 'ST']);
const KNOWN_CLASSES   = new Set(['TP', 'SP', 'MC', 'SC', 'NW', 'ST']);
const DASH_CLASSES    = ['TP', 'SP', 'MC', 'SC', '未分類'];

// ===== 閾値 =====
const MIN_IMP_FOR_CAND  = 30;
const LOW_CTR_THRESHOLD = 0.005; // 0.5%

// ===== 出力：共通列定義（DRY）=====
const COLS_COMMON = [
  { key: 'pages',        label: 'ページ'   , type: 'int'     },
  { key: 'clicks',       label: 'クリック', type: 'int'     },
  { key: 'imps',         label: '表示',     type: 'int'     },
  { key: 'ctr',          label: 'CTR',      type: 'pct'     },
  { key: 'posWavg',      label: '平均順位', type: 'num2'    },
  { key: 'deltaClicks',  label: 'Δクリック', type: 'pct'    },
  { key: 'deltaImps',    label: 'Δ表示',     type: 'pct'    },
  { key: 'zeroClickRate',label: '0クリック率', type: 'pct'  },
  { key: 'lowCtrRate',   label: '低CTR率',   type: 'pct'    },
  { key: 'recoveryRate', label: '持ち直し率', type: 'pct'  },
  { key: 'worseningRate',label: '悪化率',    type: 'pct'    },
];

// 分類テーブルも同じ列構造で出す（表示したい列だけ切り出し）
const COLS_CLASS = [
  { key: 'pages',        label: 'ページ',    type: 'int' },
  { key: 'clicks',       label: 'クリック',  type: 'int' },
  { key: 'imps',         label: '表示',      type: 'int' },
  { key: 'ctr',          label: 'CTR',       type: 'pct' },
  { key: 'posWavg',      label: '平均順位',  type: 'num2'},
  { key: 'deltaClicks',  label: 'Δクリック', type: 'pct' },
  { key: 'deltaImps',    label: 'Δ表示',     type: 'pct' },
  { key: 'zeroClickRate',label: '0クリック率', type: 'pct' },
  { key: 'lowCtrRate',   label: '低CTR率',   type: 'pct' },
  { key: 'recoveryRate', label: '持ち直し率', type: 'pct' },
  { key: 'worseningRate',label: '悪化率',    type: 'pct' },
];

/**
 * 手動実行/トリガー実行の入口
 */
function runOverviewDashboard() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const domains = getDomainsFromSettings_(ss);

  const summary = {
    all: makeAgg_(),
    seo: makeAgg_(),
    byDomainAll: {},
    byDomainSeo: {},
    badClassRows: [],
  };

  for (const domain of domains) {
    const sh = ss.getSheetByName(domain);
    if (!sh) continue;

    const values = sh.getDataRange().getValues();
    if (values.length < 2) continue;

    const header = values[0].map(h => String(h || '').trim());
    const idx = indexMap_OverView_(header);




    const domAll = makeAgg_();
    const domSeo = makeAgg_();

    for (let r = 1; r < values.length; r++) {
      const row = values[r];

      const url = getCell_(row, idx, 'URL');
      if (!url) continue;

      const rawClass = getCell_(row, idx, '分類');
      const cls = normalizeClass_(rawClass);
      const clsKey = toDashClassKey_(cls);

      const title = getCell_(row, idx, 'タイトル');

      const clicks     = toNumber_(getCell_(row, idx, 'CLI_1M')) || 0;
      const imps       = toNumber_(getCell_(row, idx, 'IMP_1M')) || 0;
      const lyClicks   = toNumber_(getCell_(row, idx, 'LY1_CLI')) || 0;
      const lyImps     = toNumber_(getCell_(row, idx, 'LY1_IMP')) || 0;
      const pos        = toNumber_(getCell_(row, idx, 'POS_1M'));

      const threeClicks = toNumber_(getCell_(row, idx, '3M_CLI')) || 0;
      const ly3Clicks   = toNumber_(getCell_(row, idx, 'LY3_CLI')) || 0;

      const delta1 = computeRatio_(clicks, lyClicks);
      const delta3 = computeRatio_(threeClicks, ly3Clicks);

      const ctr = (imps > 0) ? (clicks / imps) : null;

      // 候補判定
      const flags = {
        isCandZeroClick: (imps >= MIN_IMP_FOR_CAND && clicks === 0),
        isCandLowCtr:    (imps >= MIN_IMP_FOR_CAND && ctr !== null && ctr <= LOW_CTR_THRESHOLD),
      };

      if (cls === '未分類') {
        summary.badClassRows.push([domain, url, rawClass, title, imps, clicks]);
      }

      // 全体（除外含む）
      addRowToAgg_(summary.all, clsKey, clicks, imps, lyClicks, lyImps, pos, flags);
      addRowToAgg_(domAll,      clsKey, clicks, imps, lyClicks, lyImps, pos, flags);

      // 持ち直し/悪化（全体）
      const trend = classifyTrend_(clicks, lyClicks, delta1, delta3);
      if (trend) {
        addTrendToAgg_(summary.all, clsKey, trend);
        addTrendToAgg_(domAll,      clsKey, trend);
      }

      // SEO対象のみ（NW/ST除外）
      if (!EXCLUDE_CLASSES.has(cls)) {
        addRowToAgg_(summary.seo, clsKey, clicks, imps, lyClicks, lyImps, pos, flags);
        addRowToAgg_(domSeo,      clsKey, clicks, imps, lyClicks, lyImps, pos, flags);

        const trendSeo = classifyTrend_(clicks, lyClicks, delta1, delta3);
        if (trendSeo) {
          addTrendToAgg_(summary.seo, clsKey, trendSeo);
          addTrendToAgg_(domSeo,      clsKey, trendSeo);
        }
      }
    }

    finalizeAgg_(domAll);
    finalizeAgg_(domSeo);

    summary.byDomainAll[domain] = domAll;
    summary.byDomainSeo[domain] = domSeo;
  }

  finalizeAgg_(summary.all);
  finalizeAgg_(summary.seo);

  writeOverview_(ss, domains, summary);
  writeUnclassified_(ss, summary.badClassRows);
  upsertOverviewHistory_(ss, domains, summary);
}

/* =========================================================
 * 集計（DRY）
 * =======================================================*/

function makeAgg_() {
  const byClass = {};
  DASH_CLASSES.forEach(c => { byClass[c] = makeClassAgg_(); });
  return {
    pages: 0, clicks: 0, imps: 0,
    lyClicks: 0, lyImps: 0,
    posWsum: 0, posW: 0,

    candZeroClickPages: 0,
    candLowCtrPages: 0,

    recoveryPages: 0,
    worseningPages: 0,

    // derived
    ctr: null, posWavg: null,
    deltaClicks: null, deltaImps: null,
    zeroClickRate: null, lowCtrRate: null,
    recoveryRate: null, worseningRate: null,

    byClass,
  };
}

function makeClassAgg_() {
  return {
    pages: 0, clicks: 0, imps: 0,
    lyClicks: 0, lyImps: 0,
    posWsum: 0, posW: 0,

    candZeroClickPages: 0,
    candLowCtrPages: 0,

    recoveryPages: 0,
    worseningPages: 0,

    // derived
    ctr: null, posWavg: null,
    deltaClicks: null, deltaImps: null,
    zeroClickRate: null, lowCtrRate: null,
    recoveryRate: null, worseningRate: null,
  };
}

function addRowToAgg_(agg, clsKey, clicks, imps, lyClicks, lyImps, pos, flags) {
  const c = agg.byClass[clsKey] ? clsKey : '未分類';

  // overall
  agg.pages++;
  agg.clicks += clicks;
  agg.imps += imps;
  agg.lyClicks += lyClicks;
  agg.lyImps += lyImps;

  if (pos !== null && imps > 0) {
    agg.posWsum += pos * imps;
    agg.posW += imps;
  }

  if (flags?.isCandZeroClick) agg.candZeroClickPages++;
  if (flags?.isCandLowCtr)    agg.candLowCtrPages++;

  // byClass
  const ca = agg.byClass[c];
  ca.pages++;
  ca.clicks += clicks;
  ca.imps += imps;
  ca.lyClicks += lyClicks;
  ca.lyImps += lyImps;

  if (pos !== null && imps > 0) {
    ca.posWsum += pos * imps;
    ca.posW += imps;
  }
  if (flags?.isCandZeroClick) ca.candZeroClickPages++;
  if (flags?.isCandLowCtr)    ca.candLowCtrPages++;
}

function addTrendToAgg_(agg, clsKey, trend) {
  const c = agg.byClass[clsKey] ? clsKey : '未分類';
  if (trend === 'recovery') {
    agg.recoveryPages++;
    agg.byClass[c].recoveryPages++;
  } else if (trend === 'worsening') {
    agg.worseningPages++;
    agg.byClass[c].worseningPages++;
  }
}

/**
 * 持ち直し/悪化 判定
 * - 出現（click>0 && LY=0）=> 持ち直し
 * - 消失（click=0 && LY>0）=> 悪化
 * - それ以外：delta1/delta3が両方出るなら加速度で判定
 */
function classifyTrend_(clicks, lyClicks, delta1, delta3) {
  if (clicks > 0 && lyClicks === 0) return 'recovery';
  if (clicks === 0 && lyClicks > 0) return 'worsening';
  if (delta1 !== null && delta3 !== null) {
    if (delta1 > delta3) return 'recovery';
    if (delta1 < delta3) return 'worsening';
  }
  return null;
}

function finalizeAgg_(agg) {
  agg.ctr       = (agg.imps > 0) ? (agg.clicks / agg.imps) : null;
  agg.posWavg   = (agg.posW > 0) ? (agg.posWsum / agg.posW) : null;
  agg.deltaClicks = (agg.lyClicks > 0) ? (agg.clicks / agg.lyClicks) : null;
  agg.deltaImps   = (agg.lyImps > 0) ? (agg.imps / agg.lyImps) : null;

  agg.zeroClickRate = (agg.pages > 0) ? (agg.candZeroClickPages / agg.pages) : null;
  agg.lowCtrRate    = (agg.pages > 0) ? (agg.candLowCtrPages / agg.pages) : null;
  agg.recoveryRate  = (agg.pages > 0) ? (agg.recoveryPages / agg.pages) : null;
  agg.worseningRate = (agg.pages > 0) ? (agg.worseningPages / agg.pages) : null;

  for (const k of Object.keys(agg.byClass)) {
    const ca = agg.byClass[k];
    ca.ctr       = (ca.imps > 0) ? (ca.clicks / ca.imps) : null;
    ca.posWavg   = (ca.posW > 0) ? (ca.posWsum / ca.posW) : null;
    ca.deltaClicks = (ca.lyClicks > 0) ? (ca.clicks / ca.lyClicks) : null;
    ca.deltaImps   = (ca.lyImps > 0) ? (ca.imps / ca.lyImps) : null;

    ca.zeroClickRate = (ca.pages > 0) ? (ca.candZeroClickPages / ca.pages) : null;
    ca.lowCtrRate    = (ca.pages > 0) ? (ca.candLowCtrPages / ca.pages) : null;
    ca.recoveryRate  = (ca.pages > 0) ? (ca.recoveryPages / ca.pages) : null;
    ca.worseningRate = (ca.pages > 0) ? (ca.worseningPages / ca.pages) : null;
  }
}

/* =========================================================
 * 出力（DRY）
 * =======================================================*/

function writeOverview_(ss, domains, summary) {
  const sh = ensureSheet_(ss, OVERVIEW_SHEET);
  sh.getDataRange().clearContent(); // 書式は保持

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

  const out = [];
  const meta = { tables: [] }; // 後で書式を当てるため

  // タイトル
  out.push(['概観ダッシュボード', `更新: ${now}`]);

  // 1) 全体
  pushSection_(
    out, meta,
    '■全体サマリー',
    ['対象'].concat(COLS_COMMON.map(c => c.label)),
    [
      ['全ページ'].concat(renderAggByCols_(summary.all, COLS_COMMON)),
      ['SEO対象'].concat(renderAggByCols_(summary.seo, COLS_COMMON)),
    ]
  );

  // 2) 分類別（SEO対象）
  const classRows = DASH_CLASSES.map(cls => {
    const a = summary.seo.byClass[cls] || makeClassAgg_();
    return [cls].concat(renderAggByCols_(a, COLS_CLASS));
  });

  pushSection_(
    out, meta,
    '■分類別（SEO対象）',
    ['分類'].concat(COLS_CLASS.map(c => c.label)),
    classRows
  );

  // 3) ブランド別（SEO対象）
  const domainRows = [];
  for (const domain of domains) {
    const a = summary.byDomainSeo[domain];
    if (!a) continue;
    domainRows.push([domain].concat(renderAggByCols_(a, COLS_COMMON)));
  }

  pushSection_(
    out, meta,
    '■ブランド別（SEO対象）',
    ['ドメイン'].concat(COLS_COMMON.map(c => c.label)),
    domainRows
  );

  // 4) ブランド×分類（ページ数）
  const matrixTitleRow = out.length + 1;
  out.push(['■ブランド×分類（ページ数 / SEO対象）', '']);
  out.push(['ドメイン'].concat(DASH_CLASSES));

  const matrixStartRow = out.length + 1;
  const matrixBody = [];
  for (const domain of domains) {
    const a = summary.byDomainSeo[domain];
    if (!a) continue;
    matrixBody.push([domain].concat(DASH_CLASSES.map(c => (a.byClass[c]?.pages || 0))));
  }
  matrixBody.forEach(r => out.push(r));

  // setValues 1回
  const padded = pad2d_(out);
  ensureSheetDimensions_(sh, padded.length, padded[0].length);
  sh.getRange(1, 1, padded.length, padded[0].length).setValues(padded);
  sh.setFrozenRows(1);

  // テーブル書式（数値・%）を当てる
  applyTableFormats_(sh, meta.tables);

  // 凡例：ブランド×分類の右側に出す
  const legend = buildLegend_();
  const legendRow = matrixTitleRow ;                // タイトル直下
  const legendCol = 1 + 1 + DASH_CLASSES.length + 1;   // 行列の右に2列空ける
  ensureSheetDimensions_(sh, legendRow + legend.length + 2, legendCol + 2);
  sh.getRange(legendRow, legendCol, legend.length, 1).setValues(legend);
}

function pushSection_(out, meta, title, headerRow, rows) {
  out.push([title, '']);
  const headerIndex = out.length + 1;
  out.push(headerRow);

  const startRow = out.length + 1;
  rows.forEach(r => out.push(r));

  // テーブル範囲記録（書式用）
  meta.tables.push({
    headerRow: headerIndex,
    startRow,
    numRows: rows.length,
    startCol: 1,
    numCols: headerRow.length,
    headerRowValues: headerRow,
  });

  out.push(['', '']);
}

function renderAggByCols_(agg, colsDef) {
  return colsDef.map(c => (agg[c.key] ?? null));
}

function applyTableFormats_(sh, tables) {
  for (const t of tables) {
    // ヘッダから列タイプ判定（DRY）
    const header = t.headerRowValues;
    const types = header.map(h => {
      if (h === 'CTR' || String(h).includes('率') || String(h).startsWith('Δ')) return 'pct';
      if (h === '平均順位') return 'num2';
      if (h === 'ページ' || h === 'クリック' || h === '表示') return 'int';
      return null;
    });

    // データ範囲
    const range = sh.getRange(t.startRow, t.startCol, t.numRows, t.numCols);

    // 列ごとにnumberFormat（必要分だけ）
    for (let i = 0; i < types.length; i++) {
      const type = types[i];
      if (!type) continue;
      const colRange = sh.getRange(t.startRow, t.startCol + i, t.numRows, 1);
      if (type === 'pct')  colRange.setNumberFormat('0.00%');
      if (type === 'num2') colRange.setNumberFormat('0.00');
      if (type === 'int')  colRange.setNumberFormat('#,##0');
    }
  }
}

/* ----------------- 未分類 ----------------- */

function writeUnclassified_(ss, rows) {
  const sh = ensureSheet_(ss, OVERVIEW_UNCLASS_SHEET);
  sh.getDataRange().clearContent(); // 書式保持

  const out = [['ドメイン', 'URL', '分類(元データ)', 'タイトル', 'IMP_1M', 'CLI_1M']];
  if (rows && rows.length) rows.forEach(r => out.push(r));

  const padded = pad2d_(out);
  ensureSheetDimensions_(sh, padded.length, padded[0].length);
  sh.getRange(1, 1, padded.length, padded[0].length).setValues(padded);
  sh.setFrozenRows(1);
}

/* ----------------- 履歴（同日上書き：DRY＆高速） ----------------- */

function upsertOverviewHistory_(ss, domains, summary) {
  const sh = ensureSheet_(ss, HISTORY_SHEET);

  const header = [
    '日付', '粒度', 'ドメイン', '分類',
    'ページ数', 'クリック(1M)', '表示(1M)', 'CTR',
    '平均順位', 'Δクリック(1M)', 'Δ表示(1M)',
    '0クリック率', '低CTR率', '持ち直し率', '悪化率',
    '更新タイムスタンプ',
  ];

  // ヘッダ保証（初回だけ）
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  } else {
    const h0 = sh.getRange(1, 1, 1, header.length).getValues()[0];
    if (String(h0[0] || '') !== '日付') {
      sh.insertRowBefore(1);
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
    }
  }

  const tz = Session.getScriptTimeZone();
  const dayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const ts     = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');

  const newRows = [];
  newRows.push(historyRow_(dayKey, 'ALL', '', '', summary.all, ts));
  newRows.push(historyRow_(dayKey, 'SEO', '', '', summary.seo, ts));
  for (const domain of domains) {
    const a = summary.byDomainSeo[domain];
    if (!a) continue;
    newRows.push(historyRow_(dayKey, 'DOMAIN_SEO', domain, '', a, ts));
  }
  for (const cls of DASH_CLASSES) {
    const a = summary.seo.byClass[cls];
    if (!a) continue;
    newRows.push(historyRow_(dayKey, 'CLASS_SEO', '', cls, a, ts));
  }

  // --- 同日分を削除（A列だけを見る） ---
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const deleteRows = [];
    for (let i = 0; i < dates.length; i++) {
      if (String(dates[i] || '') === dayKey) deleteRows.push(i + 2);
    }
    // 下から消す（行ズレ防止）
    deleteRows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  }

  // --- 末尾に追加 ---
  const appendStart = sh.getLastRow() + 1;
  sh.getRange(appendStart, 1, newRows.length, header.length).setValues(newRows);

  // 表示形式（A列/P列含め確実に見えるように）
  const lastRow2 = sh.getLastRow();
  if (lastRow2 >= 2) {
    sh.getRange(2, 1, lastRow2 - 1, 1).setNumberFormat('yyyy-mm-dd');           // 日付
    sh.getRange(2, 8, lastRow2 - 1, 1).setNumberFormat('0.00%');                // CTR
    sh.getRange(2, 9, lastRow2 - 1, 1).setNumberFormat('0.00');                 // 平均順位
    sh.getRange(2, 10, lastRow2 - 1, 6).setNumberFormat('0.00%');               // Δ/率
    sh.getRange(2, 5, lastRow2 - 1, 2).setNumberFormat('#,##0');                // clicks/imps
    sh.getRange(2, 16, lastRow2 - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新TS
  }
}


function historyRow_(dayKey, scope, domain, cls, a, ts) {
  return [
    dayKey, scope, domain, cls,
    a.pages ?? 0,
    a.clicks ?? 0,
    a.imps ?? 0,
    a.ctr ?? null,
    a.posWavg ?? null,
    a.deltaClicks ?? null,
    a.deltaImps ?? null,
    a.zeroClickRate ?? null,
    a.lowCtrRate ?? null,
    a.recoveryRate ?? null,
    a.worseningRate ?? null,
    ts,
  ];
}

/* ----------------- 凡例 ----------------- */

function buildLegend_() {
  return [
    ['【凡例】'],
    ['クリック/表示：直近1か月（GSC）'],
    ['CTR：クリック ÷ 表示'],
    ['平均順位：IMP加重平均'],
    ['Δクリック：前年比（1M ÷ LY1）'],
    ['Δ表示：前年比（1M ÷ LY1）'],
    [`0クリック率：表示≥${MIN_IMP_FOR_CAND} かつ クリック=0 の割合`],
    [`低CTR率：表示≥${MIN_IMP_FOR_CAND} かつ CTR≤${(LOW_CTR_THRESHOLD*100).toFixed(1)}% の割合`],
    ['持ち直し：出現(click>0&LY=0)を優先。次にΔ1M>Δ3M'],
    ['悪化：消失(click=0&LY>0)を優先。次にΔ1M<Δ3M'],
    ['※ NW / ST はSEO対象外'],
  ];
}

/* =========================================================
 * Utilities
 * =======================================================*/

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureSheetDimensions_(sh, neededRows, neededCols) {
  if (neededCols < 1) neededCols = 1;
  if (neededRows < 1) neededRows = 1;

  const curCols = sh.getMaxColumns();
  if (curCols < neededCols) sh.insertColumnsAfter(curCols, neededCols - curCols);

  const curRows = sh.getMaxRows();
  if (curRows < neededRows) sh.insertRowsAfter(curRows, neededRows - curRows);
}

function getDomainsFromSettings_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    return ss.getSheets()
      .map(s => s.getName())
      .filter(n => n.includes('.'));
  }

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const header = v[0].map(x => String(x || '').trim());
  const iDomain = header.indexOf('ドメイン名');
  if (iDomain === -1) return [];

  const out = [];
  for (let r = 1; r < v.length; r++) {
    const d = String(v[r][iDomain] || '').trim();
    if (d) out.push(d);
  }
  return out;
}

function indexMap_OverView_(header) {
  const map = {};
  header.forEach((h, i) => {
    const key = normalizeHeader_(h);
    if (key) map[key] = i;
  });
  return map;
}

function getCell_(row, idx, key) {
  const i = idx[key];
  if (i === undefined) return '';
  return row[i];
}

function normalizeClass_(v) {
  const s = String(v || '').trim();
  if (!s || s === '-' || s === '0' || s.toLowerCase() === 'none') return '未分類';
  const u = s.toUpperCase();
  if (KNOWN_CLASSES.has(u)) return u;
  return '未分類';
}

function toDashClassKey_(cls) {
  const u = String(cls || '').trim().toUpperCase();
  if (DASH_CLASSES.includes(u)) return u;
  return '未分類';
}

function toNumber_(v) {
  if (v === null || v === '' || v === undefined) return null;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function computeRatio_(a, b) {
  if (b === null || b === undefined || b === 0) return null;
  return a / b;
}

function pad2d_(arr) {
  const w = arr.reduce((m, r) => Math.max(m, (r ? r.length : 0)), 0) || 1;
  return arr.map(r => {
    const rr = (r ? r.slice() : []);
    while (rr.length < w) rr.push('');
    return rr;
  });
}

// ===== デバッグ設定 =====
const DEBUG_VALIDATE = true;
const DEBUG_SHEET = '概観_debug';

/**
 * runOverviewDashboard の最後で呼ぶ（finalizeAgg_ の後、writeOverview_ の前後どちらでも可）
 */
function runSelfChecks_(ss, domains, summary) {
  const errs = [];

  // 全体・SEOの整合性
  errs.push(...validateAgg_(summary.all, 'ALL'));
  errs.push(...validateAgg_(summary.seo, 'SEO'));

  // ドメイン別（SEO）
  for (const d of domains) {
    const a = summary.byDomainSeo[d];
    if (!a) continue;
    errs.push(...validateAgg_(a, `DOMAIN_SEO:${d}`));
  }

  // 追加：ブランド×分類（ページ数）の行和チェック（= ドメインSEOページ数）
  for (const d of domains) {
    const a = summary.byDomainSeo[d];
    if (!a) continue;
    const rowSum = DASH_CLASSES.reduce((s, c) => s + (a.byClass?.[c]?.pages || 0), 0);
    if (rowSum !== a.pages) {
      errs.push(`[DOMAIN_SEO:${d}] ブランド×分類 行和不一致: rowSum=${rowSum}, domainPages=${a.pages}`);
    }
  }

  // 結果をシートに出す（OKでも出す）
  writeDebugSheet_(ss, errs);

  if (errs.length) {
    // 実行を止めて気づけるようにする
    throw new Error('概観 self-check failed:\n' + errs.join('\n'));
  }
}

function validateAgg_(agg, label) {
  const errs = [];

  // 1) byClass のページ合計 = agg.pages
  const sumPages = DASH_CLASSES.reduce((s, c) => s + (agg.byClass?.[c]?.pages || 0), 0);
  if (sumPages !== agg.pages) {
    errs.push(`[${label}] pages不一致: agg.pages=${agg.pages}, sum(byClass.pages)=${sumPages}`);
  }

  // 2) 候補数はページ数を超えない
  if ((agg.candZeroClickPages || 0) > (agg.pages || 0)) {
    errs.push(`[${label}] candZeroClickPagesがpages超過: cand=${agg.candZeroClickPages}, pages=${agg.pages}`);
  }
  if ((agg.candLowCtrPages || 0) > (agg.pages || 0)) {
    errs.push(`[${label}] candLowCtrPagesがpages超過: cand=${agg.candLowCtrPages}, pages=${agg.pages}`);
  }

  // 3) 持ち直し/悪化もページ数を超えない
  if ((agg.recoveryPages || 0) > (agg.pages || 0)) {
    errs.push(`[${label}] recoveryPagesがpages超過: rec=${agg.recoveryPages}, pages=${agg.pages}`);
  }
  if ((agg.worseningPages || 0) > (agg.pages || 0)) {
    errs.push(`[${label}] worseningPagesがpages超過: wor=${agg.worseningPages}, pages=${agg.pages}`);
  }

  // 4) posW（IMPの加重母数）はimpsを超えない（超えるなら加算ミス）
  if ((agg.posW || 0) > (agg.imps || 0)) {
    errs.push(`[${label}] posWがimps超過: posW=${agg.posW}, imps=${agg.imps}`);
  }

  // 5) 率が0〜1の範囲か（nullは許容）
  errs.push(...validateRate_(agg.zeroClickRate, `${label}.zeroClickRate`));
  errs.push(...validateRate_(agg.lowCtrRate, `${label}.lowCtrRate`));
  errs.push(...validateRate_(agg.recoveryRate, `${label}.recoveryRate`));
  errs.push(...validateRate_(agg.worseningRate, `${label}.worseningRate`));

  // 6) byClass 側も同様
  for (const c of DASH_CLASSES) {
    const ca = agg.byClass?.[c];
    if (!ca) continue;

    if ((ca.candZeroClickPages || 0) > (ca.pages || 0)) {
      errs.push(`[${label}.${c}] candZeroClickPagesがpages超過: cand=${ca.candZeroClickPages}, pages=${ca.pages}`);
    }
    if ((ca.candLowCtrPages || 0) > (ca.pages || 0)) {
      errs.push(`[${label}.${c}] candLowCtrPagesがpages超過: cand=${ca.candLowCtrPages}, pages=${ca.pages}`);
    }
    if ((ca.recoveryPages || 0) > (ca.pages || 0)) {
      errs.push(`[${label}.${c}] recoveryPagesがpages超過: rec=${ca.recoveryPages}, pages=${ca.pages}`);
    }
    if ((ca.worseningPages || 0) > (ca.pages || 0)) {
      errs.push(`[${label}.${c}] worseningPagesがpages超過: wor=${ca.worseningPages}, pages=${ca.pages}`);
    }
    if ((ca.posW || 0) > (ca.imps || 0)) {
      errs.push(`[${label}.${c}] posWがimps超過: posW=${ca.posW}, imps=${ca.imps}`);
    }

    errs.push(...validateRate_(ca.zeroClickRate, `${label}.${c}.zeroClickRate`));
    errs.push(...validateRate_(ca.lowCtrRate, `${label}.${c}.lowCtrRate`));
    errs.push(...validateRate_(ca.recoveryRate, `${label}.${c}.recoveryRate`));
    errs.push(...validateRate_(ca.worseningRate, `${label}.${c}.worseningRate`));
  }

  return errs;
}

function validateRate_(v, name) {
  const errs = [];
  if (v === null || v === undefined || v === '') return errs;
  if (typeof v !== 'number' || isNaN(v)) {
    errs.push(`[RATE] ${name} が数値ではありません: ${v}`);
    return errs;
  }
  if (v < 0 || v > 1) {
    errs.push(`[RATE] ${name} が範囲外(0〜1): ${v}`);
  }
  return errs;
}

function writeDebugSheet_(ss, errs) {
  const sh = ensureSheet_(ss, DEBUG_SHEET);
  sh.getDataRange().clearContent();

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

  const out = [];
  out.push(['概観_debug', `更新: ${now}`]);
  out.push(['結果', errs.length ? 'NG（エラーあり）' : 'OK']);
  out.push(['', '']);
  out.push(['エラー一覧']);
  if (errs.length) {
    errs.forEach(e => out.push([e]));
  } else {
    out.push(['整合性チェックはすべてOK']);
  }

  const padded = pad2d_(out);
  ensureSheetDimensions_(sh, padded.length, padded[0].length);
  sh.getRange(1, 1, padded.length, padded[0].length).setValues(padded);
  sh.setFrozenRows(1);
}

function normalizeHeader_(v) {
  return String(v || '')
    .replace(/^\uFEFF/, '')   // BOM除去
    .replace(/\r?\n/g, ' ')
    .trim();
}

function looksLikeUrl_(v) {
  return /^https?:\/\//i.test(String(v || '').trim());
}

function debugReadState_(ss, domain, header, idx, values) {
  const sh = ensureSheet_(ss, '概観_debug_read');
  sh.clearContents();

  const out = [];

  out.push(['domain', domain]);
  out.push(['timestamp', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')]);
  out.push(['', '']);

  out.push(['【header raw】']);
  out.push(header.map((h, i) => `${i}:${String(h)}`));
  out.push(['', '']);

  out.push(['【index map】']);
  out.push(['URL', idx['URL'] ?? '']);
  out.push(['分類', idx['分類'] ?? '']);
  out.push(['タイトル', idx['タイトル'] ?? '']);
  out.push(['メタ', idx['メタ'] ?? '']);
  out.push(['更新日(lastmod)', idx['更新日(lastmod)'] ?? '']);
  out.push(['CLI_1M', idx['CLI_1M'] ?? '']);
  out.push(['IMP_1M', idx['IMP_1M'] ?? '']);
  out.push(['LY1_CLI', idx['LY1_CLI'] ?? '']);
  out.push(['LY1_IMP', idx['LY1_IMP'] ?? '']);
  out.push(['POS_1M', idx['POS_1M'] ?? '']);
  out.push(['3M_CLI', idx['3M_CLI'] ?? '']);
  out.push(['LY3_CLI', idx['LY3_CLI'] ?? '']);
  out.push(['', '']);

  out.push([
    'rowNo',
    'read:URL',
    'read:分類',
    'read:タイトル',
    'read:メタ',
    'A_actual',
    'B_actual',
    'C_actual',
    'D_actual',
    'E_actual',
    'urlLike?',
  ]);

  const max = Math.min(values.length - 1, 20);
  for (let r = 1; r <= max; r++) {
    const row = values[r];
    const readUrl   = getCell_(row, idx, 'URL');
    const readClass = getCell_(row, idx, '分類');
    const readTitle = getCell_(row, idx, 'タイトル');
    const readMeta  = getCell_(row, idx, 'メタ');

    out.push([
      r + 1,
      String(readUrl || ''),
      String(readClass || ''),
      String(readTitle || ''),
      String(readMeta || ''),
      String(row[0] || ''),
      String(row[1] || ''),
      String(row[2] || ''),
      String(row[3] || ''),
      String(row[4] || ''),
      looksLikeUrl_(readUrl) ? 'OK' : 'NG',
    ]);
  }

  const padded = pad2d_(out);
  ensureSheetDimensions_(sh, padded.length, padded[0].length);
  sh.getRange(1, 1, padded.length, padded[0].length).setValues(padded);
  sh.setFrozenRows(1);
}

function assertRuntimeColumnMapping_(domain, idx, values) {
  const sampleMax = Math.min(values.length - 1, 20);
  let badUrl = 0;
  let badClass = 0;

  for (let r = 1; r <= sampleMax; r++) {
    const row = values[r];
    const readUrl = getCell_(row, idx, 'URL');
    const readClass = String(getCell_(row, idx, '分類') || '').trim().toUpperCase();

    if (readUrl && !looksLikeUrl_(readUrl)) badUrl++;

    if (
      readClass &&
      !['TP', 'SP', 'MC', 'SC', 'NW', 'ST', '-', '未分類'].includes(readClass)
    ) {
      badClass++;
    }
  }

  if (badUrl >= 3 || badClass >= 3) {
    throw new Error(
      `列読み取り異常: ${domain} / badUrl=${badUrl} / badClass=${badClass}`
    );
  }
}


