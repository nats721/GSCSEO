/*************************************************************
 * GSC 改善候補抽出 & クエリ順位差分統合版（+ ZeroClick 統合）
 * - Weekly / Monthly / LongTerm：差分（cur vs prev）で悪化抽出
 * - ZeroClick：表示あり × クリック0（スナップショット）で抽出
 * - Manual：手動キュー（弱いSP等を常駐）
 * - Query：表示回数上位（+ 差分系のみ順位急落）を抽出
 * - GPTプロンプト：本文（文末CTA除外）を埋め込み出力
 *
 * ★重要修正点（要望反映）
 * 1) 未分類（空欄 / "-" / "未分類"）は作業対象にしない（推定もしない）
 * 2) URLの正規化（末尾/）＋レンジ横断の重複排除（同じURLは1回だけ）
 * 3) 並べ替え（スコア順ソート）は不要 → OFF（必要なら設定でON可）
 * 4) 手動で作業対象を固定できるシートを用意：作業対象_手動（永続）
 *************************************************************/

const BASE_SHEET_ID   = '1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs';   // 読み取りのみ
const OUTPUT_SHEET_ID = '1_URzGA15RYBOmLfjsifuuMIigAqR_YaPJvMbPgQosVw';   // 出力のみ
const CACHE_SHEET_ID  = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o';   // HTMLキャッシュのSS

const DEFAULT_TOP_N = 200;
const TOP_N_BY_RANGE = {
  Weekly: 200,
  Monthly: 200,
  LongTerm: 200,
  ZeroClick: 200,
  Snippet: 200,
  Manual: 200
};
const QUERY_TOP = 100;
const EXCLUDE_RECENT_LASTMOD_DAYS = 0; // ← 好きなNに

// ★ 並べ替え（スコア順）不要：OFF（必要なら true）
const ENABLE_SORT_BY_SCORE = false;

// ★ 手動キュー（永続）／手動出力（毎回生成）
const MANUAL_WORKLIST_SHEET = "作業対象_手動";
const MANUAL_OUTPUT_SHEET   = "改善候補_Manual";

const QUERY_LIMIT = {
  Weekly:   100,
  Monthly:  100,
  LongTerm: 50,
  ZeroClick: 50,
  Manual:   50,
    Snippet:  50   // ★追加
};

const DAY = 86400 * 1000;

/* =========================
   ZeroClick 統合設定
========================= */
const ZEROCLICK_MIN_IMP = 30;      // 表示回数下限
const ZEROCLICK_DAYS    = 14;     // 直近28日スナップショット
const ZEROCLICK_LAG_DAYS = 3;     // GSC遅延を見込んで最新日から引く

/* =========================
   Snippet（スニペット不一致）設定  ★追加
========================= */
const SNIPPET_MIN_IMP    = 50;   // 表示回数下限（ノイズ抑制）
const SNIPPET_MAX_POS    = 10;   // 1〜10位を優先（不要なら100などに）
const SNIPPET_MAX_CTR    = 0.02; // CTR上限（2%以下を対象）
const SNIPPET_DAYS       = 28;   // 直近28日スナップショット
const SNIPPET_LAG_DAYS   = 3;    // GSC遅延見込み
const SNIPPET_INTRO_CHARS = 120; // 冒頭文の判定文字数

/* =========================
   非HTMLリソース除外（PDF/画像など）
========================= */
const EXCLUDE_NON_HTML_RESOURCES = true;



// PDF/画像/Office/圧縮/動画など（必要に応じて増減）
const EXCLUDE_EXTS = [
  'pdf',
  'png','jpg','jpeg','gif','webp','svg','bmp','tif','tiff','ico',
  'doc','docx','xls','xlsx','ppt','pptx','csv',
  'zip','rar','7z','gz','tgz',
  'mp3','wav','m4a',
  'mp4','mov','m4v','webm'
];

/* =========================
   ★ ページ分類（略語）説明（GPTプロンプトに埋め込む）
   - 未分類（空欄 / "-" / "未分類"）は作業対象外（推定もしない）
========================= */
const PAGE_CATEGORY_GLOSSARY = [
  "TP = トップページ（ホーム/サイトの入口）",
  "SP = サービスページ（講座・コース・申込導線）",
  "MC = メインコラム（コラム/解説/勉強法など集客記事）",
  "SC = サブコラム（MCやSPを補強するための補助コラム※サイト定義に合わせて判断）",
  "NW = ニュース/お知らせ（運用上は改善対象外として除外）",
  "ST = 企業情報等のSEO対策を行わないこ固定ページ（運用上は改善対象外として除外）",
  "- / 空欄 / 未分類 = 未分類（このスクリプトでは改善対象外として除外）"
].join("\n");

/* =========================
   ★ 禁則ワード（外部SSから読み込み）
========================= */
const FORBIDDEN_WORDS_SS_ID   = '1-LnhzG2eE0C3XR0RK6G7-2AQIzT1cBeGMHMiZkpDNls';
const FORBIDDEN_WORDS_SHEET   = '禁則ワード';
const FORBIDDEN_WORDS_COL     = 1;     // A列
const FORBIDDEN_WORDS_MAX_IN_PROMPT = 200; // プロンプトに載せる最大数（長すぎる場合の保険）


function buildPageCategoryHelp_(cat){
  const c = String(cat || '').trim();
  const cur = c ? `【ページ分類】${c}` : "【ページ分類】（未分類：このスクリプトでは対象外）";
  return `${cur}
【ページ分類の略語（このプロンプト内の定義）】
${PAGE_CATEGORY_GLOSSARY}
※未分類（空欄 / "-" / "未分類"）は作業対象に含めません。`;
}

/*************************************************************
 * 週次運用版 × 統合スコアロジック（CTR＋順位＋クリック） + ZeroClick + Manual
 *************************************************************/
function runGSCImproveExtractionWeekly() {

  const domains = loadDomainsFromExistingSheet();
  const ssOut   = SpreadsheetApp.openById(OUTPUT_SHEET_ID);

  clearOutputFour_(ssOut);
  ensureManualWorklistSheet_(ssOut);

  const cacheMeta = loadCacheMetaMaps_();               // lastmod + category
  const lastmodMap  = cacheMeta.lastmodMap;
  const categoryMap = cacheMeta.categoryMap;

  const bodyMap  = loadBodyMapCleaned_();               // 本文（文末CTA除外済）

  // 差分レンジ（クリックがある最新日で合わせる）
  const latestForDiff = detectLatestAvailableDate(domains[0]);

  // ZeroClick はクリック前提の latest を使わない（固定ラグでスナップショット）
  const zeroBase = getZeroClickBaseDate_();
  const snippetBase = getSnippetBaseDate_(); // ★追加
  const ranges = {
    weekly: makeRangeWeekly_(latestForDiff),
    monthly: makeRangeMonthly_(latestForDiff),
    long: makeRangeLong_(latestForDiff),
    zero: makeRangeZeroClick_(zeroBase),
    snippet: makeRangeSnippet_(snippetBase)  
  };

  // ★レンジ横断の重複排除（同じURLは1回だけ出す）
  const picked = new Set();

  // 1) Manual（手動キュー）を最優先で出す（弱いSP等を常駐）
  extractManualTargets_(ssOut, domains, ranges.zero.cur, MANUAL_OUTPUT_SHEET, lastmodMap, categoryMap, bodyMap, picked);
 // ★追加：スニペット（不一致検出）
  extractRangeSnippetMismatch_(ssOut, domains, ranges.snippet, "改善候補_スニペット", "Snippet",
                              lastmodMap, categoryMap, bodyMap, picked);
  // 2) ZeroClick（表示あり×クリック0）
  extractRangeZeroClick_(ssOut, domains, ranges.zero, "改善候補_ZeroClick", "ZeroClick",
                         lastmodMap, categoryMap, bodyMap, picked);

  // 3) Weekly / Monthly / LongTerm（差分系）
  extractRangeWithScore_(ssOut, domains, ranges.weekly,  "改善候補_Weekly",  "Weekly",  lastmodMap, categoryMap, bodyMap, picked);
  extractRangeWithScore_(ssOut, domains, ranges.monthly, "改善候補_Monthly", "Monthly", lastmodMap, categoryMap, bodyMap, picked);
  extractRangeWithScore_(ssOut, domains, ranges.long,    "改善候補_LongTerm","LongTerm",lastmodMap, categoryMap, bodyMap, picked);

}

/*************************************************************
 * 最新日検出（過去14日クリックありの最終日）※差分レンジ専用
 *************************************************************/
function detectLatestAvailableDate(domain) {
  const token = ScriptApp.getOAuthToken();
  const url =
    "https://searchconsole.googleapis.com/webmasters/v3/sites/" +
    encodeURIComponent(domain) + "/searchAnalytics/query";

  const payload = {
    startDate: fmt(shift(new Date(), -14)),
    endDate: fmt(new Date()),
    dimensions: ["date"],
    rowLimit: 300
  };

  const res = UrlFetchApp.fetch(url, {
    method:"post",
    contentType:"application/json",
    payload:JSON.stringify(payload),
    headers:{ Authorization:"Bearer "+token }
  });

  const json = JSON.parse(res.getContentText());
  let latest = null;

  (json.rows || []).forEach(r => {
    if (r.clicks > 0) latest = r.keys[0];
  });

  // クリックが取れない場合は安全側へ
  return latest ? new Date(latest) : shift(new Date(), -4);
}

/*************************************************************
 * ZeroClick 用の基準日（クリック非依存）
 *************************************************************/
function getZeroClickBaseDate_() {
  return shift(new Date(), -ZEROCLICK_LAG_DAYS);
}

/*************************************************************
 * Snippet 用の基準日（クリック非依存） ★追加
 *************************************************************/
function getSnippetBaseDate_() {
  return shift(new Date(), -SNIPPET_LAG_DAYS);
}

/*************************************************************
 * Snippet レンジ（スナップショット） ★追加
 *************************************************************/
function makeRangeSnippet_(base){
  return {
    cur: { start: shift(base, -(SNIPPET_DAYS-1)), end: base }
  };
}

/*************************************************************/
function loadDomainsFromExistingSheet() {
  const ss = SpreadsheetApp.openById(BASE_SHEET_ID);
  return ss.getSheets()
    .map(sh => sh.getName())
    .filter(n => n.includes('.') && n !== "設定")
    .map(n => `sc-domain:${n}`);
}

/*************************************************************/
function shift(dt,days){ return new Date(dt.getTime()+days*DAY); }

function getTopN_(rangeName){
  const n = TOP_N_BY_RANGE[rangeName];
  return (typeof n === "number" && n > 0) ? n : DEFAULT_TOP_N;
}

/*************************************************************
 * ★修正：fetchPagesMetrics
 * - URLを normalizeUrl_ して同一URL（末尾/違い等）を統合
 * - ctr は clicks / impressions で再計算
 * - position は impressions で重み付け平均（imp=0のときは1で近似）
 *************************************************************/
function fetchPagesMetrics(domain, range) {
  const token = ScriptApp.getOAuthToken();
  const url =
    "https://searchconsole.googleapis.com/webmasters/v3/sites/" +
    encodeURIComponent(domain) + "/searchAnalytics/query";

  const payload = {
    startDate: fmt(range.start),
    endDate: fmt(range.end),
    dimensions: ["page"],
    rowLimit: 20000
  };

  const res = UrlFetchApp.fetch(url, {
    method:"post",
    contentType:"application/json",
    payload:JSON.stringify(payload),
    headers:{ Authorization:"Bearer "+token }
  });

  const json = JSON.parse(res.getContentText());

  // 集約用
  const acc = {};

  (json.rows || []).forEach(r => {
    const raw = safeDecode(r.keys[0]);
    const pageUrl = normalizeUrl_(raw);

    const clicks = Number(r.clicks) || 0;
    const impressions = Number(r.impressions) || 0;
    const position = (typeof r.position === 'number' && isFinite(r.position)) ? Number(r.position) : 100;

    if (!acc[pageUrl]) acc[pageUrl] = { clicks:0, impressions:0, posSum:0, posW:0 };

    acc[pageUrl].clicks += clicks;
    acc[pageUrl].impressions += impressions;

    const w = impressions > 0 ? impressions : 1;
    acc[pageUrl].posSum += position * w;
    acc[pageUrl].posW   += w;
  });

  const map = {};
  Object.keys(acc).forEach(u => {
    const a = acc[u];
    const imp = a.impressions;
    const clk = a.clicks;
    map[u] = {
      clicks: clk,
      impressions: imp,
      ctr: imp > 0 ? (clk / imp) : 0,
      position: a.posW > 0 ? (a.posSum / a.posW) : 100
    };
  });

  return map;
}

/*************************************************************/
function fetchQueries(domain, url, range) {

  const encodedUrl = url.replace(
    /[^\w\-\.~:\/\?#\[\]@!$&'()*+,;=%]/g,
    c => encodeURIComponent(c)
  );

  const token = ScriptApp.getOAuthToken();
  const endpoint =
    "https://searchconsole.googleapis.com/webmasters/v3/sites/" +
    encodeURIComponent(domain) + "/searchAnalytics/query";

  const payload = {
    startDate: fmt(range.start),
    endDate: fmt(range.end),
    dimensions: ["query"],
    dimensionFilterGroups: [{
      filters:[{
        dimension:"page",
        operator:"equals",
        expression: encodedUrl
      }]
    }],
    rowLimit: 2000
  };

  const res = UrlFetchApp.fetch(endpoint, {
    method:"post",
    contentType:"application/json",
    payload:JSON.stringify(payload),
    headers:{ Authorization:"Bearer "+token }
  });

  const json = JSON.parse(res.getContentText());
  return (json.rows||[]).map(r => ({
    query: safeDecode(r.keys[0]),
    clicks: r.clicks || 0,
    impressions: r.impressions || 0,
    ctr: r.ctr || 0,
    pos: r.position || 100
  }));
}

/*************************************************************/
function joinQueries(cur, prev) {
  const map = {};

  cur.forEach(q => map[q.query] = { cur:q, prev:null });
  prev.forEach(p => {
    if (!map[p.query]) map[p.query] = { cur:null, prev:p };
    else map[p.query].prev = p;
  });

  return Object.keys(map).map(k=>{
    const a = map[k].cur  || { clicks:0, impressions:0, ctr:0, pos:100 };
    const b = map[k].prev || { clicks:0, impressions:0, ctr:0, pos:100 };

    return {
      query:k,
      clicksCur:a.clicks,
      clicksPrev:b.clicks,
      impCur:a.impressions,
      impPrev:b.impressions,
      ctrCur:a.ctr,
      ctrPrev:b.ctr,
      posCur:a.pos,
      posPrev:b.pos,
      posDiff:a.pos - b.pos
    };
  });
}

// ★ レンジ別しきい値（差分系のみ）
function passThreshold(rangeType, a, b) {

  const impCur  = a.impressions;
  const impPrev = b.impressions;
  const clkCur  = a.clicks;
  const clkPrev = b.clicks;

  if (rangeType === "Weekly") {
    return (
      impCur >= 30  || impPrev >= 30 ||
      clkCur >= 5   || clkPrev >= 5
    );
  }

  if (rangeType === "Monthly") {
    return (
      impCur >= 100 || impPrev >= 100 ||
      clkCur >= 10  || clkPrev >= 10
    );
  }

  if (rangeType === "LongTerm") {
    return (impCur >= 1 || impPrev >= 1);
  }

  return true;
}

/*************************************************************/
function fmt(dt){ return Utilities.formatDate(dt,"Asia/Tokyo","yyyy-MM-dd"); }
function safeDecode(str){
  try { return decodeURIComponent(str); }
  catch(e) { return String(str || ""); }
}

/*************************************************************
 * 出力：改善候補_* / クエリ詳細 / GPTプロンプト
 * - 0件でも落ちないように安全化
 *************************************************************/
function writeExtractionOutput_(ss, sheetName, top, detail, prompts){

  let sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  sh.getRange(1,1,1,17).setValues([[
    "Domain",
    "URL",
    "Clicks(cur)",
    "Clicks(prev)",
    "Click差(cur-prev)",
    "CTR(cur)",
    "CTR(prev)",
    "CTR差(cur-prev)",
    "Pos(cur)",
    "Pos(prev)",
    "Pos差(cur-prev)",
    "Impr(cur)",
    "Impr(prev)",
    "Impr差(cur-prev)",
    "Score",
    "Lastmod",
    "Range"
  ]]);

  if (!top || !top.length) return;

  const fmtPct_ = v => isFinite_(v) ? (v*100).toFixed(2) + "%" : "";
  const fmt1_   = v => isFinite_(v) ? Number(v).toFixed(1) : "";
  const fmt0_   = v => isFinite_(v) ? Number(v) : "";
  const diff_   = (a,b, fn) => (isFinite_(a) && isFinite_(b)) ? fn(a-b) : "";

  sh.getRange(2,1,top.length,17).setValues(
    top.map(r => [
      r.domain || "",
      r.url || "",
      fmt0_(r.clicksCur),
      (r.clicksPrev === null || r.clicksPrev === undefined) ? "" : fmt0_(r.clicksPrev),
      diff_(r.clicksCur, r.clicksPrev, v => String(v)),
      fmtPct_(r.ctrCur),
      (r.ctrPrev === null || r.ctrPrev === undefined) ? "" : fmtPct_(r.ctrPrev),
      diff_(r.ctrCur, r.ctrPrev, v => (v*100).toFixed(2) + "%"),
      fmt1_(r.posCur),
      (r.posPrev === null || r.posPrev === undefined) ? "" : fmt1_(r.posPrev),
      diff_(r.posCur, r.posPrev, v => Number(v).toFixed(1)),
      fmt0_(r.impCur),
      (r.impPrev === null || r.impPrev === undefined) ? "" : fmt0_(r.impPrev),
      diff_(r.impCur, r.impPrev, v => String(v)),
      isFinite_(r.score) ? Number(r.score).toFixed(2) : "",
      r.lastmod || "",
      sheetName
    ])
  );

  // ▼ クエリ詳細に追記
  let qSh = ss.getSheetByName("クエリ詳細") || ss.insertSheet("クエリ詳細");
  let rows = [];

  (detail || []).forEach(d =>
    (d.queries || []).forEach(q =>
      rows.push([
        sheetName,
        d.url || "",
        q.query || "",
        isFinite_(q.clicksCur) ? q.clicksCur : "",
        (q.clicksPrev === null || q.clicksPrev === undefined) ? "" : q.clicksPrev,
        isFinite_(q.posCur) ? Number(q.posCur).toFixed(1) : "",
        (q.posPrev === null || q.posPrev === undefined) ? "" : Number(q.posPrev).toFixed(1),
        isFinite_(q.posDiff) ? Number(q.posDiff).toFixed(1) : ""
      ])
    )
  );

  if (rows.length){
    const startRow = qSh.getLastRow() + 1;
    qSh.getRange(startRow,1,rows.length,8).setValues(rows);
  }

  // ▼ GPTプロンプトに追記
  let pSh = ss.getSheetByName("GPTプロンプト") || ss.insertSheet("GPTプロンプト");
  if (prompts && prompts.length){
    const startP = pSh.getLastRow() + 1;

    // ★行ごとに式を作る（C列=URL を参照）
    const makeLookupFormula_ = (rowNum) => `=IFERROR(
    VLOOKUP(
      C${rowNum},
      IMPORTRANGE(
        "${CACHE_SHEET_ID}",
        "HTMLキャッシュ!A:F"
      ),
      3,
      FALSE
    ),
    0
  )`;

    pSh.getRange(startP,1,prompts.length,6).setValues(
      prompts.map((p,i) => {
        const rowNum = startP + i; // ★ここが肝（実際の行番号）
        return [
          sheetName,
          (top[i] && top[i]._category) ? top[i]._category : "",
          (top[i] && top[i].url) ? top[i].url : "",
          (top[i] && top[i].lastmod) ? top[i].lastmod : "",
          p,
          makeLookupFormula_(rowNum) // ★ i ではなく rowNum
        ];
      })
    );
  }
}

/*************************************************************
 * HTMLキャッシュから lastmod / category をまとめてロード（URL正規化も併用）
 *************************************************************/
function loadCacheMetaMaps_() {
  const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  const sh = ss.getSheetByName('HTMLキャッシュ');

  const lastmodMap  = {};
  const categoryMap = {};

  const n = Math.max(sh.getLastRow()-1, 0);
  if (!n) return { lastmodMap, categoryMap };

  // A:URL, B:分類, C:lastmod
  const values = sh.getRange(2,1,n,3).getValues();

  values.forEach(([url, category, lastmod]) => {
    if (!url) return;
    const raw = String(url);
    const norm = normalizeUrl_(raw);

    if (lastmod) {
      lastmodMap[raw] = lastmod;
      if (norm !== raw) lastmodMap[norm] = lastmod;
    } else {
      lastmodMap[raw] = lastmodMap[raw] || "";
      if (norm !== raw) lastmodMap[norm] = lastmodMap[norm] || "";
    }

    if (category) {
      categoryMap[raw] = String(category);
      if (norm !== raw) categoryMap[norm] = String(category);
    }
  });

  return { lastmodMap, categoryMap };
}
/* =========================
   HTMLキャッシュ：title/meta/H1/entryBodyHt Map ★追加
   （A:URL, E:title, F:metaDesc, G:H1, H:entryBodyHt）
========================= */
let __CACHE_SNIPPET_MAPS__ = null;

function loadCacheSnippetMaps_(){
  const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  const sh = ss.getSheetByName('HTMLキャッシュ');

  const titleMap = {};
  const metaMap  = {};
  const h1Map    = {};
  const entryMap = {};

  const n = Math.max(sh.getLastRow() - 1, 0);
  if (!n) return { titleMap, metaMap, h1Map, entryMap };

  // A〜H（8列）を読む
  const values = sh.getRange(2, 1, n, 8).getValues();

  values.forEach(r => {
    const urlRaw = r[0];
    if (!urlRaw) return;

    const raw  = String(urlRaw).trim();
    const norm = normalizeUrl_(raw);

    const title = r[4] || ""; // E
    const meta  = r[5] || ""; // F
    const h1    = r[6] || ""; // G
    const entry = r[7] || ""; // H

    titleMap[raw] = title; titleMap[norm] = title;
    metaMap[raw]  = meta;  metaMap[norm]  = meta;
    h1Map[raw]    = h1;    h1Map[norm]    = h1;
    entryMap[raw] = entry; entryMap[norm] = entry;
  });

  return { titleMap, metaMap, h1Map, entryMap };
}

function getCacheSnippetMaps_(){
  if (__CACHE_SNIPPET_MAPS__ !== null) return __CACHE_SNIPPET_MAPS__;
  __CACHE_SNIPPET_MAPS__ = loadCacheSnippetMaps_();
  return __CACHE_SNIPPET_MAPS__;
}

function getCurrentTitleMetaH1Entry_(url){
  const maps = getCacheSnippetMaps_();
  const u = normalizeUrl_(url);
  return {
    title: maps.titleMap[u] || "",
    meta:  maps.metaMap[u]  || "",
    h1:    maps.h1Map[u]    || "",
    entry: maps.entryMap[u] || ""
  };
}
/*************************************************************
 * ★未分類判定（空欄 / "-" / "未分類"）
 *************************************************************/
function isUnclassifiedCategory_(cat){
  const c = String(cat || '').trim();
  return (c === "" || c === "-" || c === "未分類");
}

/*************************************************************
 * レンジごとの抽出処理（差分・統合スコア版）
 * - 対象レンジ: Weekly / Monthly / LongTerm（この関数で共通処理）
 * ★修正：
 * - 未分類は除外（推定しない）
 * - レンジ横断の重複排除（picked）
 * - 並べ替え不要：ENABLE_SORT_BY_SCORE=false
 *************************************************************/
function extractRangeWithScore_(ssOut, domains, range, sheetName, rangeName, lastmodMap, categoryMap, bodyMap, picked){
  let table = [];

  domains.forEach(domain => {

    const cur  = fetchPagesMetrics(domain, range.cur);
    const prev = fetchPagesMetrics(domain, range.prev);

    Object.keys(cur).forEach(url => {
      if (EXCLUDE_NON_HTML_RESOURCES && isNonHtmlResourceUrl_(url)) return;

      const normUrl = normalizeUrl_(url);

      // ★レンジ横断の重複排除
      if (picked && picked.has(normUrl)) return;

      const a = cur[normUrl];
      const b = prev[normUrl] || { clicks:0, impressions:0, ctr:0, position:100 };

      if (!passThreshold(rangeName, a, b)) return;

      const clicksCur  = a.clicks;
      const clicksPrev = b.clicks;
      const ctrCur     = a.ctr;
      const ctrPrev    = b.ctr;
      const posCur     = a.position;
      const posPrev    = b.position;

      const ctrDiff    = ctrCur - ctrPrev;
      const posDiff    = posCur - posPrev;
      const clicksDiff = clicksCur - clicksPrev;

      const impCur  = a.impressions;
      const impPrev = b.impressions;
      const impDiff = impCur - impPrev;

      // ★ ページ分類：未分類は作業対象外（推定もしない）
      const catRaw = (categoryMap && (categoryMap[normUrl] || categoryMap[url] || categoryMap[normalizeUrl_(url)])) || "";
      const cat = String(catRaw || '').trim();

      // ★ NW / ST / 未分類 は出力しない（最重要）
      if (cat === "NW" || cat === "ST" || isUnclassifiedCategory_(cat)) return;

      const lastmod = (lastmodMap && (lastmodMap[normUrl] || lastmodMap[url] || lastmodMap[normalizeUrl_(url)])) || "";

      if (!hasValidLastmod_(lastmod)) return;

      // ★ lastmodが新しすぎるページは除外
      if (shouldExcludeByLastmod_(lastmod)) return;

      if (rangeName === "Weekly") {
        if (Math.abs(ctrDiff) < 0.005) return;
        if (Math.abs(posDiff) < 1.0) return;
        if (Math.abs(clicksDiff) <= -1) return;
      }

      if (rangeName === "LongTerm" && clicksCur === 0 && clicksPrev === 0) return;

      let score = 0;
      let hasHardBad = false;

      if (ctrPrev > 0 && ctrCur + 0.001 < ctrPrev) {
        score += (ctrPrev - ctrCur) * 5;
        hasHardBad = true;
      }

      if (clicksPrev >= 10 && clicksCur < clicksPrev * 0.9) {
        score += (clicksPrev - clicksCur) * 0.5;
        hasHardBad = true;
      }

      if (hasHardBad && posCur - posPrev > 0.5) {
        score += (posCur - posPrev) * 3;
      }

      if (!hasHardBad || score <= 0.5) return;

      table.push({
        domain,
        url: normUrl,
        clicksCur, clicksPrev,
        impCur, impPrev,
        ctrCur, ctrPrev,
        posCur, posPrev,
        impDiff,
        score,
        ctrDiff, posDiff, clicksDiff,
        lastmod,
        _category: cat
      });

    });
  });

  // ★並べ替え不要（必要なら ENABLE_SORT_BY_SCORE=true）
  if (ENABLE_SORT_BY_SCORE) {
    table.sort((a,b) => b.score - a.score);
  }

  const top = table.slice(0, getTopN_(rangeName));

  // ★採用URLを登録（レンジ横断で重複排除）
  if (picked){
    top.forEach(r => picked.add(normalizeUrl_(r.url)));
  }

  const detail = top.map(row => {

    const curQ   = fetchQueries(row.domain, row.url, range.cur);
    const prevQ  = fetchQueries(row.domain, row.url, range.prev);
    const joined = joinQueries(curQ, prevQ);

    const limit = QUERY_LIMIT[rangeName] || 20;

    const topByClicks = joined
      .filter(q => q.impCur >= (rangeName === "Monthly" ? 20 : 0))
      .sort((a,b)=> b.impCur - a.impCur)
      .slice(0, limit);

    const topByPosDrop = joined
      .filter(q => q.clicksPrev > q.clicksCur)
      .filter(q => rangeName !== "Monthly" || q.impCur >= 20)
      .sort((a,b)=> b.posDiff - a.posDiff)
      .slice(0, limit);

    return { ...row, queries: joined, topByClicks, topByPosDrop };
  });

  const prompts = detail.map(d => {
    if (rangeName === "Weekly")  return makePromptWeekly(d, bodyMap);
    if (rangeName === "Monthly") return makePromptMonthly(d, bodyMap);
    return makePromptLongTerm(d, bodyMap);
  });

  writeExtractionOutput_(ssOut, sheetName, top, detail, prompts);
}

/*************************************************************
 * ZeroClick 抽出（表示あり × クリック0）
 * - 差分ではなくスナップショット
 * ★修正：
 * - 未分類は除外（推定しない）
 * - レンジ横断の重複排除（picked）
 * - 並べ替え不要：ENABLE_SORT_BY_SCORE=false
 *************************************************************/
function extractRangeZeroClick_(ssOut, domains, range, sheetName, rangeName,
                               lastmodMap, categoryMap, bodyMap, picked) {

  let table = [];

  domains.forEach(domain => {

    const cur = fetchPagesMetrics(domain, range.cur);

    Object.keys(cur).forEach(url => {

      if (EXCLUDE_NON_HTML_RESOURCES && isNonHtmlResourceUrl_(url)) return;

      const normUrl = normalizeUrl_(url);

      // ★レンジ横断の重複排除
      if (picked && picked.has(normUrl)) return;

      const a = cur[normUrl];

      if (a.impressions < ZEROCLICK_MIN_IMP) return;
      if (a.clicks !== 0) return;

      const catRaw = (categoryMap && (categoryMap[normUrl] || categoryMap[url])) || "";
      const cat  = String(catRaw || '').trim();

      if (cat === "NW" || cat === "ST" || isUnclassifiedCategory_(cat)) return;

      const lastmod = (lastmodMap && (lastmodMap[normUrl] || lastmodMap[url])) || "";

      // ★追加：lastmod が無い（=HTMLキャッシュに無い / 空）ものは対象外
      if (!hasValidLastmod_(lastmod)) return;

      // ★ ZeroClick でも lastmod が新しすぎるページは除外
      if (shouldExcludeByLastmod_(lastmod)) return;

      // ★スコア：IMPを主軸 + 上位順位ほど少し優先（「見えてるのに押されない」を先に）
      let score = a.impressions;
      if (a.position <= 20) score += (21 - a.position) * 10;

      table.push({
        domain,
        url: normUrl,
        clicksCur: a.clicks,
        clicksPrev: null,
        impCur: a.impressions,
        impPrev: null,
        ctrCur: a.ctr,
        ctrPrev: null,
        posCur: a.position,
        posPrev: null,
        impDiff: null,
        score,
        lastmod,
        _category: cat
      });
    });
  });

  // ★並べ替え不要（必要なら ENABLE_SORT_BY_SCORE=true）
  if (ENABLE_SORT_BY_SCORE) {
    table.sort((a,b) => b.score - a.score);
  }

  const top = table.slice(0, getTopN_(rangeName));

  // ★採用URLを登録（レンジ横断で重複排除）
  if (picked){
    top.forEach(r => picked.add(normalizeUrl_(r.url)));
  }

  const detail = top.map(row => {

    const curQ = fetchQueries(row.domain, row.url, range.cur);

    // joinQueries相当の形に寄せる（prev無し）
    const joined = curQ.map(q => ({
      query: q.query,
      clicksCur: q.clicks,
      clicksPrev: null,
      impCur: q.impressions,
      impPrev: null,
      ctrCur: q.ctr,
      ctrPrev: null,
      posCur: q.pos,
      posPrev: null,
      posDiff: null
    }));

    const limit = QUERY_LIMIT[rangeName] || 20;

    const topByClicks = joined
      .sort((a,b)=> (b.impCur || 0) - (a.impCur || 0))
      .slice(0, limit);

    return { ...row, queries: joined, topByClicks, topByPosDrop: [] };
  });

  const prompts = detail.map(d => makePromptZeroClick(d, bodyMap));
  writeExtractionOutput_(ssOut, sheetName, top, detail, prompts);
}

function normalizeTextForCompare_(s){
  return String(s||"")
    .replace(/\s+/g,"")
    .replace(/[｜\|\-ー―–—:：・，,。．.（）\(\)\[\]【】「」『』"'’`]/g,"")
    .trim();
}

function getIntroForSnippet_(html){
  // stripTags_ は既にあなたのコード内に存在（ZeroClick改修で追加済み）
  // 無い場合は stripTags_ を先に置いて下さい
  const t = stripTags_(String(html||""));
  return [...t].slice(0, SNIPPET_INTRO_CHARS).join("");
}

/*************************************************************
 * Snippet（本文と title/meta/H1/冒頭 の不一致候補） ★追加
 *************************************************************/
function extractRangeSnippetMismatch_(ssOut, domains, range, sheetName, rangeName,
                                     lastmodMap, categoryMap, bodyMap, picked){

  let table = [];

  domains.forEach(domain => {

    const cur = fetchPagesMetrics(domain, range.cur);

    Object.keys(cur).forEach(url => {

      if (EXCLUDE_NON_HTML_RESOURCES && isNonHtmlResourceUrl_(url)) return;

      const normUrl = normalizeUrl_(url);

      // 既に他レンジで拾ったURLは除外（必要なら外してOK）
      if (picked && picked.has(normUrl)) return;

      const a = cur[normUrl];
      if (!a) return;

      // スニペット対象のしきい値
      if (a.impressions < SNIPPET_MIN_IMP) return;
      if (a.position > SNIPPET_MAX_POS) return;
      if (a.ctr > SNIPPET_MAX_CTR) return;

      // 分類（NW/ST/未分類除外）
      const catRaw = (categoryMap && (categoryMap[normUrl] || categoryMap[url])) || "";
      const cat = String(catRaw || '').trim();
      if (cat === "NW" || cat === "ST" || isUnclassifiedCategory_(cat)) return;

      // lastmod 必須（HTMLキャッシュに存在しないものは除外）
      const lastmod = (lastmodMap && (lastmodMap[normUrl] || lastmodMap[url])) || "";
      if (!hasValidLastmod_(lastmod)) return;

      // ★ Snippet でも lastmod が新しすぎるページは除外
      if (shouldExcludeByLastmod_(lastmod)) return;

      // タイトル/メタ/H1/entryBody（HTMLキャッシュから取得）
      const tmh = getCurrentTitleMetaH1Entry_(normUrl);
      const title = String(tmh.title || "");
      const meta  = String(tmh.meta || "");
      const h1    = String(tmh.h1 || "");

      // 冒頭：本文Map優先、無ければ entryBodyHt
      const bodyHtml = getBodyFromMap_(bodyMap, normUrl) || "";
      const srcHtml  = bodyHtml || String(tmh.entry || "");
      const intro = getIntroForSnippet_(srcHtml);

      // GSCクエリ（表示回数上位）
      const curQ = fetchQueries(domain, normUrl, range.cur) || [];
      const joined = curQ.map(q => ({
        query: q.query,
        clicksCur: q.clicks,
        clicksPrev: null,
        impCur: q.impressions,
        impPrev: null,
        ctrCur: q.ctr,
        ctrPrev: null,
        posCur: q.pos,
        posPrev: null,
        posDiff: null
      })).sort((x,y)=> (y.impCur||0) - (x.impCur||0));

      if (!joined.length) return;

      const primaryQuery = String(joined[0].query || "").trim();
      const token = extractRepresentativeToken_(primaryQuery);
      if (!token) return;

      // 不一致判定
      const titleHas = title.includes(token);
      const metaHas  = meta.includes(token);
      const h1Has    = h1.includes(token);
      const introHas = intro.includes(token);

      const missing = (titleHas?0:1) + (metaHas?0:1) + (h1Has?0:1) + (introHas?0:1);
      const titleH1Diff = normalizeTextForCompare_(title) !== normalizeTextForCompare_(h1);

      // 「一致していない」候補として拾う条件
      if (missing === 0 && !titleH1Diff) return;

      // 並び順用スコア（露出が多いほど優先）
      let score = 0;
      score += missing * 50;
      score += titleH1Diff ? 30 : 0;
      score += (a.impressions || 0) * 0.05;

      // 追加：抽出理由文字列（プロンプト用）
      const reasonParts = [];
      if (!titleHas) reasonParts.push("titleに代表語なし");
      if (!metaHas)  reasonParts.push("metaに代表語なし");
      if (!h1Has)    reasonParts.push("H1に代表語なし");
      if (!introHas) reasonParts.push("冒頭に代表語なし");
      if (titleH1Diff) reasonParts.push("titleとH1が不一致");

      table.push({
        domain,
        url: normUrl,
        clicksCur: a.clicks,
        clicksPrev: null,
        impCur: a.impressions,
        impPrev: null,
        ctrCur: a.ctr,
        ctrPrev: null,
        posCur: a.position,
        posPrev: null,
        impDiff: null,
        score,
        lastmod,
        _category: cat,

        // detail/プロンプト用の付帯
        _token: token,
        _intro: intro,
        _h1: h1,
        _title: title,
        _meta: meta,
        _topByClicks: joined.slice(0, (QUERY_LIMIT.Snippet || 50))
      });

    });
  });

  table.sort((a,b) => (b.score||0) - (a.score||0));
  const top = table.slice(0, getTopN_(rangeName));

  if (picked){
    top.forEach(r => picked.add(normalizeUrl_(r.url)));
  }

  const detail = top.map(row => ({
    ...row,
    queries: row._topByClicks || [],
    topByClicks: row._topByClicks || [],
    topByPosDrop: []
  }));

  const prompts = detail.map(d => makePromptSnippet(d, bodyMap));
  writeExtractionOutput_(ssOut, sheetName, top, detail, prompts);
}

/*************************************************************
 * ★Manual（手動キュー）
 * - 入力：作業対象_手動（永続・クリアしない）
 *   A:有効(checkbox) / B:URL / C:メモ / D:最終出力 / E:ステータス
 * - 出力：改善候補_Manual（毎回生成）
 * - 未分類は除外（HTMLキャッシュの分類が空欄/未分類ならスキップ）
 * - lastmod新しすぎ除外は「手動キューでは適用しない」（常駐させるため）
 *************************************************************/
function ensureManualWorklistSheet_(ss){
  const sh = ss.getSheetByName(MANUAL_WORKLIST_SHEET) || ss.insertSheet(MANUAL_WORKLIST_SHEET);
  const header = ["有効","URL","メモ","最終出力","ステータス"];

  if (sh.getLastRow() < 1 || String(sh.getRange(1,1).getValue() || "").trim() === "") {
    sh.getRange(1,1,1,header.length).setValues([header]);
  } else {
    // A1が「有効」でない等でも上書きはしない（運用破壊防止）
  }

  // checkbox（失敗しても落とさない）
  try {
    const max = sh.getMaxRows();
    if (max >= 2) sh.getRange(2,1,max-1,1).insertCheckboxes();
  } catch(e){}

  return sh;
}

function loadManualTargets_(ss){
  const sh = ensureManualWorklistSheet_(ss);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { sh, list: [] };

  const values = sh.getRange(2,1,lastRow-1,5).getValues();
  const list = [];

  values.forEach((r, i) => {
    const enabled = r[0];
    const urlRaw  = String(r[1] || "").trim();
    const memo    = String(r[2] || "").trim();

    if (!urlRaw) return;

    // enabled: FALSE/false/0 のみ無効。空欄は有効扱い（入力の手間を減らす）
    const enStr = String(enabled).toLowerCase().trim();
    const isEnabled = !(enabled === false || enStr === "false" || enStr === "0");

    if (!isEnabled) return;

    list.push({
      url: normalizeUrl_(urlRaw),
      memo,
      row: i + 2
    });
  });

  return { sh, list };
}

function getHost_(url){
  const s = String(url || "").trim();
  const m = s.match(/^https?:\/\/([^\/?#]+)/i);
  return m ? m[1] : "";
}

// sc-domain のどれに属するURLか（最長サフィックス一致）
function resolveDomainForUrl_(url, domains){
  const host0 = getHost_(url);
  if (!host0) return null;

  const host = host0.toLowerCase();
  const hostNoWww = host.replace(/^www\./, "");

  let best = null;
  let bestLen = -1;

  domains.forEach(d => {
    const dn = String(d).replace(/^sc-domain:/,'').toLowerCase();
    const ok =
      host === dn || host.endsWith("." + dn) ||
      hostNoWww === dn || hostNoWww.endsWith("." + dn);
    if (ok && dn.length > bestLen) {
      best = d;
      bestLen = dn.length;
    }
  });

  return best;
}

function extractManualTargets_(ssOut, domains, rangeCur, sheetName, lastmodMap, categoryMap, bodyMap, picked){

  const wl = loadManualTargets_(ssOut);
  const shWL = wl.sh;
  const targets = wl.list;

  // 出力シートのヘッダだけは必ず作る
  writeExtractionOutput_(ssOut, sheetName, [], [], []);

  if (!targets.length) return;

  // URL→domain に振り分け
  const byDomain = {};
  const noDomain = [];

  targets.forEach(t => {
    const d = resolveDomainForUrl_(t.url, domains);
    if (!d) {
      noDomain.push(t);
      return;
    }
    (byDomain[d] = byDomain[d] || []).push(t);
  });

  // ステータス更新（軽量にセル更新）
  const nowStr = Utilities.formatDate(new Date(),"Asia/Tokyo","yyyy-MM-dd HH:mm:ss");

  // ドメイン不明
  noDomain.forEach(t => {
    try {
      shWL.getRange(t.row, 5).setValue("SKIP:ドメイン未判定（sc-domainに無い）");
    } catch(e){}
  });

  let table = [];

  Object.keys(byDomain).forEach(domain => {

    const cur = fetchPagesMetrics(domain, rangeCur);

    byDomain[domain].forEach(t => {

      const url = t.url;

      // ★レンジ横断の重複排除（Manualが最優先）
      if (picked && picked.has(url)) {
        try { shWL.getRange(t.row, 5).setValue("SKIP:他レンジで採用済"); } catch(e){}
        return;
      }

      // ★分類はHTMLキャッシュ基準。未分類は対象外（推定もしない）
      const catRaw = (categoryMap && (categoryMap[url] || categoryMap[normalizeUrl_(url)])) || "";
      const cat = String(catRaw || '').trim();

      if (cat === "NW" || cat === "ST" || isUnclassifiedCategory_(cat)) {
        try { shWL.getRange(t.row, 5).setValue("SKIP:分類未設定/除外（"+(cat||"空欄")+")"); } catch(e){}
        return;
      }

      const metrics = cur[url] || {clicks:0, impressions:0, ctr:0, position:100};
      const lastmod = (lastmodMap && (lastmodMap[url] || lastmodMap[normalizeUrl_(url)])) || "";

      // HTMLキャッシュ（本文/タイトル/メタ）が無い場合はプロンプトが無意味なので除外
      const tm = getCurrentTitleMeta(url);
      const body = getBodyFromMap_(bodyMap, url);
      if (!body && !tm.title && !tm.meta) {
        try { shWL.getRange(t.row, 5).setValue("SKIP:HTMLキャッシュ未登録"); } catch(e){}
        return;
      }

      table.push({
        domain,
        url,
        clicksCur: metrics.clicks,
        clicksPrev: null,
        impCur: metrics.impressions,
        impPrev: null,
        ctrCur: metrics.ctr,
        ctrPrev: null,
        posCur: metrics.position,
        posPrev: null,
        impDiff: null,
        // 手動キューは優先度を固定したいケースが多いので score は参考値（imp）
        score: metrics.impressions,
        lastmod,
        _category: cat,
        _memo: t.memo,
        _wlRow: t.row
      });

    });
  });

  // 手動キューは入力順が大事になりがちなので、ソートしない
  // （ENABLE_SORT_BY_SCORE は適用しない）

  const top = table.slice(0, getTopN_("Manual"));

  // ★採用URLを登録（レンジ横断で重複排除）
  if (picked){
    top.forEach(r => picked.add(normalizeUrl_(r.url)));
  }

  // クエリ取得＋プロンプト生成
  const detail = top.map(row => {

    const curQ = fetchQueries(row.domain, row.url, rangeCur);

    const joined = curQ.map(q => ({
      query: q.query,
      clicksCur: q.clicks,
      clicksPrev: null,
      impCur: q.impressions,
      impPrev: null,
      ctrCur: q.ctr,
      ctrPrev: null,
      posCur: q.pos,
      posPrev: null,
      posDiff: null
    }));

    const limit = QUERY_LIMIT.Manual || 20;

    const topByClicks = joined
      .sort((a,b)=> (b.impCur || 0) - (a.impCur || 0))
      .slice(0, limit);

    return { ...row, queries: joined, topByClicks, topByPosDrop: [] };
  });

  const prompts = detail.map(d => makePromptManual(d, bodyMap));

  // Manual 出力
  writeExtractionOutput_(ssOut, sheetName, top, detail, prompts);

  // 手動シートの最終出力・ステータス更新（採用分のみ）
  top.forEach(r => {
    try {
      shWL.getRange(r._wlRow, 4).setValue(nowStr);
      shWL.getRange(r._wlRow, 5).setValue("出力済");
    } catch(e){}
  });
}

/*************************************************************
 * レンジ作成
 *************************************************************/
function makeRangeWeekly_(latest){
  return {
    cur:  { start: shift(latest,-13),       end: latest },
    prev: { start: shift(latest,-27),       end: shift(latest,-14) }
  };
}

function makeRangeMonthly_(latest){
  return {
    cur:  { start: shift(latest,-27),       end: latest },
    prev: { start: shift(latest,-27-365),   end: shift(latest,-365) }
  };
}

function makeRangeLong_(latest){
  return {
    cur:  { start: shift(latest,-89),       end: latest },
    prev: { start: shift(latest,-89-365),   end: shift(latest,-365) }
  };
}

function makeRangeZeroClick_(base){
  // base を含む直近 ZEROCLICK_DAYS 日
  return {
    cur: { start: shift(base, -(ZEROCLICK_DAYS-1)), end: base }
  };
}

/*************************************************************
 * 出力クリア（4レンジ + Manual出力）
 * ★注意：作業対象_手動 は絶対にクリアしない
 *************************************************************/
function clearOutputFour_(ss){
  ["改善候補_Manual",
   "改善候補_Weekly","改善候補_Monthly","改善候補_LongTerm","改善候補_ZeroClick","改善候補_スニペット",
   "クエリ詳細","GPTプロンプト"
  ].forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clearContents();
  });

  ss.getSheetByName("クエリ詳細")
    .getRange(1,1,1,8)
    .setValues([["Range","URL","Query","ClicksCur","ClicksPrev","PosCur","PosPrev","posDiff"]]);

  ss.getSheetByName("GPTプロンプト")
    .getRange(1,1,1,6)
    .setValues([["Range","ページ分類","URL","更新日時","プロンプト(MonthlyとLongTermはChatGPT5.2Thinking Extendedを使用して下さい)","更新日時(最新)"]]);
}

/*************************************************************
 * プロンプト生成（Weekly / Monthly / LongTerm / ZeroClick / Manual）
 *************************************************************/

/*********************** Weekly ******************************/
function makePromptWeekly(d, bodyMap) {

  const pct = x => (x*100).toFixed(2) + "%";
  const { title: oldTitle, meta: oldMeta } = getCurrentTitleMeta(d.url);
  const diag = analyzeTitle_(oldTitle);
  const ctrGuideline = buildCtrGuideline_(diag);
  const tmGuard = buildTitleMetaGuardrails_(oldTitle, oldMeta, d.topByClicks || []);
  const bodyHtml = truncate_(getBodyFromMap_(bodyMap, d.url) || '', BODY_TEXT_MAX);

  const pageHelp = buildPageCategoryHelp_(d._category);

  return `以下のページについて、まず【本文】を確認し、
「どんな読者の、どの段階の悩み」に対して書かれているページかを把握してください。

そのうえで、直近14日間のデータから、
検索結果上で生じている「検索意図とのズレ」を特定し、
CTR改善に直結するタイトル案とメタディスクリプション案を作成してください。

【URL】${d.url}
${pageHelp}

【現在タイトル】${oldTitle}
【現在メタ】${oldMeta}

【短期変化（14日）】
CTR：${pct(d.ctrPrev)}→${pct(d.ctrCur)}（差${(d.ctrCur - d.ctrPrev).toFixed(4)}）
順位：${d.posPrev.toFixed(1)}→${d.posCur.toFixed(1)}（差${d.posDiff.toFixed(1)}）
クリック：${d.clicksPrev}→${d.clicksCur}
表示回数：${d.impPrev}→${d.impCur}（差${d.impCur - d.impPrev}）
スコア：${d.score.toFixed(2)}

【検索量の多いクエリ（14日）】
${d.topByClicks.map(q =>
`- ${q.query}｜表示${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${q.posCur.toFixed(1)}`
).join("\n")}

【順位急落クエリ（14日）】
${d.topByPosDrop.map(q =>
`- ${q.query}｜表示${q.impPrev}→${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${q.posPrev.toFixed(1)}→${q.posCur.toFixed(1)}（落差${q.posDiff.toFixed(1)}）`
).join("\n")}

【指示】
- 本文の内容と、表示されているクエリの検索意図が
  どの点で合っていないか／合っているかを明確に言語化すること
- タイトル・メタが本文内容を正しく代表しているかを評価すること
- 「誰に向けたページなのか」が一瞬で分かる表現に修正すること
- Weeklyのため、ページ構成の大改修は前提としない

【CTR改善の型（守る）】
${ctrGuideline}

【タイトル/メタ 禁則（該当がある場合のみ）】
${tmGuard}

【禁則ワード（シート連動）】
${buildForbiddenWordsBlock_()}

【出力】
1. 本文内容の要約（想定読者・悩み）
2. クエリと本文のズレの指摘
3. CTRが落ちた主因の仮説
4. タイトル案3つ（全角35字前後）
5. メタディスクリプション案3つ（全角90字前後）

【本文】
\`\`\`html
${bodyHtml}
\`\`\`
`;
}

/*********************** Monthly ******************************/
function makePromptMonthly(d, bodyMap) {

  const limit = QUERY_LIMIT.Monthly;
  const pct = x => (x*100).toFixed(2) + "%";
  const { title: oldTitle, meta: oldMeta } = getCurrentTitleMeta(d.url);
  const diag = analyzeTitle_(oldTitle);
  const ctrGuideline = buildCtrGuideline_(diag);
  const tmGuard = buildTitleMetaGuardrails_(oldTitle, oldMeta, d.topByClicks || []);

  const bodyHtml = truncate_(getBodyFromMap_(bodyMap, d.url) || '', BODY_TEXT_MAX);

  const pageHelp = buildPageCategoryHelp_(d._category);

  return `以下のページについて、まず【本文】を確認し、
「どんな塾なのか(対象/科目等)」「どの種類のページか（勉強法／用語解説／サービス紹介など）」を把握してください。

そのうえで、直近28日のデータを基に、
下記5つの改善タスクの中から “最も効果の高い1つだけ” を選択し、
具体的な改善案を提示してください。

【改善タスク候補（必ず1つだけ選ぶ）】
A. タイトル／メタの改善
B. 導入文（冒頭）の改善
C. H2/H3見出し構成の改善
D. 内容の不足部分の補完
E. 結論・要点の再整理

【URL】
${d.url}
${pageHelp}

【現在タイトル】
${oldTitle}

【現在メタ】
${oldMeta}

【中期変化（28日）】
CTR：${pct(d.ctrPrev)} → ${pct(d.ctrCur)}
順位：${d.posPrev.toFixed(1)} → ${d.posCur.toFixed(1)}
クリック：${d.clicksPrev} → ${d.clicksCur}
表示回数：${d.impPrev} → ${d.impCur}（差 ${d.impCur - d.impPrev}）
スコア：${d.score.toFixed(2)}

【主要クエリ（28日・表示回数上位${limit}件）】
${d.topByClicks.map(q =>
  `- ${q.query}｜表示${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${q.posCur.toFixed(1)}`
).join("\n")}

【出力】選択したタスクに対する具体的な改善案
1. 本文確認結果（ページ種別の把握）
2. 選ばれた改善タスクと簡単な理由（A〜Eから1つだけ）
3. 選択したタスクに対する具体的な改善案
   A:現状タイトル、更新後のタイトル案(全角35字程度)、現状メタディスクリプション、更新後のメタディスクリプション案(全角90字程度)
   B:現状の導入文（冒頭）、更新後の導入文（冒頭）
   C:現状のH2/H3見出し構成、更新後のH2/H3を明示した上で、Wordpressの投稿形式に沿った更新後の【本文】全文※タイトルメタは除く
   D:追加すべきブロック(見出し＋本文)を追加した、Wordpressの投稿形式に沿った【本文】全文※タイトルメタは除く
   E:更新後の【本文】全文※タイトルメタは除く
   ※更新後のタイトル/メタ/導入文/本文はコードブロックで囲うこと
   ※Wordpressの投稿を前提として、元ページとノリを合わせること。過度なデザインチェンジや大幅な内容改変はこれを固く禁じます。
   ※メタ発言(自身や私向けへのコメント)は本文でなくコメントに付記すること。

【CTR改善の型（守る）】
${ctrGuideline}

【タイトル/メタ 禁則（該当がある場合のみ）】
${tmGuard}

【禁則ワード（シート連動）】
${buildForbiddenWordsBlock_()}

【本文】
\`\`\`html
${bodyHtml}
\`\`\`
`;
}

/*********************** LongTerm ******************************/
function makePromptLongTerm(d, bodyMap) {

  const pct = x => (x*100).toFixed(2) + "%";
  const { title: oldTitle, meta: oldMeta } = getCurrentTitleMeta(d.url);
  const diag = analyzeTitle_(oldTitle);
  const ctrGuideline = buildCtrGuideline_(diag);
  const tmGuard = buildTitleMetaGuardrails_(oldTitle, oldMeta, d.topByClicks || []);

  const bodyHtml = truncate_(getBodyFromMap_(bodyMap, d.url) || '', BODY_TEXT_MAX);

  const pageHelp = buildPageCategoryHelp_(d._category);

  return `以下のページについて、まず【本文】を確認し、
「どんな塾なのか(対象/科目等)」「どの種類のページか（勉強法／用語解説／サービス紹介など）」を把握してください。
そのうえで、直近90日のデータを基に、
下記5つの改善タスクの中から “最も効果の高い1つだけ” を選択し、
具体的な改善案を提示してください。

【改善タスク候補（必ず1つだけ選ぶ）】
A. タイトル／メタの改善
B. 導入文（冒頭）の改善
C. H2/H3見出し構成の改善
D. 内容の不足部分の補完
E. 結論・要点の再整理

【URL】${d.url}
${pageHelp}

【現在タイトル】${oldTitle}
【現在メタ】${oldMeta}

【長期変化（90日）】
CTR：${pct(d.ctrPrev)}→${pct(d.ctrCur)}
順位：${d.posPrev.toFixed(1)}→${d.posCur.toFixed(1)}
クリック：${d.clicksPrev}→${d.clicksCur}
表示回数：${d.impPrev}→${d.impCur}（差${d.impCur - d.impPrev}）
スコア：${d.score.toFixed(2)}

【上位クエリ（90日）】
${d.topByClicks.map(q =>
`- ${q.query}｜表示${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${q.posCur.toFixed(1)}`
).join("\n")}

【改善方針】
- タイトルでは専門性・網羅性・具体性を強化し、検索結果内での位置付けを明確化
- 類似テーマのページが多い領域でも、独自の切り口や価値が直感的に伝わる構成にする
- メタディスクリプションでは、情報量・解決範囲・得られるメリットを端的に示す
- 長期的に読者ニーズとのギャップが生じている箇所を補完し、信頼性と有用性を高める

【出力】選択したタスクに対する具体的な改善案
1. 本文確認結果（ページ種別の把握）
2. 選ばれた改善タスクと簡単な理由（A〜Eから1つだけ）
3. 選択したタスクに対する具体的な改善案
   A:現状タイトル、更新後のタイトル案(全角35字程度)、現状メタディスクリプション、更新後のメタディスクリプション案(全角90字程度)
   B:現状の導入文（冒頭）、更新後の導入文（冒頭）
   C:現状のH2/H3見出し構成、更新後のH2/H3を明示した上で、Wordpressの投稿形式に沿った更新後の【本文】全文※タイトルメタは除く
   D:追加すべきブロック(見出し＋本文)を追加した、Wordpressの投稿形式に沿った【本文】全文※タイトルメタは除く
   E:更新後の【本文】全文※タイトルメタは除く
   ※更新後のタイトル/メタ/導入文/本文はコードブロックで囲うこと
   ※Wordpressの投稿を前提として、元ページとノリを合わせること。過度なデザインチェンジや大幅な内容改変はこれを固く禁じます。
   ※メタ発言(自身や私向けへのコメント)は本文でなくコメントに付記すること。

【CTR改善の型（守る）】
${ctrGuideline}

【タイトル/メタ 禁則（該当がある場合のみ）】
${tmGuard}

【禁則ワード（シート連動）】
${buildForbiddenWordsBlock_()}

【本文】
\`\`\`html
${bodyHtml}
\`\`\`
`;
}

/*********************** ZeroClick ******************************/
function makePromptZeroClick(d, bodyMap) {

  const limit = QUERY_LIMIT.ZeroClick || 30;
  const pct = x => (x*100).toFixed(2) + "%";
  const { title: oldTitle, meta: oldMeta } = getCurrentTitleMeta(d.url);

  const diag = analyzeTitle_(oldTitle);
  const ctrGuideline = buildCtrGuideline_(diag);
  const tmGuard = buildTitleMetaGuardrails_(oldTitle, oldMeta, d.topByClicks || []);
  const bodyHtml = truncate_(getBodyFromMap_(bodyMap, d.url) || '', BODY_TEXT_MAX);
  // ★ここで判定（本文先頭も使う）
  const rec = recommendZeroClickTask2_(d, oldTitle, oldMeta, bodyHtml);
  //const rec = recommendZeroClickTask_(d, oldTitle, oldMeta);
  const pageHelp = buildPageCategoryHelp_(d._category);

  return `以下のページについて、まず【本文】を確認し、
「どんな塾なのか(対象/科目等)」「どの種類のページか」「どんな読者の、どの段階の悩み」かを把握してください。

そのうえで、直近${ZEROCLICK_DAYS}日で「表示回数が一定以上あるのにクリックが0」である事実を踏まえ、
下記タスクA〜Eのうち、【推奨タスク】を“必ず1つだけ”実施してください（他はやらない）。

【改善タスク候補（必ず1つだけ）】
A. タイトル／メタの改善
B. 導入文（冒頭）の改善
C. H2/H3見出し構成の改善
D. 内容の不足部分の補完（例文/練習/FAQ 等）
E. 結論・要点の再整理

【推奨タスク】${rec.task}
【推奨理由】${rec.reason}
（参考）defRate=${(rec.defRate*100).toFixed(0)}% / issueCount=${rec.issueCount}

【URL】${d.url}
${pageHelp}

【現在タイトル】${oldTitle}
【現在メタ】${oldMeta}

【現状（直近${ZEROCLICK_DAYS}日）】
表示回数：${d.impCur}
クリック：${d.clicksCur}
CTR：${pct(d.ctrCur)}
順位：${Number(d.posCur).toFixed(1)}
（参考）分類：${d._category || ""}

【主要クエリ（表示回数上位${limit}件）】
${(d.topByClicks || []).slice(0, limit).map(q =>
`- ${q.query}｜表示${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${isFinite_(q.posCur) ? Number(q.posCur).toFixed(1) : ""}`
).join("\n")}

【CTR0専用ガイドライン（守る）】
${ctrGuideline}

【タイトル/メタ 禁則（該当がある場合のみ）】
${tmGuard}

【禁則ワード（シート連動）】
${buildForbiddenWordsBlock_()}

【出力】選択したタスクに対する具体的な改善案
1. 本文確認結果（ページ種別の把握）
2. 選ばれた改善タスクと理由（A〜Eから1つだけ）
3. 具体的な改善案（タスクに対応した形式で、コードブロックで囲う）

【本文】
\`\`\`html
${bodyHtml}
\`\`\`
`;
}

function stripTags_(html){
  return String(html||"")
    .replace(/<script[\s\S]*?<\/script>/gi,' ')
    .replace(/<style[\s\S]*?<\/style>/gi,' ')
    .replace(/<\/?[^>]+>/g,' ')
    .replace(/&nbsp;|&#160;/gi,' ')
    .replace(/&amp;/gi,'&')
    .replace(/&lt;/gi,'<')
    .replace(/&gt;/gi,'>')
    .replace(/\s+/g,' ')
    .trim();
}

function getIntroText_(bodyHtml, n){
  const t = stripTags_(bodyHtml);
  return t.slice(0, n || 240);
}

function looksGenericIntro_(intro){
  const s = String(intro||"").trim();
  // ありがちな“何も言ってない導入”を弾く
  return /^(こんにちは|はじめまして|お世話になっております|今回は|本記事では|この記事では|この記事を読むと|まずは|結論)/.test(s);
}
function recommendZeroClickTask2_(d, oldTitle, oldMeta, bodyHtml){
  const imp = Number(d.impCur || 0);
  const pos = Number(d.posCur || 100);
  const cat = String(d._category || "").trim(); // SP/MC/SC...

  const qs = (d.topByClicks || []).slice(0, 20).map(x => String(x.query || ""));
  const defCount = qs.filter(isSerpAnswerQuery_).length;
  const defRate = qs.length ? (defCount / qs.length) : 0;

  const tm = analyzeTitleMeta_(oldTitle, oldMeta, d.topByClicks || []);
  const issueCount = (tm.issues || []).length;

  const title = String(oldTitle || "");
  const meta  = String(oldMeta || "");
  const titleLen = [...title].length;
  const metaLen  = [...meta].length;

  const intro = getIntroText_(bodyHtml, 240);
  const headGeneric = looksGenericIntro_(intro);

  const repToken = (d.topByClicks && d.topByClicks.length)
    ? extractRepresentativeToken_(d.topByClicks[0].query)
    : "";

  const headHasToken  = repToken ? intro.includes(repToken) : true;
  const titleHasToken = repToken ? title.includes(repToken) : true;
  const metaHasToken  = repToken ? meta.includes(repToken)  : true;

  // 1) SERP完結寄り（意味/とは/読み方 等が多い）→ “押す理由”を作る D
  if (defRate >= 0.6){
    return { task:"D", reason:"定義/意味系が多くSERP完結寄り→例/練習/FAQで“押す理由”を作る", defRate, issueCount };
  }

  // 2) IMPが小さい → まず土台を増やす（D or C）
  if (imp < 50){
    if (pos <= 10) return { task:"D", reason:"上位だがIMPが少ない→不足補完で対象KW帯を拡張", defRate, issueCount };
    if (pos <= 20) return { task:"C", reason:"10位圏外→H2/H3で狙いを明確化し順位とIMPを上げる", defRate, issueCount };
    return { task:"D", reason:"順位も低い→不足補完で土台づくり", defRate, issueCount };
  }

  // 3) 上位なのにクリック0 → 原則A（Bは例外）
  if (pos <= 10){
    // SPは特に“スニペット勝負”なのでBを原則出さない
    if (cat === "SP"){
      return { task:"A", reason:"SPで上位なのにクリック0→スニペット（タイトル/メタ）最優先", defRate, issueCount };
    }

    // 「タイトル/メタが強い」かの簡易判定
    const titleMetaStrong =
      (issueCount === 0) &&
      (titleLen >= 28) && (metaLen >= 80) &&
      titleHasToken && metaHasToken;

    // ★Bを出すのは“導入が弱い”と判定できた時だけ
    if (titleMetaStrong && (headGeneric || !headHasToken)){
      return { task:"B", reason:"タイトル/メタは概ね良いが、導入が一般論/挨拶寄り・主要語が薄い→冒頭で刺す", defRate, issueCount };
    }

    // それ以外はAへ寄せる（これでB激減）
    return { task:"A", reason:"上位表示なのにクリック0→まずスニペット（タイトル/メタ）を強化", defRate, issueCount };
  }

  // 4) 11〜20位：構造
  if (pos <= 20){
    return { task:"C", reason:"10位圏外→構造（H2/H3）で順位を上げる", defRate, issueCount };
  }

  // 5) それ以下：要点再整理
  return { task:"E", reason:"順位が低い→要点/狙いKW帯の再整理", defRate, issueCount };
}
/*********************** Manual ******************************/
function makePromptManual(d, bodyMap){

  const limit = QUERY_LIMIT.Manual || 30;
  const pct = x => (x*100).toFixed(2) + "%";
  const { title: oldTitle, meta: oldMeta } = getCurrentTitleMeta(d.url);

  const diag = analyzeTitle_(oldTitle);
  const ctrGuideline = buildCtrGuideline_(diag);
  const tmGuard = buildTitleMetaGuardrails_(oldTitle, oldMeta, d.topByClicks || []);
  const bodyHtml = truncate_(getBodyFromMap_(bodyMap, d.url) || '', BODY_TEXT_MAX);

  const pageHelp = buildPageCategoryHelp_(d._category);
  const memo = String(d._memo || "").trim();

  return `以下のページは【手動キュー】（作業対象として固定）です。
まず【本文】を確認し、「どんな塾なのか(対象/科目等)」「どの種類のページか」「どんな読者の、どの段階の悩み」かを把握してください。

そのうえで、直近${ZEROCLICK_DAYS}日相当のスナップショットを参考に、
下記タスクA〜Eの中から “最も効果の高い1つだけ” を選んで改善案を提示してください。

【改善タスク候補（必ず1つだけ選ぶ）】
A. タイトル／メタの改善
B. 導入文（冒頭）の改善
C. H2/H3見出し構成の改善
D. 内容の不足部分の補完（例文/練習/FAQ 等）
E. 結論・要点の再整理

【URL】${d.url}
${pageHelp}
${memo ? `\n【メモ（手動）】${memo}\n` : ""}

【現在タイトル】${oldTitle}
【現在メタ】${oldMeta}

【現状（スナップショット）】
表示回数：${d.impCur}
クリック：${d.clicksCur}
CTR：${pct(d.ctrCur)}
順位：${Number(d.posCur).toFixed(1)}
更新日時（キャッシュ）：${d.lastmod || ""}

【主要クエリ（表示回数上位${limit}件）】
${(d.topByClicks || []).slice(0, limit).map(q =>
`- ${q.query}｜表示${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${isFinite_(q.posCur) ? Number(q.posCur).toFixed(1) : ""}`
).join("\n")}

【CTR改善の型（守る）】
${ctrGuideline}

【タイトル/メタ 禁則（該当がある場合のみ）】
${tmGuard}

【禁則ワード（シート連動）】
${buildForbiddenWordsBlock_()}

【出力】選択したタスクに対する具体的な改善案
1. 本文確認結果（ページ種別の把握）
2. 選ばれた改善タスクと簡単な理由（A〜Eから1つだけ）
3. 選択したタスクに対する具体的な改善案（コードブロックで囲う）
   A:タイトル案/メタ案
   B:導入文差し替え
   C/D/E:本文（必要箇所の追記・再構成）

【本文】
\`\`\`html
${bodyHtml}
\`\`\`
`;
}

/*********************** Snippet（スニペット不一致） ★追加 ******************************/
function makePromptSnippet(d, bodyMap){

  const limit = QUERY_LIMIT.Snippet || 50;

  // HTMLキャッシュ由来（H1含む）
  const tmh = getCurrentTitleMetaH1Entry_(d.url);
  const oldTitle = String(tmh.title || "");
  const oldMeta  = String(tmh.meta  || "");
  const oldH1    = String(tmh.h1    || "");

  // 本文（本文Map優先、無ければ entryBodyHt）
  const rawHtml = getBodyFromMap_(bodyMap, d.url) || tmh.entry || "";
  const bodyHtml = truncate_(rawHtml, BODY_TEXT_MAX);
  const intro = getIntroForSnippet_(rawHtml);

  const pageHelp = buildPageCategoryHelp_(d._category);

  return `以下のページは「スニペット整合性」（本文と title/meta/H1/冒頭文）が弱い可能性があります。
まず【本文】を確認し、想定読者・悩み・解決範囲を把握してください。

そのうえで、GSCの主要クエリ（表示回数上位）を根拠に、
「検索意図に対して、タイトル/メタ/H1/冒頭${SNIPPET_INTRO_CHARS}字がズレている点」を特定し、
ズレを解消するための修正案（タイトル/メタ/H1/冒頭）を提示してください。
※本文の“大改修”は前提にしません（必要なら追記ブロックは最大1つまで）。

【URL】${d.url}
${pageHelp}

【現状（スニペット要素）】
- タイトル：${oldTitle}
- メタ：${oldMeta}
- H1：${oldH1}
- 冒頭${SNIPPET_INTRO_CHARS}字：${intro}

【GSC主要クエリ（表示回数上位${limit}件）】
${(d.topByClicks || []).slice(0, limit).map(q =>
`- ${q.query}｜表示${q.impCur}｜CTR ${(q.impCur ? (q.clicksCur / q.impCur * 100).toFixed(2) : "0.00")}%｜順位${isFinite_(q.posCur) ? Number(q.posCur).toFixed(1) : ""}`
).join("\n")}

【参考：代表語（自動抽出）】${d._token || ""}

【出力（必ずこの順）】
1) 本文の要約（想定読者・悩み・解決）
2) 検索意図の整理（主要クエリから2〜3パターン）
3) 不一致の指摘（タイトル/メタ/H1/冒頭のどこがズレているか）
4) 修正方針（“どの意図に寄せるか”を1つに決める）
5) 修正案（すべてコードブロック）
   - タイトル案3つ（全角35字前後＋56字以上の網羅型を1つ）
   - メタ案3つ（全角90字前後）
   - H1案2つ
   - 冒頭${SNIPPET_INTRO_CHARS}字（2案：<p>で開始）

【禁則ワード（シート連動）】
${buildForbiddenWordsBlock_()}

【本文】
\`\`\`html
${bodyHtml}
\`\`\`
`;
}

/*************************************************************
 * Snapshotシートから「現在のタイトルとメタ」を取得（URL差異に耐性）
 *************************************************************/
function getCurrentTitleMeta(url) {

  const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  const sh = ss.getSheetByName('HTMLキャッシュ');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { title: "", meta: "" };

  const target = normalizeUrl_(url);
  const values = sh.getRange(2, 1, lastRow - 1, 6).getValues();

  for (let i = 0; i < values.length; i++) {
    const uRaw  = values[i][0];
    if (!uRaw) continue;

    const u = normalizeUrl_(String(uRaw));
    if (u !== target) continue;

    const title = values[i][4];
    const meta  = values[i][5];
    return { title: title || "", meta: meta || "" };
  }

  return { title: "", meta: "" };
}

/************************************************************
 * URL →（シートA列で一致行検索）→ H列本文 → 文末CTA除外 → 本文HTML返却
 ************************************************************/
const BODY_SHEET_ID  = CACHE_SHEET_ID;
const BODY_SHEET_GID = 1666145974;

const COL_URL  = 1; // A
const COL_BODY = 8; // H

function getBodyHtmlWithoutTailCtaFromSheet(url, opts) {
  const body = getBodyHtmlByUrlFromSheet_(url);
  if (!body) return '';
  return removeTailCtaFromBodyHtml(body, opts);
}

function getRawBodyHtmlFromSheet(url) {
  return getBodyHtmlByUrlFromSheet_(url) || '';
}

function getBodyHtmlByUrlFromSheet_(url) {
  const targetUrl = String(url || '').trim();
  if (!targetUrl) return '';

  const ss = SpreadsheetApp.openById(BODY_SHEET_ID);
  const sh = ss.getSheets().find(s => s.getSheetId() === BODY_SHEET_GID);
  if (!sh) throw new Error('対象シートが見つかりません（gid=' + BODY_SHEET_GID + '）');

  const lastRow = sh.getLastRow();
  if (lastRow < 1) return '';

  const finder = sh.getRange(1, COL_URL, lastRow, 1)
    .createTextFinder(targetUrl)
    .matchEntireCell(true)
    .findNext();

  if (!finder) return '';

  const row = finder.getRow();
  const body = sh.getRange(row, COL_BODY).getValue();
  return String(body || '');
}

/* ===================== 文末CTA除外ロジック ===================== */

function removeTailCtaFromBodyHtml(bodyHtml, opts) {
  return removeTailCtaBlocks_(String(bodyHtml || ''), opts);
}

function removeTailCtaBlocks_(bodyHtml, opts) {
  let html = String(bodyHtml || '');
  if (!html) return '';

  const o = opts || {};
  const ctaClassTokens = o.ctaClassTokens || ['veu_cta', 'veu-cta', 'vk_cta', 'vk-cta'];
  const minTailRatio = (typeof o.minTailRatio === 'number') ? o.minTailRatio : 0.65;
  const tailWindowChars = (typeof o.tailWindowChars === 'number') ? o.tailWindowChars : 6000;

  const SOCIALSET_DIV_RE =
    /<div\b[^>]*class=["'][^"']*\bveu_socialSet\b[^"']*["'][\s\S]*?<\/div>\s*(?:<!--[\s\S]*?-->)?/gi;

  html = html.replace(SOCIALSET_DIV_RE, '');

  const tokenAlt = ctaClassTokens.map(t => '\\b' + escapeRegExp_(t) + '\\b').join('|');

  const CTA_BLOCK_RE = new RegExp(
    '<([a-z0-9]+)\\b[^>]*class=["\'][^"\']*(?:' + tokenAlt + ')[^"\']*["\'][\\s\\S]*?<\\/\\1>',
    'gi'
  );

  const len = html.length;
  const tailStart = Math.max(Math.floor(len * minTailRatio), len - tailWindowChars);

  const hits = [];
  let m;
  while ((m = CTA_BLOCK_RE.exec(html)) !== null) {
    const idx = m.index;
    if (idx >= tailStart) hits.push({ start: idx, end: idx + m[0].length });
  }
  if (!hits.length) return html;

  hits.sort((a, b) => b.start - a.start);
  let out = html;
  for (let i = 0; i < hits.length; i++) {
    out = out.slice(0, hits[i].start) + out.slice(hits[i].end);
  }
  return out;
}

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/* ===================== 本文Map ===================== */

const BODY_TEXT_MAX = 100000;

function truncate_(s, maxLen) {
  const str = String(s || '');
  if (!maxLen || str.length <= maxLen) return str;
  return str.slice(0, maxLen) + '\n...(truncated)';
}

function loadBodyMapCleaned_() {
  const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  const sh = ss.getSheets().find(s => s.getSheetId() === BODY_SHEET_GID);
  if (!sh) throw new Error('本文シートが見つかりません（gid=' + BODY_SHEET_GID + '）');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const values = sh.getRange(2, 1, lastRow - 1, 8).getValues();

  const map = {};
  for (let i = 0; i < values.length; i++) {
    const url = values[i][0];
    const body = values[i][7];
    if (!url) continue;

    const raw = String(url);
    const norm = normalizeUrl_(raw);

    const cleaned = removeTailCtaFromBodyHtml(String(body || ''), {
      minTailRatio: 0.65,
      tailWindowChars: 6000
    });

    map[raw] = cleaned;
    if (norm !== raw) map[norm] = cleaned;
  }
  return map;
}

function parseLastmodToDate_(v){
  if (v instanceof Date && !isNaN(v)) return v;

  const s = String(v || '').trim();
  if (!s) return null;

  // まずは素直にDateパース
  let d = new Date(s);
  if (!isNaN(d)) return d;

  // yyyy/MM/dd or yyyy-MM-dd を手動パース（念のため）
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m){
    d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (!isNaN(d)) return d;
  }
  return null;
}

function isLastmodTooRecent_(lastmod, days){
  if (!days || days <= 0) return false;
  const d = parseLastmodToDate_(lastmod);
  if (!d) return false; // lastmod不明は「除外しない」運用にしておくのが安全

  const now = new Date();
  now.setHours(0,0,0,0);
  const cutoff = new Date(now.getTime() - days * DAY);
  return d >= cutoff; // cutoff以降（=最近更新）なら除外
}

function shouldExcludeByLastmod_(lastmod){
  return isLastmodTooRecent_(lastmod, EXCLUDE_RECENT_LASTMOD_DAYS);
}

function hasValidLastmod_(v){
  if (v instanceof Date && !isNaN(v)) return true;
  const s = String(v || '').trim();
  return s !== '';
}

/*************************************************************
 * URL正規化 & Map取得補助
 *************************************************************/
function normalizeUrl_(u) {
  return String(u || '').trim();
}


function getBodyFromMap_(bodyMap, url) {
  if (!bodyMap) return "";
  const raw = String(url || "");
  const norm = normalizeUrl_(raw);
  return bodyMap[raw] || bodyMap[norm] || "";
}

function isFinite_(v) {
  return typeof v === 'number' && isFinite(v);
}

/*************************************************************
 * タイトル診断（CTR0対策の自動分岐に使う）
 *************************************************************/
function analyzeTitle_(title) {
  const t = String(title || "");
  const hasPipe = t.includes("｜");
  const hasQuestion = /[?？]/.test(t);
  const hasLessonCode = /(Lesson|STAGE|Stage|Progress|Book|New\s*Treasure|stage\d+|lesson\d+)/i.test(t);
  const hasOnlyTextbookLike = hasLessonCode && !/(攻略|使い方|覚え方|違い|見分け方|コツ|練習|例文|図解|問題|解説)/.test(t);
  const len = [...t].length; // 日本語の概算文字数
  return { hasPipe, hasQuestion, hasLessonCode, hasOnlyTextbookLike, len };
}

/*************************************************************
 * CTR0向け「強制ガイドライン」文字列を生成
 *************************************************************/
function buildCtrGuideline_(diag) {
  const notes = [];

  if (diag.hasQuestion) {
    notes.push("- タイトルの「？」は削除し、断定形にする（例：～とは？→～の意味と使い方）");
  } else {
    notes.push("- タイトルに「？」は付けない（断定形でメリットを明示）");
  }

  if (diag.hasOnlyTextbookLike) {
    notes.push("- Lesson/Stage/Progress等の“教科書コード”をタイトル先頭に置かない（末尾の補足へ回す）");
  } else if (diag.hasLessonCode) {
    notes.push("- Lesson/Stage/Progress等の情報は残してよいが、末尾の括弧補足に寄せる");
  }

  notes.push("- 可能な限り「｜」区切りの3要素で構造化する：");
  notes.push("  ①【主要KW】 ②悩み/ベネフィット（例：よくあるミス・例文・練習問題） ③対象（例：中1英語／中学受験算数）＋教材情報は末尾補足");
  notes.push("- タイトル案は“短さ最優先”にしない。56字以上の網羅型も必ず1つ作る");
  notes.push("- 46〜55字の中途半端帯に寄りすぎない（必要なら短くするか、逆に特盛りにする）");

  return notes.join("\n");
}

function isNonHtmlResourceUrl_(url) {
  const u = String(url || '').trim();
  if (!u) return false;

  const noQ = u.split('#')[0].split('?')[0];
  const m = noQ.match(/\.([a-z0-9]+)$/i);
  if (!m) return false;

  const ext = m[1].toLowerCase();
  return EXCLUDE_EXTS.includes(ext);
}

/*************************************************************
 * タイトル＋メタ診断（ダメ更新パターンの禁則ガード生成用）
 *************************************************************/
function analyzeTitleMeta_(title, meta, topByClicks) {
  const issues = [];

  const t = String(title || "");
  const m = String(meta  || "").trim();
  const td = analyzeTitle_(t);

  if (/[?？]/.test(t)) {
    issues.push("タイトルに「？」を付けない（断定形にする）");
  }

  if (td.hasOnlyTextbookLike) {
    issues.push("Lesson/Stage/Progress等の“教科書コード”をタイトル先頭に置かない（末尾補足へ）");
  }

  if (/^(結論[:：]|【結論】|1分で|１分で)/.test(m)) {
    issues.push("メタ冒頭に「結論：」「1分で」等の見出し語を置かず、自然文で開始する");
  }

  if (/(この記事では|この記事を読むと|以下のページ|このページでは)/.test(m)) {
    issues.push("メタで「この記事では」等のメタ説明を避け、悩み→解決→得られる結果を先に書く");
  }

  if (topByClicks && topByClicks.length) {
    const q0 = String(topByClicks[0].query || "").trim();
    const key = extractRepresentativeToken_(q0);

    if (key && !t.includes(key) && !m.includes(key)) {
      issues.push(`主要クエリ由来の代表語「${key}」がタイトル/メタに見当たらないため、必ず含める`);
    }
  }

  if (td.len >= 46 && td.len <= 55) {
    issues.push("タイトルが46〜55字の中途半端帯。短くするか、56字以上の網羅型に寄せる");
  }

  return { issues, titleDiag: td };
}

function extractRepresentativeToken_(query) {
  const q = String(query || "").trim();
  if (!q) return "";

  const alnum = q.match(/[A-Za-z][A-Za-z0-9\-\+\/]*/g);
  if (alnum && alnum.length) {
    alnum.sort((a, b) => b.length - a.length);
    return alnum[0] || "";
  }

  const cleaned = q
    .replace(/[（）\(\)\[\]【】「」『』"'’`]/g, "")
    .replace(/[、。，．・\/\-\_\+\|\:：]/g, " ");

  const parts = cleaned.split(/\s+/).filter(Boolean);
  if (!parts.length) return "";

  const STOP = new Set([
    "とは","意味","使い方","覚え方","違い","見分け方","コツ","例文","練習","問題","解説","まとめ","一覧",
    "中学","中1","中2","中3","高校","小学生","中学受験","算数","国語","英語","社会","理科","大学受験","塾","講座"
  ]);

  const cand = parts.filter(p => ([...p].length >= 2) && !STOP.has(p));
  const pick = (cand.length ? cand : parts)
    .sort((a, b) => [...b].length - [...a].length)[0];

  return pick || "";
}

function buildTitleMetaGuardrails_(title, meta, topByClicks) {
  const a = analyzeTitleMeta_(title, meta, topByClicks);
  if (!a.issues.length) return "- （追加なし）";
  return a.issues.map(s => "- " + s).join("\n");
}

/*************************************************************
 * ZeroClick：SERP完結っぽいクエリ判定
 *************************************************************/
function isSerpAnswerQuery_(q){
  const s = String(q || "");
  return /(とは|意味|何|なに|読み方|定義|換算|単位|英語で|発音)/.test(s);
}

/*************************************************************
 * ZeroClick：数値で推奨タスクを決める（A〜E）
 * - imp / pos / defRate で分岐
 *************************************************************/
function recommendZeroClickTask_(d, oldTitle, oldMeta){
  const imp = Number(d.impCur || 0);
  const pos = Number(d.posCur || 100);

  const qs = (d.topByClicks || []).slice(0, 20).map(x => String(x.query || ""));
  const defCount = qs.filter(isSerpAnswerQuery_).length;
  const defRate = qs.length ? (defCount / qs.length) : 0;

  const tm = analyzeTitleMeta_(oldTitle, oldMeta, d.topByClicks || []);
  const issueCount = (tm.issues || []).length;

  if (imp < 50){
    if (pos <= 10){
      if (defRate >= 0.5){
        return { task: "D", reason: "上位だが表示が少ない＋定義/意味系が多くSERP完結寄り→例/練習/FAQで長尾を拾いIMPを増やす", defRate, issueCount };
      }
      return { task: "D", reason: "上位だが表示が少ない→不足補完（例/練習/FAQ/手順/比較）で検索面を広げIMPを増やす", defRate, issueCount };
    }

    if (pos <= 20){
      return { task: "C", reason: "表示が少なく10位圏外→H2/H3構造を再設計して狙いKW帯を明確化し、順位とIMPを上げる", defRate, issueCount };
    }

    return { task: "D", reason: "表示が少なく順位も低い→不足補完（例/FAQ/比較/練習）で対象KW帯を増やしIMPの土台を作る", defRate, issueCount };
  }

  if (pos <= 10){
    if (defRate >= 0.5){
      return { task: "D", reason: "定義/意味系が多くSERP完結寄り→例/練習/FAQを追加", defRate, issueCount };
    }
    if (issueCount >= 1){
      return { task: "A", reason: "上位表示なのにクリック0→訴求（タイトル/メタ）を最優先", defRate, issueCount };
    }
    return { task: "B", reason: "タイトル/メタに致命傷が少ない→冒頭（導入文）で刺す", defRate, issueCount };
  }

  if (pos <= 20){
    return { task: "C", reason: "10位圏外→構造（H2/H3）から順位を上げる", defRate, issueCount };
  }

  return { task: "E", reason: "順位が低い→結論・要点・狙いKW帯の再整理", defRate, issueCount };
}


// 実行中キャッシュ（毎回スプレッドシートを開かない）
let __FORBIDDEN_WORDS_CACHE__ = null;

/**
 * 禁則ワード一覧を外部スプレッドシートの「禁則ワード」シート A列から取得
 * - 空欄除外 / trim / 重複除去
 * - 先頭セルが「禁則ワード」等の見出しっぽい場合は除外
 */
function loadForbiddenWords_() {
  try {
    const ss = SpreadsheetApp.openById(FORBIDDEN_WORDS_SS_ID);
    const sh = ss.getSheetByName(FORBIDDEN_WORDS_SHEET);
    if (!sh) return [];

    const lastRow = sh.getLastRow();
    if (lastRow < 1) return [];

    const vals = sh.getRange(1, FORBIDDEN_WORDS_COL, lastRow, 1).getValues()
      .map(r => String(r[0] || '').trim())
      .filter(s => s);

    if (!vals.length) return [];

    // 見出し除外（A1が「禁則ワード」「ワード」「NGワード」などの場合）
    const head = vals[0];
    const isHeaderLike = /^(禁則ワード|NGワード|禁止ワード|ワード|一覧)$/i.test(head);
    const body = isHeaderLike ? vals.slice(1) : vals;

    // 重複除去（順序維持）
    const seen = new Set();
    const out = [];
    body.forEach(w => {
      if (!w) return;
      if (seen.has(w)) return;
      seen.add(w);
      out.push(w);
    });

    return out;

  } catch (e) {
    // 読み込み失敗でも処理は止めない（プロンプトから禁則ワード欄が空になるだけ）
    return [];
  }
}

/**
 * キャッシュ付き取得
 */
function getForbiddenWords_() {
  if (__FORBIDDEN_WORDS_CACHE__ !== null) return __FORBIDDEN_WORDS_CACHE__;
  __FORBIDDEN_WORDS_CACHE__ = loadForbiddenWords_();
  return __FORBIDDEN_WORDS_CACHE__;
}

/**
 * プロンプト埋め込み用の禁則ワードブロック生成
 */
function buildForbiddenWordsBlock_() {
  const words = getForbiddenWords_() || [];
  if (!words.length) return "（禁則ワード：設定なし / 取得できませんでした）";

  const max = FORBIDDEN_WORDS_MAX_IN_PROMPT || 200;
  const sliced = words.slice(0, max);
  const rest = words.length - sliced.length;

  const list = sliced.map(w => `- ${w}`).join("\n");
  return [
    "以下の【禁則ワード】は、タイトル/メタ/導入文/見出し/本文のいずれにも原則使用しないでください。",
    "言い換えが必要な場合は、自然な日本語に置換してください。",
    "",
    "【禁則ワード】",
    list,
    rest > 0 ? `\n…ほか ${rest} 件（省略）` : ""
  ].join("\n");
}
