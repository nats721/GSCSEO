/*************************************************************
 * GSCキーワードベース（SearchConsole Advanced Service 不要版）
 * - SearchConsole（高度なGoogleサービス）を使わず、
 *   UrlFetch + ScriptApp.getOAuthToken() で Search Console API を叩きます。
 *
 * Stage1:
 *   各ドメインの上位クエリを収集 → 意図スコア付与 → 「KW候補」へ出力
 * Stage2:
 *   「KW候補」で 対象=TRUE のクエリを Query×Page で展開 → 「KW→ページ」「KWサマリ」「カニバリ候補」
 *
 * 前提:
 * - ドメイン一覧は BASE_SHEET_ID の「設定」シート A列（ドメイン名）にある（ヘッダ: ドメイン名 / サイトマップ）
 * - 出力先は KEYWORD_SHEET_ID（別ブック）に作成される
 *************************************************************/

/** ===== あなたの環境に合わせてここだけ確認 ===== */
//const BASE_SHEET_ID   = "1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs"; // 各ブランドサイト管理＿メイン（読み取り）
const KEYWORD_SHEET_ID = "1LjGXklKL_NL5Q9rGLjgRJEazDKz4dV6-PrrynQ3dmLw"; // キーワードベース出力先（このスクリプトの出力）
const DOMAIN_SETTINGS_SHEET = "設定"; // BASE側
//const CACHE_SHEET_ID   = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o'; // キャッシュ用ブック

const CACHE_SHEET_NAME = 'HTMLキャッシュ';

// （任意）ページ分類を付与したい場合：HTMLキャッシュのSS（URL→分類）を参照
// ※不要なら空文字にしてください（分類列は空になります）
//const CACHE_SHEET_ID   = "1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o";
//const CACHE_SHEET_NAME = "HTMLキャッシュ"; // URL, 分類, lastmod,... が入っている想定（A:URL, B:分類）

/** ===== 出力シート名 ===== */
const SH_KW_SETTINGS  = "KW設定";
const SH_KW_CAND      = "KW候補";
const SH_KW_PAGES     = "KW→ページ";
const SH_KW_SUMMARY   = "KWサマリ";
const SH_KW_CANNIBAL  = "カニバリ候補";
const SH_LOG          = "ログ";

/** ===== デフォルト設定 ===== */
const KW_DEFAULT = {
  DAYS: 90,                // 収集期間（日）
  LAG_DAYS: 3,             // GSC遅延見込み（今日から引く）
  MIN_IMP: 10,             // Stage1: 最低表示回数
  MAX_QUERIES_PER_DOMAIN: 2500, // Stage1: 1ドメインあたり最大クエリ
  MAX_KEYWORDS_TOTAL: 8000,     // Stage1: 全体上限（保険）
  INTENT_MIN_SCORE: 4,     // Stage1: これ未満は原則落とす（ただし手動で対象TRUEにすればStage2可能）
  MAX_KEYWORDS_TO_EXPAND: 300,  // Stage2: 1回の実行で展開するKW数（タイムアウト回避）
  CANNIBAL_MIN_IMP: 50,    // カニバリ判定の最低表示回数（サマリ側）
  TOP_PAGES_PER_KW: 20     // 1KWあたり出力するページ数上限
};

/** ===== onOpen メニュー ===== */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("GSCキーワード")
      .addItem("（初回）KW設定シート作成/更新", "SETUP_KW_SETTINGS")
      .addSeparator()
      .addItem("Stage1: クエリ収集 → KW候補作成", "RUN_STAGE1_QUERYLIST_ALL_DOMAINS")
      .addItem("Stage2: 対象KWを Query×Page 展開", "RUN_STAGE2_EXPAND_SELECTED_KEYWORDS")
      .addSeparator()
      .addItem("Stage2: 進捗リセット（再実行用）", "KW_RESET_STAGE2_CURSOR")
      .addToUi();
  } catch (e) {
    // UI無し環境でも落とさない
  }
}

/** ===== ログ ===== */
function log_(level, msg, obj) {
  const ss = SpreadsheetApp.openById(KEYWORD_SHEET_ID);
  const sh = ss.getSheetByName(SH_LOG) || ss.insertSheet(SH_LOG);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,4).setValues([["ts","level","msg","json"]]);
  const ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
  sh.appendRow([ts, level, String(msg || ""), obj ? JSON.stringify(obj) : ""]);
}

/** ===== 日付ユーティリティ ===== */
const DAY_MS = 86400 * 1000;
function shiftDate_(d, days) { return new Date(d.getTime() + days * DAY_MS); }
function fmt_(d) { return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd"); }

/** ===== 正規化 ===== */
function normStr_(s){ return String(s || "").trim(); }
function normUrl_(u){ return String(u || "").trim(); }

/** ===== 設定シート（KW設定）作成 ===== */
function SETUP_KW_SETTINGS() {
  const ss = SpreadsheetApp.openById(KEYWORD_SHEET_ID);
  let sh = ss.getSheetByName(SH_KW_SETTINGS);
  if (!sh) sh = ss.insertSheet(SH_KW_SETTINGS);

  sh.clearContents();
  sh.getRange(1,1,1,7).setValues([[
    "key","value","memo",
    "includeWords(任意)","excludeWords(任意)","includeRegex(任意)","excludeRegex(任意)"
  ]]);

  const rows = [
    ["DAYS", KW_DEFAULT.DAYS, "収集期間（日）", "", "", "", ""],
    ["LAG_DAYS", KW_DEFAULT.LAG_DAYS, "GSC遅延見込み（日）", "", "", "", ""],
    ["MIN_IMP", KW_DEFAULT.MIN_IMP, "Stage1 最低表示回数", "", "", "", ""],
    ["MAX_QUERIES_PER_DOMAIN", KW_DEFAULT.MAX_QUERIES_PER_DOMAIN, "Stage1 1ドメイン最大クエリ数", "", "", "", ""],
    ["MAX_KEYWORDS_TOTAL", KW_DEFAULT.MAX_KEYWORDS_TOTAL, "Stage1 全体上限（保険）", "", "", "", ""],
    ["INTENT_MIN_SCORE", KW_DEFAULT.INTENT_MIN_SCORE, "Stage1 意図スコア閾値（未満は原則落とす）", "", "", "", ""],
    ["MAX_KEYWORDS_TO_EXPAND", KW_DEFAULT.MAX_KEYWORDS_TO_EXPAND, "Stage2 1回の展開KW数（タイムアウト回避）", "", "", "", ""],
    ["TOP_PAGES_PER_KW", KW_DEFAULT.TOP_PAGES_PER_KW, "Stage2 1KWあたり出力ページ上限", "", "", "", ""],
    ["CANNIBAL_MIN_IMP", KW_DEFAULT.CANNIBAL_MIN_IMP, "カニバリ判定の最低表示回数", "", "", "", ""],
  ];

  // 例：CV寄りの含有語（必要なら編集）
  const includeWords = [
    "塾","個別","個別指導","家庭教師","オンライン","講座","コース","料金","費用","月謝","体験","無料体験",
    "入塾","相談","問い合わせ","評判","口コミ","比較","対策","講習","サピックス","早稲アカ","四谷","グノ","浜学園"
  ];
  const excludeWords = [
    "とは","意味","読み方","定義","公式","答え","解き方","単位","何","なに","wiki"
  ];

  // include/exclude を 1行に詰める（簡単運用用）
  rows[0][3] = includeWords.join(" / ");
  rows[0][4] = excludeWords.join(" / ");

  sh.getRange(2,1,rows.length,7).setValues(rows);

  // 見た目
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,7);

  log_("INFO", "KW設定シートを作成/更新しました", null);
}

/** ===== KW設定読み取り ===== */
function getKwConfig_() {
  const ss = SpreadsheetApp.openById(KEYWORD_SHEET_ID);
  const sh = ss.getSheetByName(SH_KW_SETTINGS);
  const cfg = Object.assign({}, KW_DEFAULT);

  let includeWords = [];
  let excludeWords = [];
  let includeRegex = [];
  let excludeRegex = [];

  if (!sh || sh.getLastRow() < 2) {
    return { cfg, includeWords, excludeWords, includeRegex, excludeRegex };
  }

  const vals = sh.getRange(2,1,sh.getLastRow()-1,7).getValues();
  vals.forEach(r => {
    const key = normStr_(r[0]);
    const val = r[1];

    if (key && cfg.hasOwnProperty(key)) {
      const num = Number(val);
      cfg[key] = isFinite(num) ? num : val;
    }

    // includeWords/excludeWords は 1行で「 / 」区切り想定
    const inc = normStr_(r[3]);
    const exc = normStr_(r[4]);
    const incRe = normStr_(r[5]);
    const excRe = normStr_(r[6]);

    if (inc) includeWords = includeWords.concat(inc.split("/").map(s=>normStr_(s)).filter(Boolean));
    if (exc) excludeWords = excludeWords.concat(exc.split("/").map(s=>normStr_(s)).filter(Boolean));
    if (incRe) includeRegex.push(incRe);
    if (excRe) excludeRegex.push(excRe);
  });

  // 重複除去
  includeWords = uniq_(includeWords);
  excludeWords = uniq_(excludeWords);
  includeRegex = uniq_(includeRegex);
  excludeRegex = uniq_(excludeRegex);

  return { cfg, includeWords, excludeWords, includeRegex, excludeRegex };
}
function uniq_(arr){
  const out = [];
  const seen = new Set();
  (arr||[]).forEach(x=>{
    const s = String(x||"").trim();
    if (!s) return;
    if (seen.has(s)) return;
    seen.add(s);
    out.push(s);
  });
  return out;
}

/** ===== ドメイン設定読み取り（BASEの「設定」シート） ===== */
function buildAccessibleSiteUrlMap_() {
  const ssBase = SpreadsheetApp.openById(BASE_SHEET_ID);
  const sh = ssBase.getSheetByName(DOMAIN_SETTINGS_SHEET);
  if (!sh) throw new Error("BASE側に「設定」シートが見つかりません: " + DOMAIN_SETTINGS_SHEET);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2,1,lastRow-1,2).getValues(); // A:ドメイン名 B:サイトマップ
  const out = [];

  vals.forEach(r => {
    const domain = normStr_(r[0]);
    const sitemap = normStr_(r[1]);

    if (!domain || !domain.includes(".")) return;

    // sc-domain 前提
    const siteUrl = "sc-domain:" + domain.replace(/^sc-domain:/, "");
    out.push({ domain, siteUrl, sitemap });
  });

  return out;
}

/** ===== Search Console API（UrlFetch） ===== */
function gscFetch_(siteUrl, payload) {
  const token = ScriptApp.getOAuthToken();
  const endpoint = "https://searchconsole.googleapis.com/webmasters/v3/sites/"
    + encodeURIComponent(siteUrl) + "/searchAnalytics/query";

  const res = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + token }
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code >= 200 && code < 300) {
    return JSON.parse(text);
  }

  // 403/404等はログして空で返す（ドメイン設定の途中で落とさない）
  log_("ERROR", "GSC API error: " + code, { siteUrl, payload, text: text.slice(0, 1200) });
  return { rows: [] };
}

/** ===== Stage1: 上位クエリ収集 ===== */
function RUN_STAGE1_QUERYLIST_ALL_DOMAINS() {
  // 設定が無ければ作る
  const ssOut = SpreadsheetApp.openById(KEYWORD_SHEET_ID);
  if (!ssOut.getSheetByName(SH_KW_SETTINGS)) SETUP_KW_SETTINGS();

  const { cfg, includeWords, excludeWords, includeRegex, excludeRegex } = getKwConfig_();
  const sites = buildAccessibleSiteUrlMap_();

  ensureSheet_(ssOut, SH_KW_CAND, true);
  ensureSheet_(ssOut, SH_LOG, false);

  const shOut = ssOut.getSheetByName(SH_KW_CAND);
  shOut.clearContents();

  shOut.getRange(1,1,1,14).setValues([[
    "domain","siteUrl","query",
    "clicks","impressions","ctr","position",
    "intentScore","intentLabel","scoreReason",
    "対象(checkbox)","rangeStart","rangeEnd","memo"
  ]]);

  // 期間
  const base = shiftDate_(new Date(), -Number(cfg.LAG_DAYS || 0));
  const end = base;
  const start = shiftDate_(end, -(Number(cfg.DAYS || 90)-1));

  const rowsAll = [];
  let total = 0;

  sites.forEach(s => {
    const payload = {
      startDate: fmt_(start),
      endDate: fmt_(end),
      dimensions: ["query"],
      rowLimit: Number(cfg.MAX_QUERIES_PER_DOMAIN || 2000)
    };

    const json = gscFetch_(s.siteUrl, payload);
    const rows = (json.rows || []);

    rows.forEach(r => {
      if (total >= Number(cfg.MAX_KEYWORDS_TOTAL || 8000)) return;

      const q = normStr_((r.keys && r.keys[0]) ? r.keys[0] : "");
      if (!q) return;

      const imp = Number(r.impressions || 0);
      if (imp < Number(cfg.MIN_IMP || 0)) return;

      // フィルタ
      if (isExcludedQuery_(q, excludeWords, excludeRegex)) return;
      if (!isIncludedQuery_(q, includeWords, includeRegex)) {
        // include条件を設定している場合：何もヒットしないクエリは落とす
        if ((includeWords && includeWords.length) || (includeRegex && includeRegex.length)) return;
      }

      const score = scoreIntent_(q, includeWords, excludeWords);
      const intentScore = score.score;
      const intentLabel = score.label;
      const scoreReason = score.reason;

      // 閾値未満は原則落とす（ただし「includeWords/regex」無し運用なら広めに拾いたいケースがあるので条件化）
      if (Number(cfg.INTENT_MIN_SCORE || 0) > 0) {
        if (intentScore < Number(cfg.INTENT_MIN_SCORE)) return;
      }

      const clk = Number(r.clicks || 0);
      const pos = (typeof r.position === "number" && isFinite(r.position)) ? Number(r.position) : 100;
      const ctr = imp > 0 ? (clk / imp) : 0;

      rowsAll.push([
        s.domain, s.siteUrl, q,
        clk, imp, ctr, pos,
        intentScore, intentLabel, scoreReason,
        false, fmt_(start), fmt_(end), ""
      ]);
      total++;
    });
  });

  // 並べ替え（意図スコア→表示回数→クリック）
  rowsAll.sort((a,b) => {
    const s1 = Number(a[7] || 0), s2 = Number(b[7] || 0);
    if (s1 !== s2) return s2 - s1;
    const i1 = Number(a[4] || 0), i2 = Number(b[4] || 0);
    if (i1 !== i2) return i2 - i1;
    const c1 = Number(a[3] || 0), c2 = Number(b[3] || 0);
    return c2 - c1;
  });

  if (rowsAll.length) {
    shOut.getRange(2,1,rowsAll.length,14).setValues(rowsAll);

    // checkbox
    try {
      shOut.getRange(2,11,rowsAll.length,1).insertCheckboxes();
    } catch(e) {}

    // ctr/pos 表示整形
    shOut.getRange(2,6,rowsAll.length,1).setNumberFormat("0.00%");
    shOut.getRange(2,7,rowsAll.length,1).setNumberFormat("0.0");
  }

  shOut.setFrozenRows(1);
  shOut.autoResizeColumns(1,14);

  log_("INFO", "Stage1 完了: KW候補を出力しました", {
    domains: sites.length,
    keywords: rowsAll.length,
    range: { start: fmt_(start), end: fmt_(end) }
  });
}

/** ===== クエリ include/exclude ===== */
function isExcludedQuery_(q, excludeWords, excludeRegex){
  const s = String(q || "");
  (excludeWords || []).forEach(w => {
    if (!w) return;
  });
  // excludeWords（部分一致）
  for (let i=0;i<(excludeWords||[]).length;i++){
    const w = String(excludeWords[i]||"").trim();
    if (!w) continue;
    if (s.includes(w)) return true;
  }
  // excludeRegex
  for (let i=0;i<(excludeRegex||[]).length;i++){
    const re = String(excludeRegex[i]||"").trim();
    if (!re) continue;
    try {
      if (new RegExp(re).test(s)) return true;
    } catch(e){}
  }
  return false;
}
function isIncludedQuery_(q, includeWords, includeRegex){
  const s = String(q || "");
  // includeWords があれば、どれかに当たればOK
  if (includeWords && includeWords.length){
    for (let i=0;i<includeWords.length;i++){
      const w = String(includeWords[i]||"").trim();
      if (!w) continue;
      if (s.includes(w)) return true;
    }
    // includeWordsがあるのに何も当たらないなら false（ただし includeRegex があれば後で判定）
  }
  // includeRegex があれば、どれかに当たればOK
  if (includeRegex && includeRegex.length){
    for (let i=0;i<includeRegex.length;i++){
      const re = String(includeRegex[i]||"").trim();
      if (!re) continue;
      try {
        if (new RegExp(re).test(s)) return true;
      } catch(e){}
    }
  }
  // include条件が一切無い場合は true
  if ((!includeWords || !includeWords.length) && (!includeRegex || !includeRegex.length)) return true;

  return false;
}

/** ===== 意図スコア（CV寄り判定） ===== */
function scoreIntent_(query, includeWords, excludeWords){
  const q = String(query || "").trim();
  if (!q) return { score:0, label:"LOW", reason:"empty" };

  let score = 0;
  const reasons = [];

  // CV寄り強ワード（固定）
  const STRONG = [
    "料金","費用","月謝","体験","無料体験","問い合わせ","入塾","相談",
    "個別","個別指導","家庭教師","オンライン","塾","講座","コース",
    "評判","口コミ","比較","おすすめ"
  ];
  const MEDIUM = [
    "対策","講習","転塾","併用","フォロー","入試","中学受験","算数","国語","理科","社会","英語",
    "サピックス","早稲アカ","四谷","グノ","浜学園"
  ];
  const INFO = [
    "とは","意味","読み方","定義","公式","解き方","答え","単位","何","なに","wiki"
  ];

  STRONG.forEach(w=>{
    if (q.includes(w)) { score += 3; reasons.push("+"+w); }
  });
  MEDIUM.forEach(w=>{
    if (q.includes(w)) { score += 1; reasons.push("+"+w); }
  });

  // includeWords を使っている場合：当たれば少し加点（運用に寄せる）
  (includeWords || []).forEach(w=>{
    const s = String(w||"").trim();
    if (!s) return;
    if (q.includes(s)) { score += 0.5; }
  });

  // 情報系は減点（ただし、CVワードも含む場合は相殺される）
  INFO.forEach(w=>{
    if (q.includes(w)) { score -= 2; reasons.push("-"+w); }
  });
  (excludeWords || []).forEach(w=>{
    const s = String(w||"").trim();
    if (!s) return;
    if (q.includes(s)) { score -= 1; }
  });

  // 短すぎる単語は弱い（1語だけ等）
  if ([...q].length <= 2) score -= 2;

  // ラベル
  let label = "LOW";
  if (score >= 10) label = "HIGH";
  else if (score >= 6) label = "MID";

  return { score: Math.round(score*10)/10, label, reason: reasons.slice(0, 30).join(" ") };
}

/** ===== シート確保 ===== */
function ensureSheet_(ss, name, clear){
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (clear) sh.clearContents();
  return sh;
}

/** ===== Stage2: 対象KWを Query×Page に展開 ===== */
function RUN_STAGE2_EXPAND_SELECTED_KEYWORDS() {
  const ssOut = SpreadsheetApp.openById(KEYWORD_SHEET_ID);
  const shCand = ssOut.getSheetByName(SH_KW_CAND);
  if (!shCand || shCand.getLastRow() < 2) {
    throw new Error("先に Stage1（KW候補の作成）を実行してください。");
  }

  // ★Apps Scriptに「残り実行ミリ秒」を返すAPIは無いので、自前の経過時間ガードで打ち切る
  // 6分枠想定。バッファ込みで安全に抜ける
  const shouldStop = makeTimeGuard_(6 * 60 * 1000);

  const { cfg } = getKwConfig_();
  const maxKw = Number(cfg.MAX_KEYWORDS_TO_EXPAND || 300);
  const topPages = Number(cfg.TOP_PAGES_PER_KW || 20);

  // 期間（KW候補の列から取得：rangeStart/rangeEnd）
  const header = shCand.getRange(1,1,1,shCand.getLastColumn()).getValues()[0];
  const idx = indexMap_(header);

  const colSiteUrl = idx["siteUrl"] || 2;
  const colQuery   = idx["query"]   || 3;
  const colTarget  = idx["対象(checkbox)"] || idx["対象"] || 11;
  const colStart   = idx["rangeStart"] || 12;
  const colEnd     = idx["rangeEnd"]   || 13;

  const values = shCand.getRange(2,1,shCand.getLastRow()-1,shCand.getLastColumn()).getValues();

  // 対象=TRUE を優先。無ければ上位から maxKw 件
  const picked = [];
  values.forEach(r => {
    const isTarget = (r[colTarget-1] === true);
    const q = normStr_(r[colQuery-1]);
    const siteUrl = normStr_(r[colSiteUrl-1]);
    const start = r[colStart-1];
    const end   = r[colEnd-1];
    if (!q || !siteUrl) return;
    if (isTarget) picked.push({ siteUrl, query:q, rangeStart: start, rangeEnd: end });
  });

  if (!picked.length) {
    for (let i=0; i<values.length && picked.length<maxKw; i++){
      const r = values[i];
      const q = normStr_(r[colQuery-1]);
      const siteUrl = normStr_(r[colSiteUrl-1]);
      const start = r[colStart-1];
      const end   = r[colEnd-1];
      if (!q || !siteUrl) continue;
      picked.push({ siteUrl, query:q, rangeStart: start, rangeEnd: end });
    }
  }

  // カテゴリMap（任意）※CACHE_SHEET_IDが空/未設定なら使わない
  const catMap = (CACHE_SHEET_ID && CACHE_SHEET_NAME) ? loadCategoryMap_() : {};

  // Stage2進捗（タイムアウト回避のため、カーソル方式）
  const props = PropertiesService.getScriptProperties();
  const cursor = Number(props.getProperty("KW_STAGE2_CURSOR") || "0");

  // 初回（cursor=0）なら出力をクリア
  if (cursor === 0){
    ensureSheet_(ssOut, SH_KW_PAGES, true);
    ensureSheet_(ssOut, SH_KW_SUMMARY, true);
    ensureSheet_(ssOut, SH_KW_CANNIBAL, true);

    ssOut.getSheetByName(SH_KW_PAGES).getRange(1,1,1,12).setValues([[
      "query","siteUrl","page","category",
      "clicks","impressions","ctr","position",
      "rangeStart","rangeEnd","pageShare(clicks)","note"
    ]]);
  }

  let outRows = [];
  let processed = 0;

  // 保存して抜ける共通処理
  function saveAndExit_(nextCursor, reason, extra){
    props.setProperty("KW_STAGE2_CURSOR", String(nextCursor));
    flushStage2_(ssOut, outRows, catMap, Number(cfg.CANNIBAL_MIN_IMP || 50));
    log_("INFO", reason, Object.assign({
      cursor: nextCursor,
      totalPicked: picked.length,
      processed
    }, extra || {}));
  }

  const BUF_MS = 25 * 1000; // 25秒バッファで安全停止

  for (let i = cursor; i < picked.length; i++) {
    // ループに入った時点で時間が厳しければ「未処理のまま」保存して抜ける
    if (shouldStop(BUF_MS)) {
      saveAndExit_(i, "Stage2 途中保存（時間ガード・続きは再実行）", null);
      return;
    }

    const item = picked[i];

    // 範囲（Date型で入っている場合のみ採用）
    const rs = item.rangeStart;
    const re = item.rangeEnd;

    let end = (re instanceof Date) ? re : null;
    let start = (rs instanceof Date) ? rs : null;

    // 候補から取れない場合は cfg から再計算
    if (!end || !start){
      const base = shiftDate_(new Date(), -Number(cfg.LAG_DAYS || 0));
      end = base;
      start = shiftDate_(end, -(Number(cfg.DAYS || 90)-1));
    }

    // API: query固定 → page一覧
    const payload = {
      startDate: fmt_(start),
      endDate: fmt_(end),
      dimensions: ["page"],
      dimensionFilterGroups: [{
        filters: [{
          dimension: "query",
          operator: "equals",
          expression: item.query
        }]
      }],
      rowLimit: 5000
    };

    const json = gscFetch_(item.siteUrl, payload);
    const rows = (json.rows || []);

    if (!rows.length) {
      processed++;

      // 1回に maxKw 超えない
      if (processed >= maxKw) {
        saveAndExit_(i + 1, "Stage2 途中保存（上限到達・続きは再実行）", { processed });
        return;
      }

      // 時間ガード（処理後にもう一回）
      if (shouldStop(BUF_MS)) {
        saveAndExit_(i + 1, "Stage2 途中保存（時間ガード・続きは再実行）", null);
        return;
      }

      continue;
    }

    // クエリ合計（share算出用）
    let sumClicks = 0;
    rows.forEach(r => sumClicks += Number(r.clicks || 0));

    // 上位ページだけ出す
    rows.sort((a,b)=> Number(b.clicks||0) - Number(a.clicks||0));
    const limited = rows.slice(0, topPages);

    limited.forEach(r=>{
      const page = normUrl_((r.keys && r.keys[0]) ? r.keys[0] : "");
      const clk = Number(r.clicks || 0);
      const imp = Number(r.impressions || 0);
      const ctr = imp > 0 ? (clk / imp) : 0;
      const pos = (typeof r.position === "number" && isFinite(r.position)) ? Number(r.position) : 100;

      const cat = catMap[page] || catMap[normUrl_(page)] || "";
      const share = (sumClicks > 0) ? (clk / sumClicks) : 0;

      outRows.push([
        item.query, item.siteUrl, page, cat,
        clk, imp, ctr, pos,
        fmt_(start), fmt_(end),
        share,
        ""
      ]);
    });

    processed++;

    // 1回に maxKw 超えない
    if (processed >= maxKw) {
      saveAndExit_(i + 1, "Stage2 途中保存（上限到達・続きは再実行）", { processed });
      return;
    }

    // 時間ガード（処理後）
    if (shouldStop(BUF_MS)) {
      saveAndExit_(i + 1, "Stage2 途中保存（時間ガード・続きは再実行）", null);
      return;
    }
  }

  // 完走
  props.setProperty("KW_STAGE2_CURSOR", "0");
  flushStage2_(ssOut, outRows, catMap, Number(cfg.CANNIBAL_MIN_IMP || 50));
  log_("INFO", "Stage2 完了", { processed, totalPicked: picked.length });
}

/** ===== Stage2 進捗リセット ===== */
function KW_RESET_STAGE2_CURSOR(){
  PropertiesService.getScriptProperties().setProperty("KW_STAGE2_CURSOR", "0");
  log_("INFO", "Stage2カーソルをリセットしました", null);
}

/** ===== ヘッダ index map ===== */
function indexMap_(headerRow){
  const m = {};
  for (let i=0;i<headerRow.length;i++){
    const k = String(headerRow[i] || "").trim();
    if (!k) continue;
    m[k] = i+1; // 1-based
  }
  return m;
}

/** ===== URL→分類 map（任意） ===== */
function loadCategoryMap_(){
  const map = {};
  try {
    const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
    const sh = ss.getSheetByName(CACHE_SHEET_NAME);
    if (!sh) return map;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return map;

    const vals = sh.getRange(2,1,lastRow-1,2).getValues(); // A:URL B:分類
    vals.forEach(r=>{
      const url = normUrl_(r[0]);
      const cat = normStr_(r[1]);
      if (!url) return;
      map[url] = cat;
    });
    return map;
  } catch(e){
    log_("ERROR", "loadCategoryMap_ failed", { error: String(e) });
    return map;
  }
}

/** ===== Stage2 flush: KW→ページ出力 & サマリ/カニバリ生成 ===== */
function flushStage2_(ssOut, outRows, catMap, cannibalMinImp){
  if (!outRows || !outRows.length){
    // 空でもサマリシートのヘッダは作る
    ensureSummarySheets_(ssOut);
    return;
  }

  // KW→ページ 追記
  const shPages = ssOut.getSheetByName(SH_KW_PAGES) || ssOut.insertSheet(SH_KW_PAGES);
  const startRow = shPages.getLastRow() + 1;
  shPages.getRange(startRow,1,outRows.length,12).setValues(outRows);

  // 表示形式
  try {
    shPages.getRange(startRow,7,outRows.length,1).setNumberFormat("0.00%"); // ctr
    shPages.getRange(startRow,8,outRows.length,1).setNumberFormat("0.0");   // pos
    shPages.getRange(startRow,11,outRows.length,1).setNumberFormat("0.00%"); // share
  } catch(e){}

  // サマリ再生成（全行から作る：簡潔運用）
  ensureSummarySheets_(ssOut);

  const all = readAllKwPages_(ssOut);
  const summary = buildKwSummary_(all);

  writeKwSummary_(ssOut, summary);

  // カニバリ抽出
  const cannibal = summary.filter(s => s.pageCount >= 2 && s.totalImpressions >= cannibalMinImp && s.topShare < 0.80);
  writeKwCannibal_(ssOut, cannibal);

  log_("INFO", "Stage2 flush 完了", { appended: outRows.length, totalRows: all.length });
}

function ensureSummarySheets_(ssOut){
  const shSum = ssOut.getSheetByName(SH_KW_SUMMARY) || ssOut.insertSheet(SH_KW_SUMMARY);
  const shCan = ssOut.getSheetByName(SH_KW_CANNIBAL) || ssOut.insertSheet(SH_KW_CANNIBAL);

  shSum.clearContents();
  shCan.clearContents();

  shSum.getRange(1,1,1,12).setValues([[
    "query","totalClicks","totalImpressions","ctr","avgPosition(w)","pageCount",
    "topPage","topClicks","topShare","sites","rangeStart","rangeEnd"
  ]]);

  shCan.getRange(1,1,1,13).setValues([[
    "query","totalClicks","totalImpressions","ctr","avgPosition(w)","pageCount",
    "topPage","topClicks","topShare","sites","rangeStart","rangeEnd","note"
  ]]);
}

function readAllKwPages_(ssOut){
  const sh = ssOut.getSheetByName(SH_KW_PAGES);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2,1,sh.getLastRow()-1,12).getValues();
  return vals.map(r => ({
    query: r[0],
    siteUrl: r[1],
    page: r[2],
    category: r[3],
    clicks: Number(r[4]||0),
    impressions: Number(r[5]||0),
    ctr: Number(r[6]||0),
    position: Number(r[7]||100),
    rangeStart: r[8],
    rangeEnd: r[9]
  }));
}

function buildKwSummary_(rows){
  const m = {}; // query -> agg
  rows.forEach(x=>{
    const q = normStr_(x.query);
    if (!q) return;
    if (!m[q]) m[q] = {
      query:q,
      totalClicks:0,
      totalImpressions:0,
      posSum:0,
      posW:0,
      pages:{}, // page -> clicks, imp
      sites:new Set(),
      rangeStart: x.rangeStart,
      rangeEnd: x.rangeEnd
    };
    const a = m[q];
    a.totalClicks += x.clicks;
    a.totalImpressions += x.impressions;
    const w = x.impressions > 0 ? x.impressions : 1;
    a.posSum += x.position * w;
    a.posW += w;

    a.sites.add(String(x.siteUrl||""));
    const p = normUrl_(x.page);
    if (!a.pages[p]) a.pages[p] = { clicks:0, impressions:0 };
    a.pages[p].clicks += x.clicks;
    a.pages[p].impressions += x.impressions;
  });

  const out = [];
  Object.keys(m).forEach(q=>{
    const a = m[q];
    // top page
    let topPage = "";
    let topClicks = -1;
    Object.keys(a.pages).forEach(p=>{
      const c = a.pages[p].clicks;
      if (c > topClicks){
        topClicks = c;
        topPage = p;
      }
    });

    const totalClicks = a.totalClicks;
    const totalImpressions = a.totalImpressions;
    const ctr = totalImpressions > 0 ? (totalClicks / totalImpressions) : 0;
    const avgPos = a.posW > 0 ? (a.posSum / a.posW) : 100;
    const pageCount = Object.keys(a.pages).length;
    const topShare = totalClicks > 0 ? (topClicks / totalClicks) : 0;

    out.push({
      query: a.query,
      totalClicks,
      totalImpressions,
      ctr,
      avgPos,
      pageCount,
      topPage,
      topClicks: topClicks < 0 ? 0 : topClicks,
      topShare,
      sites: Array.from(a.sites).join(" / "),
      rangeStart: a.rangeStart,
      rangeEnd: a.rangeEnd
    });
  });

  out.sort((x,y)=>{
    // 表示回数→クリック
    if (y.totalImpressions !== x.totalImpressions) return y.totalImpressions - x.totalImpressions;
    return y.totalClicks - x.totalClicks;
  });

  return out;
}

function writeKwSummary_(ssOut, list){
  const sh = ssOut.getSheetByName(SH_KW_SUMMARY);
  if (!list.length) return;

  const rows = list.map(x=>[
    x.query,
    x.totalClicks,
    x.totalImpressions,
    x.ctr,
    x.avgPos,
    x.pageCount,
    x.topPage,
    x.topClicks,
    x.topShare,
    x.sites,
    x.rangeStart,
    x.rangeEnd
  ]);

  sh.getRange(2,1,rows.length,12).setValues(rows);

  try {
    sh.getRange(2,4,rows.length,1).setNumberFormat("0.00%"); // ctr
    sh.getRange(2,5,rows.length,1).setNumberFormat("0.0");   // pos
    sh.getRange(2,9,rows.length,1).setNumberFormat("0.00%"); // topShare
  } catch(e){}

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,12);
}

function writeKwCannibal_(ssOut, list){
  const sh = ssOut.getSheetByName(SH_KW_CANNIBAL);
  if (!list.length) return;

  const rows = list.map(x=>[
    x.query,
    x.totalClicks,
    x.totalImpressions,
    x.ctr,
    x.avgPos,
    x.pageCount,
    x.topPage,
    x.topClicks,
    x.topShare,
    x.sites,
    x.rangeStart,
    x.rangeEnd,
    "topShare<0.80 && pageCount>=2"
  ]);

  sh.getRange(2,1,rows.length,13).setValues(rows);

  try {
    sh.getRange(2,4,rows.length,1).setNumberFormat("0.00%");
    sh.getRange(2,5,rows.length,1).setNumberFormat("0.0");
    sh.getRange(2,9,rows.length,1).setNumberFormat("0.00%");
  } catch(e){}

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,13);
}

// 実行時間の安全打ち切り（Apps Scriptには残りミリ秒APIが無いので自前）
function makeTimeGuard_(maxMillis) {
  const start = Date.now();
  return function shouldStop_(bufferMillis) {
    const buf = (typeof bufferMillis === "number") ? bufferMillis : 20 * 1000; // 20秒バッファ
    return (Date.now() - start) >= (maxMillis - buf);
  };
}