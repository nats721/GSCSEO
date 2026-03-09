/************************************************************
 * 内部リンク現状スナップショット（CTA分類対応・完全版）
 *
 * - canonical は元シート URL を唯一の正とする（出力は必ず raw URL）
 * - link_type:
 *    content    : 通常の文中リンク
 *    cta_inline : 本文内ボタン等（btn/button/cta系）
 *    cta_auto   : <section class="veu_cta"> 内（自動CTA）
 * - SNSシェア（veu_socialSet）＆ intent/sharer URL は除外
 * - ページ別は「全URLを母集合」とし、リンク0も必ず表示
 * - http→https で存在するものは to_exists を https 側で true に寄せつつ、
 *   別シートに「HTTP残存リンク」リストを出す
 *
 * 出力：
 * ① 内部リンク_一覧（9列：末尾に link_type）
 * ② 内部リンク_ページ別（設計リンク指標＋内訳）※0も出す
 * ③ 内部リンク_被リンク（content / cta_inline / cta_auto 別）
 * ④ 内部リンク_HTTP残存（暫時修正用）
 ************************************************************/

function exportInternalLinksAll() {

  /* ========= 設定 ========= */

  const SRC_SS_ID      = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o';
  const SRC_SHEET_GID  = 1666145974;
  const DST_SS_ID      = '1U-XK7TVV8BWxo27iRqySKmh6SmlyZncjoC26ZAX8LjM';

  const SHEET_LIST     = '内部リンク_一覧';
  const SHEET_PAGE     = '内部リンク_ページ別';
  const SHEET_INBOUND  = '内部リンク_被リンク';
  const SHEET_HTTP     = '内部リンク_HTTP残存';

  // hrefレベル除外（保険）
  const EXCLUDE_HREF_PATTERNS = [
    'facebook.com/sharer.php',
    'twitter.com/intent/tweet',
    'twitter.com/intent',
    'bsky.app/intent/compose'
  ];

  Logger.log('実行開始');

  /* ========= シート取得 ========= */

  const srcSS = SpreadsheetApp.openById(SRC_SS_ID);
  const srcSheet = srcSS.getSheets().find(s => s.getSheetId() === SRC_SHEET_GID);
  if (!srcSheet) throw new Error('元シートが見つかりません（gid確認）');

  const dstSS  = SpreadsheetApp.openById(DST_SS_ID);
  const shList = resetSheet_(dstSS, SHEET_LIST);
  const shPage = resetSheet_(dstSS, SHEET_PAGE);
  const shIn   = resetSheet_(dstSS, SHEET_INBOUND);
  const shHttp = resetSheet_(dstSS, SHEET_HTTP);

  /* ========= 元データ取得 ========= */

  const values = srcSheet.getDataRange().getValues();
  const header = values.shift();

  const idx = {};
  for (let i = 0; i < header.length; i++) idx[String(header[i])] = i;

  const IDX_URL     = idx['URL'];
  const IDX_CLASS   = idx['分類'];
  const IDX_LASTMOD = idx['lastmod'];
  const IDX_BODY    = idx['entryBodyHtml'];

  if ([IDX_URL, IDX_CLASS, IDX_LASTMOD, IDX_BODY].some(v => v === undefined)) {
    throw new Error('元シートの列名が一致しません（URL / 分類 / lastmod / entryBodyHtml）');
  }

  /* ========= canonical URL マップ構築 =========
   * key   : 正規化済みURL候補（raw / decode / encode / http<->https候補）
   * value : { rawUrl, 分類 }
   *
   * 重要：
   * - 出力は rawUrl（元シート表記）を使う
   * - 正規化は照合キーのためだけに使う
   */

  const urlMap = new Map();

  for (let i = 0; i < values.length; i++) {
    const raw = values[i][IDX_URL];
    if (!raw) continue;

    const rawUrl = String(raw);             // ← 末尾/ありの元表記
    const klass  = values[i][IDX_CLASS];

    const candidates = buildLookupCandidates_(rawUrl);

    const payload = { rawUrl: rawUrl, 分類: klass };
    for (let k = 0; k < candidates.length; k++) {
      urlMap.set(candidates[k], payload);
    }
  }

  /* ========= 抽出 ========= */

  const A_TAG_RE = /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;

  // 文末自動CTA（VK ExUnit系）
  const CTA_SECTION_RE = /<section[^>]*class=["'][^"']*veu_cta[^"']*["'][\s\S]*?<\/section>/gi;

  // SNSシェアブロック（VK/VEU系）
  const SOCIALSET_DIV_RE = /<div[^>]*class=["'][^"']*veu_socialSet[^"']*["'][\s\S]*?<\/div>/gi;

  const linkRows = [];

  // HTTP残存（暫時修正タスク用）
  const httpRows = [];
  const httpSeen = new Set(); // 重複除去キー

  for (let i = 0; i < values.length; i++) {

    const fromUrlRaw = values[i][IDX_URL];
    const bodyHtml   = values[i][IDX_BODY];
    if (!fromUrlRaw || !bodyHtml) continue;

    const fromRaw    = String(fromUrlRaw);     // ← 出力は必ずこの raw
    const fromHost   = getHost_(fromRaw);
    if (!fromHost) continue;

    const fromClass   = values[i][IDX_CLASS];
    const fromLastmod = values[i][IDX_LASTMOD];
    const html        = String(bodyHtml);

    // --- 1) cta_auto 抽出：<section class="veu_cta"> 内だけ ---
    const ctaSections = html.match(CTA_SECTION_RE) || [];
    for (let s = 0; s < ctaSections.length; s++) {
      const sec = ctaSections[s];
      let m;
      A_TAG_RE.lastIndex = 0;
      while ((m = A_TAG_RE.exec(sec)) !== null) {
        pushLink_(m, 'cta_auto');
      }
    }

    // --- 2) 本文（content / cta_inline）抽出 ---
    let cleanHtml = html.replace(CTA_SECTION_RE, '').replace(SOCIALSET_DIV_RE, '');

    let m;
    A_TAG_RE.lastIndex = 0;
    while ((m = A_TAG_RE.exec(cleanHtml)) !== null) {
      const aHtml = m[0];
      const isInlineCta = /class=["'][^"']*(btn|button|cta)[^"']*/i.test(aHtml);
      pushLink_(m, isInlineCta ? 'cta_inline' : 'content');
    }

    function pushLink_(m, linkType) {

      const rawHref = String(m[1] || '').trim();
      if (!rawHref) return;

      // 明示除外
      if (rawHref.startsWith('#')) return;
      if (/^(mailto|tel|javascript):/i.test(rawHref)) return;

      // SNS intent/sharer URL 除外（hrefレベル）
      const hrefLower = rawHref.toLowerCase();
      for (let k = 0; k < EXCLUDE_HREF_PATTERNS.length; k++) {
        if (hrefLower.indexOf(EXCLUDE_HREF_PATTERNS[k]) !== -1) return;
      }

      const resolved = resolveUrl_(fromRaw, rawHref);
      if (!resolved) return;

      // 照合候補を作る（decode/encode + http/https 揺れ）
      const candidates = buildLookupCandidates_(resolved);

      let hit = null;
      for (let c = 0; c < candidates.length; c++) {
        const h = urlMap.get(candidates[c]);
        if (h) { hit = h; break; }
      }

      const exists = !!hit;

      // 出力URLは必ず rawUrl（元シート表記）に寄せる
      // 見つからない場合のみ resolved（正規化はしない：事実の記録）
      const finalTo = exists ? hit.rawUrl : normalizeUrlPreserveSchemeAndHost_(resolved);
      const toClass = exists ? hit.分類 : '';

      const toHost = getHost_(finalTo);
      const isInternal = (toHost === fromHost);

      const anchor = stripHtml_(String(m[2] || ''));

      linkRows.push([
        fromRaw,
        fromLastmod,
        fromClass,
        finalTo,
        toClass,
        anchor,
        isInternal,
        exists,
        linkType
      ]);

      // HTTP残存：内部リンクで、httpで書かれているが https canonical が存在するもの
      // ※finalTo は https(rawUrl) に寄るので、ここは rawHref/resolved を基準に拾う
      if (isInternal) {
        const resolvedNorm = normalizeUrlPreserveSchemeAndHost_(resolved);
        if (resolvedNorm.toLowerCase().startsWith('http://')) {
          // https にした照合キーが urlMap に存在するか
          const httpsCandidate = 'https://' + resolvedNorm.substring(7);
          const httpsCandidates = buildLookupCandidates_(httpsCandidate);

          let httpsHit = null;
          for (let cc = 0; cc < httpsCandidates.length; cc++) {
            const hh = urlMap.get(httpsCandidates[cc]);
            if (hh) { httpsHit = hh; break; }
          }

          if (httpsHit) {
            const key = fromRaw + '||' + rawHref + '||' + linkType;
            if (!httpSeen.has(key)) {
              httpSeen.add(key);
              httpRows.push([
                fromRaw,
                fromLastmod,
                fromClass,
                linkType,
                rawHref,
                resolvedNorm,
                httpsHit.rawUrl
              ]);
            }
          }
        }
      }
    }
  }

  /* ========= ① 内部リンク_一覧 ========= */

  shList.getRange(1, 1, 1, 9).setValues([[
    'from_URL','lastmod(from)','from_分類',
    'to_URL','to_分類','anchor',
    'is_internal','to_exists','link_type'
  ]]);

  if (linkRows.length) {
    shList.getRange(2, 1, linkRows.length, 9).setValues(linkRows);
  }

  /* ========= ② ページ別集計（全URL母集合、0も表示） =========
   * 設計リンク = content + cta_inline（cta_autoは除外）
   * ただし内訳として content / cta_inline / cta_auto も出す
   */

  const pageMap = new Map();

  // ★母集合：元シート全URLで初期化（リンク0も必ず出す）
  for (let i = 0; i < values.length; i++) {
    const u = values[i][IDX_URL];
    if (!u) continue;

    const rawUrl = String(u);
    if (!pageMap.has(rawUrl)) {
      pageMap.set(rawUrl, {
        分類: values[i][IDX_CLASS],

        // 設計指標（cta_auto除外）
        design_internal: 0,
        design_external: 0,
        design_broken: 0,
        design_uniq: {},

        // 内訳
        internal_content: 0,
        internal_cta_inline: 0,
        internal_cta_auto: 0,
        external_content: 0,
        external_cta_inline: 0,
        external_cta_auto: 0,
        broken_content: 0,
        broken_cta_inline: 0,
        broken_cta_auto: 0
      });
    }
  }

  for (let i = 0; i < linkRows.length; i++) {
    const r = linkRows[i];

    const from     = r[0];  // raw
    const toUrl    = r[3];
    const isInternal = r[6] === true;
    const exists   = r[7] === true;
    const linkType = r[8];

    const o = pageMap.get(from);
    if (!o) continue; // 念のため

    // 内訳カウント（全link_type）
    if (isInternal) {
      if (linkType === 'content') o.internal_content++;
      else if (linkType === 'cta_inline') o.internal_cta_inline++;
      else if (linkType === 'cta_auto') o.internal_cta_auto++;

      if (!exists) {
        if (linkType === 'content') o.broken_content++;
        else if (linkType === 'cta_inline') o.broken_cta_inline++;
        else if (linkType === 'cta_auto') o.broken_cta_auto++;
      }
    } else {
      if (linkType === 'content') o.external_content++;
      else if (linkType === 'cta_inline') o.external_cta_inline++;
      else if (linkType === 'cta_auto') o.external_cta_auto++;
    }

    // 設計指標（cta_auto除外）
    if (linkType !== 'cta_auto') {
      if (isInternal) {
        o.design_internal++;
        o.design_uniq[toUrl] = true;
        if (!exists) o.design_broken++;
      } else {
        o.design_external++;
      }
    }
  }

  const pageRows = [];
  pageMap.forEach((v, url) => {
    pageRows.push([
      url,
      v.分類,

      v.design_internal,
      v.design_external,
      Object.keys(v.design_uniq).length,
      v.design_broken,

      v.internal_content,
      v.internal_cta_inline,
      v.internal_cta_auto,
      v.external_content,
      v.external_cta_inline,
      v.external_cta_auto
    ]);
  });

  shPage.getRange(1, 1, 1, 12).setValues([[
    'URL','分類',
    '内部リンク数_設計(content+cta_inline)','外部リンク数_設計(content+cta_inline)',
    'ユニーク内部リンク数_設計','内部リンク切れ数_設計',
    '内部_content','内部_cta_inline','内部_cta_auto',
    '外部_content','外部_cta_inline','外部_cta_auto'
  ]]);

  if (pageRows.length) {
    shPage.getRange(2, 1, pageRows.length, 12).setValues(pageRows);
  }

  /* ========= ③ 被リンク集計 =========
   * internal かつ to_exists=true のものだけ集計
   * content / cta_inline / cta_auto を分けて出す（既存仕様維持）
   */

  const inboundMap = new Map();

  for (let i = 0; i < linkRows.length; i++) {
    const r = linkRows[i];

    const from = r[0];
    const to   = r[3];
    const toClass = r[4];
    const isInternal = r[6] === true;
    const exists = r[7] === true;
    const linkType = r[8];

    if (!isInternal || !exists) continue;

    if (!inboundMap.has(to)) {
      inboundMap.set(to, {
        分類: toClass,
        content_count: 0, content_src: {},
        cta_inline_count: 0, cta_inline_src: {},
        cta_auto_count: 0, cta_auto_src: {}
      });
    }

    const o = inboundMap.get(to);

    if (linkType === 'content') {
      o.content_count++;
      o.content_src[from] = true;
    } else if (linkType === 'cta_inline') {
      o.cta_inline_count++;
      o.cta_inline_src[from] = true;
    } else if (linkType === 'cta_auto') {
      o.cta_auto_count++;
      o.cta_auto_src[from] = true;
    }
  }

  const inboundRows = [];
  inboundMap.forEach((v, url) => {
    inboundRows.push([
      url,
      v.分類,

      v.content_count,
      Object.keys(v.content_src).length,

      v.cta_inline_count,
      Object.keys(v.cta_inline_src).length,

      v.cta_auto_count,
      Object.keys(v.cta_auto_src).length
    ]);
  });

  shIn.getRange(1, 1, 1, 8).setValues([[
    'to_URL','to_分類',
    '被リンク数_content','リンク元ユニーク数_content',
    '被リンク数_cta_inline','リンク元ユニーク数_cta_inline',
    '被リンク数_cta_auto','リンク元ユニーク数_cta_auto'
  ]]);

  if (inboundRows.length) {
    shIn.getRange(2, 1, inboundRows.length, 8).setValues(inboundRows);
  }

  /* ========= ④ HTTP残存リンク ========= */

  shHttp.getRange(1, 1, 1, 7).setValues([[
    'from_URL','lastmod(from)','from_分類',
    'link_type','href_found','resolved_http','canonical_https'
  ]]);

  if (httpRows.length) {
    shHttp.getRange(2, 1, httpRows.length, 7).setValues(httpRows);
  }

  SpreadsheetApp.flush();

  Logger.log(
    '完了：リンク行=' + linkRows.length +
    ' / ページ集計=' + pageRows.length +
    ' / 被リンク集計=' + inboundRows.length +
    ' / HTTP残存=' + httpRows.length
  );
}

/* ================== utilities（このファイル単体で完結） ================== */

function resetSheet_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  return sh;
}

// 照合用キー：#・?除去、末尾/除去
function normalizeUrl_(url) {
  let u = String(url || '').trim();
  if (!u) return u;

  const h = u.indexOf('#');
  if (h >= 0) u = u.substring(0, h);

  const q = u.indexOf('?');
  if (q >= 0) u = u.substring(0, q);

  if (u.length > 8 && u.endsWith('/')) u = u.slice(0, -1);
  return u;
}

// 出力用の保険：#・?だけ除去（末尾/は触らない）
function normalizeUrlPreserveSchemeAndHost_(url) {
  let u = String(url || '').trim();
  if (!u) return u;

  const h = u.indexOf('#');
  if (h >= 0) u = u.substring(0, h);

  const q = u.indexOf('?');
  if (q >= 0) u = u.substring(0, q);

  return u;
}

function safeDecode_(u) {
  try {
    return decodeURIComponent(String(u));
  } catch (e) {
    return String(u);
  }
}

/**
 * 絶対URLの path 部分のみを安全に encode する（scheme/hostは触らない）
 */
function safeEncodeUrl_(absUrl) {
  const s = String(absUrl || '').trim();

  const m = s.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!m) return s;

  const origin = m[1];
  const path = m[2] || '';

  const encodedPath = path.split('/').map((seg, i) => {
    if (i === 0) return '';
    const dec = safeDecode_(seg);
    return encodeURIComponent(dec);
  }).join('/');

  return origin + encodedPath;
}

function getHost_(absUrl) {
  const m = String(absUrl || '').match(/^https?:\/\/([^\/]+)/i);
  return m ? m[1].toLowerCase() : '';
}

// URL クラス不使用の絶対URL解決（../ も処理）
function resolveUrl_(baseAbs, href) {
  const base = String(baseAbs || '');
  const h = String(href || '').trim();
  if (!h) return '';

  if (/^https?:\/\//i.test(h)) return h;

  if (/^\/\//.test(h)) {
    const scheme = base.match(/^https?:/i);
    return (scheme ? scheme[0] : 'https:') + h;
  }

  const bm = base.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!bm) return '';

  const origin = bm[1];
  let basePath = bm[2] || '/';

  if (!basePath.endsWith('/')) {
    const i = basePath.lastIndexOf('/');
    basePath = (i >= 0) ? basePath.substring(0, i + 1) : '/';
  }

  if (h.startsWith('/')) return origin + normalizePath_(h);

  return origin + normalizePath_(basePath + h);
}

function normalizePath_(path) {
  const parts = String(path || '').split('/');
  const stack = [];

  for (let i = 0; i < parts.length; i++) {
    const seg = parts[i];

    if (seg === '' && i === 0) { stack.push(''); continue; }
    if (seg === '' || seg === '.') continue;

    if (seg === '..') {
      if (stack.length > 1) stack.pop();
      continue;
    }
    stack.push(seg);
  }

  let out = stack.join('/');
  if (!out.startsWith('/')) out = '/' + out;
  return out;
}

function stripHtml_(s) {
  return String(s || '').replace(/<[^>]+>/g, '').trim();
}

/**
 * 照合候補生成：
 * - normalize（末尾/除去）
 * - decode / encode 変換
 * - http<->https 揺れも候補に含める
 */
function buildLookupCandidates_(absUrl) {
  const out = [];
  const seen = new Set();

  function add(u) {
    const k = normalizeUrl_(u);
    if (!k) return;
    if (seen.has(k)) return;
    seen.add(k);
    out.push(k);
  }

  const raw = String(absUrl || '').trim();
  if (!raw) return out;

  // ベース
  add(raw);

  // decode/encode
  const n = normalizeUrl_(raw);
  const d = normalizeUrl_(safeDecode_(n));
  const e = normalizeUrl_(safeEncodeUrl_(d));
  add(n); add(d); add(e);

  // http<->https 揺れ
  if (n.toLowerCase().startsWith('http://')) {
    const https = 'https://' + n.substring(7);
    add(https);
    add(safeDecode_(https));
    add(safeEncodeUrl_(safeDecode_(https)));
  } else if (n.toLowerCase().startsWith('https://')) {
    const http = 'http://' + n.substring(8);
    add(http);
    add(safeDecode_(http));
    add(safeEncodeUrl_(safeDecode_(http)));
  }

  return out;
}






