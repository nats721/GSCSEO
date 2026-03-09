/**
 * キャッシュ列（11列固定）
 *  A: URL
 *  B: 分類
 *  C: lastmod（サイトマップ由来／YYYY/MM/DD HH:MM:SS）
 *  D: fetch日時
 *  E: title
 *  F: metaDesc
 *  G: H1
 *  H: entryBodyHtml
 *  I: ogTitle
 *  J: ogDesc
 *  K: jsonLd
 */

const MAIN_SHEET_ID    = '1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs'; // 設定シートがあるブック
const CACHE_SHEET_ID   = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o'; // キャッシュ用ブック

const CACHE_SHEET_NAME = 'HTMLキャッシュ';
const CACHE_DIFF_LOG   = 'CacheDiffLog';

/************************************************************
 * ▼ 共通ユーティリティ
 ************************************************************/
function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function stripTags_(s) {
  return s ? s.replace(/<[^>]+>/g, '') : '';
}

/**
 * サイトマップの lastmod を JST の yyyy/MM/dd HH:mm:ss に正規化
 */
function normalizeLastmod_(raw) {
  if (!raw) return '';
  const d = new Date(raw);
  if (isNaN(d)) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

/************************************************************
 * ▼ entry-body 抽出（divネスト追跡・内部のみ）
 ************************************************************/
function extractEntryBodyStrict_(html) {
  if (!html) return '';

  // <div class="entry-body"...> を探す
  const startRe = /<div[^>]+class=["'][^"']*entry-body[^"']*["'][^>]*>/i;
  const startMatch = html.match(startRe);
  if (!startMatch) return '';

  const startTag   = startMatch[0];
  const startIndex = html.indexOf(startTag);
  if (startIndex === -1) return '';

  const contentStart = startIndex + startTag.length;

  let depth = 1;
  let i = contentStart;
  const len = html.length;

  while (i < len) {
    const openIdx  = html.indexOf('<div', i);
    const closeIdx = html.indexOf('</div', i);

    if (closeIdx === -1) break;

    if (openIdx !== -1 && openIdx < closeIdx) {
      depth++;
      i = openIdx + 4;
    } else {
      depth--;
      if (depth === 0) {
        const contentEnd = closeIdx; // 対応する </div> の直前まで
        const raw = html.substring(contentStart, contentEnd);
        // 先頭の空白・改行だけ削除
        return raw.replace(/^\s+/, '');
      }
      i = closeIdx + 5;
    }
  }

  // フォールバック：終端まで
  const fallback = html.substring(contentStart);
  return fallback.replace(/^\s+/, '');
}

/************************************************************
 * ▼ HTML解析：title / metaDesc / H1 / OGP / JSON-LD
 ************************************************************/
function extractTitle_(html) {
  const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  return m ? m[1].trim() : '';
}

function extractMetaDesc_(html) {
  const m = html.match(/<meta[^>]+name=["']description["'][^>]*content=["']([^"']*)["']/i);
  return m ? m[1].trim() : '';
}

function extractH1_(html) {
  const m = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  return m ? stripTags_(m[1]).trim() : '';
}

function extractOgMeta_(html, prop) {
  const re = new RegExp(
    `<meta[^>]+property=["']${prop}["'][^>]*content=["']([^"']*)["']`, 'i'
  );
  const m = html.match(re);
  return m ? m[1].trim() : '';
}

function extractJsonLd_(html) {
  const blocks = [];
  const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    blocks.push(m[1].trim());
  }
  return blocks.join('\n---\n');
}

/************************************************************
 * ▼ 設定シート → サイトマップ一覧取得
 *  シート「設定」
 *   A列: ドメイン名
 *   B列: サイトマップURL
 ************************************************************/
function getSitemapSettings_() {
  const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
  const sh = ss.getSheetByName('設定');
  if (!sh) throw new Error('設定シートがありません: 設定');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const list = [];

  values.forEach(([domain, sitemapUrl]) => {
    if (!domain || !sitemapUrl) return;
    list.push({ domain, sitemapUrl });
  });

  return list;
}

/************************************************************
 * ▼ サイトマップ（index含む）を再帰的に巡回し URL+lastmod 収集
 ************************************************************/
function collectUrlsFromSitemapRecursive_(sitemapUrl, urlMap, visited) {
  if (!sitemapUrl) return;
  if (visited.has(sitemapUrl)) return;
  visited.add(sitemapUrl);

  try {
    const res = UrlFetchApp.fetch(sitemapUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() >= 300) {
      Logger.log('Sitemap fetch失敗: ' + sitemapUrl + ' / ' + res.getResponseCode());
      return;
    }

    const xml = res.getContentText();
    const doc = XmlService.parse(xml);
    const root = doc.getRootElement();
    const children = root.getChildren();

    children.forEach(child => {
      const tagName = child.getName().toLowerCase();
      const ns = child.getNamespace();

      if (tagName === 'sitemap') {
        // サイトマップインデックス
        const locEl = child.getChild('loc', ns);
        if (locEl) {
          collectUrlsFromSitemapRecursive_(locEl.getText(), urlMap, visited);
        }
        return;
      }

      if (tagName === 'url') {
        // 通常の URL エントリ
        const locEl = child.getChild('loc', ns);
        if (!locEl) return;

        const locRaw = locEl.getText();
        const loc = decodeURIComponent(locRaw);

        const lmEl = child.getChild('lastmod', ns);
        const lmRaw = lmEl ? lmEl.getText() : '';
        const lastmod = normalizeLastmod_(lmRaw);

        urlMap.set(loc, lastmod);
      }
    });

  } catch (e) {
    Logger.log('Sitemap解析失敗: ' + sitemapUrl + ' / ' + e);
  }
}

/**
 * すべてのサイトマップから {url, lastmod} の配列を取得
 */
function getAllUrlsFromSitemaps_() {
  const settings = getSitemapSettings_();
  const urlMap = new Map();
  const visited = new Set();

  settings.forEach(({ sitemapUrl }) => {
    collectUrlsFromSitemapRecursive_(sitemapUrl, urlMap, visited);
  });

  const list = [];
  urlMap.forEach((lastmod, url) => {
    list.push({ url, lastmod });
  });

  return list;
}

/************************************************************
 * ▼ キャッシュシート / 差分ログシート
 ************************************************************/
function getOrCreateCacheSheet_() {
  const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  let sh = ss.getSheetByName(CACHE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CACHE_SHEET_NAME);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'URL','分類','lastmod','fetch日時',
      'title','metaDesc','H1',
      'entryBodyHtml','ogTitle','ogDesc','jsonLd'
    ]);
  }
  return sh;
}

function getOrCreateCacheDiffLogSheet_() {
  const ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  let sh = ss.getSheetByName(CACHE_DIFF_LOG);
  if (!sh) {
    sh = ss.insertSheet(CACHE_DIFF_LOG);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'URL','lastmod','fetch日時',
      'title','metaDesc','H1',
      'entryBodyHtml','ogTitle','ogDesc','jsonLd'
    ]);
  }
  return sh;
}

/**
 * キャッシュシート → URL → {row, lastmod, rowValues}
 */
function buildCacheMap_(cacheSh) {
  const map = new Map();
  const lastRow = cacheSh.getLastRow();
  if (lastRow < 2) return map;

  const values = cacheSh.getRange(2, 1, lastRow - 1, 11).getValues();
  values.forEach((row, i) => {
    const url = row[0];
    if (!url) return;
    map.set(url, {
      row: i + 2,
      lastmodTs: lastmodToTs_(row[2]), // ★ 比較用
      rowValues: row
    });

  });

  return map;
}

/************************************************************
 * ▼ キャッシュ1行生成（10列）
 ************************************************************/
function makeCacheRow_(url, lastmod, html) {

  const title    = extractTitle_(html);
  const metaDesc = extractMetaDesc_(html);
  const h1       = extractH1_(html);
  const entryHtml = extractEntryBodyStrict_(html);
  const ogTitle  = extractOgMeta_(html, 'og:title');
  const ogDesc   = extractOgMeta_(html, 'og:description');
  const jsonLd   = extractJsonLd_(html);

  return [
    url,
    '',            // ← 分類（SP / MC / SC / TP / NW / ST など）
    lastmod,
    now_(),
    title,
    metaDesc,
    h1,
    entryHtml,
    ogTitle,
    ogDesc,
    jsonLd
  ];
}

/************************************************************
 * ▼ 差分ログ：10列そのまま書き込み、差分セルだけ色付け
 ************************************************************/
function appendLogWithHighlights_(logSh, oldRow, newRow) {
  const rowIdx = logSh.getLastRow() + 1;
  logSh.getRange(rowIdx, 1, 1, newRow.length).setValues([newRow]);

  const highlight = '#fff2cc';

  for (let col = 1; col <= newRow.length; col++) {
    const oldVal = String(oldRow[col - 1] || '');
    const newVal = String(newRow[col - 1] || '');
    if (oldVal !== newVal) {
      logSh.getRange(rowIdx, col).setBackground(highlight);
    }
  }
}

/************************************************************
 * ▼ 初期キャッシュ構築（1回きり）
 *   - サイトマップに出てくる全URLを対象
 *   - まだキャッシュに無いURLだけ fetch
 *   - 4分30秒で安全停止（GAS制限対策）
 ************************************************************/
function initCacheFromSitemaps() {

  const MAX_SECONDS = 270; // ★ 4分30秒
  const start = Date.now();

  const cacheSh = getOrCreateCacheSheet_();
  const cacheMap = buildCacheMap_(cacheSh);

  const allUrls = getAllUrlsFromSitemaps_();

  let created = 0;

  for (const { url, lastmod } of allUrls) {
    
    if(shouldSkipUrl_(url)) continue;

    // ★ GAS制限対策：4分30秒で安全終了
    const elapsed = (Date.now() - start) / 1000;
    if (elapsed > MAX_SECONDS) {
      Logger.log(`安全停止：${created}件作成（${elapsed.toFixed(1)}秒経過）`);
      return;
    }

    // 既にキャッシュにあるURLはスキップ
    if (cacheMap.has(url)) continue;

    try {
      const res = UrlFetchApp.fetch(encodeURI(url), { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() >= 300) {
        Logger.log('初回fetch失敗: ' + url + ' / ' + res.getResponseCode());
        continue;
      }

      const html = res.getContentText('UTF-8');
      const row = makeCacheRow_(url, lastmod, html);
      cacheSh.appendRow(row);
      created++;

    } catch (e) {
      Logger.log('初回fetch例外: ' + url + ' / ' + e);
    }
  }

  Logger.log('initCacheFromSitemaps 完了：新規 ' + created + ' 件');
}


/************************************************************
 * ▼ 差分更新（定期実行用）
 *   - サイトマップ lastmod とキャッシュ lastmod を比較
 *   - 異なるURLだけ fetch
 *   - CacheDiffLog に新しい行を保存し、変わったセルをハイライト
 *   - 4分30秒で安全停止
 ************************************************************/
function updateCacheDiffFromSitemaps() {



  const MAX_SECONDS = 270; // ★ 4分30秒
  const start = Date.now();

  const cacheSh  = getOrCreateCacheSheet_();
  const logSh    = getOrCreateCacheDiffLogSheet_();
  const cacheMap = buildCacheMap_(cacheSh);

  const allUrls = getAllUrlsFromSitemaps_();

  // ★ lastmod 出現回数を集計（今回実行分）
  const lastmodCountMap = {};
  allUrls.forEach(({ lastmod }) => {
    if (!lastmod) return;
    lastmodCountMap[lastmod] = (lastmodCountMap[lastmod] || 0) + 1;
  });


  let updated = 0;
  let created = 0;

  for (const { url, lastmod } of allUrls) {
    
    if(shouldSkipUrl_(url)) continue;
    // ★ 実行時間チェック
    const elapsed = (Date.now() - start) / 1000;
    if (elapsed > MAX_SECONDS) {
      Logger.log(
        `安全停止：更新 ${updated} 件 / 新規 ${created} 件（${elapsed.toFixed(1)}秒）`
      );
      return;
    }

    const info = cacheMap.get(url);

    /* =========================
     * 新規URL
     * ========================= */
    if (!info) {
      try {
        const res = UrlFetchApp.fetch(encodeURI(url), {
          muteHttpExceptions: true,
          followRedirects: true
        });

        if (res.getResponseCode() >= 300) {
          Logger.log('新規fetch失敗: ' + url + ' / ' + res.getResponseCode());
          continue;
        }

        const html   = res.getContentText('UTF-8');
        const newRow = makeCacheRow_(url, lastmod, html);

        cacheSh.appendRow(newRow);
        created++;

      } catch (e) {
        Logger.log('新規fetch例外: ' + url + ' / ' + e);
      }

      continue;
    }

    /* =========================
     * lastmod 変化なし → スキップ
     * ========================= */
    const sitemapLastmodTs = lastmodToTs_(lastmod);

    // ★ lastmod が同一ならスキップ
    if (info.lastmodTs === sitemapLastmodTs) {
      continue;
    }

    // ★ 上位階層かつ lastmod が一括更新ならスキップ
    if (
      isTopLevelUrl_(url) &&
      lastmod &&
      lastmodCountMap[lastmod] >= 2
    ) {
      // 伝播更新とみなす
      continue;
    }


    /* =========================
     * lastmod 変化あり → 差分更新
     * ========================= */
    try {
      const res = UrlFetchApp.fetch(encodeURI(url), {
        muteHttpExceptions: true,
        followRedirects: true
      });

      if (res.getResponseCode() >= 300) {
        Logger.log('差分fetch失敗: ' + url + ' / ' + res.getResponseCode());
        continue;
      }

      const html   = res.getContentText('UTF-8');
      const newRow = makeCacheRow_(url, lastmod, html);
      const oldRow = info.rowValues;

      // ★ B列（分類）を旧行から引き継ぐ
      newRow[1] = oldRow[1];

      // 差分ログ（変更セルをハイライト）
      appendLogWithHighlights_(logSh, oldRow, newRow);

      // キャッシュ上書き
      cacheSh
        .getRange(info.row, 1, 1, newRow.length)
        .setValues([newRow]);
      
           


      updated++;

    } catch (e) {
      Logger.log('差分fetch例外: ' + url + ' / ' + e);
    }
  }

  Logger.log(
    'updateCacheDiffFromSitemaps 完了：更新 ' +
    updated + ' 件 / 新規 ' + created + ' 件'
  );
}
// ★ lastmod を「比較用 timestamp」に正規化
function lastmodToTs_(raw) {
  if (!raw) return 0;
  const d = new Date(raw);
  return isNaN(d) ? 0 : d.getTime();
}

function shouldSkipUrl_(url) {
  if (!url) return true;

  // 既存ルール
  if (url.includes('/author/')) return true;

  // ① 明示無視（完全一致：host+path）
  if (SKIP_URL_SET.has(normalizeSkipUrl_(url))) return true;

  // ② 深さ1のハブURLを自動除外（/column/ /blog/ /info/ 等）
  if (isHubTopPage_(url)) return true;

  return false;
}



function isTopLevelUrl_(url) {
  try {
    const path = new URL(url).pathname;
    const depth = path.split('/').filter(Boolean).length;
    return depth <= 1; // / または /blog/ /column/ など
  } catch (e) {
    return false;
  }
}
/************************************************************
 * ▼ 無視URL（処理対象外）＋ ハブURL（深さ1）自動除外
 ************************************************************/

// 明示的に無視したいURL（ドメイン順）
const SKIP_URLS = [
  'https://dokkai-labo.tokyo/blog/',
  'https://dokkai-labo.tokyo/col_blog/',
  'https://dokkai-labo.tokyo/info/',

  'https://eigo-rewrite.com/column/',

  'https://eng-support.com/column/',
  'https://eng-support.com/info/',

  'https://jukureiwa.com/info/',
  'https://jukureiwa.com/tips/',

  'https://juku-escot.com/c_blogs/',
  'https://juku-escot.com/column/',

  // www/非www 揺れ対策
  'https://www.kokugo-trigger.com/column/',
  'https://kokugo-trigger.com/column/',

  'https://koten-kakitsubata.jp/blog/',

  'https://natsui-sansu-juku.com/col_blog/',
  'https://natsui-sansu-juku.com/column/',

  'https://rika-quest.com/column/',
];

// 深さ1の「一覧・ハブ」スラッグ（/column/ /blog/ /info/ 等）を自動除外
const HUB_SLUGS = new Set([
  'column',
  'blog',
  'info',
  'tips',
  'col_blog',
  'c_blogs',
]);

function normalizeSkipUrl_(u) {
  if (!u) return '';
  try {
    const url = new URL(u);
    let path = url.pathname || '/';
    // 末尾スラッシュ吸収（/ だけは残す）
    if (path.length > 1 && path.endsWith('/')) path = path.slice(0, -1);
    // protocol は無視して host + path で比較
    return `${url.host}${path}`;
  } catch (e) {
    // URL()で落ちた場合の保険
    return String(u)
      .replace(/[?#].*$/, '')
      .replace(/\/$/, '')
      .replace(/^https?:\/\//, '')
      .replace(/\/{2,}/g, '/');
  }
}

const SKIP_URL_SET = new Set(SKIP_URLS.map(normalizeSkipUrl_));

function isHubTopPage_(url) {
  try {
    const path = new URL(url).pathname || '/';
    const segs = path.split('/').filter(Boolean); // ["column"] 等
    // 深さ1だけを対象（/column/ や /blog/ 等）
    if (segs.length !== 1) return false;
    return HUB_SLUGS.has(segs[0]);
  } catch (e) {
    return false;
  }
}

