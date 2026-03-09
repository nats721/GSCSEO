/*************************************************************
 * 全URLメタ情報 Snapshot 取得スクリプト（最終版）
 *
 * 元ブック：各ブランドサイト管理__メイン
 *   ID: 1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs
 *   - シート「設定」A列：ドメイン名
 *   - 各ドメイン名シートの A列：URL
 *
 * 出力ブック：Snapshot保存用
 *   ID: 1SCYQd7_cobOYqT0r_F0IqD4tE4rYn5y8Eup4PJTH8es
 *   シート名：Snapshot
 *
 * 1URLごとに取得する項目：
 *  A: URL
 *  B: タイトル（<title>）
 *  C: メタディスクリプション（meta[name=description]）
 *  D: H1（最初の<h1>）
 *  E: H2数（本文領域のみ）
 *  F: H3数（本文領域のみ）
 *  G: 本文文字数（タグ除去後：本文領域のみ）
 *  H: 内部リンク数（同一ドメイン or 相対パス：本文領域のみ／同一hrefは1回）
 *  I: CTA数（本文領域内のCTAブロック数）
 *       - <section class="...veu_cta..."> を1つで1カウント
 *       - style に「border: npx solid ...」を持つ <div> 内に
 *         <a href="..."> が1つ以上あるブロックを1カウント
 *  J: OGPタイトル（og:title）
 *  K: OGP説明（og:description）
 *  L: OGP画像URL（og:image）
 *  M: 構造化データJSON-LD数（application/ld+json の script 個数）
 *  N: 取得日時（Asia/Tokyo）
 *************************************************************/

const SNAP_MGMT_SHEET_ID   = '1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs';
const SNAP_RESULT_SHEET_ID = '1SCYQd7_cobOYqT0r_F0IqD4tE4rYn5y8Eup4PJTH8es';
const SNAP_SHEET_NAME      = 'Snapshot';

const SNAP_MAX_SECONDS = 270;   // 約4分30秒
const SNAP_BATCH_SIZE  = 50;    // まとめて書き込む行数

/*************************************************************
 * エントリーポイント
 *************************************************************/
function runSnapshotCrawl() {

  const startTime = new Date().getTime();

  const mgmtSS = SpreadsheetApp.openById(SNAP_MGMT_SHEET_ID);
  const snapSS = SpreadsheetApp.openById(SNAP_RESULT_SHEET_ID);
  const snapSh = SNAP_getOrCreateSnapshotSheet_(snapSS);

  // 既存URL（重複取得防止）
  const existingUrlSet = SNAP_getExistingUrlSet_(snapSh);

  // 「設定」シートからドメイン一覧取得
  const settingSh = mgmtSS.getSheetByName('設定');
  if (!settingSh) {
    Logger.log('設定シートが見つかりません');
    return;
  }

  const lastRow = settingSh.getLastRow();
  if (lastRow < 2) {
    Logger.log('設定シートにドメインがありません');
    return;
  }

  const domainValues = settingSh.getRange(2, 1, lastRow - 1, 1).getValues();
  const domains = domainValues
    .map(r => r[0])
    .filter(v => typeof v === 'string' && v.trim() !== '');

  if (domains.length === 0) {
    Logger.log('有効なドメインがありません');
    return;
  }

  let rowsToAppend   = [];
  let processedCount = 0;

  // forEach ではなく for...of（途中 return で安全終了するため）
  for (const domainName of domains) {

    const domainSheet = mgmtSS.getSheetByName(domainName);
    if (!domainSheet) {
      Logger.log('ドメインシートが見つかりません: ' + domainName);
      continue;
    }

    const lastUrlRow = domainSheet.getLastRow();
    if (lastUrlRow < 2) {
      Logger.log('URLがありません: ' + domainName);
      continue;
    }

    const urlValues = domainSheet.getRange(2, 1, lastUrlRow - 1, 1).getValues();
    const urls = urlValues
      .map(r => r[0])
      .filter(u => typeof u === 'string' && u.trim() !== '');

    for (const url of urls) {

      // タイムアウト安全終了
      if (SNAP_isTimeOver_(startTime)) {
        Logger.log('タイムアウト安全終了: ' + processedCount + ' URLまで処理');

        if (rowsToAppend.length > 0) {
          SNAP_appendRows_(snapSh, rowsToAppend);
        }
        return;
      }

      // 既に取得済みならスキップ
      if (existingUrlSet.has(url)) {
        continue;
      }

      try {
        const row = SNAP_fetchSnapshotForUrl_(url);
        if (row) {
          rowsToAppend.push(row);
          existingUrlSet.add(url);
          processedCount++;
        }
      } catch (e) {
        Logger.log('ERROR: ' + url + ' で例外発生: ' + e);
      }

      // バッチ書き込み
      if (rowsToAppend.length >= SNAP_BATCH_SIZE) {
        SNAP_appendRows_(snapSh, rowsToAppend);
        rowsToAppend = [];
      }
    }
  }

  // 端数分を書き込み
  if (rowsToAppend.length > 0) {
    SNAP_appendRows_(snapSh, rowsToAppend);
  }

  Logger.log('Snapshot 完了: ' + processedCount + ' URL');
}

/*************************************************************
 * タイムアウト判定
 *************************************************************/
function SNAP_isTimeOver_(startTime) {
  const elapsedSec = (new Date().getTime() - startTime) / 1000;
  return elapsedSec > SNAP_MAX_SECONDS;
}

/*************************************************************
 * Snapshotシート生成（なければ作成）
 *************************************************************/
function SNAP_getOrCreateSnapshotSheet_(ss) {
  let sh = ss.getSheetByName(SNAP_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SNAP_SHEET_NAME);
  }

  // ヘッダがまだ無い場合のみヘッダ行を作成
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 14).setValues([[
      'URL',                          // A
      'タイトル',                     // B
      'メタディスクリプション',       // C
      'H1',                           // D
      'H2数',                         // E
      'H3数',                         // F
      '本文文字数',                   // G
      '内部リンク数',                 // H
      'CTA数',                        // I
      'OGPタイトル',                  // J
      'OGP説明',                      // K
      'OGP画像URL',                   // L
      '構造化データJSON-LD数',       // M
      '取得日時'                      // N
    ]]);
  }

  return sh;
}

/*************************************************************
 * Snapshotに既存のURLセットを作成
 *************************************************************/
function SNAP_getExistingUrlSet_(sh) {
  const lastRow = sh.getLastRow();
  const set = new Set();
  if (lastRow < 2) return set;

  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // A列(URL)
  vals.forEach(r => {
    const url = r[0];
    if (url) set.add(String(url));
  });
  return set;
}

/*************************************************************
 * 行追加（バッチ）
 *************************************************************/
function SNAP_appendRows_(sh, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

/*************************************************************
 * 1URL分のSnapshotを取得
 *************************************************************/
function SNAP_fetchSnapshotForUrl_(url) {

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('HTTPエラー: ' + code + ' URL=' + url);
    return null;
  }

  const html     = res.getContentText('UTF-8');
  const mainHtml = SNAP_extractMainContent_(html);  // 本文領域のみ

  const title    = SNAP_extractTitle_(html);
  const metaDesc = SNAP_extractMetaDescription_(html);
  const h1       = SNAP_extractH1_(html);

  const h2Count  = SNAP_countTag_(mainHtml, 'h2');
  const h3Count  = SNAP_countTag_(mainHtml, 'h3');

  const textLen  = SNAP_estimateTextLength_(mainHtml);
  const iLinks   = SNAP_countInternalLinks_(mainHtml, url);
  //const ctaCount = SNAP_detectCTA_(mainHtml);
  const ctaCount = SNAP_detectCTA_(html,mainHtml);

  const ogTitle  = SNAP_extractOgMeta_(html, 'og:title');
  const ogDesc   = SNAP_extractOgMeta_(html, 'og:description');
  const ogImg    = SNAP_extractOgMeta_(html, 'og:image');

  const jsonld   = SNAP_countJsonLd_(html);
  const modified = SNAP_extractModifiedDateFromJsonLd_(html);

  return [
    url,
    title,
    metaDesc,
    h1,
    h2Count,
    h3Count,
    textLen,
    iLinks,
    ctaCount,
    ogTitle,
    ogDesc,
    ogImg,
    jsonld,
    modified   // ←これが更新日時(JSON-LD)
  ];

}

function SNAP_extractModifiedDateFromJsonLd_(html) {
  const regex = /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let match;

  while ((match = regex.exec(html)) !== null) {
    try {
      const json = JSON.parse(match[1]);

      // 1階層・Article系
      if (json.dateModified) return json.dateModified;

      // @graph 対応
      if (json['@graph']) {
        for (const item of json['@graph']) {
          if (item.dateModified) return item.dateModified;
        }
      }

    } catch (_) {}
  }

  return ''; // 見つからなければ空
}


/*************************************************************
 * 本文領域抽出
 * - div.entry-body があれば「全部結合」
 * - なければ main / article
 * - それもなければ全文
 *************************************************************/
function SNAP_extractMainContent_(html) {
  if (!html) return '';

  // entry-body 開始
  const startRe =
    /<div[^>]+class=["'][^"']*entry-body[^"']*["'][^>]*>/i;
  const startMatch = html.match(startRe);
  if (!startMatch) return '';

  const startTag = startMatch[0];
  const startIndex = html.indexOf(startTag) + startTag.length;

  // 本文終了マーカー候補（Lightning 固定順）
  const endMarkers = [
    '<section class="veu_cta',
    '<div class="veu_socialSet'
  ];

  let endIndex = -1;

  for (const marker of endMarkers) {
    const pos = html.indexOf(marker, startIndex);
    if (pos !== -1) {
      if (endIndex === -1 || pos < endIndex) {
        endIndex = pos;
      }
    }
  }

  // Fallback：それでも無い rare ケース（固定ページなど）
  if (endIndex === -1) {
    const footerPos = html.indexOf('<div class="entry-footer"', startIndex);
    if (footerPos !== -1) endIndex = footerPos;
    else endIndex = html.length;
  }

  return html.substring(startIndex, endIndex);
}





function SNAP_detectCTA_(html, mainHtml) {
  if (!html) return 0;

  let count = 0;

  /***********************************************************
   * ① Lightning / ExUnit CTA（veu_cta） → 全文対象
   ***********************************************************/
  const exUnitRe =
    /<(?:div|section)[^>]+class=["'][^"']*\bveu_cta\b[^"']*["'][^>]*>/gi;

  const exMatches = html.match(exUnitRe);
  if (exMatches) {
    count += exMatches.length;
  }

  /***********************************************************
   * ② 独自CTA（border: npx solid ...） → 本文(mainHtml)のみ対象
   ***********************************************************/
  if (mainHtml) {
    const customCtaRe =
      /<div[^>]+style=["'][^"']*border[^:]*:\s*\d+px\s+solid[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi;

    let m;
    while ((m = customCtaRe.exec(mainHtml)) !== null) {
      const blockHtml = m[1];

      // 本文内にリンクがあるブロックだけカウント
      if (/<a[^>]+href=["'][^"']+["'][^>]*>[\s\S]*?<\/a>/i.test(blockHtml)) {
        count++;
      }
    }
  }

  /***********************************************************
   * ③ カード型CTA（Bootstrap/UI系） → 本文(mainHtml)のみ対象
   ***********************************************************/
  if (mainHtml) {
    const cardCtaRe =
      /<div[^>]+class=["'][^"']*\bcard\b[^"']*["'][^>]*>[\s\S]*?<a[^>]+class=["'][^"']*\bbtn\b[^"']*["'][^>]*>[\s\S]*?<\/a>/gi;

    const cardMatches = mainHtml.match(cardCtaRe);
    if (cardMatches) {
      count += cardMatches.length;
    }
  }


  return count;
}



/*************************************************************
 * OGP meta 抽出（og:title / og:description / og:image）
 *************************************************************/
function SNAP_extractOgMeta_(html, propertyName) {
  if (!html) return '';

  const pattern = '<meta[^>]+property=[\'"]' + propertyName + '[\'"][^>]*>';
  const re = new RegExp(pattern, 'i');
  const m = html.match(re);
  if (!m) return '';

  const tag = m[0];
  const cm = tag.match(/content=["']([\s\S]*?)["']/i);
  if (!cm) return '';

  return SNAP_decodeHtml_(cm[1].trim());
}

/*************************************************************
 * JSON-LD 構造化データ数カウント
 *************************************************************/
function SNAP_countJsonLd_(html) {
  if (!html) return 0;
  const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script>/gi;
  let count = 0;
  while (re.exec(html) !== null) {
    count++;
  }
  return count;
}

/*************************************************************
 * 内部リンク数カウント（本文HTMLを対象・同一hrefは1回だけ）
 *************************************************************/
function SNAP_countInternalLinks_(html, pageUrl) {
  if (!html || !pageUrl) return 0;

  const hostMatch = pageUrl.match(/^https?:\/\/([^\/]+)/i);
  if (!hostMatch) return 0;
  const host = hostMatch[1];

  const re = /<a [^>]*href=["']([^"']+)["'][^>]*>/gi;
  const seen = new Set();
  let count = 0;
  let m;

  while ((m = re.exec(html)) !== null) {
    let href = m[1];
    if (!href) continue;

    href = href.trim();
    if (
      href.startsWith('#') ||
      href.toLowerCase().startsWith('mailto:') ||
      href.toLowerCase().startsWith('tel:') ||
      href.toLowerCase().startsWith('javascript:')
    ) {
      continue;
    }

    let isInternal = false;

    if (href.startsWith('http')) {
      const hm = href.match(/^https?:\/\/([^\/]+)/i);
      if (hm && hm[1] === host) {
        isInternal = true;
      }
    } else if (href.startsWith('/')) {
      // ルート相対パスは内部リンク扱い
      isInternal = true;
    } else {
      // "./" "../" 等はここでは除外（必要なら足す）
      continue;
    }

    if (!isInternal) continue;

    const key = href.split('#')[0];
    if (!seen.has(key)) {
      seen.add(key);
      count++;
    }
  }

  return count;
}

/*************************************************************
 * 汎用 HTML Utility
 *************************************************************/
function SNAP_extractTitle_(html) {
  const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!m) return '';
  return SNAP_decodeHtml_(SNAP_stripTags_(m[1]).trim());
}

function SNAP_extractMetaDescription_(html) {
  const m = html.match(/<meta[^>]+name=["']description["'][^>]*>/i);
  if (!m) return '';
  const tag = m[0];
  const cm = tag.match(/content=["']([\s\S]*?)["']/i);
  if (!cm) return '';
  return SNAP_decodeHtml_(cm[1].trim());
}

function SNAP_extractH1_(html) {
  const m = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  if (!m) return '';
  return SNAP_decodeHtml_(SNAP_stripTags_(m[1]).trim());
}

function SNAP_countTag_(html, tagName) {
  if (!html) return 0;
  const re = new RegExp('<' + tagName + '\\b[^>]*>', 'gi');
  const matches = html.match(re);
  return matches ? matches.length : 0;
}

function SNAP_estimateTextLength_(html) {
  if (!html) return 0;

  let text = html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '');

  text = SNAP_stripTags_(text);
  text = SNAP_decodeHtml_(text);
  text = text.replace(/\s+/g, ' ').trim();

  return text.length;
}

function SNAP_stripTags_(str) {
  if (!str) return '';
  return str.replace(/<[^>]+>/g, '');
}

function SNAP_decodeHtml_(str) {
  if (!str) return '';
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, '\'')
    .replace(/&nbsp;/g, ' ');
}



