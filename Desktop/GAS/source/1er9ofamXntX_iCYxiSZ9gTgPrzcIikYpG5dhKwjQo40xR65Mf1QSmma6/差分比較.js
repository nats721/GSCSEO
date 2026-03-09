/*************************************************************
 * Snapshot：差分監視対応・最新状態1行保持・ChangeLog記録版
 *
 * 前提：
 * - SNAP_MGMT_SHEET_ID / SNAP_RESULT_SHEET_ID / SNAP_SHEET_NAME /
 *   SNAP_MAX_SECONDS は別ファイル等で定義済み
 * - SNAP_extractMainContent_ / SNAP_extractTitle_ など
 *   HTML解析系ユーティリティ関数も既存のものを利用
 *
 * 仕様：
 * - Snapshotシート：URLごとに常に「最新状態を1行だけ」保持（15列）
 * - ChangeLogシート：タイトル等13項目に変化があった時だけ1行追記
 * - 取得日時(列N=14)：差分があった時だけ更新
 * - チェック日時(列O=15)：クロール実行ごとに必ず更新
 * - 処理順：未登録URL → checkedAt（チェック日時）が古いURLの順
 *************************************************************/

const DIFF_SHEET_NAME = 'ChangeLog';

/*************************************************************
 * メイン処理：差分監視＋Snapshot上書き
 *************************************************************/
function runSnapshotCrawl2() {

  const startTime = Date.now();
  const mgmtSS = SpreadsheetApp.openById(SNAP_MGMT_SHEET_ID);
  const snapSS = SpreadsheetApp.openById(SNAP_RESULT_SHEET_ID);

  const snapSh = SNAP_getOrCreateSnapshotSheet_(snapSS);
  const diffSh = SNAP_getOrCreateDiffSheet_(snapSS);

  // URL → {rowIndex, fetchedAt, checkedAt}
  const existingUrlMap = SNAP_buildExistingMap_(snapSh);

  // 管理ブックから全URL取得
  const allUrls = SNAP_collectAllUrls_(mgmtSS);

  // ★ 未登録URL → checkedAtが古い順 に並び替え
  const orderedUrls = SNAP_sortByCheckedAt_(allUrls, existingUrlMap);

  for (const url of orderedUrls) {

    // タイムアウト安全終了
    if ((Date.now() - startTime) / 1000 > SNAP_MAX_SECONDS) {
      Logger.log('安全終了（タイムアウト）');
      return;
    }

    const newRow = SNAP_fetchSnapshotForUrl_(url);
    if (!newRow) continue;

    const rowInfo = existingUrlMap.get(url);
    const nowTs = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    if (rowInfo) {
      // 既存行あり → 差分比較
      const oldRow = snapSh.getRange(rowInfo.rowIndex, 1, 1, 15).getValues()[0];

      // 取得日時・チェック日時を除く 1〜12列だけで差分判定
      const isChanged = SNAP_writeDiffIfChanged_(diffSh, oldRow, newRow);

      // newRow をベースに書き込み用行データを作成
      const writeRow = [...newRow];

      // 取得日時（列14）：内容に変化があったときだけ更新
      writeRow[13] = isChanged ? nowTs : oldRow[13];

      // チェック日時（列15）：毎回更新
      writeRow[14] = nowTs;

      snapSh.getRange(rowInfo.rowIndex, 1, 1, 15).setValues([writeRow]);

    } else {
      // 新規URL（初回登録）→ 差分比較不要
      newRow[13] = nowTs; // 取得日時
      newRow[14] = nowTs; // チェック日時

      snapSh.appendRow(newRow);

      const newIndex = snapSh.getLastRow();
      existingUrlMap.set(url, {
        rowIndex:  newIndex,
        fetchedAt: newRow[13],
        checkedAt: newRow[14]
      });
    }
  }

  Logger.log('Snapshot 完了');
}

/*************************************************************
 * 管理ブックから全URLを取得
 *************************************************************/
function SNAP_collectAllUrls_(mgmtSS) {
  const settingSh = mgmtSS.getSheetByName('設定');
  if (!settingSh) return [];

  const last = settingSh.getLastRow();
  if (last < 2) return [];

  const domains = settingSh.getRange(2, 1, last - 1, 1).getValues().flat();
  const urls = [];

  domains.forEach(domain => {
    if (!domain) return;
    const sh = mgmtSS.getSheetByName(domain);
    if (!sh) return;

    const lr = sh.getLastRow();
    if (lr < 2) return;

    const arr = sh.getRange(2, 1, lr - 1, 1).getValues().flat();
    arr.forEach(u => {
      if (typeof u === 'string' && u.trim() !== '') {
        urls.push(u.trim());
      }
    });
  });

  return urls;
}

/*************************************************************
 * Snapshotシート（15列仕様）を取得 or 作成
 * 旧仕様（14列）だった場合は15列目に「チェック日時」を追加
 *************************************************************/
function SNAP_getOrCreateSnapshotSheet_(ss) {
  let sh = ss.getSheetByName(SNAP_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SNAP_SHEET_NAME);
  }

  if (sh.getLastRow() === 0) {
    // 新規シート → 15列ヘッダを作成
    sh.getRange(1, 1, 1, 15).setValues([[
      'URL', 'タイトル', 'メタディスクリプション', 'H1',
      'H2数', 'H3数', '本文文字数', '内部リンク数', 'CTA数',
      'OGPタイトル', 'OGP説明', 'OGP画像URL', '構造化データJSON-LD数',
      '取得日時', 'チェック日時'
    ]]);
  } else {
    // 旧シートで 15列目ヘッダが無ければ追加
    const lastCol = sh.getLastColumn();
    if (lastCol < 15) {
      sh.getRange(1, 15).setValue('チェック日時');
    }
  }

  return sh;
}

/*************************************************************
 * ChangeLog シート生成
 *************************************************************/
function SNAP_getOrCreateDiffSheet_(ss) {
  let sh = ss.getSheetByName(DIFF_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DIFF_SHEET_NAME);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 6).setValues([[
      'URL', '項目', '旧値', '新値', '差分日時', 'メモ'
    ]]);
  }

  return sh;
}

/*************************************************************
 * Snapshot 既存行 → URL→{rowIndex,fetchedAt,checkedAt} にマップ化
 *************************************************************/
function SNAP_buildExistingMap_(sh) {
  const map  = new Map();
  const last = sh.getLastRow();
  if (last < 2) return map;

  const rows = sh.getRange(2, 1, last - 1, 15).getValues();

  rows.forEach((r, i) => {
    const url = r[0];
    if (!url) return;
    map.set(url, {
      rowIndex:  i + 2,
      fetchedAt: r[13],  // 取得日時
      checkedAt: r[14]   // チェック日時
    });
  });

  return map;
}

/*************************************************************
 * URL一覧を「未登録 → checkedAtの古い順」に並び替え
 *************************************************************/
function SNAP_sortByCheckedAt_(allUrls, existingUrlMap) {

  const list = allUrls.map(url => {
    const info = existingUrlMap.get(url);

    // ★ checkedAt を「timestamp（ミリ秒）」に正規化
    let tsNum = 0;
    if (info && info.checkedAt) {
      const v = info.checkedAt;
      tsNum = (v instanceof Date)
        ? v.getTime()
        : new Date(v).getTime();  // 文字列でもOKに
    }

    return {
      url: url,
      hasRow: !!info,
      tsNum: tsNum        // ← 数値で比較する
    };
  });

  list.sort((a,b) => {
    if (!a.hasRow && b.hasRow) return -1;
    if (!b.hasRow && a.hasRow) return 1;

    if (!a.hasRow && !b.hasRow) return 0;

    // checkedAt が無い（tsNum=0）のものを古い扱い
    return a.tsNum - b.tsNum;
  });

  return list.map(x => x.url);
}


/*************************************************************
 * 差分を書き込む
 * 取得日時・チェック日時は比較対象外
 *************************************************************/
function SNAP_writeDiffIfChanged_(diffSh, oldRow, newRow) {
  const columns = [
    'URL',                   // 0
    'タイトル',              // 1
    'メタディスクリプション',// 2
    'H1',                    // 3
    'H2数',                  // 4
    'H3数',                  // 5
    '本文文字数',            // 6
    '内部リンク数',          // 7
    'CTA数',                 // 8
    'OGPタイトル',           // 9
    'OGP説明',               //10
    'OGP画像URL',           //11
    '構造化データJSON-LD数' //12
    // 13: 取得日時（比較しない）
    // 14: チェック日時（比較しない）
  ];

  const url = newRow[0];
  const ts  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  const rows = [];
  let changed = false;

  // 比較するのは 1〜12列のみ（タイトル〜JSON-LD数）
  for (let i = 1; i <= 12; i++) {
    if (String(oldRow[i]) !== String(newRow[i])) {
      changed = true;
      rows.push([
        url,
        columns[i],
        oldRow[i],
        newRow[i],
        ts,
        ''
      ]);
    }
  }

  if (rows.length > 0) {
    diffSh.getRange(diffSh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }

  return changed;
}

/*************************************************************
 * fetch → HTML解析 → 15項目配列に整形
 * （HTML解析系ユーティリティは既存のものを利用）
 *************************************************************/
function SNAP_fetchSnapshotForUrl_(url) {
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (res.getResponseCode() >= 300) return null;

    const html     = res.getContentText('UTF-8');
    const mainHtml = SNAP_extractMainContent_(html);  // 本文領域のみ

    return [
      url,
      SNAP_extractTitle_(html),
      SNAP_extractMetaDescription_(html),
      SNAP_extractH1_(html),
      SNAP_countTag_(mainHtml, 'h2'),
      SNAP_countTag_(mainHtml, 'h3'),
      SNAP_estimateTextLength_(mainHtml),
      SNAP_countInternalLinks_(mainHtml, url),
      SNAP_detectCTA_(html, mainHtml),
      SNAP_extractOgMeta_(html, 'og:title'),
      SNAP_extractOgMeta_(html, 'og:description'),
      SNAP_extractOgMeta_(html, 'og:image'),
      SNAP_countJsonLd_(html),
      '',   // 取得日時（ここでは空。runSnapshotCrawl2側でセット）
      ''    // チェック日時（ここでは空。runSnapshotCrawl2側でセット）
    ];

  } catch (e) {
    Logger.log('ERROR fetch: ' + url + ' / ' + e);
    return null;
  }
}
