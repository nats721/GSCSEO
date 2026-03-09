/************************************************************
 * 階層マスタ同期（安全装置付き）
 *
 * - HTMLキャッシュ を正とする
 * - ST / NW / "-" は除外
 * - 新規URLは追記
 * - 対象外URLは削除
 * - HTMLキャッシュから有効URLが取れない場合は何もしない
 ************************************************************/

const CACHE_SS_ID  = '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o';
const CACHE_SHEET  = 'HTMLキャッシュ';

const MASTER_SS_ID = '1U-XK7TVV8BWxo27iRqySKmh6SmlyZncjoC26ZAX8LjM';
const MASTER_SHEET = '階層マスタ';

// 除外分類
const EXCLUDE_TYPES = new Set(['TP','ST', 'NW', '-']);

function syncHierarchyMaster() {
  const cacheSS  = SpreadsheetApp.openById(CACHE_SS_ID);
  const masterSS = SpreadsheetApp.openById(MASTER_SS_ID);

  const cacheSh  = cacheSS.getSheetByName(CACHE_SHEET);
  const masterSh = masterSS.getSheetByName(MASTER_SHEET);

  // --- シートが取れない場合は何もしない ---
  if (!cacheSh || !masterSh) {
    Logger.log('対象シートが取得できないため処理中断');
    return;
  }

  const cacheValues  = cacheSh.getDataRange().getValues();
  const masterValues = masterSh.getDataRange().getValues();

  // --- HTMLキャッシュ側の有効URL抽出 ---
  const validUrls = new Set();

  for (let i = 1; i < cacheValues.length; i++) {
    const url  = cacheValues[i][0];
    const type = cacheValues[i][1];

    if (!url) continue;
    if (EXCLUDE_TYPES.has(type)) continue;

    validUrls.add(url);
  }

  // --- 有効URLが1件もない場合は何もしない ---
  if (validUrls.size === 0) {
    Logger.log('有効URLが0件のため処理を行いません');
    return;
  }

  // --- 階層マスタ側URL ---
  const masterUrlMap = new Map(); // url -> rowIndex

  for (let i = 1; i < masterValues.length; i++) {
    const url = masterValues[i][0];
    if (!url) continue;
    masterUrlMap.set(url, i + 1); // sheet row
  }

  // --- 追加対象 ---
  const appendRows = [];
  validUrls.forEach(url => {
    if (!masterUrlMap.has(url)) {
      appendRows.push([url]);
    }
  });

  if (appendRows.length > 0) {
    masterSh
      .getRange(masterSh.getLastRow() + 1, 1, appendRows.length, 1)
      .setValues(appendRows);
  }

  // --- 削除対象 ---
  const deleteRows = [];

  masterUrlMap.forEach((row, url) => {
    if (!validUrls.has(url)) {
      deleteRows.push(row);
    }
  });

  // 下から削除（行ズレ防止）
  deleteRows
    .sort((a, b) => b - a)
    .forEach(row => masterSh.deleteRow(row));
}
