/*************************************************************
 * GSC 改善「効果検証」専用スクリプト
 * - 元の管理シート：各ブランドサイト管理__メイン
 * - 対象：更新日が「GSC最新日から 10〜15日前」のURL
 * - 比較期間：0〜7日前（after） vs 8〜14日前（before）
 * - 結果は評価用スプレッドシートに append
 *************************************************************/

// ★元の管理ブック（各ブランドサイト管理__メイン）
const EVAL_MGMT_SHEET_ID = '1yPT0gT_Hc3tFW3yFhiOFq7O6LxJ9E1NGmmQEBWtVRvs';

// ★評価出力用ブック
const EVAL_RESULT_SHEET_ID = '1Jg24JG8rWeNI8H9KVZNMEBHUQRuQnTIDUP9kMCNA_OE';

// 評価を書き込むシート名（なければ作る）
const EVAL_SHEET_NAME = '評価';

// 日付計算用
const EVAL_DAY_MS = 86400 * 1000;

/*************************************************************
 * ENTRY POINT
 *************************************************************/
function runGSCUpdateEvaluation() {
  const mgmtSS  = SpreadsheetApp.openById(EVAL_MGMT_SHEET_ID);
  const evalSS  = SpreadsheetApp.openById(EVAL_RESULT_SHEET_ID);
  const evalSh  = EVAL_getOrCreateEvalSheet_(evalSS);

  // ドメイン一覧（シート名がドメインのものだけ）
  const domainSheets = mgmtSS.getSheets()
    .map(sh => sh.getName())
    .filter(n => n.includes('.') && n !== '設定');

  if (domainSheets.length === 0) {
    Logger.log('ドメインシートが見つかりません');
    return;
  }

  // GSCの最新利用可能日を、最初のドメインから取得
  const firstProperty = 'sc-domain:' + domainSheets[0];
  const latestDate = EVAL_detectLatestAvailableDate_(firstProperty);
  Logger.log('latest GSC date = ' + EVAL_fmtDate_(latestDate));

  // 「更新から14〜21日経過」判定用レンジ
  const lastmodFrom = EVAL_shiftDate_(latestDate, -40); // 15日前
  const lastmodTo   = EVAL_shiftDate_(latestDate, -7); // 10日前

  // 管理ブックから対象URLを抽出
  const targets = [];
  domainSheets.forEach(domainName => {
    const sh = mgmtSS.getSheetByName(domainName);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // A:URL, D:更新日(lastmod) を想定
    const values = sh.getRange(2, 1, lastRow - 1, 5).getValues();

    values.forEach((row, idx) => {
      const url     = row[0];
      const lastmod = row[4]; // D列

      if (!url || !(lastmod instanceof Date)) return;

      if (lastmod >= lastmodFrom && lastmod <= lastmodTo) {
        targets.push({
          domain: domainName,
          url,
          lastmod
        });
      }
    });
  });

  if (targets.length === 0) {
    Logger.log('評価対象となるURLがありません（更新日10〜15日前なし）');
    return;
  }

  // 更新日でソート
  targets.sort((a, b) => a.lastmod - b.lastmod);

  // すでに評価済みの URL+更新日 はスキップ（重複防止）
  const existingKeys = EVAL_getExistingKeys_(evalSh); // Set<"url|yyy-MM-dd">

  const beforeRange = {
    start: EVAL_shiftDate_(latestDate, -14), // 8〜14日前
    end:   EVAL_shiftDate_(latestDate, -8)
  };
  const afterRange = {
    start: EVAL_shiftDate_(latestDate, -7),  // 0〜7日前
    end:   latestDate
  };

  // // ---- 14日 vs 14日 方式 ----
  // const beforeRange = {
  //   start: EVAL_shiftDate_(latestDate, -28), // 14〜27日前
  //   end:   EVAL_shiftDate_(latestDate, -14)
  // };
  // const afterRange = {
  //   start: EVAL_shiftDate_(latestDate, -14), // 0〜13日前
  //   end:   latestDate
  // };


  const latestDateStr = EVAL_fmtDate_(latestDate);
  const rowsToAppend = [];

  targets.forEach(t => {
    const key = t.url + '|' + EVAL_fmtDate_(t.lastmod);
    if (existingKeys.has(key)) {
      // 既に同じURL+更新日で記録済みならスキップ
      return;
    }

    const property = 'sc-domain:' + t.domain;

    const before = EVAL_fetchPageMetrics_(property, t.url, beforeRange);
    const after  = EVAL_fetchPageMetrics_(property, t.url, afterRange);

    const diffClicks = after.clicks - before.clicks;
    const diffCtr    = after.ctr    - before.ctr;
    const diffPos    = before.position - after.position; // +なら順位改善

    const result = EVAL_judgeLabel_(before, after, diffClicks, diffCtr, diffPos);
    
    rowsToAppend.push([
      t.domain,
      t.url,
      EVAL_fmtDateTime_(t.lastmod),
      latestDateStr,
      EVAL_fmtDate_(beforeRange.start),
      EVAL_fmtDate_(beforeRange.end),
      EVAL_fmtDate_(afterRange.start),
      EVAL_fmtDate_(afterRange.end),
      before.clicks,
      after.clicks,
      before.ctr,
      after.ctr,
      before.position,
      after.position,
      diffClicks,
      diffCtr,
      diffPos,
      result,
      key  // 内部用キー（URL+更新日）
    ]);
  });

  if (rowsToAppend.length === 0) {
    Logger.log('新規に書き込む評価行がありません');
    return;
  }

  // 追記
  const startRow = evalSh.getLastRow() + 1;
  evalSh.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
        .setValues(rowsToAppend);

  Logger.log('評価行を ' + rowsToAppend.length + ' 件追加しました');
}

/*************************************************************
 * 評価シートの取得・ヘッダ生成
 *************************************************************/
function EVAL_getOrCreateEvalSheet_(ss) {
  let sh = ss.getSheetByName(EVAL_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(EVAL_SHEET_NAME);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 19).setValues([[
      'ドメイン',
      'URL',
      '記事更新日',
      'GSCデータ基準日',
      '改善前：開始日（8〜14日前）',
      '改善前：終了日',
      '改善後：開始日（0〜7日前）',
      '改善後：終了日',
      'クリック数（改善前）',
      'クリック数（改善後）',
      'CTR（改善前）',
      'CTR（改善後）',
      '平均順位（改善前）',
      '平均順位（改善後）',
      'クリック差分',
      'CTR差分',
      '順位改善度（+で改善）',
      '判定',
      '識別キー（内部管理用）'
    ]]);

  }

  return sh;
}

/*************************************************************
 * 既存評価行のキー集合を取得（重複防止用）
 *************************************************************/
function EVAL_getExistingKeys_(sh) {
  const lastRow = sh.getLastRow();
  const set = new Set();
  if (lastRow < 2) return set;

  // Key は 19列目に入れている
  const vals = sh.getRange(2, 19, lastRow - 1, 1).getValues();
  vals.forEach(r => {
    const key = r[0];
    if (key) set.add(String(key));
  });
  return set;
}

/*************************************************************
 * GSC最新日検出（過去14日ベース）
 *************************************************************/
function EVAL_detectLatestAvailableDate_(property) {
  const token = ScriptApp.getOAuthToken();
  const url =
    'https://searchconsole.googleapis.com/webmasters/v3/sites/' +
    encodeURIComponent(property) + '/searchAnalytics/query';

  const today = new Date();
  const payload = {
    startDate: EVAL_fmtDate_(EVAL_shiftDate_(today, -14)),
    endDate:   EVAL_fmtDate_(today),
    dimensions: ['date'],
    rowLimit: 300
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + token }
  });

  const json = JSON.parse(res.getContentText());
  let latest = null;
  (json.rows || []).forEach(r => {
    if (r.clicks > 0) {
      latest = r.keys[0]; // 'YYYY-MM-DD'
    }
  });

  return latest ? new Date(latest) : EVAL_shiftDate_(today, -4);
}

/*************************************************************
 * ページ別メトリクス取得（Before / After）
 *************************************************************/
function EVAL_fetchPageMetrics_(property, url, range) {
  const token = ScriptApp.getOAuthToken();
  const endpoint =
    'https://searchconsole.googleapis.com/webmasters/v3/sites/' +
    encodeURIComponent(property) + '/searchAnalytics/query';

  // 日本語URLなどを部分エンコード
  const encodedUrl = url.replace(
    /[^\w\-\.~:\/\?#\[\]@!$&'()*+,;=%]/g,
    c => encodeURIComponent(c)
  );

  const payload = {
    startDate: EVAL_fmtDate_(range.start),
    endDate:   EVAL_fmtDate_(range.end),
    dimensions: ['page'],
    dimensionFilterGroups: [{
      filters: [{
        dimension: 'page',
        operator: 'equals',
        expression: encodedUrl
      }]
    }],
    rowLimit: 1
  };

  const res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + token }
  });

  const json = JSON.parse(res.getContentText());
  const row = (json.rows && json.rows[0]) || null;

  if (!row) {
    return {
      clicks: 0,
      impressions: 0,
      ctr: 0,
      position: 0
    };
  }

  return {
    clicks: row.clicks || 0,
    impressions: row.impressions || 0,
    ctr: row.ctr || 0,
    position: row.position || 0
  };
}

/*************************************************************
 * ユーティリティ
 *************************************************************/
function EVAL_shiftDate_(dt, days) {
  return new Date(dt.getTime() + days * EVAL_DAY_MS);
}

function EVAL_fmtDate_(dt) {
  return Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function EVAL_fmtDateTime_(dt) {
  return Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}


/*************************************************************
 * 判定ロジック（細かい分類）
 *************************************************************/
function EVAL_judgeLabel_(before, after, diffClicks, diffCtr, diffPos) {
  const totalClicks = before.clicks + after.clicks;

  // クリックがほとんどない → そもそも判定しづらい
  if (totalClicks < 5) {
    return 'データ少（クリック合計5未満・暫定）';
  }

  // 大きく改善
  if (diffClicks >= 3 && diffCtr > 0 && diffPos > 0) {
    return '総合改善（クリック・CTR・順位すべて向上）';
  }

  // クリックもCTRも増えているが、順位はあまり動いていない
  if (diffClicks >= 3 && diffCtr > 0) {
    return '流入改善（クリック・CTR向上／順位はほぼ変わらず）';
  }

  // 大きく悪化
  if (diffClicks <= -3 && diffCtr < 0 && diffPos < 0) {
    return '総合悪化（クリック・CTR・順位すべて悪化）';
  }

  // クリック減＋CTR悪化（順位はそこまで変わっていない）
  if (diffClicks <= -3 && diffCtr < 0) {
    return '流入減少（クリック・CTR悪化／順位は要確認）';
  }

  // 順位は上がっているのにCTRが上がっていない → タイトル側の問題ぽい
  if (diffPos > 3 && diffCtr <= 0) {
    return '順位改善だがCTR頭打ち（タイトル要検討）';
  }

  // 順位は下がっているのにCTRが落ちていない → ニーズは強いページ
  if (diffPos < -3 && diffCtr >= 0) {
    return '順位悪化だがCTR維持（ニーズは強い）';
  }

  return '変化小（要検証）';
}

