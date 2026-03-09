/*************************************************************
 * GPTプロンプト → OpenAI Responses API 逐次実行版
 * -----------------------------------------------------------
 * 前提シート:
 *   シート名: GPTプロンプト
 *   A: Range
 *   B: ページ分類
 *   C: URL
 *   D: 更新日時
 *   E: プロンプト
 *   F: 更新日時(最新) ※既存数式列
 *
 * このスクリプトが使う追加列:
 *   G: 実行対象      (checkbox)
 *   H: 実行ステータス (WAIT / RUNNING / DONE / ERROR)
 *   I: 実行日時
 *   J: APIレスポンスID
 *   K: 出力本文
 *   L: エラー内容
 *
 * 注意:
 * - OpenAI APIキーはスクリプトプロパティに OPENAI_API_KEY として保存
 * - モデルは一律 gpt-5.1-chat-latest
 * - Batch は使わない
 *************************************************************/

const OUTPUT_PROMPT_SHEET_ID = '1_URzGA15RYBOmLfjsifuuMIigAqR_YaPJvMbPgQosVw';
const PROMPT_SHEET_NAME = 'GPTプロンプト';

const OPENAI_API_URL = 'https://api.openai.com/v1/responses';
const OPENAI_MODEL = 'gpt-5.4';

// 実行制御
const DEFAULT_ROWS_PER_RUN = 8;       // 1回で処理する最大件数
const MAX_RUNTIME_MS = 5 * 60 * 1000; // GAS 6分制限を見越して短め
const REQUEST_INTERVAL_MS = 1200;     // 通常時の待機
const MAX_RETRY = 5;                  // 429 / 5xx リトライ回数
const INITIAL_BACKOFF_MS = 3000;      // 初回バックオフ

// 列定義
const COL = {
  RANGE: 1,        // A
  CATEGORY: 2,     // B
  URL: 3,          // C
  UPDATED_AT: 4,   // D
  PROMPT: 5,       // E
  UPDATED_LATEST: 6, // F
  ENABLED: 7,      // G
  STATUS: 8,       // H
  EXECUTED_AT: 9,  // I
  RESPONSE_ID: 10, // J
  OUTPUT: 11,      // K
  ERROR: 12        // L
};

/* =========================================================
 * 公開関数
 * =======================================================*/

/**
 * 初回セットアップ
 * - G〜L列ヘッダ作成
 * - G列にcheckbox設定
 */
function preparePromptExecutionSheet() {
  const sh = getPromptSheet_();

  sh.getRange(1, COL.ENABLED, 1, 6).setValues([[
    '実行対象',
    '実行ステータス',
    '実行日時',
    'APIレスポンスID',
    '出力本文',
    'エラー内容'
  ]]);

  const maxRows = sh.getMaxRows();
  if (maxRows >= 2) {
    sh.getRange(2, COL.ENABLED, maxRows - 1, 1).insertCheckboxes();
  }

  sh.setFrozenRows(1);
}

/**
 * 新規待機行を実行
 * 対象: G=TRUE かつ Hが空欄 or WAIT
 */
function runPromptQueue() {
  runPromptQueueInternal_({
    mode: 'new',
    limit: DEFAULT_ROWS_PER_RUN
  });
}

/**
 * エラー行だけ再実行
 * 対象: G=TRUE かつ H=ERROR
 */
function retryErrorPromptQueue() {
  runPromptQueueInternal_({
    mode: 'error',
    limit: DEFAULT_ROWS_PER_RUN
  });
}

/**
 * RUNNING のまま止まった行を WAIT に戻す
 */
function resetRunningToWait() {
  const sh = getPromptSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const values = sh.getRange(2, COL.STATUS, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const status = String(values[i][0] || '').trim();
    if (status === 'RUNNING') {
      sh.getRange(rowNum, COL.STATUS).setValue('WAIT');
    }
  }
}

/**
 * DONE を空欄に戻す（再実行したい時用）
 */
function clearDoneStatus() {
  const sh = getPromptSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const values = sh.getRange(2, COL.STATUS, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const status = String(values[i][0] || '').trim();
    if (status === 'DONE') {
      sh.getRange(rowNum, COL.STATUS, 1, 5).clearContent(); // H〜L
    }
  }
}

/* =========================================================
 * 内部メイン
 * =======================================================*/

function runPromptQueueInternal_(opts) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    preparePromptExecutionSheet();

    const sh = getPromptSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const startTime = Date.now();
    const mode = opts.mode || 'new';
    const limit = Number(opts.limit || DEFAULT_ROWS_PER_RUN);

    const numRows = lastRow - 1;
    const values = sh.getRange(2, 1, numRows, COL.ERROR).getValues();

    let processed = 0;

    for (let i = 0; i < values.length; i++) {
      if (processed >= limit) break;
      if (Date.now() - startTime > MAX_RUNTIME_MS) break;

      const rowNum = i + 2;
      const row = values[i];

      const enabled = row[COL.ENABLED - 1];
      const status = String(row[COL.STATUS - 1] || '').trim();
      const prompt = String(row[COL.PROMPT - 1] || '').trim();

      if (enabled !== true) continue;
      if (!prompt) continue;

      if (mode === 'new') {
        if (!(status === '' || status === 'WAIT')) continue;
      } else if (mode === 'error') {
        if (status !== 'ERROR') continue;
      } else {
        continue;
      }

      try {
        markRowRunning_(sh, rowNum);

        const result = callOpenAIResponsesWithRetry_(prompt);
        const outputText = extractOutputText_(result.json);

        if (!outputText) {
          throw new Error('レスポンス本文を抽出できませんでした');
        }

        sh.getRange(rowNum, COL.RESPONSE_ID).setValue(result.responseId || result.requestId || '');
        sh.getRange(rowNum, COL.OUTPUT).setValue(outputText);
        sh.getRange(rowNum, COL.STATUS).setValue('DONE');
        sh.getRange(rowNum, COL.ERROR).clearContent();

      } catch (e) {
        sh.getRange(rowNum, COL.STATUS).setValue('ERROR');
        sh.getRange(rowNum, COL.ERROR).setValue(buildErrorMessage_(e));
      }

      processed++;
      SpreadsheetApp.flush();
      Utilities.sleep(REQUEST_INTERVAL_MS);
    }

  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 * OpenAI API
 * =======================================================*/

function callOpenAIResponsesWithRetry_(prompt) {
  const apiKey = getOpenAiApiKey_();

  const payload = {
    model: OPENAI_MODEL,
    input: prompt,
    store: false
  };

  let lastError = null;

  for (let attempt = 1; attempt <= MAX_RETRY; attempt++) {
    const res = UrlFetchApp.fetch(OPENAI_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const bodyText = res.getContentText();
    const headers = getHeadersSafe_(res);
    const requestId = findHeaderValue_(headers, 'x-request-id');

    let json = null;
    try {
      json = JSON.parse(bodyText);
    } catch (e) {
      json = null;
    }

    if (code >= 200 && code < 300) {
      return {
        code: code,
        json: json || {},
        requestId: requestId,
        responseId: json && json.id ? json.id : ''
      };
    }

    const errMsg = extractApiErrorMessage_(json, bodyText, code, requestId);
    lastError = new Error(errMsg);

    // 429 と 5xx はリトライ
    if (code === 429 || code >= 500) {
      const sleepMs = INITIAL_BACKOFF_MS * Math.pow(2, attempt - 1);
      Utilities.sleep(sleepMs);
      continue;
    }

    // それ以外は即終了
    throw lastError;
  }

  throw lastError || new Error('OpenAI API呼び出しに失敗しました');
}

function extractOutputText_(json) {
  if (!json) return '';

  if (typeof json.output_text === 'string' && json.output_text.trim()) {
    return json.output_text.trim();
  }

  const chunks = [];

  if (Array.isArray(json.output)) {
    json.output.forEach(item => {
      if (!item) return;

      if (Array.isArray(item.content)) {
        item.content.forEach(c => {
          if (!c) return;
          if (c.type === 'output_text' && typeof c.text === 'string' && c.text.trim()) {
            chunks.push(c.text.trim());
          }
        });
      }
    });
  }

  return chunks.join('\n\n').trim();
}

/* =========================================================
 * シート操作
 * =======================================================*/

function getPromptSheet_() {
  const ss = SpreadsheetApp.openById(OUTPUT_PROMPT_SHEET_ID);
  const sh = ss.getSheetByName(PROMPT_SHEET_NAME);
  if (!sh) {
    throw new Error('シート「' + PROMPT_SHEET_NAME + '」が見つかりません');
  }
  return sh;
}

function markRowRunning_(sh, rowNum) {
  sh.getRange(rowNum, COL.STATUS).setValue('RUNNING');
  sh.getRange(rowNum, COL.EXECUTED_AT).setValue(new Date());
  sh.getRange(rowNum, COL.ERROR).clearContent();
}

/* =========================================================
 * ユーティリティ
 * =======================================================*/

function getOpenAiApiKey_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('スクリプトプロパティに OPENAI_API_KEY が設定されていません');
  }
  return apiKey;
}

function getHeadersSafe_(res) {
  try {
    return res.getAllHeaders ? res.getAllHeaders() : res.getHeaders();
  } catch (e) {
    return {};
  }
}

function findHeaderValue_(headers, targetKey) {
  const t = String(targetKey || '').toLowerCase();
  for (const k in headers) {
    if (String(k).toLowerCase() === t) {
      return headers[k];
    }
  }
  return '';
}

function extractApiErrorMessage_(json, rawText, code, requestId) {
  let detail = '';

  if (json && json.error) {
    if (typeof json.error === 'string') {
      detail = json.error;
    } else if (json.error.message) {
      detail = json.error.message;
    } else {
      detail = JSON.stringify(json.error);
    }
  } else {
    detail = String(rawText || '');
  }

  const rid = requestId ? ' / request_id=' + requestId : '';
  return 'OpenAI API error: ' + code + ' / ' + detail + rid;
}

function buildErrorMessage_(e) {
  if (!e) return 'Unknown error';
  if (e.stack) return String(e.message || e) + '\n' + String(e.stack);
  return String(e.message || e);
}