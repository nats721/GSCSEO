/************************************************************
 * 後続処理：階層マスタに対して「親→子」「子→親」リンク不足を抽出し、
 * entryBodyHtml（元シート）を使って ChatGPT プロンプトを生成する
 *
 * 重要：
 * - 本文HTMLは切り捨てない（トリミング禁止）
 * - Sheetsセル上限のため、プロンプト全文は Drive に txt 保存し、URLを出力する
 * - プロンプトに入れる本文から「文末CTA」「SNSブロック」「snsbox」を除外
 ************************************************************/

/** ========= 設定（exportInternalLinksAll と一致させる） ========= */
const HIER_POST_CFG = {
  SRC_SS_ID: '1ZygntHC1J6IxzKbTidmOyf5nUyW4GCLurHCjexrLm_o',
  SRC_SHEET_GID: 1666145974,
  DST_SS_ID: '1U-XK7TVV8BWxo27iRqySKmh6SmlyZncjoC26ZAX8LjM',

  HIER_SHEET: '階層マスタ',
  LINKS_SHEET: '内部リンク_一覧',

  OUT_P2C: '階層リンク不足_親→子',
  OUT_P2C_DETAIL: '階層リンク不足_親→子_詳細',
  OUT_C2P: '階層リンク不足_子→親',
  OUT_C2P_DETAIL: '階層リンク不足_子→親_詳細',

  // 断片出力（セル上限回避の保険）
  PROMPT_PARTS_SHEET: 'PromptParts',
  PROMPT_PART_CHUNK: 200000, // 1セルに入る安全圏（本文は切り捨てない＝分割する）

  // Driveへプロンプト全文保存（推奨）
  SAVE_PROMPT_TO_DRIVE: false,
  DRIVE_FOLDER_NAME: '内部リンク_プロンプト',

  EXCLUDE_CLASSES: new Set(['NW', 'ST']),
  ALLOW_LINK_TYPES: new Set(['content', 'cta_inline']), // cta_autoは設計に含めない
};

/** ワンボタン：スナップショット→後続処理 */
function runAll_ExportThenHierarchyPrompts_AllDirections() {
  exportInternalLinksAll();
  runHierarchyLinkAuditAndPrompts_AllDirections();
}

/** 後続処理本体 */
function runHierarchyLinkAuditAndPrompts_AllDirections() {
  const dstSS = SpreadsheetApp.openById(HIER_POST_CFG.DST_SS_ID);

  const hierSh  = mustGetSheetByName_(dstSS, HIER_POST_CFG.HIER_SHEET);
  const linksSh = mustGetSheetByName_(dstSS, HIER_POST_CFG.LINKS_SHEET);

  const hierVals = hierSh.getDataRange().getValues();
  const linkVals = linksSh.getDataRange().getValues();
  if (hierVals.length < 2) throw new Error('階層マスタが空です');
  if (linkVals.length < 2) throw new Error('内部リンク_一覧が空です（exportInternalLinksAll の後に実行してください）');

  const hIdx = headerIndexMap_(hierVals[0]);
  const lIdx = headerIndexMap_(linkVals[0]);

  // 階層マスタ
  const H_URL    = mustIdx_(hIdx, 'URL');
  const H_PARENT = mustIdx_(hIdx, '親URL');
  const H_CLASS  = optIdx_(hIdx, '分類');
  const H_TITLE  = optIdx_(hIdx, 'タイトル');
  const H_PCLASS = optIdx_(hIdx, '親分類');
  const H_PTITLE = optIdx_(hIdx, '親タイトル');

  // 内部リンク_一覧
  const L_FROM     = mustIdx_(lIdx, 'from_URL');
  const L_TO       = mustIdx_(lIdx, 'to_URL');
  const L_INTERNAL = mustIdx_(lIdx, 'is_internal');
  const L_EXISTS   = mustIdx_(lIdx, 'to_exists');
  const L_TYPE     = mustIdx_(lIdx, 'link_type');

  // 既存リンクセット（content/cta_inlineのみ）
  const linkSet = new Set();
  for (let i = 1; i < linkVals.length; i++) {
    const r = linkVals[i];
    if (r[L_INTERNAL] !== true) continue;
    if (r[L_EXISTS] !== true) continue;
    const lt = String(r[L_TYPE] || '').trim();
    if (!HIER_POST_CFG.ALLOW_LINK_TYPES.has(lt)) continue;

    const from = normKeyUrl_(r[L_FROM]);
    const to   = normKeyUrl_(r[L_TO]);
    if (!from || !to) continue;
    linkSet.add(from + '\u0000' + to);
  }

  // 本文HTML（entryBodyHtml）
  const bodyMap = loadEntryBodyHtmlMap_(HIER_POST_CFG.SRC_SS_ID, HIER_POST_CFG.SRC_SHEET_GID);

  // 欠落抽出
  const p2cMap = new Map(); // 親→子：親URLキー
  const c2pMap = new Map(); // 子→親：子URLキー

  const p2cDetailRows = [];
  const c2pDetailRows = [];

  for (let i = 1; i < hierVals.length; i++) {
    const r = hierVals[i];

    const childRaw  = String(r[H_URL] || '').trim();
    const parentRaw = String(r[H_PARENT] || '').trim();
    if (!childRaw || !parentRaw) continue;

    const childK  = normKeyUrl_(childRaw);
    const parentK = normKeyUrl_(parentRaw);
    if (!childK || !parentK) continue;
    if (childK === parentK) continue;

    const cClass = H_CLASS  != null ? String(r[H_CLASS]  || '').trim() : '';
    const pClass = H_PCLASS != null ? String(r[H_PCLASS] || '').trim() : '';
    if (cClass && HIER_POST_CFG.EXCLUDE_CLASSES.has(cClass)) continue;
    if (pClass && HIER_POST_CFG.EXCLUDE_CLASSES.has(pClass)) continue;

    const cTitle = H_TITLE  != null ? String(r[H_TITLE]  || '').trim() : '';
    const pTitle = H_PTITLE != null ? String(r[H_PTITLE] || '').trim() : '';

    // 親→子 欠落
    const hasP2C = linkSet.has(parentK + '\u0000' + childK);
    if (!hasP2C) {
      if (!p2cMap.has(parentRaw)) {
        p2cMap.set(parentRaw, { parentUrl: parentRaw, parentClass: pClass, parentTitle: pTitle, children: [] });
      }
      p2cMap.get(parentRaw).children.push({ childUrl: childRaw, childClass: cClass, childTitle: cTitle });

      p2cDetailRows.push([parentRaw, pClass, pTitle, childRaw, cClass, cTitle]);
    }

    // 子→親 欠落
    const hasC2P = linkSet.has(childK + '\u0000' + parentK);
    if (!hasC2P) {
      if (!c2pMap.has(childRaw)) {
        c2pMap.set(childRaw, {
          childUrl: childRaw, childClass: cClass, childTitle: cTitle,
          parentUrl: parentRaw, parentClass: pClass, parentTitle: pTitle
        });
      }
      c2pDetailRows.push([childRaw, cClass, cTitle, parentRaw, pClass, pTitle]);
    }
  }

  // PromptParts初期化
  const partsSh = resetSheetByName_(dstSS, HIER_POST_CFG.PROMPT_PARTS_SHEET);
  partsSh.getRange(1, 1, 1, 5).setValues([['key','URL','part_no','kind','text']]);

  // Driveフォルダ
  const folder = HIER_POST_CFG.SAVE_PROMPT_TO_DRIVE ? getOrCreateFolder_(HIER_POST_CFG.DRIVE_FOLDER_NAME) : null;

  // 親→子プロンプト生成（親ページ本文を渡す）
  const p2cRows = [];
  p2cMap.forEach((bundle) => {
    const rawHtml = bodyMap.get(normKeyUrl_(bundle.parentUrl)) || '';
    const cleanedHtml = cleanBodyForPrompt_(rawHtml); // ← 切り捨てない（除外のみ）

    const prompt = buildParentToChildPrompt_(bundle, cleanedHtml);

    const key = makeKey_('P2C', bundle.parentUrl);
    writePromptParts_(partsSh, bundle.parentUrl, key, 'prompt', prompt);

    let fileUrl = '';
    if (folder) fileUrl = saveTextToFolder_(folder, key + '.txt', prompt);

    p2cRows.push([
      bundle.parentUrl, bundle.parentClass, bundle.parentTitle,
      bundle.children.length,
      bundle.children.map(c => c.childUrl).join('\n'),
      key,
      fileUrl
    ]);
  });

  // 子→親プロンプト生成（子ページ本文を渡す）
  const c2pRows = [];
  c2pMap.forEach((x) => {
    const rawHtml = bodyMap.get(normKeyUrl_(x.childUrl)) || '';
    const cleanedHtml = cleanBodyForPrompt_(rawHtml); // ← 切り捨てない（除外のみ）

    const prompt = buildChildToParentPrompt_(x, cleanedHtml);

    const key = makeKey_('C2P', x.childUrl);
    writePromptParts_(partsSh, x.childUrl, key, 'prompt', prompt);

    let fileUrl = '';
    if (folder) fileUrl = saveTextToFolder_(folder, key + '.txt', prompt);

    c2pRows.push([
      x.childUrl, x.childClass, x.childTitle,
      x.parentUrl, x.parentClass, x.parentTitle,
      key,
      fileUrl
    ]);
  });

  // 出力（親→子）
  const outP2C = resetSheetByName_(dstSS, HIER_POST_CFG.OUT_P2C);
  outP2C.getRange(1, 1, 1, 7).setValues([[
    '親URL','親分類','親タイトル','不足子数','不足子URL一覧','prompt_key','prompt_txt_url(Drive)'
  ]]);
  if (p2cRows.length) {
    outP2C.getRange(2, 1, p2cRows.length, 7).setValues(p2cRows);
    outP2C.getRange(2, 5, p2cRows.length, 1).setWrap(true);
    outP2C.setFrozenRows(1);
    //outP2C.autoResizeColumns(1, 7);
  }

  const outP2CD = resetSheetByName_(dstSS, HIER_POST_CFG.OUT_P2C_DETAIL);
  outP2CD.getRange(1, 1, 1, 6).setValues([[
    '親URL','親分類','親タイトル','子URL（親を持つ）','子分類','子タイトル'
  ]]);
  if (p2cDetailRows.length) {
    outP2CD.getRange(2, 1, p2cDetailRows.length, 6).setValues(p2cDetailRows);
    outP2CD.setFrozenRows(1);
    //outP2CD.autoResizeColumns(1, 6);
  }

  // 出力（子→親）
  const outC2P = resetSheetByName_(dstSS, HIER_POST_CFG.OUT_C2P);
  outC2P.getRange(1, 1, 1, 8).setValues([[
    '子URL','子分類','子タイトル','親URL','親分類','親タイトル','prompt_key','prompt_txt_url(Drive)'
  ]]);
  if (c2pRows.length) {
    outC2P.getRange(2, 1, c2pRows.length, 8).setValues(c2pRows);
    outC2P.setFrozenRows(1);
    //outC2P.autoResizeColumns(1, 8);
  }

  const outC2PD = resetSheetByName_(dstSS, HIER_POST_CFG.OUT_C2P_DETAIL);
  outC2PD.getRange(1, 1, 1, 6).setValues([[
    '子URL','子分類','子タイトル','親URL','親分類','親タイトル'
  ]]);
  if (c2pDetailRows.length) {
    outC2PD.getRange(2, 1, c2pDetailRows.length, 6).setValues(c2pDetailRows);
    outC2PD.setFrozenRows(1);
    //outC2PD.autoResizeColumns(1, 6);
  }

  SpreadsheetApp.flush();
}

/** ========= プロンプト（親→子） ========= */
function buildParentToChildPrompt_(bundle, parentHtmlCleaned) {
  const childrenLines = bundle.children.map((c, i) => {
    const t = c.childTitle ? `（${c.childTitle}）` : '';
    return `${i + 1}. ${c.childUrl}${t}`;
  }).join('\n');

  const lines = [];
  lines.push('以下はサイトの階層マスタに基づく「親→子」内部リンク不足です。');
  lines.push('親ページ本文（HTML）に、下記の子ページへのリンクを追加してください。');
  lines.push('');
  lines.push('【親ページ】');
  lines.push(`- URL: ${bundle.parentUrl}`);
  if (bundle.parentTitle) lines.push(`- タイトル: ${bundle.parentTitle}`);
  if (bundle.parentClass) lines.push(`- 分類: ${bundle.parentClass}`);
  lines.push('');
  lines.push('【リンク追加が必要な子ページ一覧（この親を持つ）】');
  lines.push(childrenLines);
  lines.push('');
  lines.push('【重要：リンク挿入ルール】');
  lines.push('- まず既存本文の流れを壊さない位置に自然挿入（段落内 / 箇条書き / 関連箇所の直後）。');
  lines.push('- どうしても自然挿入できない場合のみ、本文内に最小限の導線セクションを新設してよい（例：H2「関連ページ」/「次に読む」）。');
  lines.push('- ただし内容を膨らませない（導線セクションも1行＋リンク一覧程度）。');
  lines.push('- 文末CTA（veu_cta）やSNSブロック、snsboxに入れて逃げるのは禁止。本文導線として成立させる。');
  lines.push('- 出力は「修正版HTML全文（親ページの本文HTML）」のみ。差分説明は不要。');
  lines.push('- 同一段落にリンクを詰め込みすぎない（目安：1段落あたり1〜2リンク）。');
  lines.push('- アンカーテキストは「こちら」禁止。子ページのテーマが分かる文言にする。');
  lines.push('- 子リンクが多い場合は、既存H2の該当箇所に分散して挿入し、リンク集1箇所に偏らせない。');

  lines.push('');
  lines.push('【親ページ本文HTML（文末CTA/SNS/snsboxは除外済み）】');
  lines.push('```html');
  lines.push(String(parentHtmlCleaned || '')); // ★切り捨て禁止：そのまま入れる
  lines.push('```');

  return lines.join('\n');
}

/** ========= プロンプト（子→親） ========= */
function buildChildToParentPrompt_(x, childHtmlCleaned) {
  const lines = [];
  lines.push('以下はサイトの階層マスタに基づく「子→親」内部リンク不足です。');
  lines.push('子ページ本文（HTML）に、親ページへのリンクを追加してください。');
  lines.push('');
  lines.push('【子ページ】');
  lines.push(`- URL: ${x.childUrl}`);
  if (x.childTitle) lines.push(`- タイトル: ${x.childTitle}`);
  if (x.childClass) lines.push(`- 分類: ${x.childClass}`);
  lines.push('');
  lines.push('【親ページ（リンク先）】');
  lines.push(`- 親URL: ${x.parentUrl}`);
  if (x.parentTitle) lines.push(`- 親タイトル: ${x.parentTitle}`);
  if (x.parentClass) lines.push(`- 親分類: ${x.parentClass}`);
  lines.push('');
  lines.push('【重要：リンク挿入ルール】');
  lines.push('- まず既存本文の流れを壊さない位置に自然挿入（導入直後 / まとめ直前 / 関連段落内）。');
  lines.push('- どうしても自然挿入できない場合のみ、本文内に最小限の導線セクションを新設してよい。');
  lines.push('- ただし内容を膨らませない（導線セクションも1行＋リンク1本程度）。');
  lines.push('- 文末CTA（veu_cta）やSNSブロック、snsboxに入れて逃げるのは禁止。本文導線として成立させる。');
  lines.push('- 出力は「修正版HTML全文（子ページの本文HTML）」のみ。差分説明は不要。');
  lines.push('- 親リンクは原則1本。パンくず等と重複しても“本文導線”として意味がある位置に置く。');
  lines.push('- 親リンクのアンカーは「全体像はこちら」など、親がハブであることが伝わる文言にする。');
  lines.push('');
  lines.push('【子ページ本文HTML（文末CTA/SNS/snsboxは除外済み）】');
  lines.push('```html');
  lines.push(String(childHtmlCleaned || '')); // ★切り捨て禁止：そのまま入れる
  lines.push('```');

  return lines.join('\n');
}

/** ========= 本文クレンジング（除外のみ。切り捨てはしない） ========= */
function cleanBodyForPrompt_(html) {
  const src = String(html || '');
  if (!src) return '';

  const CTA_SECTION_RE   = /<section[^>]*class=["'][^"']*veu_cta[^"']*["'][\s\S]*?<\/section>/gi;
  const SOCIALSET_DIV_RE = /<div[^>]*class=["'][^"']*veu_socialSet[^"']*["'][\s\S]*?<\/div>/gi;

  // snsbox / sharebox 系（div/section/aside）
  const SNSBOX_RE = /<(div|section|aside)\b[^>]*(?:class|id)=["'][^"']*(?:snsbox|sns-box|sns_box|sharebox|share-box|share_box)[^"']*["'][\s\S]*?<\/\1>/gi;

  return src
    .replace(CTA_SECTION_RE, '')
    .replace(SOCIALSET_DIV_RE, '')
    .replace(SNSBOX_RE, '')
    .trim();
}

/** ========= entryBodyHtmlマップ ========= */
function loadEntryBodyHtmlMap_(srcSsId, srcGid) {
  const srcSS = SpreadsheetApp.openById(srcSsId);
  const srcSheet = srcSS.getSheets().find(s => s.getSheetId() === srcGid);
  if (!srcSheet) throw new Error('元シートが見つかりません（gid確認）');

  const values = srcSheet.getDataRange().getValues();
  const header = values.shift();
  const idx = {};
  for (let i = 0; i < header.length; i++) idx[String(header[i])] = i;

  const IDX_URL  = idx['URL'];
  const IDX_BODY = idx['entryBodyHtml'];
  if (IDX_URL === undefined || IDX_BODY === undefined) {
    throw new Error('元シートの列名が一致しません（URL / entryBodyHtml）');
  }

  const map = new Map();
  for (let i = 0; i < values.length; i++) {
    const raw = values[i][IDX_URL];
    if (!raw) continue;
    const urlKey = normKeyUrl_(raw);
    if (!urlKey) continue;
    if (!map.has(urlKey)) map.set(urlKey, String(values[i][IDX_BODY] || ''));
  }
  return map;
}

/** ========= PromptParts（切り捨て禁止：分割保存） ========= */
function writePromptParts_(partsSh, targetUrl,key, kind, text) {
  const s = String(text || '');
  const chunk = HIER_POST_CFG.PROMPT_PART_CHUNK;

  const rows = [];
  let partNo = 1;
  for (let i = 0; i < s.length; i += chunk) {
    rows.push([targetUrl,key, partNo, kind, s.substring(i, i + chunk)]);
    partNo++;
  }
  if (!rows.length) rows.push([key, 1, kind, '']);

  partsSh.getRange(partsSh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
}

/** ========= Drive保存（全文を保存。切り捨てなし） ========= */
function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function saveTextToFolder_(folder, fileName, content) {
  // 同名があれば削除して作り直し（確実に更新）
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  const file = folder.createFile(fileName, String(content || ''), MimeType.PLAIN_TEXT);
  return file.getUrl();
}

/** ========= utils ========= */
function makeKey_(prefix, url) {
  const safe = normKeyUrl_(url).replace(/[^a-zA-Z0-9]/g, '_').slice(-80);
  const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  return `${prefix}_${safe}_${ts}`;
}

function mustGetSheetByName_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シートが見つかりません: ${name}`);
  return sh;
}

function resetSheetByName_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  return sh;
}

function headerIndexMap_(headerRow) {
  const m = {};
  for (let i = 0; i < headerRow.length; i++) {
    const k = String(headerRow[i] || '').trim();
    if (!k) continue;
    if (m[k] == null) m[k] = i;
  }
  return m;
}

function mustIdx_(map, colName) {
  const i = map[colName];
  if (i == null) throw new Error(`必要列が見つかりません: ${colName}`);
  return i;
}

function optIdx_(map, colName) {
  const i = map[colName];
  return i == null ? null : i;
}

// exportInternalLinksAll の normalizeUrl_ と同等（#/?除去 + 末尾/除去）
function normKeyUrl_(url) {
  let u = String(url || '').trim();
  if (!u) return '';
  const h = u.indexOf('#');
  if (h >= 0) u = u.substring(0, h);
  const q = u.indexOf('?');
  if (q >= 0) u = u.substring(0, q);
  if (u.length > 8 && u.endsWith('/')) u = u.slice(0, -1);
  return u;
}
