/**
 * Sheets.gs — スプレッドシートへのアクセス層
 *
 * 他のファイルは SpreadsheetApp を直接触らない。ここだけを差し替えれば
 * 保存先を Notion / BigQuery / DB に変えられる(Notionに載せ替える場合もこのファイルだけで済む)。
 *
 * 列は「ヘッダー名」で引くので、列を挿入・並べ替えしてもコードは壊れない。
 */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。initProject() を実行してください。');
  return sh;
}

/** ヘッダー名 -> 0始まりの列インデックス */
function headerMap_(sh) {
  const last = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, last).getValues()[0];
  const map = {};
  head.forEach(function (h, i) { if (h !== '') map[String(h).trim()] = i; });
  return map;
}

/**
 * シート全体を読む
 * @return {{sh:GoogleAppsScript.Spreadsheet.Sheet, map:Object, values:Array<Array>, firstRow:number}}
 *   values は2行目以降のデータ。行番号は firstRow + index。
 */
function readSheet_(name) {
  const sh = sheet_(name);
  const map = headerMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const values = lastRow < 2 ? [] : sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { sh: sh, map: map, values: values, firstRow: 2 };
}

/** 行配列 -> {ヘッダー名: 値} のオブジェクト */
function rowToObj_(row, map) {
  const o = {};
  for (const k in map) o[k] = row[map[k]];
  return o;
}

/** オブジェクトを1行追記する。存在しない列名は無視される */
function appendRow_(name, obj) {
  const sh = sheet_(name);
  const map = headerMap_(sh);
  const width = sh.getLastColumn();
  const row = new Array(width).fill('');
  for (const k in obj) {
    if (map[k] !== undefined) row[map[k]] = obj[k];
  }
  sh.appendRow(row);
  return sh.getLastRow();
}

/** 指定行を部分更新する。obj に無い列は触らない */
function updateRow_(name, rowIndex, obj) {
  const sh = sheet_(name);
  const map = headerMap_(sh);
  for (const k in obj) {
    if (map[k] === undefined) continue;
    sh.getRange(rowIndex, map[k] + 1).setValue(obj[k]);
  }
}

/** 監査用の生ログ。追記専用なので絶対に競合しないし、いくら増えても他に影響しない */
function logRaw_(empId, type, result, ua) {
  try {
    const o = {};
    o[CFG.C.RAW_TS] = new Date();
    o[CFG.C.RAW_ID] = empId || '';
    o[CFG.C.RAW_TYPE] = type || '';
    o[CFG.C.RAW_RESULT] = result || '';
    o[CFG.C.RAW_UA] = (ua || '').substring(0, 300);
    appendRow_(CFG.SH.RAW, o);
  } catch (e) {
    console.error('生ログ書き込み失敗: ' + e.message); // ログ失敗で打刻を止めない
  }
}

// ============ 社員マスタ ============

function listEmployees_() {
  const d = readSheet_(CFG.SH.EMP);
  const out = [];
  d.values.forEach(function (row, i) {
    const o = rowToObj_(row, d.map);
    if (!o[CFG.C.EMP_ID]) return;
    out.push({
      row: d.firstRow + i,
      empId: String(o[CFG.C.EMP_ID]).trim(),
      name: String(o[CFG.C.EMP_NAME] || '').trim(),
      type: o[CFG.C.EMP_TYPE] || '時給',
      wage: Number(o[CFG.C.EMP_WAGE] || 0),
      base: Number(o[CFG.C.EMP_BASE] || 0),
      active: o[CFG.C.EMP_ACTIVE] === true || o[CFG.C.EMP_ACTIVE] === 'TRUE',
      token: o[CFG.C.EMP_TOKEN] || ''
    });
  });
  return out;
}

/**
 * 社員1件を取得。打刻のたびに全件読むのを避けるため6時間キャッシュする。
 * マスタを編集したら clearEmployeeCache() を実行(編集トリガーで自動実行される)。
 */
function getEmployee_(empId) {
  const cache = CacheService.getScriptCache();
  const key = 'emp_' + empId;
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const list = listEmployees_();
  const found = list.filter(function (e) { return e.empId === empId; })[0];
  if (!found) return null;
  cache.put(key, JSON.stringify(found), 21600);
  return found;
}

function clearEmployeeCache() {
  const keys = listEmployees_().map(function (e) { return 'emp_' + e.empId; });
  if (keys.length) CacheService.getScriptCache().removeAll(keys);
}

// ============ 勤怠ログ ============

/**
 * 指定日・指定社員の勤怠行を探す。
 * 「日付+社員ID」のキーだけ2列読むので、行数が増えても軽い。
 */
function findAttendanceRow_(empId, dateStr) {
  const sh = sheet_(CFG.SH.ATT);
  const map = headerMap_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const dateCol = map[CFG.C.ATT_DATE] + 1;
  const idCol = map[CFG.C.ATT_ID] + 1;
  const n = lastRow - 1;
  const dates = sh.getRange(2, dateCol, n, 1).getValues();
  const ids = sh.getRange(2, idCol, n, 1).getValues();

  // 当日の行は末尾付近にあるので後ろから走査する
  for (let i = n - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() !== empId) continue;
    if (toDateStr_(dates[i][0]) !== dateStr) continue;
    return { row: 2 + i, sh: sh, map: map, values: sh.getRange(2 + i, 1, 1, sh.getLastColumn()).getValues()[0] };
  }
  return null;
}

/** セルの値(Date でも文字列でも)を 'yyyy-MM-dd' に正規化 */
function toDateStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd');
  return String(v).trim().substring(0, 10).replace(/\//g, '-');
}
