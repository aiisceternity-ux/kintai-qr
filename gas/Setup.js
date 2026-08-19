/**
 * Setup.gs — 初期化・メニュー・トリガー
 *
 * 使い方: GASエディタで initProject を1回実行する。以上。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤怠')
    .addItem('QRコードを発行/印刷', 'showQrDialog')
    .addSeparator()
    .addItem('先月分の給与を集計', 'menuCloseLastMonth')
    .addItem('月を指定して給与を集計', 'menuCloseMonth')
    .addSeparator()
    .addItem('選択した行の実働を再計算', 'recalcSelectedRows')
    .addItem('退勤もれをチェック', 'checkUnclosed')
    .addSeparator()
    .addItem('初期セットアップ', 'initProject')
    .addItem('自動実行トリガーを設置', 'installTriggers')
    .addToUi();
}

/** シート・書式・入力規則・秘密鍵を一式そろえる。何度実行しても壊れない */
function initProject() {
  // ウェブアプリから開けるように、対象スプレッドシートのIDを控える(これが無いと打刻APIが動かない)
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) setProp_('SHEET_ID', active.getId());

  const ss = ss_();
  ss.setSpreadsheetTimeZone(CFG.TZ);

  for (const name in SCHEMA) {
    const def = SCHEMA[name];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    // ヘッダー
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers])
      .setFontWeight('bold').setBackground('#efefef');
    sh.setFrozenRows(1);

    const map = headerMap_(sh);

    // 表示形式
    for (const col in def.formats) {
      if (map[col] === undefined) continue;
      sh.getRange(2, map[col] + 1, sh.getMaxRows() - 1, 1).setNumberFormat(def.formats[col]);
    }
    // 列幅
    for (const col in def.widths) {
      if (map[col] === undefined) continue;
      sh.setColumnWidth(map[col] + 1, def.widths[col]);
    }
  }

  // 入力規則
  const emp = ss.getSheetByName(CFG.SH.EMP);
  const em = headerMap_(emp);
  emp.getRange(2, em[CFG.C.EMP_TYPE] + 1, emp.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['時給', '月給'], true).build());
  emp.getRange(2, em[CFG.C.EMP_ACTIVE] + 1, emp.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  const att = ss.getSheetByName(CFG.SH.ATT);
  const am = headerMap_(att);
  att.getRange(2, am[CFG.C.ATT_STATUS] + 1, att.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['勤務中', '退勤済', '要確認'], true).build());

  const pay = ss.getSheetByName(CFG.SH.PAY);
  const pm = headerMap_(pay);
  pay.getRange(2, pm[CFG.C.PAY_FIXED] + 1, pay.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  // 秘密鍵(未設定なら生成)
  if (!prop_('QR_SECRET', false)) setProp_('QR_SECRET', randomKey_(48));
  if (!prop_('ADMIN_KEY', false)) setProp_('ADMIN_KEY', randomKey_(16));

  // 社員が0人ならサンプルを1件入れる
  if (listEmployees_().length === 0) {
    const o = {};
    o[CFG.C.EMP_ID] = 'A001';
    o[CFG.C.EMP_NAME] = '田中 太郎';
    o[CFG.C.EMP_TYPE] = '時給';
    o[CFG.C.EMP_WAGE] = 1200;
    o[CFG.C.EMP_BASE] = 0;
    o[CFG.C.EMP_ACTIVE] = true;
    appendRow_(CFG.SH.EMP, o);
  }

  issueAllTokens();

  console.log([
    'セットアップ完了',
    'SHEET_ID = ' + prop_('SHEET_ID', false),
    'ADMIN_KEY = ' + prop_('ADMIN_KEY'),
    'この後: デプロイ > 新しいデプロイ > 種類「ウェブアプリ」',
    '  次のユーザーとして実行: 自分 / アクセス: 全員',
    '発行された /exec URL を web/scan.html の API_URL に貼る'
  ].join('\n'));
}

/** 社員マスタの全員ぶんのQRトークンを生成して書き戻す */
function issueAllTokens() {
  const list = listEmployees_();
  list.forEach(function (e) {
    const t = makeToken_(e.empId);
    if (e.token === t) return;
    const o = {};
    o[CFG.C.EMP_TOKEN] = t;
    updateRow_(CFG.SH.EMP, e.row, o);
  });
  clearEmployeeCache();
  return list.length;
}

/** 自動実行トリガーを設置(重複しないよう既存の同名を消してから) */
function installTriggers() {
  const wanted = ['closeLastMonth', 'checkUnclosed', 'onEditHandler'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (wanted.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });

  // 毎月1日の 4:00 に前月を締める
  ScriptApp.newTrigger('closeLastMonth').timeBased()
    .onMonthDay(1).atHour(4).create();

  // 毎日 3:00 に退勤もれをチェック
  ScriptApp.newTrigger('checkUnclosed').timeBased()
    .everyDays(1).atHour(3).create();

  // 社員マスタを直したらキャッシュを捨てる
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss_()).onEdit().create();

  console.log('トリガーを設置しました');
}

function onEditHandler(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== CFG.SH.EMP) return;
  clearEmployeeCache();
  issueAllTokens(); // 社員を追加したらトークンも自動発行
}

function randomKey_(len) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let s = '';
  const bytes = Utilities.computeHmacSha256Signature(Utilities.getUuid(), Utilities.getUuid());
  for (let i = 0; i < len; i++) {
    s += chars.charAt(Math.abs(bytes[i % bytes.length] + i * 7) % chars.length);
  }
  return s;
}

// ---- メニュー用のラッパー ----

function showQrDialog() {
  const t = HtmlService.createTemplateFromFile('admin');
  t.employees = listEmployees_();
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(900).setHeight(650), 'QRコード発行'
  );
}

function menuCloseLastMonth() {
  const r = closeLastMonth();
  SpreadsheetApp.getUi().alert(
    r.month + ' の集計が完了しました。\n\n' +
    '書き込み: ' + r.written + '件\n確定済みのためスキップ: ' + r.skipped + '件\n' +
    '要確認で除外した勤怠: ' + r.warn + '件'
  );
}

function menuCloseMonth() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('対象月を yyyy-MM で入力してください（例: 2026-07）');
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const ym = res.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) { ui.alert('形式が違います'); return; }
  const r = closeMonth(ym);
  ui.alert(ym + ' の集計が完了しました。\n\n書き込み: ' + r.written + '件\nスキップ: ' + r.skipped + '件');
}
