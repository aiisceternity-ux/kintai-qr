/**
 * Setup.gs — 初期化・メニュー・トリガー
 *
 * 使い方: GASエディタで initProject を1回実行する。以上。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤怠')
    .addItem('ダッシュボードを開く', 'showDashboard')
    .addItem('ダッシュボードのURLを表示', 'showDashboardUrl')
    .addItem('QRコードを発行/印刷', 'showQrDialog')
    .addSeparator()
    .addItem('先月分の給与を集計', 'menuCloseLastMonth')
    .addItem('月を指定して給与を集計', 'menuCloseMonth')
    .addSeparator()
    .addItem('選択した行の実働を再計算', 'recalcSelectedRows')
    .addItem('退勤もれをチェック', 'checkUnclosed')
    .addSeparator()
    .addItem('社員を削除', 'menuDeleteEmployee')
    .addItem('テストデータを全消去', 'menuClearTestData')
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

  // サンプル社員は入れない。実データと紛れて「消したのに残っている」事故の元になる。
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


// ---- 保守用 ----

/** 社員IDを指定して社員マスタから行ごと削除する */
function menuDeleteEmployee() {
  const ui = SpreadsheetApp.getUi();
  const list = listEmployees_();
  if (!list.length) { ui.alert('社員が登録されていません。'); return; }

  const res = ui.prompt(
    '社員を削除',
    '削除する社員IDを入力してください。\n\n登録中:\n' +
      list.map(function (e) { return '  ' + e.empId + '  ' + e.name + (e.active ? '' : '（在籍外）'); }).join('\n'),
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const id = res.getResponseText().trim();
  if (!id) return;
  const n = deleteEmployeeRow_(id);
  ui.alert(n ? (id + ' を削除しました。') : (id + ' は見つかりませんでした。'));
}

/** 勤怠ログ・給与・打刻生ログを空にする。運用開始前のテストデータ一掃用 */
function menuClearTestData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'テストデータを全消去',
    '「' + CFG.SH.ATT + '」「' + CFG.SH.PAY + '」「' + CFG.SH.RAW + '」のデータ行をすべて消します。\n' +
    '社員マスタは消しません。この操作は元に戻せません。\n\n実行しますか?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const a = clearDataRows_(CFG.SH.ATT);
  const p = clearDataRows_(CFG.SH.PAY);
  const r = clearDataRows_(CFG.SH.RAW);
  CacheService.getScriptCache().removeAll(
    listEmployees_().map(function (e) { return 'cd_' + e.empId; }));
  ui.alert('消去しました。\n' + CFG.SH.ATT + ': ' + a + '行\n' + CFG.SH.PAY + ': ' + p + '行\n' + CFG.SH.RAW + ': ' + r + '行');
}


/** スプレッドシート内でダッシュボードを開く */
function showDashboard() {
  const t = HtmlService.createTemplateFromFile('dash');
  t.key = prop_('ADMIN_KEY');
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(900).setHeight(680), '勤怠ダッシュボード');
}

/** スマホで開くためのURLを表示する。合言葉入りなので取り扱い注意 */
function showDashboardUrl() {
  const url = ScriptApp.getService().getUrl() + '?page=dash&key=' + prop_('ADMIN_KEY');
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,sans-serif;font-size:13px;line-height:1.7;padding:8px">' +
    '<p>このURLをスマホのブラウザで開くとダッシュボードが見られます。<br>' +
    '<b>合言葉が含まれているので、共有先に注意してください。</b></p>' +
    '<textarea style="width:100%;height:90px;font-size:12px" onclick="this.select()">' + url + '</textarea>' +
    '<p style="color:#666">URLが漏れた場合は、スクリプトプロパティの ADMIN_KEY を変更すれば無効化できます。</p></div>'
  ).setWidth(560).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'ダッシュボードのURL');
}
