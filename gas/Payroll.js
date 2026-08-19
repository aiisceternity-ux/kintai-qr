/**
 * Payroll.gs — 月次の給与集計と、打刻忘れの検知
 */

/** 月初トリガー用。前月分を締める */
function closeLastMonth() {
  const d = new Date();
  const prev = new Date(d.getFullYear(), d.getMonth() - 1, 1);
  const ym = Utilities.formatDate(prev, CFG.TZ, 'yyyy-MM');
  const r = closeMonth(ym);
  console.log(ym + ' の給与を集計: ' + r.written + '件');
  return r;
}

/**
 * 指定月(yyyy-MM)の勤怠を集計して給与シートに書く。
 * 既に「確定」チェックが入っている明細は上書きしない(締めた後の事故防止)。
 */
function closeMonth(ym) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const att = readSheet_(CFG.SH.ATT);
    const emps = listEmployees_();
    const empMap = {};
    emps.forEach(function (e) { empMap[e.empId] = e; });

    // --- 勤怠を社員ごとに集計 ---
    const agg = {}; // empId -> {days, hours, ot}
    let warn = 0;
    att.values.forEach(function (row) {
      const o = rowToObj_(row, att.map);
      const dateStr = toDateStr_(o[CFG.C.ATT_DATE]);
      if (dateStr.substring(0, 7) !== ym) return;

      const id = String(o[CFG.C.ATT_ID] || '').trim();
      if (!id) return;
      if (o[CFG.C.ATT_STATUS] === '要確認') { warn++; return; } // 壊れた行は集計に混ぜない

      const work = Number(o[CFG.C.ATT_WORK] || 0);
      if (!work) return;

      if (!agg[id]) agg[id] = { days: 0, hours: 0, ot: 0 };
      agg[id].days += 1;
      agg[id].hours += work;
      agg[id].ot += Number(o[CFG.C.ATT_OT] || 0);
    });

    // --- 既存の給与明細を引く(確定済みは触らない) ---
    const pay = readSheet_(CFG.SH.PAY);
    const existing = {}; // empId -> {row, fixed}
    pay.values.forEach(function (row, i) {
      const o = rowToObj_(row, pay.map);
      if (String(o[CFG.C.PAY_MONTH]).trim() !== ym) return;
      existing[String(o[CFG.C.PAY_ID]).trim()] = {
        row: pay.firstRow + i,
        fixed: o[CFG.C.PAY_FIXED] === true || o[CFG.C.PAY_FIXED] === 'TRUE'
      };
    });

    // --- 書き込み ---
    const now = new Date();
    let written = 0, skipped = 0;

    emps.forEach(function (emp) {
      const a = agg[emp.empId];
      const isMonthly = String(emp.type).indexOf('月給') >= 0;
      if (!a && !isMonthly) return;            // 時給者で勤怠ゼロなら明細を作らない
      const days = a ? a.days : 0;
      const hours = a ? +a.hours.toFixed(2) : 0;
      const ot = a ? +a.ot.toFixed(2) : 0;
      const amount = calcAmount_(emp, hours, ot);

      const o = {};
      o[CFG.C.PAY_MONTH]  = ym;
      o[CFG.C.PAY_ID]     = emp.empId;
      o[CFG.C.PAY_NAME]   = emp.name;
      o[CFG.C.PAY_DAYS]   = days;
      o[CFG.C.PAY_HOURS]  = hours;
      o[CFG.C.PAY_OT]     = ot;
      o[CFG.C.PAY_AMOUNT] = amount;
      o[CFG.C.PAY_GEN]    = now;

      const ex = existing[emp.empId];
      if (ex && ex.fixed) { skipped++; return; }
      if (ex) {
        updateRow_(CFG.SH.PAY, ex.row, o);
      } else {
        o[CFG.C.PAY_FIXED] = false;
        appendRow_(CFG.SH.PAY, o);
      }
      written++;
    });

    return { month: ym, written: written, skipped: skipped, warn: warn };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 支給額の確定ロジック。
 * 時給者: 通常時間 × 時給 + 残業時間 × 時給 × 割増率
 * 月給者: 基本給の固定支給
 * ※深夜割増(22-5時)・社会保険・源泉徴収はここでは扱わない。README参照。
 */
function calcAmount_(emp, hours, ot) {
  if (String(emp.type).indexOf('月給') >= 0) return Math.round(emp.base || 0);
  const regular = Math.max(0, hours - ot);
  return Math.round(regular * emp.wage + ot * emp.wage * CFG.OVERTIME_RATE);
}

/**
 * 退勤打刻の漏れを検知する。毎日深夜のトリガーで動かす。
 * 前日以前で「勤務中」のまま残っている行を「要確認」にして通知する。
 */
function checkUnclosed() {
  const today = businessDate_(new Date());
  const att = readSheet_(CFG.SH.ATT);
  const found = [];

  att.values.forEach(function (row, i) {
    const o = rowToObj_(row, att.map);
    if (o[CFG.C.ATT_STATUS] !== '勤務中') return;
    const dateStr = toDateStr_(o[CFG.C.ATT_DATE]);
    if (!dateStr || dateStr >= today) return; // 当日ぶんはまだ勤務中で正常

    const up = {};
    up[CFG.C.ATT_STATUS] = '要確認';
    updateRow_(CFG.SH.ATT, att.firstRow + i, up);
    found.push(dateStr + ' ' + o[CFG.C.ATT_NAME] + '(' + o[CFG.C.ATT_ID] + ')');
  });

  if (found.length && CFG.ALERT_TO) {
    MailApp.sendEmail(
      CFG.ALERT_TO,
      '[勤怠] 退勤打刻もれ ' + found.length + '件',
      '以下の勤務が「勤務中」のまま残っていたため「要確認」にしました。\n' +
      '手で退勤時刻を入れて、実働・残業を recalcRow で再計算してください。\n\n' +
      found.join('\n') + '\n\n' + ss_().getUrl()
    );
  }
  console.log('未退勤: ' + found.length + '件');
  return found;
}

/**
 * 手修正した行の実働・残業を計算し直す。
 * 勤怠ログのその行を選択してメニューから実行する。
 */
function recalcSelectedRows() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== CFG.SH.ATT) {
    SpreadsheetApp.getUi().alert('「' + CFG.SH.ATT + '」シートで、直した行を選択してから実行してください。');
    return;
  }
  const map = headerMap_(sh);
  const rng = sh.getActiveRange();
  let n = 0;

  for (let r = rng.getRow(); r < rng.getRow() + rng.getNumRows(); r++) {
    if (r < 2) continue;
    const vals = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
    const o = rowToObj_(vals, map);
    const inV = o[CFG.C.ATT_IN], outV = o[CFG.C.ATT_OUT];
    if (!(inV instanceof Date) || !(outV instanceof Date)) continue;

    const manual = (o[CFG.C.ATT_BREAK] === '' || o[CFG.C.ATT_BREAK] === null) ? null : Number(o[CFG.C.ATT_BREAK]);
    const calc = calcWork_(inV, outV, manual);
    const up = {};
    up[CFG.C.ATT_BREAK]   = calc.breakMin;
    up[CFG.C.ATT_WORK]    = calc.work;
    up[CFG.C.ATT_OT]      = calc.overtime;
    up[CFG.C.ATT_STATUS]  = calc.work < 0 ? '要確認' : '退勤済';
    up[CFG.C.ATT_UPDATED] = new Date();
    updateRow_(CFG.SH.ATT, r, up);
    n++;
  }
  SpreadsheetApp.getUi().alert(n + ' 行を再計算しました。');
}
