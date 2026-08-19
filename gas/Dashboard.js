/**
 * Dashboard.js — 管理用ダッシュボードのデータAPI
 *
 * 画面(dash.html)からは google.script.run で呼ぶ。CORSもキーの露出も無い。
 * ただし合言葉(ADMIN_KEY)は各関数でも検証する。画面を経由せず直接叩かれた場合の保険。
 *
 * 給与額は返さない。この画面は勤怠の把握が目的で、金額は給与シートで見る。
 */

function verifyAdmin_(key) {
  if (key !== prop_('ADMIN_KEY')) throw new Error('権限がありません');
}

/** 今日の勤務状況。全在籍者について 勤務中/退勤済/未出勤 を返す */
function dashToday(key) {
  verifyAdmin_(key);
  const now = new Date();
  const today = businessDate_(now);
  const byEmp = attendanceIndexByDate_(today);

  const rows = listEmployees_().filter(function (e) { return e.active; }).map(function (e) {
    const r = byEmp[e.empId];
    if (!r) {
      return { empId: e.empId, name: e.name, status: '未出勤', in: '', out: '', elapsed: '', work: null };
    }
    const inD = r[CFG.C.ATT_IN], outD = r[CFG.C.ATT_OUT];
    if (inD instanceof Date && !(outD instanceof Date)) {
      const h = (now.getTime() - inD.getTime()) / 3600000;
      return {
        empId: e.empId, name: e.name, status: '勤務中',
        in: fmt_(inD, 'HH:mm'), out: '',
        elapsed: Math.floor(h) + '時間' + Math.round((h % 1) * 60) + '分',
        work: null
      };
    }
    return {
      empId: e.empId, name: e.name,
      status: r[CFG.C.ATT_STATUS] || '退勤済',
      in: inD instanceof Date ? fmt_(inD, 'HH:mm') : '',
      out: outD instanceof Date ? fmt_(outD, 'HH:mm') : '',
      elapsed: '',
      work: r[CFG.C.ATT_WORK] === '' ? null : Number(r[CFG.C.ATT_WORK])
    };
  });

  const count = { working: 0, done: 0, absent: 0 };
  rows.forEach(function (r) {
    if (r.status === '勤務中') count.working++;
    else if (r.status === '未出勤') count.absent++;
    else count.done++;
  });

  return { date: today, now: fmt_(now, 'HH:mm'), rows: rows, count: count };
}

/** 要確認アラート。打刻もれと、明らかにおかしい実働を拾う */
function dashAlerts(key) {
  verifyAdmin_(key);
  const d = readSheet_(CFG.SH.ATT);
  const out = [];

  d.values.forEach(function (row, i) {
    const o = rowToObj_(row, d.map);
    const date = toDateStr_(o[CFG.C.ATT_DATE]);
    if (!date) return;

    const work = o[CFG.C.ATT_WORK] === '' ? null : Number(o[CFG.C.ATT_WORK]);
    const reasons = [];
    if (o[CFG.C.ATT_STATUS] === '要確認') reasons.push('退勤打刻もれ');
    if (work !== null && work < 0) reasons.push('実働がマイナス');
    if (work !== null && work > 16) reasons.push('実働16時間超');
    if (o[CFG.C.ATT_IN] instanceof Date && !(o[CFG.C.ATT_OUT] instanceof Date)
        && date < businessDate_(new Date())) reasons.push('退勤が空のまま');
    if (!reasons.length) return;

    out.push({
      row: d.firstRow + i,
      date: date,
      empId: String(o[CFG.C.ATT_ID] || ''),
      name: String(o[CFG.C.ATT_NAME] || ''),
      in: o[CFG.C.ATT_IN] instanceof Date ? fmt_(o[CFG.C.ATT_IN], 'HH:mm') : '',
      out: o[CFG.C.ATT_OUT] instanceof Date ? fmt_(o[CFG.C.ATT_OUT], 'HH:mm') : '',
      work: work,
      // 同じ行に複数の異常が乗ることがあるので、まとめて出す
      reason: reasons.join(' / ')
    });
  });

  out.sort(function (a, b) { return a.date < b.date ? 1 : -1; }); // 新しい順
  return { total: out.length, rows: out.slice(0, 100) };
}

/** 個人別の月間カレンダー。ym は 'yyyy-MM' */
function dashCalendar(key, empId, ym) {
  verifyAdmin_(key);
  const emps = listEmployees_();
  const emp = emps.filter(function (e) { return e.empId === String(empId); })[0];
  if (!emp) throw new Error('社員が見つかりません: ' + empId);

  const parts = ym.split('-');
  const y = Number(parts[0]), m = Number(parts[1]);
  const lastDay = new Date(y, m, 0).getDate(); // m月0日 = m月の末日

  // その月の勤怠を日付キーで引けるようにする
  const d = readSheet_(CFG.SH.ATT);
  const map = {};
  d.values.forEach(function (row) {
    const o = rowToObj_(row, d.map);
    if (String(o[CFG.C.ATT_ID] || '') !== String(empId)) return;
    const date = toDateStr_(o[CFG.C.ATT_DATE]);
    if (date.substring(0, 7) !== ym) return;
    map[date] = o;
  });

  const days = [];
  let totalWork = 0, totalOt = 0, workedDays = 0;

  for (let day = 1; day <= lastDay; day++) {
    const date = ym + '-' + ('0' + day).slice(-2);
    const o = map[date];
    const cell = {
      date: date, day: day,
      dow: new Date(y, m - 1, day).getDay(),
      in: '', out: '', work: null, status: ''
    };
    if (o) {
      cell.in = o[CFG.C.ATT_IN] instanceof Date ? fmt_(o[CFG.C.ATT_IN], 'HH:mm') : '';
      cell.out = o[CFG.C.ATT_OUT] instanceof Date ? fmt_(o[CFG.C.ATT_OUT], 'HH:mm') : '';
      cell.work = o[CFG.C.ATT_WORK] === '' ? null : Number(o[CFG.C.ATT_WORK]);
      cell.status = o[CFG.C.ATT_STATUS] || '';
      if (cell.work) {
        totalWork += cell.work;
        totalOt += Number(o[CFG.C.ATT_OT] || 0);
        workedDays++;
      }
    }
    days.push(cell);
  }

  return {
    empId: emp.empId, name: emp.name, ym: ym, days: days,
    summary: { days: workedDays, work: +totalWork.toFixed(2), ot: +totalOt.toFixed(2) },
    employees: emps.filter(function (e) { return e.active; })
                   .map(function (e) { return { empId: e.empId, name: e.name }; })
  };
}

/** 指定日の勤怠を社員IDで引けるオブジェクトにする */
function attendanceIndexByDate_(dateStr) {
  const d = readSheet_(CFG.SH.ATT);
  const idx = {};
  d.values.forEach(function (row) {
    const o = rowToObj_(row, d.map);
    if (toDateStr_(o[CFG.C.ATT_DATE]) !== dateStr) return;
    idx[String(o[CFG.C.ATT_ID] || '')] = o;
  });
  return idx;
}

/** 画面の初期表示に必要なものをまとめて返す(往復を減らす) */
function dashInit(key) {
  verifyAdmin_(key);
  const emps = listEmployees_().filter(function (e) { return e.active; })
    .map(function (e) { return { empId: e.empId, name: e.name }; });
  return {
    employees: emps,
    thisMonth: fmt_(new Date(), 'yyyy-MM'),
    sheetUrl: ss_().getUrl()
  };
}
