/* GASサービスをスタブして Punch.gs / Payroll.gs のロジックを実走させる検証 */
const fs = require('fs'), vm = require('vm'), path = require('path');
const ROOT = process.argv[2];

// ---- 時刻の固定 ----
const RealDate = Date; let FAKE = null;
class FakeDate extends RealDate {
  constructor(...a) { (a.length === 0 && FAKE) ? super(FAKE) : super(...a); }
}
FakeDate.now = () => (FAKE ? new RealDate(FAKE).getTime() : RealDate.now());
const at = s => { FAKE = s; };

// ---- GASサービスのスタブ ----
const crypto = require('crypto');
const props = { QR_SECRET: 'test-secret-key-0123456789', ADMIN_KEY: 'admin' };
const cache = new Map();

function fmtTZ(d, tz, p) {
  const g = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year:'numeric', month:'2-digit',
    day:'2-digit', hour:'2-digit', minute:'2-digit', second:'2-digit', hour12:false })
    .formatToParts(d).reduce((o, x) => (o[x.type] = x.value, o), {});
  return p.replace('yyyy', g.year).replace('MM', g.month).replace('dd', g.day)
          .replace('HH', g.hour === '24' ? '00' : g.hour).replace('mm', g.minute).replace('ss', g.second);
}

const sandbox = {
  console, Date: FakeDate, Intl, JSON, Math, String, Number, Array, Object, isNaN, parseInt,
  Utilities: {
    formatDate: fmtTZ,
    computeHmacSha256Signature: (v, k) => Array.from(crypto.createHmac('sha256', k).update(v).digest()),
    base64EncodeWebSafe: b => Buffer.from(b).toString('base64').replace(/\+/g,'-').replace(/\//g,'_'),
    getUuid: () => crypto.randomUUID(), sleep: () => {}
  },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k,v)=>props[k]=v }) },
  CacheService: { getScriptCache: () => ({
      get: k => { const e = cache.get(k); return e && e.exp > FakeDate.now() ? e.v : null; },
      put: (k,v,s) => cache.set(k, { v, exp: FakeDate.now() + s*1000 }), removeAll: ks => ks.forEach(k=>cache.delete(k)) }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) },
  MailApp: { sendEmail: (...a) => console.log('  [mail]', a[1]) },
  SpreadsheetApp: { getUi: () => ({ alert: m => console.log('  [alert]', m) }) }
};
vm.createContext(sandbox);

// ---- 実コードを読み込む(Sheets.gsだけメモリ実装に差し替え) ----
for (const f of ['Config.js', 'Punch.js', 'Payroll.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'gas', f), 'utf8'), sandbox, { filename: f });
}

// ---- インメモリのシート実装(Sheets.gs 相当) ----
vm.runInContext(`
const DB = { rows: {} };
DB.rows[CFG.SH.ATT] = []; DB.rows[CFG.SH.PAY] = []; DB.rows[CFG.SH.RAW] = [];
const EMPLOYEES = [
  { row: 2, empId: 'A001', name: '田中 太郎', type: '時給', wage: 1200, base: 0, active: true, token: '' },
  { row: 3, empId: 'A002', name: '鈴木 花子', type: '月給', wage: 0, base: 280000, active: true, token: '' },
  { row: 4, empId: 'A003', name: '退職 済',   type: '時給', wage: 1000, base: 0, active: false, token: '' }
];
function listEmployees_() { return EMPLOYEES; }
function getEmployee_(id) { return EMPLOYEES.filter(e => e.empId === id)[0] || null; }
function clearEmployeeCache() {}
function rowToObj_(row) { return row; }
function readSheet_(name) { return { values: DB.rows[name], map: {}, firstRow: 2 }; }
function appendRow_(name, o) { DB.rows[name].push(Object.assign({}, o)); return DB.rows[name].length + 1; }
function updateRow_(name, rowIndex, o) { Object.assign(DB.rows[name][rowIndex - 2], o); }
function logRaw_() {}
function toDateStr_(v) { return v instanceof Date ? fmt_(v, 'yyyy-MM-dd') : String(v || '').substring(0, 10); }
function findAttendanceRow_(empId, dateStr) {
  const rows = DB.rows[CFG.SH.ATT];
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][CFG.C.ATT_ID] === empId && toDateStr_(rows[i][CFG.C.ATT_DATE]) === dateStr)
      return { row: 2 + i, values: rows[i], map: {} };
  }
  return null;
}
function ss_() { return { getUrl: () => '' }; }
`, sandbox);

// ---- テスト ----
let pass = 0, fail = 0;
const eq = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log((ok ? '  ✅ ' : '  ❌ ') + label + (ok ? '' : `\n      got=${JSON.stringify(got)}\n     want=${JSON.stringify(want)}`));
  ok ? pass++ : fail++;
};
const run = code => vm.runInContext(code, sandbox);
const tokenOf = id => run(`makeToken_('${id}')`);

console.log('\n■ QRトークンの署名検証');
eq('正しいトークンは社員IDを返す', run(`verifyToken_(makeToken_('A001'))`), 'A001');
eq('署名を1文字改ざんすると弾かれる', run(`verifyToken_(makeToken_('A001').slice(0,-1) + 'X')`), null);
eq('社員IDだけ差し替えても弾かれる', run(`verifyToken_('AT1.A002.' + makeToken_('A001').split('.')[2])`), null);
eq('でたらめな文字列は弾かれる', run(`verifyToken_('hello')`), null);

console.log('\n■ 打刻フロー（A001: 時給1200円）');
at('2026-07-01T08:55:00+09:00');
let r = run(`punch(${JSON.stringify(tokenOf('A001'))})`);
eq('1回目 → 出勤 08:55', [r.ok, r.type, r.time, r.name], [true, '出勤', '08:55', '田中 太郎']);

r = run(`punch(${JSON.stringify(tokenOf('A001'))})`);
eq('直後の再スキャンはクールダウンで無視', r.ok, false);

at('2026-07-01T18:10:00+09:00');
r = run(`punch(${JSON.stringify(tokenOf('A001'))})`);
// 拘束 9h15m → 8h超なので休憩60分控除 → 実働 8.25h、残業 0.25h
eq('2回目 → 退勤、休憩60分控除で実働8.25h', [r.type, r.time, r.work], ['退勤', '18:10', 8.25]);

at('2026-07-01T18:40:00+09:00');
r = run(`punch(${JSON.stringify(tokenOf('A001'))})`);
eq('3回目 → 退勤時刻の打ち直し（実働8.75h）', [r.type, r.work], ['退勤', 8.75]);
eq('行は1日1本のまま', run(`DB.rows[CFG.SH.ATT].length`), 1);

console.log('\n■ 別日・別社員・在籍外');
at('2026-07-02T09:00:00+09:00');
run(`punch(${JSON.stringify(tokenOf('A001'))})`);
at('2026-07-02T14:00:00+09:00');
r = run(`punch(${JSON.stringify(tokenOf('A001'))})`);
eq('5時間勤務は休憩控除なし', r.work, 5);

at('2026-07-02T09:00:00+09:00');
r = run(`punch(${JSON.stringify(tokenOf('A003'))})`);
eq('在籍フラグOFFの社員は打刻不可', [r.ok, r.message.indexOf('在籍') >= 0], [false, true]);

console.log('\n■ 月次給与集計');
at('2026-08-01T04:00:00+09:00');
const res = run(`closeMonth('2026-07')`);
const pay = run(`JSON.stringify(DB.rows[CFG.SH.PAY])`);
const rows = JSON.parse(pay);
const a1 = rows.filter(x => x['社員ID'] === 'A001')[0];
const a2 = rows.filter(x => x['社員ID'] === 'A002')[0];
// 7/1: 実働8.75(残業0.75)、7/2: 実働5.0 → 計13.75h、残業0.75h
// 通常13.0h×1200 + 残業0.75h×1200×1.25 = 15600 + 1125 = 16725
eq('A001 出勤日数2日 / 総実働13.75h / 残業0.75h', [a1['出勤日数'], a1['総実働'], a1['残業']], [2, 13.75, 0.75]);
eq('A001 支給額 16,725円（残業1.25倍込み）', a1['支給額'], 16725);
eq('A002 月給者は勤怠ゼロでも基本給280,000円', a2['支給額'], 280000);
eq('A003 在籍外・勤怠ゼロの時給者は明細を作らない', rows.filter(x => x['社員ID'] === 'A003').length, 0);

console.log('\n■ 締め後の保護');
run(`DB.rows[CFG.SH.PAY][0]['確定'] = true; DB.rows[CFG.SH.PAY][0]['支給額'] = 99999;`);
run(`closeMonth('2026-07')`);
eq('「確定」済みの明細は再集計で上書きされない', run(`DB.rows[CFG.SH.PAY][0]['支給額']`), 99999);

console.log('\n■ 退勤打刻もれの検知');
at('2026-07-03T09:00:00+09:00');
run(`punch(${JSON.stringify(tokenOf('A001'))})`);  // 出勤したまま放置
at('2026-07-04T03:00:00+09:00');
const un = run(`JSON.stringify(checkUnclosed())`);
eq('前日の「勤務中」を検知して1件返す', JSON.parse(un).length, 1);
eq('該当行は「要確認」になる', run(`DB.rows[CFG.SH.ATT].filter(r => r['ステータス'] === '要確認').length`), 1);

console.log('\n' + (fail ? `❌ ${fail} 件失敗 / ` : '') + `✅ ${pass} 件成功\n`);
process.exit(fail ? 1 : 0);
