/**
 * Punch.gs — 打刻の中核ロジックと Web エンドポイント
 *
 * エンドポイント:
 *   POST /exec  body: {"token":"XXX"}  -> JSON で打刻(本命)
 *   GET  /exec?token=XXX&callback=cb   -> JSONP で打刻(CORS回避の保険)
 *   GET  /exec?page=admin&key=YYY      -> QR一括発行の管理画面
 *   GET  /exec                         -> 読み取りページ(GASホスト版)
 */

// ============ Web エンドポイント ============

function doPost(e) {
  let token = '', ua = '';
  try {
    const body = JSON.parse(e.postData.contents);
    token = body.token || '';
    ua = body.ua || '';
  } catch (err) {
    token = (e.parameter && e.parameter.token) || '';
  }
  return json_(punch(token, ua));
}

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.page === 'admin') {
    if (p.key !== prop_('ADMIN_KEY')) {
      return HtmlService.createHtmlOutput('<h1>403</h1><p>キーが違います</p>');
    }
    const t = HtmlService.createTemplateFromFile('admin');
    t.employees = listEmployees_();
    return t.evaluate()
      .setTitle('QRコード発行')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (p.page === 'dash') {
    // 画面(シェル)自体はキー無しで配信してよい。中身のデータは Dashboard.js が
    // 毎回 ADMIN_KEY を検証するので、合言葉を知らない相手には何も返らない。
    // URLに合言葉を埋めると、共有やコピペのたびに漏れる危険があるため画面で入力させる。
    const t = HtmlService.createTemplateFromFile('dash');
    t.key = p.key || '';
    return t.evaluate()
      .setTitle('勤怠ダッシュボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }

  // 自己診断。設定漏れを外部から確認するためのもの(氏名などの中身は出さない)
  if (p.page === 'diag') return json_(diag_());

  if (p.token) {
    const result = punch(p.token, p.ua || '');
    if (p.callback) {
      return ContentService
        .createTextOutput(p.callback + '(' + JSON.stringify(result) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json_(result);
  }

  return HtmlService.createTemplateFromFile('scan').evaluate()
    .setTitle('勤怠打刻')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 設定が揃っているかを返す。値そのものは返さない */
function diag_() {
  const out = { ok: false, sheetIdSet: false, secretSet: false, sheets: [], employees: null, error: null };
  try {
    out.sheetIdSet = !!prop_('SHEET_ID', false);
    out.secretSet = !!prop_('QR_SECRET', false);
    out.tz = CFG.TZ;
    out.now = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
    const ss = ss_();
    out.sheets = ss.getSheets().map(function (s) { return s.getName(); });
    const emps = listEmployees_();
    out.employees = {
      total: emps.length,
      active: emps.filter(function (e) { return e.active; }).length,
      // 氏名は返さない。IDだけでは署名が作れないので打刻はできない
      ids: emps.map(function (e) { return e.empId + (e.active ? '' : '(在籍外)'); })
    };
    out.ok = true;
  } catch (err) {
    out.error = err.message;
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ 打刻本体 ============

/**
 * @param {string} token QRの中身 "AT1.社員ID.署名"
 * @param {string} ua    端末情報(監査ログ用)
 * @return {{ok:boolean, type?:string, name?:string, time?:string, work?:number, message:string}}
 */
function punch(token, ua) {
  // 同時アクセスで同じ日の行が2本できるのを防ぐ
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { ok: false, message: '混み合っています。もう一度かざしてください' };
  }
  try {
    const empId = verifyToken_(token);
    if (!empId) {
      logRaw_('', '不正', '署名検証NG: ' + String(token).substring(0, 40), ua);
      return { ok: false, message: 'このQRコードは無効です' };
    }

    const emp = getEmployee_(empId);
    if (!emp) {
      logRaw_(empId, '不正', '社員マスタに無し', ua);
      return { ok: false, message: '社員が見つかりません (' + empId + ')' };
    }
    if (!emp.active) {
      logRaw_(empId, '拒否', '在籍フラグOFF', ua);
      return { ok: false, message: emp.name + ' さんは在籍中ではありません' };
    }

    // クールダウン: 連続スキャンの二度打ちを無視
    const cache = CacheService.getScriptCache();
    const ckey = 'cd_' + empId;
    if (cache.get(ckey)) {
      return { ok: false, name: emp.name, message: '打刻済みです（' + CFG.COOLDOWN_SEC + '秒以内の再打刻は無視されます）' };
    }
    cache.put(ckey, '1', CFG.COOLDOWN_SEC);

    const now = new Date();
    const hhmm = fmt_(now, 'HH:mm');
    const bizDate = businessDate_(now);
    const hit = findAttendanceRow_(empId, bizDate);

    // --- 出勤 ---
    if (!hit) {
      const o = {};
      o[CFG.C.ATT_DATE]    = new Date(bizDate + 'T00:00:00+09:00');
      o[CFG.C.ATT_ID]      = empId;
      o[CFG.C.ATT_NAME]    = emp.name;
      o[CFG.C.ATT_IN]      = now;
      o[CFG.C.ATT_STATUS]  = '勤務中';
      o[CFG.C.ATT_PUNCHES] = 1;
      o[CFG.C.ATT_UPDATED] = now;
      appendRow_(CFG.SH.ATT, o);
      logRaw_(empId, '出勤', 'OK ' + hhmm, ua);
      return { ok: true, type: '出勤', name: emp.name, time: hhmm, message: 'おはようございます' };
    }

    const cur = rowToObj_(hit.values, hit.map);
    const inVal = cur[CFG.C.ATT_IN];

    // 出勤が空の行が残っていたら、出勤として埋める
    if (!(inVal instanceof Date)) {
      const fix = {};
      fix[CFG.C.ATT_IN] = now;
      fix[CFG.C.ATT_STATUS] = '勤務中';
      fix[CFG.C.ATT_UPDATED] = now;
      updateRow_(CFG.SH.ATT, hit.row, fix);
      logRaw_(empId, '出勤', 'OK(補正) ' + hhmm, ua);
      return { ok: true, type: '出勤', name: emp.name, time: hhmm, message: 'おはようございます' };
    }

    // --- 退勤 / 退勤時刻の打ち直し ---
    const manualBreak = (cur[CFG.C.ATT_BREAK] === '' || cur[CFG.C.ATT_BREAK] === null) ? null : Number(cur[CFG.C.ATT_BREAK]);
    const calc = calcWork_(inVal, now, manualBreak);
    const punches = Number(cur[CFG.C.ATT_PUNCHES] || 1) + 1;
    const already = cur[CFG.C.ATT_OUT] instanceof Date;

    const up = {};
    up[CFG.C.ATT_OUT]     = now;
    up[CFG.C.ATT_BREAK]   = calc.breakMin;
    up[CFG.C.ATT_WORK]    = calc.work;
    up[CFG.C.ATT_OT]      = calc.overtime;
    up[CFG.C.ATT_STATUS]  = calc.work < 0 ? '要確認' : '退勤済';
    up[CFG.C.ATT_PUNCHES] = punches;
    up[CFG.C.ATT_UPDATED] = now;
    updateRow_(CFG.SH.ATT, hit.row, up);
    logRaw_(empId, '退勤', (already ? '打ち直し ' : 'OK ') + hhmm + ' 実働' + calc.work + 'h', ua);

    return {
      ok: true,
      type: '退勤',
      name: emp.name,
      time: hhmm,
      work: calc.work,
      message: already
        ? '退勤時刻を ' + hhmm + ' に更新しました（実働 ' + calc.work + 'h）'
        : 'お疲れさまでした（実働 ' + calc.work + 'h）'
    };
  } catch (err) {
    console.error(err);
    logRaw_('', 'エラー', err.message, ua);
    return { ok: false, message: 'エラー: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// ============ トークン(QRの中身) ============

/** QR文字列を生成: "AT1.A001.<署名16文字>" */
function makeToken_(empId) {
  return 'AT1.' + empId + '.' + sign_(empId);
}

/** 署名を検証。OKなら社員IDを返す。NGなら null */
function verifyToken_(token) {
  if (!token) return null;
  const parts = String(token).trim().split('.');
  if (parts.length !== 3 || parts[0] !== 'AT1') return null;
  const empId = parts[1];
  return sign_(empId) === parts[2] ? empId : null;
}

function sign_(empId) {
  const raw = Utilities.computeHmacSha256Signature(empId, prop_('QR_SECRET'));
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '').substring(0, 16);
}

// ============ 計算 ============

/**
 * 実働時間の計算
 * @return {{work:number, overtime:number, breakMin:number}} 単位は時間(小数)
 */
function calcWork_(inDate, outDate, manualBreakMin) {
  const gross = (outDate.getTime() - inDate.getTime()) / 3600000; // 拘束時間(h)

  let breakMin = manualBreakMin;
  if (breakMin === null || breakMin === undefined || isNaN(breakMin)) {
    breakMin = CFG.AUTO_BREAK ? autoBreak_(gross) : 0;
  }

  const work = roundHours_(gross - breakMin / 60);
  const overtime = Math.max(0, +(work - CFG.DAILY_REGULAR_HOURS).toFixed(2));
  return { work: +work.toFixed(2), overtime: overtime, breakMin: breakMin };
}

function autoBreak_(grossHours) {
  for (const r of CFG.AUTO_BREAK_RULES) {
    if (grossHours > r.overHours) return r.breakMin;
  }
  return 0;
}

function roundHours_(h) {
  if (CFG.TIME_ROUNDING === 'floor15') return Math.floor(h * 4) / 4;
  if (CFG.TIME_ROUNDING === 'round15') return Math.round(h * 4) / 4;
  return h;
}

// ============ 日時ユーティリティ ============

function fmt_(d, pattern) {
  return Utilities.formatDate(d, CFG.TZ, pattern);
}

/** 業務日を返す。DAY_BOUNDARY より前の打刻は前日扱い(深夜勤務対応) */
function businessDate_(d) {
  const parts = CFG.DAY_BOUNDARY.split(':');
  const offsetMs = (Number(parts[0]) * 60 + Number(parts[1])) * 60000;
  return fmt_(new Date(d.getTime() - offsetMs), 'yyyy-MM-dd');
}
