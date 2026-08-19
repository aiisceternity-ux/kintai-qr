/**
 * Config.gs — 設定と定数
 *
 * 秘密情報はコードに直書きせず、スクリプトプロパティに入れる。
 *   GASエディタ > 左の歯車(プロジェクトの設定) > スクリプト プロパティ
 *
 * 必須プロパティ:
 *   QR_SECRET : QR署名用の共通鍵。ランダム32文字以上。漏れるとQRを偽造される
 *   ADMIN_KEY : 管理画面(QR発行)を開くための合言葉
 *
 * どちらも Setup.gs の initProject() が未設定なら自動生成する。
 */

const CFG = {
  TZ: 'Asia/Tokyo',

  // ---- シート名 ----
  SH: {
    EMP: '社員マスタ',
    ATT: '勤怠ログ',
    PAY: '給与',
    RAW: '打刻生ログ'
  },

  // ---- 打刻まわり ----
  /** 同一社員の連続打刻を無視する秒数。QRの二度読み事故を防ぐ */
  COOLDOWN_SEC: 60,
  /** 業務日の境界(HH:mm)。深夜勤務があるなら '05:00' などにすると翌2時の退勤が前日扱いになる */
  DAY_BOUNDARY: '00:00',

  // ---- 休憩の自動控除(労基法34条ベース) ----
  /** 拘束6時間超で45分、8時間超で60分を自動控除。休憩も打刻させるなら false */
  AUTO_BREAK: true,
  AUTO_BREAK_RULES: [
    { overHours: 8, breakMin: 60 },
    { overHours: 6, breakMin: 45 }
  ],

  // ---- 給与計算 ----
  DAILY_REGULAR_HOURS: 8,   // 1日の法定労働時間。超過分は割増
  OVERTIME_RATE: 1.25,      // 時間外割増率
  /** 実働時間の丸め: 'none' | 'floor15' | 'round15' */
  TIME_ROUNDING: 'none',

  // ---- 未退勤アラート ----
  /** 退勤打刻忘れを検知したらこのアドレスに通知。空なら送らない */
  ALERT_TO: '',

  // ---- 列名(シートのヘッダーを変えたらここも変える) ----
  C: {
    // 社員マスタ
    EMP_ID: '社員ID',
    EMP_NAME: '氏名',
    EMP_TYPE: '雇用区分',
    EMP_WAGE: '時給',
    EMP_BASE: '基本給',
    EMP_ACTIVE: '在籍',
    EMP_TOKEN: 'QRトークン',

    // 勤怠ログ
    ATT_DATE: '日付',
    ATT_ID: '社員ID',
    ATT_NAME: '氏名',
    ATT_IN: '出勤',
    ATT_OUT: '退勤',
    ATT_BREAK: '休憩(分)',
    ATT_WORK: '実働',
    ATT_OT: '残業',
    ATT_STATUS: 'ステータス',
    ATT_PUNCHES: '打刻回数',
    ATT_UPDATED: '更新日時',

    // 給与
    PAY_MONTH: '対象月',
    PAY_ID: '社員ID',
    PAY_NAME: '氏名',
    PAY_DAYS: '出勤日数',
    PAY_HOURS: '総実働',
    PAY_OT: '残業',
    PAY_AMOUNT: '支給額',
    PAY_FIXED: '確定',
    PAY_GEN: '生成日時',

    // 打刻生ログ(追記専用の監査ログ)
    RAW_TS: 'タイムスタンプ',
    RAW_ID: '社員ID',
    RAW_TYPE: '種別',
    RAW_RESULT: '結果',
    RAW_UA: '端末'
  }
};

/** シートの定義。initProject() がこの通りに作る */
const SCHEMA = {};
SCHEMA[CFG.SH.EMP] = {
  headers: [CFG.C.EMP_ID, CFG.C.EMP_NAME, CFG.C.EMP_TYPE, CFG.C.EMP_WAGE, CFG.C.EMP_BASE, CFG.C.EMP_ACTIVE, CFG.C.EMP_TOKEN],
  formats: { '時給': '#,##0', '基本給': '#,##0' },
  widths:  { 'QRトークン': 260, '氏名': 120 }
};
SCHEMA[CFG.SH.ATT] = {
  headers: [CFG.C.ATT_DATE, CFG.C.ATT_ID, CFG.C.ATT_NAME, CFG.C.ATT_IN, CFG.C.ATT_OUT, CFG.C.ATT_BREAK, CFG.C.ATT_WORK, CFG.C.ATT_OT, CFG.C.ATT_STATUS, CFG.C.ATT_PUNCHES, CFG.C.ATT_UPDATED],
  formats: { '日付': 'yyyy-mm-dd', '出勤': 'HH:mm', '退勤': 'HH:mm', '実働': '0.00', '残業': '0.00', '更新日時': 'yyyy-mm-dd HH:mm:ss' },
  widths:  {}
};
SCHEMA[CFG.SH.PAY] = {
  headers: [CFG.C.PAY_MONTH, CFG.C.PAY_ID, CFG.C.PAY_NAME, CFG.C.PAY_DAYS, CFG.C.PAY_HOURS, CFG.C.PAY_OT, CFG.C.PAY_AMOUNT, CFG.C.PAY_FIXED, CFG.C.PAY_GEN],
  formats: { '総実働': '0.00', '残業': '0.00', '支給額': '¥#,##0', '生成日時': 'yyyy-mm-dd HH:mm:ss' },
  widths:  {}
};
SCHEMA[CFG.SH.RAW] = {
  headers: [CFG.C.RAW_TS, CFG.C.RAW_ID, CFG.C.RAW_TYPE, CFG.C.RAW_RESULT, CFG.C.RAW_UA],
  formats: { 'タイムスタンプ': 'yyyy-mm-dd HH:mm:ss' },
  widths:  { '端末': 320, '結果': 220 }
};

function prop_(key, required) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v && required !== false) {
    throw new Error('スクリプトプロパティ ' + key + ' が未設定です。Setup.gs の initProject() を実行してください。');
  }
  return v;
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}
