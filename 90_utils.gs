// 90_utils.gs
// ==============================
// 共通ユーティリティ関数
// ==============================

/**
 * null / undefined / 空文字まわりの安全な文字列化
 *
 * - null / undefined → ''
 * - それ以外        → String(v).trim()
 */
function safeStr_(v) {
  if (v == null) return '';        // null と undefined をまとめて排除
  return String(v).trim();
}

/**
 * "1,980円" などの文字列から数値部分だけ抜き出して Number にする。
 * 数字がなければ '' を返す。
 *
 * - null / undefined / '' → ''
 * - "1,980円" → 1980
 * - "500/330" なども前処理と組み合わせて利用可
 */
function parseNumberOrEmpty_(value) {
  if (value == null || value === '') return '';
  const text = String(value);
  const numStr = text.replace(/[^\d.-]/g, '');
  if (!numStr) return '';
  const num = Number(numStr);
  if (isNaN(num)) return '';
  return num;
}
/**
 * 指定シートの「英名ヘッダー行」から headerName を探して列番号を返す。
 * 見つからなければ、一番右に新しい列を作成してヘッダーを設定する。
 *
 * 想定：
 * - HEADER_ROW は 00_constants.gs 側で 2 に設定されている
 *   （1行目: 和名 / 2行目: 英名）
 * - 契約系シート（contract_master / contract_logic_rules）で利用
 */
function getOrCreateHeaderColumn_(sheet, headerName) {
  const headerRow = HEADER_ROW; // 00_constants のエイリアス（= CM_HEADER_ROW）

  // 英名ヘッダー行を取得
  const lastCol = sheet.getLastColumn();
  const headerRange = sheet.getRange(headerRow, 1, 1, lastCol);
  const headerValues = headerRange.getValues()[0];

  // 既存ヘッダーから検索
  for (let i = 0; i < headerValues.length; i++) {
    if (headerValues[i] === headerName) {
      return i + 1; // 列番号（1始まり）
    }
  }

  // 無ければヘッダー右側に追加
  const newCol = headerValues.length + 1;

  // 英名ヘッダーに追加
  sheet.getRange(headerRow, newCol).setValue(headerName);

  // 和名ヘッダーにも仮値を入れておく（HEADER_ROW > 1 のときだけ）
  if (headerRow > 1) {
    sheet.getRange(headerRow - 1, newCol).setValue(headerName);
  }

  return newCol;
}
/**
 * 編集内容を change_log シートに1行追記する共通ロガー
 *
 * 想定:
 * - LOG_SHEET_NAME でログシート名を指定（00_constants.gs）
 * - HEADER_ROW 行に英名ヘッダー（course_id など）がある
 * - COURSE_ID_HEADER に "course_id" が入っている
 */
function logChange_(e, sheet, row) {
  const ss = sheet.getParent();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

  // ログシートが無ければ自動作成
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, 8).setValues([[
      'timestamp',
      'sheet_name',
      'row',
      'course_id',
      'column_a1',
      'field_name',
      'old_value',
      'new_value'
    ]]);
  }

  const range     = e.range;
  const sheetName = sheet.getName();

  // course_id の列位置を取得して値を取り出す
  const courseIdCol = getOrCreateHeaderColumn_(sheet, COURSE_ID_HEADER);
  const courseId    = safeStr_(sheet.getRange(row, courseIdCol).getValue());

  // 編集セルの英名（HEADER_ROW 行）
  const fieldName = safeStr_(sheet.getRange(HEADER_ROW, range.getColumn()).getValue());

  // old / new 値
  const oldValue = (typeof e.oldValue !== 'undefined') ? e.oldValue : '';

  let newValue;
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    // 単一セル編集
    newValue = range.getValue();
  } else {
    // 複数セル編集
    newValue = '(multiple cells)';
  }

  // タイムスタンプ（文字列化しておく）
  const now = new Date();
  const tz  = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  const logRow = [
    timestamp,
    sheetName,
    row,
    courseId,
    range.getA1Notation(),
    fieldName,
    oldValue,
    newValue
  ];

  logSheet.appendRow(logRow);
}

/**
 * code_master から item に該当する value の一覧を取得する
 * 例: getMasterValues_('category') → ['subscription', 'single', ...]
 *
 * 返り値は value の一次元配列
 */
function getMasterValues_(itemName) {
  if (!itemName) return [];

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_CODE_MASTER);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();  
  // [item, value, desc]

  return data
    .filter(r => safeStr_(r[0]) === itemName && safeStr_(r[1]) !== '')
    .map(r => safeStr_(r[1]));
}


/**
 * 指定セルに現在時刻（last_updated）をセットする
 *
 * - 時刻はスクリプトのタイムゾーンに従う
 * - フォーマットは contract_master / logic_rules と統一（yyyy-MM-dd HH:mm:ss）
 */
function updateLastUpdated_(sheet, row, col) {
  if (!sheet || !row || !col) return; // 安全対策

  const now = new Date();
  const tz  = Session.getScriptTimeZone();
  const formatted = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  const cell = sheet.getRange(row, col);
  cell.setValue(formatted);
  cell.setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

