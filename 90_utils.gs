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
