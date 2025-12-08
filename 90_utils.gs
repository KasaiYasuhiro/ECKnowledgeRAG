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
