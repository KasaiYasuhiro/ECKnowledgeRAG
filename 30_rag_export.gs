// 30_rag_export.gs
// ナレッジDB → RAG用CSVエクスポート関連

/**
 * CSVを安全に生成する関数
 * （カンマ・改行・ダブルクォートを正しくエスケープ）
 */
function toCsv_(rows) {
  return rows
    .map(row => row.map(field => {
      if (field === null || field === undefined) return '';
      const str = String(field);
      // ダブルクォート → "" に変換
      const escaped = str.replace(/"/g, '""');
      // カンマ or 改行 を含む場合は "" で囲む
      if (/[,"\n]/.test(escaped)) {
        return `"${escaped}"`;
      }
      return escaped;
    }).join(','))
    .join('\n');
}

/**
 * ナレッジDBからRAG用CSVをエクスポートする
 * - ステータスが「承認済み」の行だけを対象
 * - テキストを1フィールドにまとめて出力
 * - Driveフォルダ「05_RAG連携」にCSVを保存
 */
function exportRagCsv() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_KNOWLEDGE_DB); // ← 定数を利用（'ナレッジDB'）

  if (!sheet) {
    throw new Error('ナレッジDB シートが見つかりません');
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length <= 1) {
    throw new Error('ナレッジDBにデータ行がありません');
  }

  // 1行目をヘッダーとして想定（KNOW_HEADER_ROW = 1）
  const header = values[0];
  const dataRows = values.slice(1);

  // 出力CSVのヘッダー
  const output = [];
  output.push([
    'id',
    'title',
    'text',
    'main_category',
    'sub_categories',
    'type',
    'product',
    'client',
    'updated_at',
    'status',
    'source_url',
    'tags'
  ]);

  dataRows.forEach(row => {
    const KNOW_ID    = row[0];
    const TITLE      = row[1];
    const SUMMARY    = row[2];
    const MAIN_CAT   = row[3];
    const SUB_CATS   = row[4];
    const TYPE       = row[5];
    const PRODUCT    = row[6];
    const CLIENT     = row[7];
    const CORE       = row[8];
    const DELTA      = row[9];
    const NG         = row[10];
    const REASON     = row[11];
    const SOURCE_URL = row[12];
    const DUE        = row[13];
    const UPDATED_AT = row[14];
    const AUTHOR     = row[15];
    const STATUS     = row[16];

    // 承認済みだけ出力
    if (STATUS !== '承認済み') return;
    if (!CORE) return; // 空行はスキップ

    // 本文結合
    const textParts = [];

    if (TITLE)   textParts.push('【タイトル】\n' + TITLE);
    if (SUMMARY) textParts.push('【概要】\n' + SUMMARY);
    if (CORE)    textParts.push('【本文（共通ルール）】\n' + CORE);
    if (DELTA)   textParts.push('【差分・例外ルール】\n' + DELTA);
    if (NG)      textParts.push('【禁止事項・注意事項】\n' + NG);
    if (REASON)  textParts.push('【更新理由】\n' + REASON);
    if (DUE)     textParts.push('【希望反映期限】\n' + DUE);
    if (AUTHOR)  textParts.push('【登録者】\n' + AUTHOR);

    const text = textParts.join('\n\n');

    // タグ
    const tags = [
      MAIN_CAT,
      SUB_CATS,
      TYPE,
      PRODUCT,
      CLIENT
    ].filter(String).join(', ');

    // 日付整形
    let updatedStr = '';
    if (UPDATED_AT instanceof Date) {
      updatedStr = Utilities.formatDate(UPDATED_AT, 'Asia/Tokyo', 'yyyy-MM-dd\'T\'HH:mm:ss');
    } else if (UPDATED_AT) {
      updatedStr = UPDATED_AT;
    }

    output.push([
      KNOW_ID,
      TITLE,
      text,
      MAIN_CAT,
      SUB_CATS,
      TYPE,
      PRODUCT,
      CLIENT,
      updatedStr,
      STATUS,
      SOURCE_URL,
      tags
    ]);
  });

  if (output.length <= 1) {
    throw new Error('エクスポート対象の承認済みナレッジがありません');
  }

  const csvString = toCsv_(output);

  const folderName = '05_RAG連携';
  const folder = getOrCreateFolderByName_(folderName);

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `knowledge_rag_${timestamp}.csv`;

  const blob = Utilities.newBlob(csvString, 'text/csv', fileName);
  folder.createFile(blob);
}

/**
 * 指定名のフォルダを取得（なければ作成）
 */
function getOrCreateFolderByName_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) {
    return it.next();
  }
  return DriveApp.createFolder(name);
}

/**************************************************
 * RAG用CSV出力（30_rag_export.gs）
 **************************************************/

/**
 * 05_RAG連携 フォルダを取得（なければ作成）
 * ・スプレッドシートと同じ親フォルダ配下に作成
 * ・親フォルダが見つからなければマイドライブ直下に作成
 */
function getOrCreateRagFolder_() {
  var parentFolder = getSpreadsheetParentFolder_(); // 共通ユーティリティ
  return getOrCreateChildFolder_(parentFolder, RAG_FOLDER_NAME); // 00_constants.gs 側の定数を利用
}

/**
 * RAGエクスポート用ファイルを上書きしたい場合に使用する削除ヘルパー
 * - 実装は共通ユーティリティのラッパー
 */
function deleteRagFileIfExists_(fileName) {
  var ragFolder = getOrCreateRagFolder_();
  removeFilesByName_(ragFolder, fileName);
}

