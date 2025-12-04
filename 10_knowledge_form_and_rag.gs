// 10_knowledge_form_and_rag.gs

function handleKnowledgeFormSubmit(e) {
  // もともとの onFormSubmit の中身をここにコピー
    if (!e || !e.values) {
    // フォームトリガー以外から呼ばれた場合は何もしない
    return;
  }

    // フォームの回答が紐づいているスプレッドシート
  const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 本番ナレッジDBシート
  const dbSheet = ss.getSheetByName(SHEET_KNOWLEDGE_DB); // 'ナレッジDB'
  if (!dbSheet) {
    throw new Error('ナレッジDB シートが見つかりません');
  }
  // 回答内容（配列）を取得
  const v = e.values; // [タイムスタンプ, Q1, Q2, ...] など

  // ... 以降、元 onFormSubmit 内の処理をそのまま貼り付け ...
  
  // 既存行数から連番IDを作成（ヘッダー行を含む）
  const lastRow = dbSheet.getLastRow();  // 例：ヘッダーのみなら 1
  const idNumber = lastRow;              // 1 → N0001, 2 → N0002 ...
  const knowId = 'N' + Utilities.formatString('%04d', idNumber);

  // ナレッジDBに追記する1行分のデータを組み立て
  const row = [
    knowId,     // A: KNOW_ID
    v[1],       // B: タイトル（Q1）
    v[2],       // C: 概要（Q2）
    v[3],       // D: 主カテゴリ（Q3）
    v[4],       // E: 副カテゴリ（Q4）
    v[5],       // F: 種別（Q5）
    v[7],       // G: 商材（Q7）
    v[6],       // H: クライアント（Q6）
    v[8],       // I: 本文_Core（Q8）
    v[9],       // J: 本文_Delta（Q9）
    v[10],      // K: 禁止事項（Q10）
    v[11],      // L: 更新理由（Q11）
    v[12],      // M: 参照URL（Q12）
    v[14],      // N: 希望反映期限（Q14）
    new Date(), // O: 最終更新日（スクリプト側で現在時刻）
    v[13],      // P: 登録者（Q13）
    '承認待ち'  // Q: ステータス（初期値）
  ];

  // ナレッジDBシートの末尾に追加
  dbSheet.appendRow(row);
}

// 90_utils.gs でも 10_knowledge_form_and_rag.gs でもOK（契約系なので 20番台でも可）
function handleLastUpdatedOnEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 対象シート以外は無視
  if (TARGET_SHEETS.indexOf(sheetName) === -1) return;

  const row = range.getRow();
  const col = range.getColumn();

  // 1・2行目（和名 / 英名ヘッダー）は無視
  if (row <= HEADER_ROW) return;

  // last_updated 列の位置（存在しなければ作成）
  const lastUpdatedCol = getOrCreateHeaderColumn_(sheet, LAST_UPDATED_HEADER);

  // last_updated 自身を編集した場合は何もしない（ループ防止）
  if (col === lastUpdatedCol) return;

  // 編集があったらタイムスタンプを更新
  sheet.getRange(row, lastUpdatedCol).setValue(new Date());

  
}


