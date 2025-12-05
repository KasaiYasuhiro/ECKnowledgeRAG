/**************************************************
 * 契約マスタ／解約ロジック スナップショット履歴
 **************************************************/
/**
 * ヘルパー: 履歴シートを取得（なければ作成）
 * 1・2行目に元シートのヘッダーをコピーしておく
 */
function getOrCreateHistorySheet_(sourceSheetName, historySheetName) {
  const ss = SpreadsheetApp.getActive();
  let hist = ss.getSheetByName(historySheetName);
  if (!hist) {
    const src = ss.getSheetByName(sourceSheetName);
    if (!src) throw new Error(`${sourceSheetName} シートが見つかりません`);

    hist = ss.insertSheet(historySheetName);

    // 1〜2行目のヘッダーをコピー
    const lastCol = src.getLastColumn();
    const header  = src.getRange(1, 1, 2, lastCol).getValues();
    // 履歴シートの1〜2行目はヘッダーではなく「メタ情報」を置きたいので、
    // 3行目から元シートのヘッダーを置く。
    hist.getRange(3, 1, 2, lastCol).setValues(header);

    // 1行目に履歴用のメタヘッダー
    hist.getRange(1, 1, 1, 4).setValues([[
      'snapshot_ts',
      'source_sheet',
      'course_id',
      'version_note'
    ]]);
  }
  return hist;
}
/**
 * 共通ヘルパー：指定シートの指定行を履歴シートにコピー
 */
function snapshotRowToHistory_(sourceSheetName, historySheetName, row, versionNote) {
  const ss   = SpreadsheetApp.getActive();
  const src  = ss.getSheetByName(sourceSheetName);
  if (!src) throw new Error(`${sourceSheetName} シートが見つかりません`);

  const hist = getOrCreateHistorySheet_(sourceSheetName, historySheetName);

  const lastCol = src.getLastColumn();
  const values  = src.getRange(row, 1, 1, lastCol).getValues()[0];

  // course_id の列位置（既にCOL_COURSE_ID定数があるので流用可）
  const courseId = src.getRange(row, COL_COURSE_ID).getValue();

  const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  const lastHistRow = hist.getLastRow();
  const writeRow    = lastHistRow + 1;

  // 1〜4列：スナップショットメタ情報
  hist.getRange(writeRow, 1, 1, 4).setValues([[
    ts,
    sourceSheetName,
    courseId,
    versionNote || ''
  ]]);

  // 5列目以降に元行を丸コピー
  hist.getRange(writeRow, 5, 1, lastCol).setValues([values]);
}
/**
 * 選択行の contract_master をスナップショット保存
 */
function snapshotSelectedContractRowToHistory() {
  const ss   = SpreadsheetApp.getActive();
  const ui   = SpreadsheetApp.getUi();
  const sh   = ss.getSheetByName(SHEET_CM);
  if (!sh) {
    ui.alert('contract_master シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('スナップショットしたい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= 2) {
    ui.alert('3行目以降のデータ行を選択してください');
    return;
  }

  const res = ui.prompt(
    'バージョンメモ（任意）',
    '今回のスナップショットの理由や変更内容をメモしておくと後で便利です',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const note = res.getResponseText();
  snapshotRowToHistory_(SHEET_CM, SHEET_CM_HISTORY, row, note);
  ui.alert(`contract_master 行 ${row} のスナップショットを保存しました`);
}
/**
 * 選択行の contract_logic_rules をスナップショット保存
 */
function snapshotSelectedLogicRowToHistory() {
  const ss   = SpreadsheetApp.getActive();
  const ui   = SpreadsheetApp.getUi();
  const sh   = ss.getSheetByName(SHEET_CL);
  if (!sh) {
    ui.alert('contract_logic_rules シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('スナップショットしたい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= 2) {
    ui.alert('3行目以降のデータ行を選択してください');
    return;
  }

  const res = ui.prompt(
    'バージョンメモ（任意）',
    '今回のスナップショットの理由や変更内容をメモしておくと後で便利です',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const note = res.getResponseText();
  snapshotRowToHistory_(SHEET_CL, SHEET_CL_HISTORY, row, note);
  ui.alert(`contract_logic_rules 行 ${row} のスナップショットを保存しました`);
}
