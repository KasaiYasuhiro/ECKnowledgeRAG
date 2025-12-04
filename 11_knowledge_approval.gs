/**************************************************
 * ナレッジDB 承認フロー（ナレッジDBシート側で操作）
 **************************************************/

/**
 * 選択行のナレッジを「承認済み」にする
 */
function approveSelectedKnowledge() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sh  = ss.getSheetByName(SHEET_KNOWLEDGE_DB);
  if (!sh) {
    ui.alert('ナレッジDB シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('ナレッジDB シートで承認したい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= KNOW_HEADER_ROW) {
    ui.alert('データ行（ヘッダー行より下）を選択してください');
    return;
  }

  changeKnowledgeStatus_(sh, row, '承認済み', '');
  ui.alert(`行 ${row} のナレッジを「承認済み」にしました`);
}

/**
 * 選択行のナレッジを「差し戻し」にする（理由入力付き）
 */
function rejectSelectedKnowledge() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sh  = ss.getSheetByName(SHEET_KNOWLEDGE_DB);
  if (!sh) {
    ui.alert('ナレッジDB シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('ナレッジDB シートで差し戻したい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= KNOW_HEADER_ROW) {
    ui.alert('データ行（ヘッダー行より下）を選択してください');
    return;
  }

  const res = ui.prompt(
    '差し戻し理由の入力',
    '差し戻し理由を入力してください（必須ではありません）',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return; // キャンセル
  }

  const reason = res.getResponseText() || '';
  changeKnowledgeStatus_(sh, row, '差し戻し', reason);
  ui.alert(`行 ${row} のナレッジを「差し戻し」にしました`);
}

/**
 * ナレッジDB：ステータス変更の共通処理
 * ・ステータス列を更新
 * ・管理メモ列に「日時 / 実行者 / 状態 / コメント」を追記
 */
function changeKnowledgeStatus_(sheet, row, newStatus, comment) {
  // ステータス更新
  sheet.getRange(row, COL_KNOW_STATUS).setValue(newStatus);

  // 管理メモの追記
  const adminCell = sheet.getRange(row, COL_KNOW_ADMIN_NOTE);
  const oldNote   = adminCell.getValue() || '';

  const ts   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || 'unknown_user';

  let line = `[${ts}] ${user} がステータスを「${newStatus}」に変更`;
  if (comment) {
    line += `／コメント: ${comment}`;
  }

  const newNote = oldNote ? (oldNote + '\n' + line) : line;
  adminCell.setValue(newNote);
}

/**
 * KNOW_ID を指定してナレッジDB側のステータスを更新する
 */
function setKnowledgeStatusById_(knowId, newStatus, comment) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_KNOWLEDGE_DB);
  if (!sh) {
    throw new Error('ナレッジDB シートが見つかりません');
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= KNOW_HEADER_ROW) return;

  // A列（KNOW_ID）を検索
  const idRange = sh.getRange(KNOW_HEADER_ROW + 1, 1, lastRow - KNOW_HEADER_ROW, 1);
  const ids     = idRange.getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(knowId)) {
      const row = KNOW_HEADER_ROW + 1 + i;
      changeKnowledgeStatus_(sh, row, newStatus, comment);
      return;
    }
  }

  throw new Error(`KNOW_ID "${knowId}" が見つかりません`);
}



/**************************************************
 * 更新承認フロー表 と ナレッジDB の連携（ナレッジ版）
 **************************************************/

/**
 * 更新承認フロー表で選択行を「承認（ナレッジ）」にする
 * ・更新承認フロー表のステータス更新
 * ・ナレッジDB側も承認済みにする
 */
function approveFromApprovalSheet() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sh  = ss.getSheetByName(SHEET_APPROVAL_FLOW);
  if (!sh) {
    ui.alert('更新承認フロー シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('承認したい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= AF_HEADER_ROW) {
    ui.alert('データ行（ヘッダー行より下）を選択してください');
    return;
  }

  const type   = sh.getRange(row, COL_AF_TYPE).getValue();
  const knowId = sh.getRange(row, COL_AF_TARGET).getValue();

  if (type !== 'knowledge') {
    ui.alert('現在は「種別 = knowledge」の行のみ自動承認対象としています');
    return;
  }
  if (!knowId) {
    ui.alert('対象キー（KNOW_ID）が空です');
    return;
  }

  // ナレッジDB側を承認済みに
  setKnowledgeStatusById_(knowId, '承認済み', '');

  // 承認フロー表のステータス更新
  const user = Session.getActiveUser().getEmail() || 'unknown_user';
  const ts   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  sh.getRange(row, COL_AF_STATUS).setValue('承認済み');
  sh.getRange(row, COL_AF_APPROVER).setValue(user);
  sh.getRange(row, COL_AF_UPDATED_AT).setValue(ts);
}

/**
 * 更新承認フロー表で選択行を「差し戻し（ナレッジ）」にする
 * ・更新承認フロー表のステータス更新
 * ・ナレッジDB側も差し戻しにする
 */
function rejectFromApprovalSheet() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sh  = ss.getSheetByName(SHEET_APPROVAL_FLOW);
  if (!sh) {
    ui.alert('更新承認フロー シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('差し戻したい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= AF_HEADER_ROW) {
    ui.alert('データ行（ヘッダー行より下）を選択してください');
    return;
  }

  const type   = sh.getRange(row, COL_AF_TYPE).getValue();
  const knowId = sh.getRange(row, COL_AF_TARGET).getValue();

  if (type !== 'knowledge') {
    ui.alert('現在は「種別 = knowledge」の行のみ自動差し戻し対象としています');
    return;
  }
  if (!knowId) {
    ui.alert('対象キー（KNOW_ID）が空です');
    return;
  }

  // 差し戻し理由を入力
  const res = ui.prompt(
    '差し戻し理由の入力',
    '差し戻し理由を入力してください（必須ではありません）',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const reason = res.getResponseText() || '';

  // ナレッジDB側を差し戻しに
  setKnowledgeStatusById_(knowId, '差し戻し', reason);

  // 承認フロー表のステータス更新
  const user = Session.getActiveUser().getEmail() || 'unknown_user';
  const ts   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  sh.getRange(row, COL_AF_STATUS).setValue('差し戻し');
  sh.getRange(row, COL_AF_APPROVER).setValue(user);
  sh.getRange(row, COL_AF_UPDATED_AT).setValue(ts);

  const oldNote = sh.getRange(row, COL_AF_NOTE).getValue() || '';
  const line    = `[${ts}] ${user} 差し戻し理由: ${reason}`;
  sh.getRange(row, COL_AF_NOTE).setValue(oldNote ? (oldNote + '\n' + line) : line);
}
