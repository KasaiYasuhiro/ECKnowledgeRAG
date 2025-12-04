/**************************************
 * メニュー追加（LLMツール配下に統合）
 **************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // --- サブメニュー定義 ---

  // 契約マスタツール
  const menuContractMaster = ui.createMenu('契約マスタ');
  menuContractMaster
    .addItem('コース表 → マスタ反映（新規のみ追加）', 'updateContractsFromCourseSource_AppendOnly')
    .addItem('コース表 → マスタ反映（既存も上書き）', 'updateContractsFromCourseSource_Overwrite')
    .addItem('fee_table_master 差額情報を補完', 'fillFeeTableDiffFromCourseSource');

  // ⚠ 注意タグ
  const menuWarning = ui.createMenu('⚠ 注意タグ');
  menuWarning
    .addItem('選択コースの注意タグを編集', 'openWarningTagSidebar');

  // 💳 支払い区分
  const menuPayment = ui.createMenu('💳 支払い区分');
  menuPayment
    .addItem('選択コースの支払い区分を編集', 'openPaymentTypeSidebar');

  // 契約ロジック
  const menuLogic = ui.createMenu('契約ロジック');
  menuLogic
    .addItem('解約金条件テンプレ反映…', 'fillExitFeeConditionTemplates');

  // ✅ マスタチェック
  const menuCheck = ui.createMenu('✅ マスタチェック');
  menuCheck
    .addItem('選択行をチェック（contract_master）', 'validateSelectedContractRow')
    .addItem('解約ロジック行をチェック（contract_logic_rules）', 'validateSelectedLogicRow')
    .addItem('contract_master 全行をレポート出力', 'validateAllContractRows')
    .addItem('contract_logic_rules 全行をレポート出力', 'validateAllLogicRows');

  // 📦 バージョン履歴
  const menuHistory = ui.createMenu('📦 バージョン履歴');
  menuHistory
    .addItem('選択行のスナップショット（contract_master）', 'snapshotSelectedContractRowToHistory')
    .addItem('選択行のスナップショット（contract_logic_rules）', 'snapshotSelectedLogicRowToHistory');

  // 📚 ナレッジ承認
  const menuKnowledgeApproval = ui.createMenu('📚 ナレッジ承認');
  menuKnowledgeApproval
    .addItem('選択ナレッジを承認', 'approveSelectedKnowledge')
    .addItem('選択ナレッジを差し戻し', 'rejectSelectedKnowledge');

  // ✅ 更新承認フロー
  const menuApprovalFlow = ui.createMenu('✅ 更新承認フロー');
  menuApprovalFlow
    .addItem('選択行を承認（ナレッジ）', 'approveFromApprovalSheet')
    .addItem('選択行を差し戻し（ナレッジ）', 'rejectFromApprovalSheet');

  // 📤 RAGエクスポート
  const menuRagExport = ui.createMenu('📤 RAGエクスポート');
  menuRagExport
    .addItem('契約マスタRAG CSV出力（生情報）', 'exportContractsRagCsv')
    .addItem('契約マスタRAG CSV出力（1コース1行・要約版）', 'exportContractsRagLongformCsv');

  // --- メインメニューに統合 ---
  const mainMenu = ui.createMenu('📘 LLMツール');
  mainMenu
    .addSubMenu(menuContractMaster)
    .addSubMenu(menuWarning)
    .addSubMenu(menuPayment)
    .addSubMenu(menuLogic)
    .addSubMenu(menuCheck)
    .addSubMenu(menuHistory)
    .addSubMenu(menuKnowledgeApproval)
    .addSubMenu(menuApprovalFlow)
    .addSubMenu(menuRagExport)
    .addToUi();
}

// コード.gs 側の onFormSubmit をこちらに移動
// そしてトリガーもこちらに設定する
// 10_knowledge_form_and_rag.gs 側の onFormSubmit は削除
// （onFormSubmit は特殊関数名なので、同じ名前が複数あるとエラーになる）
function onFormSubmit(e) {
  handleKnowledgeFormSubmit(e);
}

/**
 * セル編集時の自動処理
 */
function onEdit(e) {
  handleLastUpdatedOnEdit(e);
}