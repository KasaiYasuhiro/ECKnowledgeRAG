// 00_constants.gs
// ==============================
// 共通設定
// ==============================

// 「編集監視」の対象シート（last_updated 更新など）
const CM_TARGET_SHEETS = ['contract_master', 'contract_logic_rules'];

// 契約系シート共通のヘッダー行（和名/英名が2行ある前提）
const CM_HEADER_ROW = 2;

// 編集時に自動更新するヘッダー名
const CM_LAST_UPDATED_HEADER = 'last_updated';

// 契約マスター・ロジックで共通して使うキー項目
const CM_COURSE_ID_HEADER = 'course_id';

// 変更ログ用シート名
const CM_LOG_SHEET_NAME = 'change_log';


// ==============================
// シート名（マスター系）
// ==============================

// 契約マスター
const SHEET_CONTRACT_MASTER = 'contract_master';

// 契約警告フラグ
const SHEET_CONTRACT_WARNING_FLAGS = 'contract_warning_flags';

// コードマスター
const SHEET_CODE_MASTER = 'code_master';

// コース原本（契約自動生成の元）
const SHEET_COURSE_SOURCE = 'course_master_source';

// 契約ロジックルール
const SHEET_CONTRACT_LOGIC = 'contract_logic_rules';

// 料金テーブル
const SHEET_FEE_TABLE = 'fee_table_master';

// コース原本（料金テーブル用に参照）
const SHEET_COURSE_SOURCE_FEE = 'course_master_source';


// ==============================
// 契約マスター：列定義
// ==============================
//
// ※ contract_master の列だけをまとめておく
//
const CM_COLUMNS = {
  COMPANY_ID:   2,  // client_company_id
  COURSE_ID:    7,  // course_id
  PAYMENT_TYPE: 14, // payment_type
  // 他にもあればここに追加
};


// ==============================
// その他：料金テーブル関連
// ==============================

// fee_table_master の region デフォルト
const CM_DEFAULT_REGION = 'all';


// ==============================
// ナレッジDB・RAG用
// ==============================

// ナレッジDB本体
const SHEET_KNOWLEDGE_DB = 'ナレッジDB';

// ナレッジDBのヘッダー行
const KNOW_HEADER_ROW = 1;

// ナレッジDB：ステータス列（Q列）
const KNOW_COL_STATUS = 17;

// ナレッジDB：管理メモ列（R列）
const KNOW_COL_ADMIN_NOTE = 18;


// ==============================
// 契約マスター／ロジック履歴
// ==============================

// 元シート
const SHEET_CM          = 'contract_master';
const SHEET_CL          = 'contract_logic_rules';

// 履歴シート
const SHEET_CM_HISTORY  = 'contract_master_history';
const SHEET_CL_HISTORY  = 'contract_logic_history';


// ==============================
// 更新承認フロー
// ==============================

const SHEET_APPROVAL_FLOW = '更新承認フロー';

// 承認フローシートのヘッダー行
const AF_HEADER_ROW = 1;

// 承認フローシートの列定義
const AF_COLUMNS = {
  ID:        1, // 管理ID（KNOW_ID）
  TYPE:      2, // 種別（"knowledge" など）
  TARGET:    3, // 対象キー（今は KNOW_IDと同じ想定）
  STATUS:    5, // ステータス
  APPROVER:  6, // 承認者
  UPDATED_AT:7, // 最終更新日時
  NOTE:      8, // 備考
};
