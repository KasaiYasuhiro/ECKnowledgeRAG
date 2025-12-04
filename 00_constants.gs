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
const SHEET_LOGIC_RULES   = 'contract_logic_rules';  // 既存コードで使っている名前
const SHEET_CONTRACT_LOGIC = SHEET_LOGIC_RULES;      // 新しい名前としても使いたい場合の別名


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


// ==============================
// 既存コードとの互換用エイリアス
// ==============================
//
// これのおかげで、コード.gs 側はすぐには書き換えなくても動きます。
// 徐々に CM_ 系に置き換えていけばOK。
// ==============================

// onEdit / バリデーションなどで使っていた旧名称
const TARGET_SHEETS     = CM_TARGET_SHEETS;
const HEADER_ROW        = CM_HEADER_ROW;
const LAST_UPDATED_HEADER = CM_LAST_UPDATED_HEADER;
const COURSE_ID_HEADER  = CM_COURSE_ID_HEADER;
const LOG_SHEET_NAME    = CM_LOG_SHEET_NAME;

// fee_table 用の旧 DEFAULT_REGION 名称
const DEFAULT_REGION    = CM_DEFAULT_REGION;

// contract_master の列番号（旧名称）
const COL_COMPANY_ID   = CM_COLUMNS.COMPANY_ID;
const COL_COURSE_ID    = CM_COLUMNS.COURSE_ID;
const COL_PAYMENT_TYPE = CM_COLUMNS.PAYMENT_TYPE;

// ナレッジDB列の旧名称
const COL_KNOW_STATUS    = KNOW_COL_STATUS;
const COL_KNOW_ADMIN_NOTE = KNOW_COL_ADMIN_NOTE;

// 承認フローテーブルの旧名称
const COL_AF_ID         = AF_COLUMNS.ID;
const COL_AF_TYPE       = AF_COLUMNS.TYPE;
const COL_AF_TARGET     = AF_COLUMNS.TARGET;
const COL_AF_STATUS     = AF_COLUMNS.STATUS;
const COL_AF_APPROVER   = AF_COLUMNS.APPROVER;
const COL_AF_UPDATED_AT = AF_COLUMNS.UPDATED_AT;
const COL_AF_NOTE       = AF_COLUMNS.NOTE;
