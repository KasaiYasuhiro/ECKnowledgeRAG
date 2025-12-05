/**************************************************
 * 注意タグ（contract_warning_tags）関連
 * - code_master: item = "contract_warning_tags"
 * - contract_warning_flags: 会社ID＋コースIDごとのタグ紐づけ
 **************************************************/

// ---- シート構造用の定数 ----------------------------------

// contract_master シートは既存の HEADER_ROW / COL_COMPANY_ID / COL_COURSE_ID を利用

// contract_warning_flags シート
// 1行目: 和名, 2行目: 英名 という前提
const WARNING_FLAGS_HEADER_ROW      = HEADER_ROW;          // = 2
const WARNING_FLAGS_FIRST_DATA_ROW  = WARNING_FLAGS_HEADER_ROW + 1; // = 3
const WARNING_FLAGS_COL_COMPANY_ID  = 1; // A列 client_company_id
const WARNING_FLAGS_COL_COURSE_ID   = 2; // B列 course_id
const WARNING_FLAGS_COL_TAG         = 3; // C列 warning_tag
const WARNING_FLAGS_COL_REMARK      = 4; // D列 remarks

// code_master シート
// 1行目: ヘッダー, 2行目以降データ という前提
const CODE_MASTER_FIRST_DATA_ROW    = 2;
const CODE_MASTER_ITEM_COL          = 1; // A列: item
const CODE_MASTER_VALUE_COL         = 2; // B列: value
const CODE_MASTER_DESC_COL          = 3; // C列: desc

// 注意タグマスタ item 名
const WARNING_TAG_ITEM_KEY          = 'contract_warning_tags';


/**************************************
 * 注意タグ サイドバーを開く
 **************************************/
function openWarningTagSidebar() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  if (!sheet) {
    ui.alert('contract_master シートが見つかりません。');
    return;
  }

  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert('contract_master シートで、編集したいコースの行を選択してから実行してください。');
    return;
  }

  const row = range.getRow();
  // 1行目：和名, 2行目：英名 という前提
  if (row <= HEADER_ROW) {
    ui.alert('データ行（' + (HEADER_ROW + 1) + '行目以降）を選択してください。');
    return;
  }

  const companyId = sheet.getRange(row, COL_COMPANY_ID).getValue(); // client_company_id
  const courseId  = sheet.getRange(row, COL_COURSE_ID).getValue();  // course_id

  if (!companyId || !courseId) {
    ui.alert('選択した行に 会社ID または コースID がありません。');
    return;
  }

  // タグマスタ取得
  const tags = getAllWarningTagMaster_();
  if (tags.length === 0) {
    ui.alert('code_master に item="' + WARNING_TAG_ITEM_KEY + '" のデータがありません。');
    return;
  }

  // すでに付与されているタグを取得
  const selectedTags = getSelectedTagsForCourse_(companyId, courseId);

  // テンプレに埋め込み
  const template = HtmlService.createTemplateFromFile('WarningTagSidebar');
  template.companyId    = companyId;
  template.courseId     = courseId;
  template.tags         = tags;
  template.selectedTags = selectedTags;

  const html = template
    .evaluate()
    .setTitle('注意タグ設定')
    .setWidth(320);

  ui.showSidebar(html);
}


/**************************************
 * code_master から注意タグ一覧取得
 * item = "contract_warning_tags" を対象
 **************************************/
function getAllWarningTagMaster_() {
  const ss        = SpreadsheetApp.getActive();
  const codeSheet = ss.getSheetByName(SHEET_CODE_MASTER);
  if (!codeSheet) return [];

  const lastRow = codeSheet.getLastRow();
  if (lastRow < CODE_MASTER_FIRST_DATA_ROW) return [];

  const numRows = lastRow - CODE_MASTER_FIRST_DATA_ROW + 1;
  const data = codeSheet
    .getRange(CODE_MASTER_FIRST_DATA_ROW, CODE_MASTER_ITEM_COL, numRows, 3)
    .getValues(); // [ [item, value, desc], ... ]

  const tags = data
    .filter(r => String(r[CODE_MASTER_ITEM_COL - 1]) === WARNING_TAG_ITEM_KEY && r[CODE_MASTER_VALUE_COL - 1])
    .map(r => ({
      value: String(r[CODE_MASTER_VALUE_COL - 1]),
      desc:  String(r[CODE_MASTER_DESC_COL - 1] || '')
    }));

  return tags;
}


/**************************************
 * あるコースに既に設定されている注意タグ一覧を取得
 **************************************/
function getSelectedTagsForCourse_(companyId, courseId) {
  const ss        = SpreadsheetApp.getActive();
  const flagSheet = ss.getSheetByName(SHEET_WARNING_FLAGS);
  if (!flagSheet) return [];

  const lastRow = flagSheet.getLastRow();
  if (lastRow < WARNING_FLAGS_FIRST_DATA_ROW) return [];

  const numRows = lastRow - WARNING_FLAGS_FIRST_DATA_ROW + 1;
  const data = flagSheet
    .getRange(WARNING_FLAGS_FIRST_DATA_ROW, WARNING_FLAGS_COL_COMPANY_ID, numRows, 3)
    .getValues();
  // [ [client_company_id, course_id, warning_tag], ... ]

  const cid = String(companyId);
  const course = String(courseId);

  const tags = data
    .filter(r =>
      String(r[WARNING_FLAGS_COL_COMPANY_ID - 1]) === cid &&
      String(r[WARNING_FLAGS_COL_COURSE_ID  - 1]) === course &&
      r[WARNING_FLAGS_COL_TAG - 1]
    )
    .map(r => String(r[WARNING_FLAGS_COL_TAG - 1]));

  return tags;
}


/**************************************
 * 注意タグの選択結果を保存（HTML から呼ばれる）
 **************************************/
function saveCourseTags(companyId, courseId, selectedTags) {
  const ss        = SpreadsheetApp.getActive();
  const flagSheet = ss.getSheetByName(SHEET_WARNING_FLAGS);

  if (!flagSheet) {
    throw new Error('contract_warning_flags シートが見つかりません。');
  }

  const cid    = String(companyId);
  const course = String(courseId);

  // まず既存のタグ行を削除
  deleteWarningFlagRowsForCourse_(flagSheet, cid, course);

  // 新しいタグを登録
  if (Array.isArray(selectedTags) && selectedTags.length > 0) {
    const rowsToInsert = selectedTags.map(tag => [
      cid,
      course,
      String(tag),
      '' // remarks
    ]);
    const startRow = Math.max(flagSheet.getLastRow() + 1, WARNING_FLAGS_FIRST_DATA_ROW);
    flagSheet.getRange(startRow, WARNING_FLAGS_COL_COMPANY_ID, rowsToInsert.length, 4).setValues(rowsToInsert);
  }
}


/**************************************
 * 指定されたコースの注意タグ行を全削除
 **************************************/
function deleteWarningFlagRowsForCourse_(flagSheet, companyId, courseId) {
  const lastRow = flagSheet.getLastRow();
  if (lastRow < WARNING_FLAGS_FIRST_DATA_ROW) return;

  const numRows = lastRow - WARNING_FLAGS_FIRST_DATA_ROW + 1;
  const data = flagSheet
    .getRange(WARNING_FLAGS_FIRST_DATA_ROW, WARNING_FLAGS_COL_COMPANY_ID, numRows, 3)
    .getValues();

  // 後ろから削除しないと行番号がずれるので注意
  for (let i = data.length - 1; i >= 0; i--) {
    const rowIndex = WARNING_FLAGS_FIRST_DATA_ROW + i;
    const row      = data[i];

    const cidVal    = String(row[WARNING_FLAGS_COL_COMPANY_ID - 1]);
    const courseVal = String(row[WARNING_FLAGS_COL_COURSE_ID  - 1]);

    if (cidVal === companyId && courseVal === courseId) {
      flagSheet.deleteRow(rowIndex);
    }
  }
}
