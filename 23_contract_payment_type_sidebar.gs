// 23_contract_payment_type_sidebar.gs
// contract_master の「支払い区分」サイドバー関連

/**************************************
 * 支払い区分 サイドバーを開く
 **************************************/
function openPaymentTypeSidebar() {
  const ss    = SpreadsheetApp.getActive();
  const ui    = SpreadsheetApp.getUi();
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
  // 1行目:和名, 2行目:英名
  if (row <= CM_HEADER_ROW) {
    ui.alert('データ行（' + (CM_HEADER_ROW + 1) + '行目以降）を選択してください。');
    return;
  }

  const companyId = sheet.getRange(row, COL_COMPANY_ID).getValue();      // client_company_id
  const courseId  = sheet.getRange(row, COL_COURSE_ID).getValue();       // course_id
  const currentPaymentType = sheet.getRange(row, COL_PAYMENT_TYPE).getValue(); // payment_type

  if (!companyId || !courseId) {
    ui.alert('選択した行に 会社ID または コースID がありません。');
    return;
  }

  // 支払い区分マスタ取得（code_master 内の item = "payment_type"）
  const paymentTypes = getAllPaymentTypeMaster_();
  if (paymentTypes.length === 0) {
    ui.alert('code_master に item="payment_type" のデータがありません。');
    return;
  }

  // 既にセットされている支払い区分を配列に
  let selectedPaymentTypes = [];
  if (currentPaymentType) {
    selectedPaymentTypes = String(currentPaymentType)
      .split(';')
      .map(s => s.trim())
      .filter(s => s);
  }

  const template = HtmlService.createTemplateFromFile('PaymentTypeSidebar');
  template.rowIndex             = row;
  template.companyId            = companyId;
  template.courseId             = courseId;
  template.paymentTypes         = paymentTypes;
  template.selectedPaymentTypes = selectedPaymentTypes;

  const html = template
    .evaluate()
    .setTitle('支払い区分の設定')
    .setWidth(320);

  ui.showSidebar(html);
}
