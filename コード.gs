
/**
 * 変更を change_log に記録
 */
function logChange_(e, sheet, row) {
  const ss = sheet.getParent();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

  // ログシートが無ければ自動作成
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, 8).setValues([[
      'timestamp',
      'sheet_name',
      'row',
      'course_id',
      'column_a1',
      'field_name',
      'old_value',
      'new_value'
    ]]);
  }

  const range = e.range;
  const sheetName = sheet.getName();

  // course_id の列位置を取得して値を取り出す
  const courseIdCol = getOrCreateHeaderColumn_(sheet, COURSE_ID_HEADER);
  const courseId = sheet.getRange(row, courseIdCol).getValue();

  // 編集セルの英名（2行目）
  const fieldName = sheet.getRange(HEADER_ROW, range.getColumn()).getValue();

  // old/new 値を取得
  const oldValue = (typeof e.oldValue !== 'undefined') ? e.oldValue : '';
  const newValue = (() => {
    const vals = range.getValues();
    if (vals.length === 1 && vals[0].length === 1) {
      return vals[0][0];
    }
    return '(multiple cells)';
  })();

  // ログ行を作成
  const logRow = [
    new Date(),
    sheetName,
    row,
    courseId,
    range.getA1Notation(),
    fieldName,
    oldValue,
    newValue
  ];

  logSheet.appendRow(logRow);
  logSheet.getRange(logSheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}


/**************************************************
 * ここから 支払い区分（payment_type）関連
 **************************************************/

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
  if (row <= 2) {
    ui.alert('データ行（3行目以降）を選択してください。');
    return;
  }

  const companyId = sheet.getRange(row, COL_COMPANY_ID).getValue();   // client_company_id
  const courseId  = sheet.getRange(row, COL_COURSE_ID).getValue();    // course_id
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
      .split(";")
      .map(s => s.trim())
      .filter(s => s);
  }

  const template = HtmlService.createTemplateFromFile('PaymentTypeSidebar');
  template.rowIndex            = row;
  template.companyId           = companyId;
  template.courseId            = courseId;
  template.paymentTypes        = paymentTypes;
  template.selectedPaymentTypes = selectedPaymentTypes;

  const html = template
    .evaluate()
    .setTitle('支払い区分の設定')
    .setWidth(320);

  ui.showSidebar(html);
}

/**************************************
 * code_master から payment_type 一覧取得
 * item = "payment_type" を対象
 **************************************/
function getAllPaymentTypeMaster_() {
  const ss        = SpreadsheetApp.getActive();
  const codeSheet = ss.getSheetByName(SHEET_CODE_MASTER);
  if (!codeSheet) return [];

  const lastRow = codeSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = codeSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  // [ [item, value, desc], ... ]
  const types = data
    .filter(r => String(r[0]) === 'payment_type' && r[1])
    .map(r => ({
      value: String(r[1]),
      desc:  String(r[2] || '')
    }));
  return types;
}

/**************************************
 * 支払い区分の選択結果を保存（HTML から呼ばれる）
 **************************************/
function savePaymentTypes(rowIndex, selectedTypes) {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  if (!sheet) {
    throw new Error('contract_master シートが見つかりません。');
  }

  let joined = '';
  if (Array.isArray(selectedTypes) && selectedTypes.length > 0) {
    joined = selectedTypes
      .map(s => String(s).trim())
      .filter(s => s)
      .join(';');
  }

  sheet.getRange(rowIndex, COL_PAYMENT_TYPE).setValue(joined);
}

/**************************************************
 * contract_logic_rules の exit_fee_condition_detail を
 * exit_fee_calc_method に応じてテンプレ自動入力するスクリプト
 *
 * 対象シート:
 *   - シート名: contract_logic_rules
 *   - 1行目: 和名ヘッダ
 *   - 2行目: 英名ヘッダ
 *   - 3行目以降: データ
 *
 * 参照カラム（列番号は 1 始まり）:
 *   A: last_updated
 *   B: client_company_id
 *   C: course_id
 *   K: exit_fee_amount
 *   L: exit_fee_calc_method
 *   N: exit_fee_condition_detail
 **************************************************/

/**
 * exit_fee_calc_method に応じて
 * exit_fee_condition_detail にテンプレを自動入力する
 */
function fillExitFeeConditionTemplates() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('contract_logic_rules');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('contract_logic_rules にデータ行がありません（3行目以降）。');
    return;
  }

  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  // 列番号（1始まり）をわかりやすく定義
  const COL_COURSE_ID      = 3;  // C: course_id
  const COL_EXIT_FEE_AMT   = 11; // K: exit_fee_amount
  const COL_EXIT_FEE_METHOD= 12; // L: exit_fee_calc_method
  const COL_EXIT_FEE_COND  = 14; // N: exit_fee_condition_detail

  let updateCount = 0;

  values.forEach((row, idx) => {
    const courseId   = row[COL_COURSE_ID - 1];
    const exitFeeAmt = row[COL_EXIT_FEE_AMT - 1];
    const method     = row[COL_EXIT_FEE_METHOD - 1];
    const currentCond= row[COL_EXIT_FEE_COND - 1];

    // course_id 空行はスキップ
    if (!courseId) return;

    // すでに exit_fee_condition_detail が入力されている場合は上書きしない
    if (currentCond && String(currentCond).trim() !== '') return;

    // exit_fee_calc_method に応じてテンプレ選択
    let templateText = '';

    if (method === 'tiered') {
      // 🔹 テンプレ③：段階制（tiered）・7回お約束など
      templateText =
        '本コースの解約金は段階制（tiered）であり、受取回数に応じて金額が変動します。\n' +
        '初回・2回目・3回目以降で金額が大きく異なるため、必ず fee_table_master を参照してください。\n\n' +
        '【計算方法】\n' +
        '・受取回数（order_count）ごとに差額金（diff_amount）を設定しています。\n' +
        '・支払方法（payment_type）によって、金額が加算・変更される場合があります。\n' +
        '・地域（北海道・沖縄など）で送料が加算される場合があります。\n\n' +
        '【参照場所】\n' +
        'fee_table_master\n' +
        '(client_company_id × course_id × payment_type × order_count × region)\n\n' +
        '※contract_logic_rules には計算原則のみを記載し、金額は fee_table_master に一元管理します。';

    } else if (method === 'fixed') {
      // 🔹 テンプレ②：固定額中心（exit_fee_amount を使う前提）
      const amtStr = (exitFeeAmt !== '' && exitFeeAmt != null)
        ? String(exitFeeAmt)
        : '（別途 fee_table_master を参照）';

      templateText =
        '本コースの解約金は原則として固定額で運用します。\n' +
        '通常は ' + amtStr + ' 円を基準としますが、実際の請求金額は fee_table_master 上の diff_amount と整合させて管理します。\n\n' +
        '【初回解約】\n' +
        '・初回のみ受取で解約する場合の解約金は固定額です。\n' +
        '・金額は exit_fee_amount の値、または fee_table_master の order_count=1 を参照します。\n\n' +
        '【2回目以降】\n' +
        '・2回目以降に解約する場合の解約金は、必要に応じて fee_table_master の\n' +
        '  (client_company_id × course_id × payment_type × order_count × region)\n' +
        '  に基づき算出します。\n\n' +
        '※最新の金額は fee_table_master を正とし、exit_fee_amount は代表値として扱います。';
    } else {
      // percentage / none / 空欄などは自動入力しない（手入力想定）
      return;
    }

    // 選ばれたテンプレを row 配列にセット
    row[COL_EXIT_FEE_COND - 1] = templateText;
    updateCount++;
  });

  // 変更があった行だけまとめて書き戻す
  if (updateCount > 0) {
    sheet.getRange(3, 1, values.length, lastCol).setValues(values);
    ui.alert('解約金条件テンプレを ' + updateCount + ' 行に反映しました。');
  } else {
    ui.alert('自動反映対象の行がありませんでした。\n' +
             '（course_id 空行、または exit_fee_condition_detail が既に入力済みの行のみでした。）');
  }
}




/**************************************************
 * contract_master バリデーション関連
 **************************************************/

// シート・列情報（既存と揃える）
//const COL_COMPANY_ID   = 2;  // client_company_id
//const COL_COURSE_ID    = 7;  // course_id
//const COL_PAYMENT_TYPE = 14; // payment_type

/**
 * code_master から item ごとの value の一覧を取得する
 * 例: getMasterValues_('category') → ['subscription','single',...]
 */
function getMasterValues_(itemName) {
  const ss        = SpreadsheetApp.getActive();
  const codeSheet = ss.getSheetByName(SHEET_CODE_MASTER);
  if (!codeSheet) return [];

  const lastRow = codeSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = codeSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data
    .filter(r => String(r[0]) === itemName && r[1])
    .map(r => String(r[1]));
}

/**
 * セルに入っている ; 区切りの値が、すべて許可された value かチェック
 */
function validateMultiValueField_(label, rawValue, allowedValues, errors) {
  if (!rawValue) return;
  const allowedSet = new Set(allowedValues);
  const parts = String(rawValue)
    .split(';')
    .map(s => s.trim())
    .filter(s => s);

  parts.forEach(p => {
    if (!allowedSet.has(p)) {
      errors.push(`${label} に不正な値があります: "${p}"（code_master に未登録）`);
    }
  });
}

/**
 * 単一値が許可された value かチェック
 */
function validateSingleValueField_(label, rawValue, allowedValues, errors) {
  if (!rawValue) return;
  const allowedSet = new Set(allowedValues);
  const v = String(rawValue).trim();
  if (v && !allowedSet.has(v)) {
    errors.push(`${label} に不正な値があります: "${v}"（code_master に未登録）`);
  }
}

/**
 * 数値フィールドのチェック（空欄は許容）
 */
function validateNumberField_(label, rawValue, errors) {
  if (rawValue === "" || rawValue === null) return;
  const n = Number(rawValue);
  if (isNaN(n)) {
    errors.push(`${label} は数値で入力してください（現在の値: "${rawValue}"）`);
  }
}

/**
 * 選択行の contract_master レコードをチェック
 */
function validateSelectedContractRow() {
  const ss    = SpreadsheetApp.getActive();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  if (!sheet) {
    ui.alert('contract_master シートが見つかりません。');
    return;
  }

  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert('contract_master シートでチェックしたい行を選択してください。');
    return;
  }

  const row = range.getRow();
  if (row <= 2) {
    ui.alert('3行目以降のデータ行を選択してください。');
    return;
  }

  // 行データ取得（1行分）
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const errors = [];

  // 必須チェック
  const clientCompanyId   = values[1];  // B: client_company_id
  const clientCompanyName = values[2];  // C: client_company_name
  const courseName        = values[3];  // D: course_name
  const category          = values[4];  // E: category
  const courseId          = values[6];  // G: course_id
  const contractType      = values[7];  // H: contract_type
  const guaranteeType     = values[10]; // K: guarantee_type
  const paymentTypeRaw    = values[13]; // N: payment_type

  if (!clientCompanyId)   errors.push('会社ID が未入力です。');
  if (!clientCompanyName) errors.push('会社名 が未入力です。');
  if (!courseName)        errors.push('コース名 が未入力です。');
  if (!category)          errors.push('カテゴリ（category）が未入力です。');
  if (!courseId)          errors.push('コースID（course_id）が未入力です。');
  if (!contractType)      errors.push('契約種別（contract_type）が未入力です。');
  if (!guaranteeType)     errors.push('保証種別（guarantee_type）が未入力です。');
  // ★ 支払い区分は「任意」にするので、ここは削除 ★
  // if (!paymentTypeRaw)    errors.push('支払い区分（payment_type）が未入力です。');

  // マスタ存在チェック
  const mCategory     = getMasterValues_('category');
  const mContractType = getMasterValues_('contract_type');
  const mCommitRule   = getMasterValues_('commit_rule');
  const mExitFeeRule  = getMasterValues_('exit_fee_rule');
  const mGuarantee    = getMasterValues_('guarantee_type');
  const mSalesCh      = getMasterValues_('sales_channels');
  const mFulfill      = getMasterValues_('fulfillment_rule');
  const mPayType      = getMasterValues_('payment_type');
  const mPayCat       = getMasterValues_('payment_method_category');
  const mInstall      = getMasterValues_('installment_available');
  const mBillingInt   = getMasterValues_('billing_interval');
  const mBoolFlag     = ['TRUE','FALSE']; // 一部フラグ列用簡易チェック（基準は大文字）

  // 単一値チェック
  validateSingleValueField_('カテゴリ（category）', category, mCategory, errors);
  validateSingleValueField_('契約種別（contract_type）', contractType, mContractType, errors);
  validateSingleValueField_('回数縛り条件（commit_rule）', values[8], mCommitRule, errors);
  validateSingleValueField_('解約金条件（exit_fee_rule）', values[9], mExitFeeRule, errors);
  validateSingleValueField_('保証種別（guarantee_type）', guaranteeType, mGuarantee, errors);
  validateSingleValueField_('定期サイクル（fulfillment_rule）', values[12], mFulfill, errors);
  validateSingleValueField_('課金間隔（billing_interval）', values[20], mBillingInt, errors);

  // 複数値チェック（;区切り）
  validateMultiValueField_('申込可能チャネル（sales_channels）', values[11], mSalesCh, errors);
  validateMultiValueField_('支払い区分（payment_type）', paymentTypeRaw, mPayType, errors);
  validateMultiValueField_('支払い方法カテゴリ（payment_method_category）', values[14], mPayCat, errors);

  // TRUE/FALSE 系（installment_available / initial_gift_flag / 各種フラグ）
  const installAvailable = values[15]; // P: installment_available
  const initialGift      = values[23]; // X: initial_gift_flag
  const isUpsellTarget   = values[25]; // Z: is_upsell_target
  const hasTrigger       = values[26]; // AA: has_trigger_keyword
  const coupon50         = values[27]; // AB: use_coupon_50_off
  const coupon30         = values[28]; // AC: use_coupon_30_off
  const hasPointCancel   = values[29]; // AD: has_point_cancel_request

  if (installAvailable && mInstall.length > 0) {
    validateSingleValueField_('分割払い可否（installment_available）', installAvailable, mInstall, errors);
  } else if (installAvailable) {
    validateSingleValueField_('分割払い可否（installment_available）', installAvailable, mBoolFlag, errors);
  }

  const boolLabelsAndValues = [
    ['初回特典プレゼント同梱フラグ（initial_gift_flag）', initialGift],
    ['アップセル対象フラグ（is_upsell_target）', isUpsellTarget],
    ['トリガーワード有無（has_trigger_keyword）', hasTrigger],
    ['50％OFFクーポン提案有無（use_coupon_50_off）', coupon50],
    ['30％OFFクーポン提案有無（use_coupon_30_off）', coupon30],
    ['ポイント解約提案有無（has_point_cancel_request）', hasPointCancel]
  ];

  boolLabelsAndValues.forEach(([label, v]) => {
    if (v !== "" && v !== null) {
      const vv = String(v).trim();
      const vvUpper = vv.toUpperCase(); // 大文字化して比較
      if (vv && mBoolFlag.indexOf(vvUpper) === -1) {
        errors.push(`${label} は TRUE か FALSE で入力してください（現在の値: "${vv}"）`);
      }
    }
  });

  // 数値チェック（価格・縛り回数）
  validateNumberField_('初回/単品価格（first_price）', values[16], errors);
  validateNumberField_('2回目特別価格（second_price）', values[17], errors);
  validateNumberField_('定期の通常価格（recurring_price）', values[18], errors);
  validateNumberField_('初回縛り回数（first_commit_count）', values[21], errors);
  validateNumberField_('累計縛り回数（total_commit_count）', values[22], errors);

  // 結果表示
  if (errors.length === 0) {
    ui.alert(`行 ${row}：問題は見つかりませんでした ✅`);
  } else {
    ui.alert(
      `行 ${row} のチェック結果`,
      errors.join('\n・ '),
      ui.ButtonSet.OK
    );
  }
}

/**
 * contract_master から course_id 一覧を取得（Setで返す）
 */
function getAllCourseIdSet_() {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  if (!sheet) return new Set();

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return new Set();

  // G列 = course_id（3行目以降）
  const range = sheet.getRange(3, 7, lastRow - 2, 1);
  const values = range.getValues();
  const set = new Set();
  values.forEach(r => {
    const v = r[0];
    if (v !== "" && v !== null) {
      set.add(String(v));
    }
  });
  return set;
}

/**************************************************
 * contract_logic_rules バリデーション
 **************************************************/

function validateSelectedLogicRow() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('contract_logic_rules');

  if (!sheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== 'contract_logic_rules') {
    ui.alert('contract_logic_rules シートで、チェックしたい行を選択してから実行してください。');
    return;
  }

  const range = activeSheet.getActiveRange();
  if (!range) {
    ui.alert('contract_logic_rules シートで、チェックしたい行を選択してください。');
    return;
  }

  const row = range.getRow();
  if (row <= 2) {
    ui.alert('3行目以降のデータ行を選択してください。');
    return;
  }

  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const errors = [];

  // 列マッピング（0始まり index）
    const col = {
    last_updated: 0,
    client_company_id: 1,
    course_id: 2,
    cancel_deadline: 3,
    cancel_deadline_logic: 4,
    holiday_handling: 5,
    oos_cancel_deadline_rule: 6,
    skip_rule: 7,
    long_pause_rule: 8,
    cancel_deadline_rule_in_long_holiday: 9,
    exit_fee_amount: 10,
    exit_fee_calc_method: 11,
    exit_fee_detail: 12,
    exit_fee_condition_detail: 13,
    exit_fee_waiver_condition: 14,
    exit_fee_notice_template: 15,
    upsell_exit_fee_flag: 16,
    upsell_exit_logic_detail: 17,
    refund_guarantee_flag: 18,
    refund_guarantee_term: 19,
    refund_guarantee_condition_detail: 20,
    guarantee_return_required: 21,
    guarantee_return_deadline: 22,
    exception_return_rule: 23,
    exception_return_deadline: 24,
    cooling_off_flag: 25,
    cooling_off_term_type: 26,
    cooling_off_term: 27,
    cooling_off_condition_detail: 28,
    first_order_cancelable_before_ship: 29,
    first_order_cancel_condition: 30,
    recurring_order_cancelable_before_ship: 31,
    recurring_order_cancel_condition: 32,
    cancel_after_ship: 33,
    cancel_explanation_template: 34,
    customer_misunderstanding_points: 35,
    cancel_deadline_rule_when_oos: 36
  };


  const courseId = values[col.course_id];

  // 必須チェック
  if (!courseId) {
    errors.push('コースID（course_id）が未入力です。');
  }

  // course_id が contract_master に存在するか
  const courseSet = getAllCourseIdSet_();
  if (courseId && !courseSet.has(String(courseId))) {
    errors.push(`コースID "${courseId}" は contract_master に存在しません。タイポか未登録の可能性があります。`);
  }

  // マスタ値（code_master）一覧取得
  const mHolidayHandling = getMasterValues_('holiday_handling');
  const mExitFeeCalc     = getMasterValues_('exit_fee_calc_method');
  const mUpsellExitFlag  = getMasterValues_('upsell_exit_fee_flag');
  const mRefundFlag      = getMasterValues_('refund_guarantee_flag');
  const mGuaranteeReturn = getMasterValues_('guarantee_return_required');
  const mCoFlag          = getMasterValues_('cooling_off_flag');
  const mCoTermType      = getMasterValues_('cooling_off_term_type');
  const mFirstCancelFlag = getMasterValues_('first_order_cancelable_before_ship');
  const mRecurCancelFlag = getMasterValues_('recurring_order_cancelable_before_ship');
  const mCancelAfterShip = getMasterValues_('cancel_after_ship');

  const mBoolFlag = ['TRUE','FALSE'];

  // 単一値チェック
  validateSingleValueField_(
    '解約期限の土日祝の扱い（holiday_handling）',
    values[col.holiday_handling],
    mHolidayHandling,
    errors
  );

  validateSingleValueField_(
    '解約金計算方法（exit_fee_calc_method）',
    values[col.exit_fee_calc_method],
    mExitFeeCalc,
    errors
  );

  // TRUE/FALSE 系 & フラグ系（code_master があればそちら優先）
  function checkFlagWithMasterOrBool(label, value, masterList) {
    if (value === "" || value === null) return;
    const v = String(value).trim();
    const vUpper = v.toUpperCase();

    if (masterList && masterList.length > 0) {
      validateSingleValueField_(label, v, masterList, errors);
    } else {
      if (mBoolFlag.indexOf(vUpper) === -1) {
        errors.push(`${label} は TRUE か FALSE で入力してください（現在の値: "${v}"）`);
      }
    }
  }

  checkFlagWithMasterOrBool(
    'アップセル解約金有無（upsell_exit_fee_flag）',
    values[col.upsell_exit_fee_flag],
    mUpsellExitFlag
  );

  checkFlagWithMasterOrBool(
    '返金保証フラグ（refund_guarantee_flag）',
    values[col.refund_guarantee_flag],
    mRefundFlag
  );

  checkFlagWithMasterOrBool(
    '返金保証で返品が必要か（guarantee_return_required）',
    values[col.guarantee_return_required],
    mGuaranteeReturn
  );

  checkFlagWithMasterOrBool(
    'クーリングオフフラグ（cooling_off_flag）',
    values[col.cooling_off_flag],
    mCoFlag
  );

  checkFlagWithMasterOrBool(
    '初回発送前キャンセル可否（first_order_cancelable_before_ship）',
    values[col.first_order_cancelable_before_ship],
    mFirstCancelFlag
  );

  checkFlagWithMasterOrBool(
    '継続分の発送前キャンセル可否（recurring_order_cancelable_before_ship）',
    values[col.recurring_order_cancelable_before_ship],
    mRecurCancelFlag
  );

  validateSingleValueField_(
    'クーリングオフ期間区分（cooling_off_term_type）',
    values[col.cooling_off_term_type],
    mCoTermType,
    errors
  );

  validateSingleValueField_(
    '発送後キャンセル可否（cancel_after_ship）',
    values[col.cancel_after_ship],
    mCancelAfterShip,
    errors
  );

  // 数値チェック（exit_fee_amount のみ）
  validateNumberField_(
    '解約金額（exit_fee_amount）',
    values[col.exit_fee_amount],
    errors
  );

  // 結果表示
  if (errors.length === 0) {
    ui.alert(`行 ${row}：問題は見つかりませんでした ✅`);
  } else {
    ui.alert(
      `contract_logic_rules 行 ${row} のチェック結果`,
      errors.join('\n・ '),
      ui.ButtonSet.OK
    );
  }
}

/**************************************************
 * contract_master 全行バリデーション → レポート出力
 **************************************************/

function validateAllContractRows() {
  const ss    = SpreadsheetApp.getActive();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  if (!sheet) {
    ui.alert('contract_master シートが見つかりません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3) {
    ui.alert('contract_master にデータ行がありません（3行目以降）。');
    return;
  }

  // レポートシートを用意
  const REPORT_SHEET_NAME = 'master_validation_report';
  let reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(REPORT_SHEET_NAME);
  } else {
    reportSheet.clear();
  }

  reportSheet.getRange(1, 1, 1, 5).setValues([[
    'sheet',
    'row',
    'client_company_id',
    'course_id',
    'message'
  ]]);

  const courseIdColIndex = 7; // G列 = course_id

  // code_master 側のマスタを事前に一回だけ読む
  const mCategory     = getMasterValues_('category');
  const mContractType = getMasterValues_('contract_type');
  const mCommitRule   = getMasterValues_('commit_rule');
  const mExitFeeRule  = getMasterValues_('exit_fee_rule');
  const mGuarantee    = getMasterValues_('guarantee_type');
  const mSalesCh      = getMasterValues_('sales_channels');
  const mFulfill      = getMasterValues_('fulfillment_rule');
  const mPayType      = getMasterValues_('payment_type');
  const mPayCat       = getMasterValues_('payment_method_category');
  const mInstall      = getMasterValues_('installment_available');
  const mBillingInt   = getMasterValues_('billing_interval');
  const mBoolFlag     = ['TRUE', 'FALSE'];

  const allValues = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const reportRows = [];

  allValues.forEach((rowValues, idx) => {
    const rowNum = idx + 3;
    const errors = [];

    const clientCompanyId   = rowValues[1];  // B
    const clientCompanyName = rowValues[2];  // C
    const courseName        = rowValues[3];  // D
    const category          = rowValues[4];  // E
    const courseId          = rowValues[6];  // G

  // ★ ここで「事実上未入力の行」はスキップする ★
  const isEffectivelyEmpty =
    (clientCompanyId === "" || clientCompanyId === null) &&
    (clientCompanyName === "" || clientCompanyName === null) &&
    (courseName === "" || courseName === null) &&
    (category === "" || category === null) &&
    (courseId === "" || courseId === null);

  if (isEffectivelyEmpty) {
    return; // この行はチェック対象外
  }


    const contractType      = rowValues[7];  // H
    const guaranteeType     = rowValues[10]; // K
    const paymentTypeRaw    = rowValues[13]; // N

    // 必須
    if (!clientCompanyId)   errors.push('会社ID が未入力です。');
    if (!clientCompanyName) errors.push('会社名 が未入力です。');
    if (!courseName)        errors.push('コース名 が未入力です。');
    if (!category)          errors.push('カテゴリ（category）が未入力です。');
    if (!courseId)          errors.push('コースID（course_id）が未入力です。');
    if (!contractType)      errors.push('契約種別（contract_type）が未入力です。');
    if (!guaranteeType)     errors.push('保証種別（guarantee_type）が未入力です。');
    // ★ 支払い区分は「任意」にするので、ここは削除 ★
    // if (!paymentTypeRaw)    errors.push('支払い区分（payment_type）が未入力です。');

    // 単一値
    validateSingleValueField_('カテゴリ（category）', category, mCategory, errors);
    validateSingleValueField_('契約種別（contract_type）', contractType, mContractType, errors);
    validateSingleValueField_('回数縛り条件（commit_rule）', rowValues[8], mCommitRule, errors);
    validateSingleValueField_('解約金条件（exit_fee_rule）', rowValues[9], mExitFeeRule, errors);
    validateSingleValueField_('保証種別（guarantee_type）', guaranteeType, mGuarantee, errors);
    validateSingleValueField_('定期サイクル（fulfillment_rule）', rowValues[12], mFulfill, errors);
    validateSingleValueField_('課金間隔（billing_interval）', rowValues[20], mBillingInt, errors);

    // 複数値
    validateMultiValueField_('申込可能チャネル（sales_channels）', rowValues[11], mSalesCh, errors);
    validateMultiValueField_('支払い区分（payment_type）', paymentTypeRaw, mPayType, errors);
    validateMultiValueField_('支払い方法カテゴリ（payment_method_category）', rowValues[14], mPayCat, errors);

    // TRUE/FALSE 系
    const installAvailable = rowValues[15]; // P
    const initialGift      = rowValues[23]; // X
    const isUpsellTarget   = rowValues[25]; // Z
    const hasTrigger       = rowValues[26]; // AA
    const coupon50         = rowValues[27]; // AB
    const coupon30         = rowValues[28]; // AC
    const hasPointCancel   = rowValues[29]; // AD

    if (installAvailable && mInstall.length > 0) {
      validateSingleValueField_('分割払い可否（installment_available）', installAvailable, mInstall, errors);
    } else if (installAvailable) {
      const vvUpper = String(installAvailable).trim().toUpperCase();
      if (mBoolFlag.indexOf(vvUpper) === -1) {
        errors.push(`分割払い可否（installment_available）は TRUE か FALSE で入力してください（現在の値: "${installAvailable}"）`);
      }
    }

    [
      ['初回特典プレゼント同梱フラグ（initial_gift_flag）', initialGift],
      ['アップセル対象フラグ（is_upsell_target）', isUpsellTarget],
      ['トリガーワード有無（has_trigger_keyword）', hasTrigger],
      ['50％OFFクーポン提案有無（use_coupon_50_off）', coupon50],
      ['30％OFFクーポン提案有無（use_coupon_30_off）', coupon30],
      ['ポイント解約提案有無（has_point_cancel_request）', hasPointCancel]
    ].forEach(([label, v]) => {
      if (v !== "" && v !== null) {
        const vv = String(v).trim();
        const vvUpper = vv.toUpperCase();
        if (vv && mBoolFlag.indexOf(vvUpper) === -1) {
          errors.push(`${label} は TRUE か FALSE で入力してください（現在の値: "${vv}"）`);
        }
      }
    });

    // 数値
    validateNumberField_('初回/単品価格（first_price）', rowValues[16], errors);
    validateNumberField_('2回目特別価格（second_price）', rowValues[17], errors);
    validateNumberField_('定期の通常価格（recurring_price）', rowValues[18], errors);
    validateNumberField_('初回縛り回数（first_commit_count）', rowValues[21], errors);
    validateNumberField_('累計縛り回数（total_commit_count）', rowValues[22], errors);

    if (errors.length > 0) {
      const cid = clientCompanyId || '';
      const coid = courseId || '';
      errors.forEach(msg => {
        reportRows.push([
          SHEET_CONTRACT_MASTER,
          rowNum,
          cid,
          coid,
          msg
        ]);
      });
    }
  });

  if (reportRows.length === 0) {
    reportSheet.getRange(2, 1, 1, 5).setValues([[
      SHEET_CONTRACT_MASTER,
      '',
      '',
      '',
      'エラーは見つかりませんでした ✅'
    ]]);
    ui.alert('全ての行をチェックしました。エラーはありませんでした。');
  } else {
    reportSheet.getRange(2, 1, reportRows.length, 5).setValues(reportRows);
    ui.alert(
      `contract_master 全行のチェックが完了しました。\n` +
      `エラー件数: ${reportRows.length}\n` +
      `詳細はシート "${REPORT_SHEET_NAME}" をご確認ください。`
    );
  }
}

/**************************************************
 * contract_logic_rules 全行バリデーション → レポート出力
 **************************************************/

function validateAllLogicRows() {
  const ss    = SpreadsheetApp.getActive();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('contract_logic_rules');
  if (!sheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3) {
    ui.alert('contract_logic_rules にデータ行がありません（3行目以降）。');
    return;
  }

  // レポートシート準備
  const REPORT_SHEET_NAME = 'logic_validation_report';
  let reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(REPORT_SHEET_NAME);
  } else {
    reportSheet.clear();
  }

  reportSheet.getRange(1, 1, 1, 5).setValues([[
    'sheet',
    'row',
    'client_company_id',
    'course_id',
    'message'
  ]]);

  // 列マッピング（0始まり index）
  const col = {
    last_updated: 0,
    client_company_id: 1,
    course_id: 2,
    cancel_deadline: 3,
    cancel_deadline_logic: 4,
    holiday_handling: 5,
    oos_cancel_deadline_rule: 6,
    skip_rule: 7,
    long_pause_rule: 8,
    cancel_deadline_rule_in_long_holiday: 9,
    exit_fee_amount: 10,
    exit_fee_calc_method: 11,
    exit_fee_detail: 12,
    exit_fee_condition_detail: 13,
    exit_fee_waiver_condition: 14,
    exit_fee_notice_template: 15,
    upsell_exit_fee_flag: 16,
    upsell_exit_logic_detail: 17,
    refund_guarantee_flag: 18,
    refund_guarantee_term: 19,
    refund_guarantee_condition_detail: 20,
    guarantee_return_required: 21,
    guarantee_return_deadline: 22,
    exception_return_rule: 23,
    exception_return_deadline: 24,
    cooling_off_flag: 25,
    cooling_off_term_type: 26,
    cooling_off_term: 27,
    cooling_off_condition_detail: 28,
    first_order_cancelable_before_ship: 29,
    first_order_cancel_condition: 30,
    recurring_order_cancelable_before_ship: 31,
    recurring_order_cancel_condition: 32,
    cancel_after_ship: 33,
    cancel_explanation_template: 34,
    customer_misunderstanding_points: 35,
    cancel_deadline_rule_when_oos: 36
  };

  // master から course_id 一覧
  const courseSet = getAllCourseIdSet_();

  // code_master からマスタ値
  const mHolidayHandling = getMasterValues_('holiday_handling');
  const mExitFeeCalc     = getMasterValues_('exit_fee_calc_method');
  const mUpsellExitFlag  = getMasterValues_('upsell_exit_fee_flag');
  const mRefundFlag      = getMasterValues_('refund_guarantee_flag');
  const mGuaranteeReturn = getMasterValues_('guarantee_return_required');
  const mCoFlag          = getMasterValues_('cooling_off_flag');
  const mCoTermType      = getMasterValues_('cooling_off_term_type');
  const mFirstCancelFlag = getMasterValues_('first_order_cancelable_before_ship');
  const mRecurCancelFlag = getMasterValues_('recurring_order_cancelable_before_ship');
  const mCancelAfterShip = getMasterValues_('cancel_after_ship');

  const mBoolFlag = ['TRUE', 'FALSE'];

  const allValues = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const reportRows = [];

  // TRUE/FALSE or master付きフラグをチェックするヘルパー
  function checkFlagWithMasterOrBool(label, value, masterList, errors) {
    if (value === "" || value === null) return;
    const v = String(value).trim();
    const vUpper = v.toUpperCase();

    if (masterList && masterList.length > 0) {
      validateSingleValueField_(label, v, masterList, errors);
    } else {
      if (mBoolFlag.indexOf(vUpper) === -1) {
        errors.push(`${label} は TRUE か FALSE で入力してください（現在の値: "${v}"）`);
      }
    }
  }

  allValues.forEach((rowValues, idx) => {
    const rowNum = idx + 3;
    const errors = [];

    const clientCompanyId = rowValues[col.client_company_id];
    const courseId        = rowValues[col.course_id];
    const cancelDeadline  = rowValues[col.cancel_deadline];
    const cancelLogic     = rowValues[col.cancel_deadline_logic];

    // ★ 完全に「まだ使ってない行」はスキップ（company_id & course_id & deadline系が空）
    const isEffectivelyEmpty =
      (clientCompanyId === "" || clientCompanyId === null) &&
      (courseId === "" || courseId === null) &&
      (cancelDeadline === "" || cancelDeadline === null) &&
      (cancelLogic === "" || cancelLogic === null);

    if (isEffectivelyEmpty) {
      return;
    }

    // ===== 必須 & 整合性チェック =====
    if (!clientCompanyId) {
      errors.push('会社ID（client_company_id）が未入力です。');
    }
    if (!courseId) {
      errors.push('コースID（course_id）が未入力です。');
    } else if (!courseSet.has(String(courseId))) {
      errors.push(`コースID "${courseId}" は contract_master に存在しません。タイポか未登録の可能性があります。`);
    }

    // ===== マスタ値チェック =====
    validateSingleValueField_(
      '解約期限の土日祝の扱い（holiday_handling）',
      rowValues[col.holiday_handling],
      mHolidayHandling,
      errors
    );

    validateSingleValueField_(
      '解約金計算方法（exit_fee_calc_method）',
      rowValues[col.exit_fee_calc_method],
      mExitFeeCalc,
      errors
    );

    validateSingleValueField_(
      'クーリングオフ期間区分（cooling_off_term_type）',
      rowValues[col.cooling_off_term_type],
      mCoTermType,
      errors
    );

    validateSingleValueField_(
      '発送後キャンセル可否（cancel_after_ship）',
      rowValues[col.cancel_after_ship],
      mCancelAfterShip,
      errors
    );

    // ===== フラグ系 =====
    checkFlagWithMasterOrBool(
      'アップセル解約金有無（upsell_exit_fee_flag）',
      rowValues[col.upsell_exit_fee_flag],
      mUpsellExitFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '返金保証の有無（refund_guarantee_flag）',
      rowValues[col.refund_guarantee_flag],
      mRefundFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '返金保証利用時の返品要否（guarantee_return_required）',
      rowValues[col.guarantee_return_required],
      mGuaranteeReturn,
      errors
    );

    checkFlagWithMasterOrBool(
      'クーリングオフ可否（cooling_off_flag）',
      rowValues[col.cooling_off_flag],
      mCoFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '初回発送前キャンセル可否（first_order_cancelable_before_ship）',
      rowValues[col.first_order_cancelable_before_ship],
      mFirstCancelFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '継続分の発送前キャンセル可否（recurring_order_cancelable_before_ship）',
      rowValues[col.recurring_order_cancelable_before_ship],
      mRecurCancelFlag,
      errors
    );

    // ===== 数値チェック =====
    validateNumberField_(
      '解約金額（exit_fee_amount）',
      rowValues[col.exit_fee_amount],
      errors
    );

    // ===== エラーをレポート行として保存 =====
    if (errors.length > 0) {
      const cid  = clientCompanyId || '';
      const coid = courseId || '';
      errors.forEach(msg => {
        reportRows.push([
          'contract_logic_rules',
          rowNum,
          cid,
          coid,
          msg
        ]);
      });
    }
  });

  if (reportRows.length === 0) {
    reportSheet.getRange(2, 1, 1, 5).setValues([[
      'contract_logic_rules',
      '',
      '',
      '',
      'エラーは見つかりませんでした ✅'
    ]]);
    ui.alert('contract_logic_rules 全行をチェックしました。エラーはありませんでした。');
  } else {
    reportSheet.getRange(2, 1, reportRows.length, 5).setValues(reportRows);
    ui.alert(
      `contract_logic_rules 全行のチェックが完了しました。\n` +
      `エラー件数: ${reportRows.length}\n` +
      `詳細はシート "${REPORT_SHEET_NAME}" をご確認ください。`
    );
  }
}


/**************************************************
 * RAG用CSV出力 共通ヘルパー
 **************************************************/


/**
 * 配列 rows([[c1,c2,...], ...]) からCSV文字列を生成
 */
function toCsvString_(rows) {
  return rows
    .map(row =>
      row
        .map(v => {
          const s = v === null || v === undefined ? '' : String(v);
          const escaped = s.replace(/"/g, '""');
          return `"${escaped}"`;
        })
        .join(',')
    )
    .join('\r\n');
}

/**
 * 05_RAG連携 フォルダを取得（なければ作成）
 * ・スプレッドシートと同じ親フォルダ配下に作る試み
 * ・ダメならマイドライブ直下に作成
 */
function getOrCreateRagFolder_() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ssFile.getParents();
  let parentFolder = null;
  if (parents.hasNext()) {
    parentFolder = parents.next();
  } else {
    parentFolder = DriveApp.getRootFolder();
  }

  // 親フォルダ配下に同名フォルダがあるか検索
  const folders = parentFolder.getFoldersByName(RAG_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  // なければ作成
  return parentFolder.createFolder(RAG_FOLDER_NAME);
}

/**
 * 指定フォルダ内の同名ファイルを削除
 */
function deleteFileIfExists_(folder, fileName) {
  const it = folder.getFilesByName(fileName);
  while (it.hasNext()) {
    const f = it.next();
    folder.removeFile(f);
  }
}

/**************************************************
 * 契約マスタ＋解約ロジック → RAG CSV出力
 **************************************************/

function exportContractsRagCsv() {
  const ss           = SpreadsheetApp.getActive();
  const masterSheet  = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  const logicSheet   = ss.getSheetByName('contract_logic_rules');
  const ui           = SpreadsheetApp.getUi();

  if (!masterSheet) {
    ui.alert('contract_master シートが見つかりません。');
    return;
  }
  if (!logicSheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  // --- 1) contract_master からコースの基本情報マップを作る
  const lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster < 3) {
    ui.alert('contract_master にデータ行がありません（3行目以降）。');
    return;
  }

  const lastColMaster = masterSheet.getLastColumn();
  const masterValues  = masterSheet.getRange(3, 1, lastRowMaster - 2, lastColMaster).getValues();

  // course_id → course情報
  const courseMap = {}; // course_id -> {companyId, companyName, courseName, lastUpdated}
  const masterRowsForRag = [];

  masterValues.forEach(row => {
    const lastUpdated   = row[0];  // A: last_updated
    const companyId     = row[1];  // B: client_company_id
    const companyName   = row[2];  // C: client_company_name
    const courseName    = row[3];  // D: course_name
    const category      = row[4];  // E: category
    const subcategories = row[5];  // F: subcategories
    const courseId      = row[6];  // G: course_id
    const contractType  = row[7];  // H: contract_type
    const commitRule    = row[8];  // I: commit_rule
    const exitFeeRule   = row[9];  // J: exit_fee_rule
    const guaranteeType = row[10]; // K: guarantee_type
    const salesChannels = row[11]; // L: sales_channels
    const fulfillRule   = row[12]; // M: fulfillment_rule
    const paymentType   = row[13]; // N: payment_type
    const paymentCat    = row[14]; // O: payment_method_category
    const installment   = row[15]; // P: installment_available
    const firstPrice    = row[16]; // Q: first_price
    const secondPrice   = row[17]; // R: second_price
    const recurringPrice = row[18]; // S: recurring_price
    const productBundle = row[19]; // T: product_bundle
    const billingInt    = row[20]; // U: billing_interval
    const firstCommit   = row[21]; // V: first_commit_count
    const totalCommit   = row[22]; // W: total_commit_count
    const initialGift   = row[23]; // X: initial_gift_flag
    const upsellType    = row[24]; // Y: upsell_type
    const isUpsellTarget = row[25]; // Z: is_upsell_target
    const hasTrigger    = row[26]; // AA: has_trigger_keyword
    const coupon50      = row[27]; // AB: use_coupon_50_off
    const coupon30      = row[28]; // AC: use_coupon_30_off
    const hasPointCancel = row[29]; // AD: has_point_cancel_request
    const warningTags   = row[30]; // AE: contract_warning_tags
    const termsLink     = row[31]; // AF: terms_link
    const remarks       = row[32]; // AG: remarks

    if (!courseId) {
      return; // course_id ない行はRAG対象外
    }

    courseMap[String(courseId)] = {
      companyId:   companyId,
      companyName: companyName,
      courseName:  courseName,
      lastUpdated: lastUpdated
    };

    // --- text を組み立て（RAG用）
    const textParts = [];

    // ---------- サマリー ----------
    const summaryLines = [];
    summaryLines.push('この文書は、以下のコースに関する契約条件と解約・返金ロジックの要約です。');
    if (companyName) summaryLines.push(`・会社名: ${companyName}`);
    if (companyId)   summaryLines.push(`・会社ID: ${companyId}`);
    if (courseName)  summaryLines.push(`・コース名: ${courseName}`);
    summaryLines.push(`・コースID: ${courseId}`);

    textParts.push(summaryLines.join('\n'));

    // ---------- 1. 基本情報 ----------
    textParts.push('\n【1. 基本情報（カテゴリ・支払い・サイクル）】');
    if (m.category)      textParts.push(`・カテゴリ: ${m.category}`);
    if (m.subcategories) textParts.push(`・サブカテゴリ: ${m.subcategories}`);
    if (m.contractType)  textParts.push(`・契約種別: ${m.contractType}`);
    if (m.guaranteeType) textParts.push(`・保証種別: ${m.guaranteeType}`);

    if (m.salesChannels) textParts.push(`・申込チャネル: ${m.salesChannels}`);
    if (m.fulfillRule)   textParts.push(`・定期サイクル: ${m.fulfillRule}`);
    if (m.billingInt)    textParts.push(`・課金間隔: ${m.billingInt}`);

    if (m.paymentType) textParts.push(`・選択可能な支払い区分(payment_type): ${m.paymentType}`);
    if (m.paymentCat)  textParts.push(`・支払い方法カテゴリ(payment_method_category): ${m.paymentCat}`);
    if (m.installment !== "" && m.installment !== null) {
      textParts.push(`・分割払い可否(installment_available): ${m.installment}`);
    }

    if (m.firstPrice !== "" && m.firstPrice !== null) {
      textParts.push(`・初回/単品価格(first_price): ${m.firstPrice}`);
    }
    if (m.secondPrice !== "" && m.secondPrice !== null) {
      textParts.push(`・2回目特別価格(second_price): ${m.secondPrice}`);
    }
    if (m.recurringPrice !== "" && m.recurringPrice !== null) {
      textParts.push(`・定期通常価格(recurring_price): ${m.recurringPrice}`);
    }

    if (m.productBundle) {
      textParts.push(`・商品構成(product_bundle): ${m.productBundle}`);
    }

    // ---------- 2. 契約・縛り条件 ----------
    textParts.push('\n【2. 契約・縛り条件】');
    if (m.commitRule) {
      textParts.push(`・回数縛り区分(commit_rule): ${m.commitRule}`);
    }
    if (m.firstCommit !== "" && m.firstCommit !== null) {
      textParts.push(`・初回縛り回数（最短受取回数 first_commit_count）: ${m.firstCommit}`);
    }
    if (m.totalCommit !== "" && m.totalCommit !== null) {
      textParts.push(`・累計縛り回数（total_commit_count）: ${m.totalCommit}`);
    }

    if (m.initialGift !== "" && m.initialGift !== null) {
      textParts.push(`・初回特典プレゼント同梱フラグ(initial_gift_flag): ${m.initialGift}`);
    }
    if (m.upsellType) {
      textParts.push(`・アップセル種別(upsell_type): ${m.upsellType}`);
    }
    if (m.isUpsellTarget !== "" && m.isUpsellTarget !== null) {
      textParts.push(`・アップセル対象フラグ(is_upsell_target): ${m.isUpsellTarget}`);
    }
    if (m.hasTrigger !== "" && m.hasTrigger !== null) {
      textParts.push(`・トリガーワード有無(has_trigger_keyword): ${m.hasTrigger}`);
    }

    if (m.coupon50 !== "" && m.coupon50 !== null) {
      textParts.push(`・50％OFFクーポン提案有無(use_coupon_50_off): ${m.coupon50}`);
    }
    if (m.coupon30 !== "" && m.coupon30 !== null) {
      textParts.push(`・30％OFFクーポン提案有無(use_coupon_30_off): ${m.coupon30}`);
    }
    if (m.hasPointCancel !== "" && m.hasPointCancel !== null) {
      textParts.push(`・ポイント解約提案有無(has_point_cancel_request): ${m.hasPointCancel}`);
    }

    if (m.warningTags) {
      textParts.push(`・注意事項タグ(contract_warning_tags): ${m.warningTags}`);
    }

    if (m.termsLink) {
      textParts.push(`・規約リンク(terms_link): ${m.termsLink}`);
    }
    if (m.remarks) {
      textParts.push(`・備考(remarks): ${m.remarks}`);
    }

    // ---------- 3. 解約受付・スキップ・欠品 ----------
    textParts.push('\n【3. 解約受付期限・スキップ・欠品時のルール】');
    if (l.cancelDeadline) {
      textParts.push(`・解約受付期限(cancel_deadline): ${l.cancelDeadline}`);
    }
    if (l.cancelDeadlineLogic) {
      textParts.push(`・受付期限判定ロジック(cancel_deadline_logic): ${l.cancelDeadlineLogic}`);
    }
    if (l.holidayHandling) {
      textParts.push(`・解約期限の土日祝扱い(holiday_handling): ${l.holidayHandling}`);
    }
    if (l.longHolidayRule) {
      textParts.push(`・長期連休の解約締切ルール(cancel_deadline_rule_in_long_holiday): ${l.longHolidayRule}`);
    }

    if (l.skipRule) {
      textParts.push(`・スキップ反映ルール(skip_rule): ${l.skipRule}`);
    }
    if (l.longPauseRule) {
      textParts.push(`・長期休止時のルール(long_pause_rule): ${l.longPauseRule}`);
    }

    if (l.oosCancelRule) {
      textParts.push(`・欠品時の受付期限ルール概要(oos_cancel_deadline_rule): ${l.oosCancelRule}`);
    }
    if (l.cancelRuleWhenOos) {
      textParts.push(`・欠品時の解約締切ルール詳細(cancel_deadline_rule_when_oos): ${l.cancelRuleWhenOos}`);
    }

    // ---------- 4. 解約金 ----------
    textParts.push('\n【4. 解約金】');
    if (m.exitFeeRule) {
      textParts.push(`・解約金ルール区分(exit_fee_rule): ${m.exitFeeRule}`);
    }
    if (l.exitFeeCalcMethod) {
      textParts.push(`・解約金計算方法(exit_fee_calc_method): ${l.exitFeeCalcMethod}`);
    }
    if (l.exitFeeAmount !== "" && l.exitFeeAmount !== null) {
      textParts.push(`・解約金額（代表値または固定値 exit_fee_amount）: ${l.exitFeeAmount}`);
    }
    if (l.exitFeeDetail) {
      textParts.push(`・解約金概要(exit_fee_detail): ${l.exitFeeDetail}`);
    }
    if (l.exitFeeCondDetail) {
      textParts.push(`・解約金発生条件の詳細(exit_fee_condition_detail): ${l.exitFeeCondDetail}`);
    }
    if (l.exitFeeWaiverCond) {
      textParts.push(`・解約金免除条件(exit_fee_waiver_condition): ${l.exitFeeWaiverCond}`);
    }
    if (l.exitFeeNoticeTemplate) {
      textParts.push(`・解約金案内トーク（CS向けテンプレ exit_fee_notice_template）:\n${l.exitFeeNoticeTemplate}`);
    }

    // ---------- 5. アップセル部分の解約 ----------
    textParts.push('\n【5. アップセル部分の解約】');
    if (m.upsellType) {
      textParts.push(`・アップセル種別(upsell_type): ${m.upsellType}`);
    }
    if (l.upsellExitFeeFlag !== "" && l.upsellExitFeeFlag !== null) {
      textParts.push(`・アップセル解約金有無(upsell_exit_fee_flag): ${l.upsellExitFeeFlag}`);
    }
    if (l.upsellExitLogicDetail) {
      textParts.push(`・アップセルの解約ロジック詳細(upsell_exit_logic_detail): ${l.upsellExitLogicDetail}`);
    }

    // ---------- 6. 返金保証・返品 ----------
    textParts.push('\n【6. 返金保証・返品】');
    if (l.refundFlag !== "" && l.refundFlag !== null) {
      textParts.push(`・返金保証の有無(refund_guarantee_flag): ${l.refundFlag}`);
    }
    if (l.refundTerm) {
      textParts.push(`・返金保証期間(refund_guarantee_term): ${l.refundTerm}`);
    }
    if (l.refundCondDetail) {
      textParts.push(`・返金保証の詳細条件(refund_guarantee_condition_detail): ${l.refundCondDetail}`);
    }

    if (l.guaranteeReturnReq !== "" && l.guaranteeReturnReq !== null) {
      textParts.push(`・返金保証利用時の返品要否(guarantee_return_required): ${l.guaranteeReturnReq}`);
    }
    if (l.guaranteeReturnDeadline) {
      textParts.push(`・返金保証で返品が必要な場合の期限(guarantee_return_deadline): ${l.guaranteeReturnDeadline}`);
    }

    if (l.exceptionReturnRule) {
      textParts.push(`・特例返品ルール(exception_return_rule): ${l.exceptionReturnRule}`);
    }
    if (l.exceptionReturnDeadline) {
      textParts.push(`・特例返品の期限(exception_return_deadline): ${l.exceptionReturnDeadline}`);
    }

    // ---------- 7. クーリングオフ ----------
    textParts.push('\n【7. クーリングオフ】');
    if (l.coFlag !== "" && l.coFlag !== null) {
      textParts.push(`・クーリングオフ可否(cooling_off_flag): ${l.coFlag}`);
    }
    if (l.coTermType) {
      textParts.push(`・クーリングオフ期間区分(cooling_off_term_type): ${l.coTermType}`);
    }
    if (l.coTerm) {
      textParts.push(`・クーリングオフの具体的な期間数値(cooling_off_term): ${l.coTerm}`);
    }
    if (l.coCondDetail) {
      textParts.push(`・クーリングオフ詳細条件(cooling_off_condition_detail): ${l.coCondDetail}`);
    }

    // ---------- 8. 発送前・発送後キャンセルと注意点 ----------
    textParts.push('\n【8. 発送前・発送後キャンセルと注意点】');
    if (l.firstCancelBeforeShip !== "" && l.firstCancelBeforeShip !== null) {
      textParts.push(`・初回発送前キャンセル可否(first_order_cancelable_before_ship): ${l.firstCancelBeforeShip}`);
    }
    if (l.firstCancelCondition) {
      textParts.push(`・初回発送前キャンセル条件(first_order_cancel_condition): ${l.firstCancelCondition}`);
    }

    if (l.recurCancelBeforeShip !== "" && l.recurCancelBeforeShip !== null) {
      textParts.push(`・継続分の発送前キャンセル可否(recurring_order_cancelable_before_ship): ${l.recurCancelBeforeShip}`);
    }
    if (l.recurCancelCondition) {
      textParts.push(`・継続分発送前キャンセル条件(recurring_order_cancel_condition): ${l.recurCancelCondition}`);
    }

    if (l.cancelAfterShip) {
      textParts.push(`・発送後キャンセル可否(cancel_after_ship): ${l.cancelAfterShip}`);
    }
    if (l.cancelExplainTemplate) {
      textParts.push(`・解約説明テンプレ（顧客案内用 cancel_explanation_template）:\n${l.cancelExplainTemplate}`);
    }

    textParts.push('・顧客が誤認しやすいポイント:');
    if (l.misunderstandingPoints) {
      textParts.push(l.misunderstandingPoints);
    } else {
      textParts.push('　特に明示されたものはありません。');
    }

    const text = textParts.join('\n');


    const docId = `contract_master:${companyId || ''}:${courseId}`;

    masterRowsForRag.push([
      docId,
      companyId || '',
      companyName || '',
      courseId,
      'contract_master',
      lastUpdated || '',
      text
    ]);
  });

  // --- 2) contract_logic_rules からロジック情報のRAG行を作る
  const lastRowLogic = logicSheet.getLastRow();
  const lastColLogic = logicSheet.getLastColumn();
  const logicValues  = logicSheet.getRange(3, 1, lastRowLogic - 2, lastColLogic).getValues();

  const logicRowsForRag = [];

  logicValues.forEach(row => {
    const courseId   = row[0];  // A: course_id
    if (!courseId) return;

    const cancelDeadline        = row[1];  // B
    const cancelDeadlineLogic   = row[2];  // C
    const holidayHandling       = row[3];  // D
    const oosCancelRule         = row[4];  // E
    const skipRule              = row[5];  // F
    const longPauseRule         = row[6];  // G
    const longHolidayRule       = row[7];  // H
    const exitFeeAmount         = row[8];  // I
    const exitFeeCalcMethod     = row[9];  // J
    const exitFeeDetail         = row[10]; // K
    const exitFeeCondDetail     = row[11]; // L
    const exitFeeWaiverCond     = row[12]; // M
    const exitFeeNoticeTemplate = row[13]; // N
    const upsellExitFeeFlag     = row[14]; // O
    const upsellExitLogicDetail = row[15]; // P
    const refundFlag            = row[16]; // Q
    const refundTerm            = row[17]; // R
    const refundCondDetail      = row[18]; // S
    const guaranteeReturnReq    = row[19]; // T
    const guaranteeReturnDeadline = row[20]; // U
    const exceptionReturnRule   = row[21]; // V
    const exceptionReturnDeadline = row[22]; // W
    const coFlag                = row[23]; // X
    const coTermType            = row[24]; // Y
    const coTerm                = row[25]; // Z
    const coCondDetail          = row[26]; // AA
    const firstCancelBeforeShip = row[27]; // AB
    const firstCancelCondition  = row[28]; // AC
    const recurCancelBeforeShip = row[29]; // AD
    const recurCancelCondition  = row[30]; // AE
    const cancelAfterShip       = row[31]; // AF
    const cancelExplainTemplate = row[32]; // AG
    const misunderstandingPoints = row[33]; // AH
    const cancelRuleWhenOos     = row[34]; // AI

    const courseInfo = courseMap[String(courseId)] || {};
    const companyId   = courseInfo.companyId   || '';
    const companyName = courseInfo.companyName || '';
    const lastUpdated = courseInfo.lastUpdated || '';

    const textParts = [];

    textParts.push(`【解約・返金ロジック】`);
    if (companyName) textParts.push(`会社名: ${companyName}`);
    if (companyId)   textParts.push(`会社ID: ${companyId}`);
    textParts.push(`コースID: ${courseId}`);

    if (cancelDeadline)      textParts.push(`解約受付期限: ${cancelDeadline}`);
    if (cancelDeadlineLogic) textParts.push(`受付期限判定ロジック: ${cancelDeadlineLogic}`);
    if (holidayHandling)     textParts.push(`解約期限の土日祝の扱い: ${holidayHandling}`);
    if (longHolidayRule)     textParts.push(`長期連休の解約締切ルール: ${longHolidayRule}`);
    if (oosCancelRule)       textParts.push(`欠品時の受付期限ルール: ${oosCancelRule}`);
    if (cancelRuleWhenOos)   textParts.push(`欠品時の解約締切ルール（詳細）: ${cancelRuleWhenOos}`);

    if (skipRule)        textParts.push(`スキップ反映ルール: ${skipRule}`);
    if (longPauseRule)   textParts.push(`長期休止時のルール: ${longPauseRule}`);

    if (exitFeeCalcMethod) textParts.push(`解約金計算方法: ${exitFeeCalcMethod}`);
    if (exitFeeAmount !== "" && exitFeeAmount !== null) textParts.push(`解約金額: ${exitFeeAmount}`);
    if (exitFeeDetail)         textParts.push(`解約金詳細: ${exitFeeDetail}`);
    if (exitFeeCondDetail)     textParts.push(`解約金発生条件詳細: ${exitFeeCondDetail}`);
    if (exitFeeWaiverCond)     textParts.push(`解約金免除条件: ${exitFeeWaiverCond}`);
    if (exitFeeNoticeTemplate) textParts.push(`解約金案内トーク: ${exitFeeNoticeTemplate}`);

    if (upsellExitFeeFlag !== "" && upsellExitFeeFlag !== null) textParts.push(`アップセル解約金有無: ${upsellExitFeeFlag}`);
    if (upsellExitLogicDetail) textParts.push(`アップセルの解約ロジック詳細: ${upsellExitLogicDetail}`);

    if (refundFlag !== "" && refundFlag !== null) textParts.push(`返金保証フラグ: ${refundFlag}`);
    if (refundTerm)            textParts.push(`返金保証期間: ${refundTerm}`);
    if (refundCondDetail)      textParts.push(`返金保証の詳細条件: ${refundCondDetail}`);
    if (guaranteeReturnReq !== "" && guaranteeReturnReq !== null) textParts.push(`返金保証で返品が必要か: ${guaranteeReturnReq}`);
    if (guaranteeReturnDeadline) textParts.push(`返品期限: ${guaranteeReturnDeadline}`);

    if (exceptionReturnRule)     textParts.push(`特例返品ルール: ${exceptionReturnRule}`);
    if (exceptionReturnDeadline) textParts.push(`特例返品の期限: ${exceptionReturnDeadline}`);

    if (coFlag !== "" && coFlag !== null) textParts.push(`クーリングオフフラグ: ${coFlag}`);
    if (coTermType) textParts.push(`クーリングオフ期間区分: ${coTermType}`);
    if (coTerm)     textParts.push(`クーリングオフ期間: ${coTerm}`);
    if (coCondDetail) textParts.push(`クーリングオフ詳細条件: ${coCondDetail}`);

    if (firstCancelBeforeShip !== "" && firstCancelBeforeShip !== null) {
      textParts.push(`初回発送前キャンセル可否: ${firstCancelBeforeShip}`);
    }
    if (firstCancelCondition) textParts.push(`初回発送前キャンセル条件: ${firstCancelCondition}`);

    if (recurCancelBeforeShip !== "" && recurCancelBeforeShip !== null) {
      textParts.push(`継続分の発送前キャンセル可否: ${recurCancelBeforeShip}`);
    }
    if (recurCancelCondition) textParts.push(`継続分発送前キャンセル条件: ${recurCancelCondition}`);

    if (cancelAfterShip)       textParts.push(`発送後キャンセル可否: ${cancelAfterShip}`);
    if (cancelExplainTemplate) textParts.push(`解約説明テンプレ: ${cancelExplainTemplate}`);
    if (misunderstandingPoints) textParts.push(`顧客が誤認しやすいポイント: ${misunderstandingPoints}`);

    const text = textParts.join('\n');

    const docId = `contract_logic_rules:${companyId || ''}:${courseId}`;

    logicRowsForRag.push([
      docId,
      companyId,
      companyName,
      courseId,
      'contract_logic_rules',
      lastUpdated,
      text
    ]);
  });

  // --- 3) CSVを組み立てて Drive に保存
  const header = [
    'doc_id',
    'client_company_id',
    'client_company_name',
    'course_id',
    'source',
    'last_updated',
    'text'
  ];

  const allRows = [header].concat(masterRowsForRag, logicRowsForRag);
  const csvStr  = toCsvString_(allRows);

  const folder = getOrCreateRagFolder_();
  deleteFileIfExists_(folder, RAG_FILE_NAME);

  const blob = Utilities.newBlob(csvStr, 'text/csv', RAG_FILE_NAME);
  folder.createFile(blob);

  ui.alert(`RAG用CSVを出力しました。\nフォルダ: ${RAG_FOLDER_NAME}\nファイル: ${RAG_FILE_NAME}`);
}

/**************************************************
 * 契約マスタ＋解約ロジック → RAG CSV（1コース1行・要約版）
 **************************************************/

function exportContractsRagLongformCsv() {
  const ss           = SpreadsheetApp.getActive();
  const masterSheet  = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  const logicSheet   = ss.getSheetByName('contract_logic_rules');
  const ui           = SpreadsheetApp.getUi();

  if (!masterSheet) {
    ui.alert('contract_master シートが見つかりません。');
    return;
  }
  if (!logicSheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  // --- 1) contract_master からコース基本情報マップを作成
  const lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster < 3) {
    ui.alert('contract_master にデータ行がありません（3行目以降）。');
    return;
  }
  const lastColMaster = masterSheet.getLastColumn();
  const masterValues  = masterSheet.getRange(3, 1, lastRowMaster - 2, lastColMaster).getValues();

  const masterMap = {}; // course_id -> {...}
  masterValues.forEach(row => {
    const lastUpdated   = row[0];  // A: last_updated
    const companyId     = row[1];  // B: client_company_id
    const companyName   = row[2];  // C: client_company_name
    const courseName    = row[3];  // D: course_name
    const category      = row[4];  // E: category
    const subcategories = row[5];  // F: subcategories
    const courseId      = row[6];  // G: course_id
    const contractType  = row[7];  // H: contract_type
    const commitRule    = row[8];  // I: commit_rule
    const exitFeeRule   = row[9];  // J: exit_fee_rule
    const guaranteeType = row[10]; // K: guarantee_type
    const salesChannels = row[11]; // L: sales_channels
    const fulfillRule   = row[12]; // M: fulfillment_rule
    const paymentType   = row[13]; // N: payment_type
    const paymentCat    = row[14]; // O: payment_method_category
    const installment   = row[15]; // P: installment_available
    const firstPrice    = row[16]; // Q: first_price
    const secondPrice   = row[17]; // R: second_price
    const recurringPrice = row[18]; // S: recurring_price
    const productBundle = row[19]; // T: product_bundle
    const billingInt    = row[20]; // U: billing_interval
    const firstCommit   = row[21]; // V: first_commit_count
    const totalCommit   = row[22]; // W: total_commit_count
    const initialGift   = row[23]; // X: initial_gift_flag
    const upsellType    = row[24]; // Y: upsell_type
    const isUpsellTarget = row[25]; // Z: is_upsell_target
    const hasTrigger    = row[26]; // AA: has_trigger_keyword
    const coupon50      = row[27]; // AB: use_coupon_50_off
    const coupon30      = row[28]; // AC: use_coupon_30_off
    const hasPointCancel = row[29]; // AD: has_point_cancel_request
    const warningTags   = row[30]; // AE: contract_warning_tags
    const termsLink     = row[31]; // AF: terms_link
    const remarks       = row[32]; // AG: remarks

    if (!courseId) return;

    masterMap[String(courseId)] = {
      lastUpdated,
      companyId,
      companyName,
      courseName,
      category,
      subcategories,
      contractType,
      commitRule,
      exitFeeRule,
      guaranteeType,
      salesChannels,
      fulfillRule,
      paymentType,
      paymentCat,
      installment,
      firstPrice,
      secondPrice,
      recurringPrice,
      productBundle,
      billingInt,
      firstCommit,
      totalCommit,
      initialGift,
      upsellType,
      isUpsellTarget,
      hasTrigger,
      coupon50,
      coupon30,
      hasPointCancel,
      warningTags,
      termsLink,
      remarks
    };
  });

  // --- 2) contract_logic_rules からロジック情報マップを作成
  const lastRowLogic = logicSheet.getLastRow();
  const lastColLogic = logicSheet.getLastColumn();
  const logicValues  = logicSheet.getRange(3, 1, lastRowLogic - 2, lastColLogic).getValues();

  const logicMap = {}; // course_id -> {...}
  logicValues.forEach(row => {
    const lastUpdatedLogic = row[0];  // A: last_updated
    const companyId        = row[1];  // B: client_company_id
    const courseId         = row[2];  // C: course_id
    if (!courseId) return;

    const cancelDeadline        = row[3];  // D
    const cancelDeadlineLogic   = row[4];  // E
    const holidayHandling       = row[5];  // F
    const oosCancelRule         = row[6];  // G
    const skipRule              = row[7];  // H
    const longPauseRule         = row[8];  // I
    const longHolidayRule       = row[9];  // J
    const exitFeeAmount         = row[10]; // K
    const exitFeeCalcMethod     = row[11]; // L
    const exitFeeDetail         = row[12]; // M
    const exitFeeCondDetail     = row[13]; // N
    const exitFeeWaiverCond     = row[14]; // O
    const exitFeeNoticeTemplate = row[15]; // P
    const upsellExitFeeFlag     = row[16]; // Q
    const upsellExitLogicDetail = row[17]; // R
    const refundFlag            = row[18]; // S
    const refundTerm            = row[19]; // T
    const refundCondDetail      = row[20]; // U
    const guaranteeReturnReq    = row[21]; // V
    const guaranteeReturnDeadline = row[22]; // W
    const exceptionReturnRule   = row[23]; // X
    const exceptionReturnDeadline = row[24]; // Y
    const coFlag                = row[25]; // Z
    const coTermType            = row[26]; // AA
    const coTerm                = row[27]; // AB
    const coCondDetail          = row[28]; // AC
    const firstCancelBeforeShip = row[29]; // AD
    const firstCancelCondition  = row[30]; // AE
    const recurCancelBeforeShip = row[31]; // AF
    const recurCancelCondition  = row[32]; // AG
    const cancelAfterShip       = row[33]; // AH
    const cancelExplainTemplate = row[34]; // AI
    const misunderstandingPoints = row[35]; // AJ
    const cancelRuleWhenOos     = row[36]; // AK

    logicMap[String(courseId)] = {
      lastUpdatedLogic,
      companyId,
      cancelDeadline,
      cancelDeadlineLogic,
      holidayHandling,
      oosCancelRule,
      skipRule,
      longPauseRule,
      longHolidayRule,
      exitFeeAmount,
      exitFeeCalcMethod,
      exitFeeDetail,
      exitFeeCondDetail,
      exitFeeWaiverCond,
      exitFeeNoticeTemplate,
      upsellExitFeeFlag,
      upsellExitLogicDetail,
      refundFlag,
      refundTerm,
      refundCondDetail,
      guaranteeReturnReq,
      guaranteeReturnDeadline,
      exceptionReturnRule,
      exceptionReturnDeadline,
      coFlag,
      coTermType,
      coTerm,
      coCondDetail,
      firstCancelBeforeShip,
      firstCancelCondition,
      recurCancelBeforeShip,
      recurCancelCondition,
      cancelAfterShip,
      cancelExplainTemplate,
      misunderstandingPoints,
      cancelRuleWhenOos
    };
  });

  // --- 3) 1コース1行の要約テキストを組み立て
  const header = [
    'doc_id',
    'client_company_id',
    'client_company_name',
    'course_id',
    'course_name',
    'source',
    'last_updated',
    'text'
  ];
  const rows = [header];

  Object.keys(masterMap).forEach(courseId => {
    const m = masterMap[courseId];
    const l = logicMap[courseId] || {};

    const companyId   = m.companyId   || l.companyId || '';
    const companyName = m.companyName || '';
    const courseName  = m.courseName  || '';

    let lastUpdated = m.lastUpdated || l.lastUpdatedLogic || '';
    if (lastUpdated instanceof Date) {
      lastUpdated = Utilities.formatDate(
        lastUpdated,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      );
    }

    const textParts = [];

    // ---------- サマリー ----------
    const summaryLines = [];
    summaryLines.push('この文書は、以下のコースに関する契約条件と解約・返金ロジックの要約です。');
    if (companyName) summaryLines.push(`・会社名: ${companyName}`);
    if (companyId)   summaryLines.push(`・会社ID: ${companyId}`);
    if (courseName)  summaryLines.push(`・コース名: ${courseName}`);
    summaryLines.push(`・コースID: ${courseId}`);
    textParts.push(summaryLines.join('\n'));

    // ---------- 1. 基本情報 ----------
    textParts.push('\n【1. 基本情報（カテゴリ・支払い・サイクル）】');
    if (m.category)      textParts.push(`・カテゴリ: ${m.category}`);
    if (m.subcategories) textParts.push(`・サブカテゴリ: ${m.subcategories}`);
    if (m.contractType)  textParts.push(`・契約種別: ${m.contractType}`);
    if (m.guaranteeType) textParts.push(`・保証種別: ${m.guaranteeType}`);

    if (m.salesChannels) textParts.push(`・申込チャネル: ${m.salesChannels}`);
    if (m.fulfillRule)   textParts.push(`・定期サイクル: ${m.fulfillRule}`);
    if (m.billingInt)    textParts.push(`・課金間隔: ${m.billingInt}`);

    if (m.paymentType) textParts.push(`・選択可能な支払い区分(payment_type): ${m.paymentType}`);
    if (m.paymentCat)  textParts.push(`・支払い方法カテゴリ(payment_method_category): ${m.paymentCat}`);
    if (m.installment !== "" && m.installment !== null) {
      textParts.push(`・分割払い可否(installment_available): ${m.installment}`);
    }

    if (m.firstPrice !== "" && m.firstPrice !== null) {
      textParts.push(`・初回/単品価格(first_price): ${m.firstPrice}`);
    }
    if (m.secondPrice !== "" && m.secondPrice !== null) {
      textParts.push(`・2回目特別価格(second_price): ${m.secondPrice}`);
    }
    if (m.recurringPrice !== "" && m.recurringPrice !== null) {
      textParts.push(`・定期通常価格(recurring_price): ${m.recurringPrice}`);
    }

    if (m.productBundle) {
      textParts.push(`・商品構成(product_bundle): ${m.productBundle}`);
    }

    // ---------- 2. 契約・縛り条件 ----------
    textParts.push('\n【2. 契約・縛り条件】');
    if (m.commitRule) {
      textParts.push(`・回数縛り区分(commit_rule): ${m.commitRule}`);
    }
    if (m.firstCommit !== "" && m.firstCommit !== null) {
      textParts.push(`・初回縛り回数（最短受取回数 first_commit_count）: ${m.firstCommit}`);
    }
    if (m.totalCommit !== "" && m.totalCommit !== null) {
      textParts.push(`・累計縛り回数（total_commit_count）: ${m.totalCommit}`);
    }

    if (m.initialGift !== "" && m.initialGift !== null) {
      textParts.push(`・初回特典プレゼント同梱フラグ(initial_gift_flag): ${m.initialGift}`);
    }
    if (m.upsellType) {
      textParts.push(`・アップセル種別(upsell_type): ${m.upsellType}`);
    }
    if (m.isUpsellTarget !== "" && m.isUpsellTarget !== null) {
      textParts.push(`・アップセル対象フラグ(is_upsell_target): ${m.isUpsellTarget}`);
    }
    if (m.hasTrigger !== "" && m.hasTrigger !== null) {
      textParts.push(`・トリガーワード有無(has_trigger_keyword): ${m.hasTrigger}`);
    }

    if (m.coupon50 !== "" && m.coupon50 !== null) {
      textParts.push(`・50％OFFクーポン提案有無(use_coupon_50_off): ${m.coupon50}`);
    }
    if (m.coupon30 !== "" && m.coupon30 !== null) {
      textParts.push(`・30％OFFクーポン提案有無(use_coupon_30_off): ${m.coupon30}`);
    }
    if (m.hasPointCancel !== "" && m.hasPointCancel !== null) {
      textParts.push(`・ポイント解約提案有無(has_point_cancel_request): ${m.hasPointCancel}`);
    }

    if (m.warningTags) {
      textParts.push(`・注意事項タグ(contract_warning_tags): ${m.warningTags}`);
    }

    if (m.termsLink) {
      textParts.push(`・規約リンク(terms_link): ${m.termsLink}`);
    }
    if (m.remarks) {
      textParts.push(`・備考(remarks): ${m.remarks}`);
    }

    // ---------- 3. 解約受付・スキップ・欠品 ----------
    textParts.push('\n【3. 解約受付期限・スキップ・欠品時のルール】');
    if (l.cancelDeadline) {
      textParts.push(`・解約受付期限(cancel_deadline): ${l.cancelDeadline}`);
    }
    if (l.cancelDeadlineLogic) {
      textParts.push(`・受付期限判定ロジック(cancel_deadline_logic): ${l.cancelDeadlineLogic}`);
    }
    if (l.holidayHandling) {
      textParts.push(`・解約期限の土日祝扱い(holiday_handling): ${l.holidayHandling}`);
    }
    if (l.longHolidayRule) {
      textParts.push(`・長期連休の解約締切ルール(cancel_deadline_rule_in_long_holiday): ${l.longHolidayRule}`);
    }

    if (l.skipRule) {
      textParts.push(`・スキップ反映ルール(skip_rule): ${l.skipRule}`);
    }
    if (l.longPauseRule) {
      textParts.push(`・長期休止時のルール(long_pause_rule): ${l.longPauseRule}`);
    }

    if (l.oosCancelRule) {
      textParts.push(`・欠品時の受付期限ルール概要(oos_cancel_deadline_rule): ${l.oosCancelRule}`);
    }
    if (l.cancelRuleWhenOos) {
      textParts.push(`・欠品時の解約締切ルール詳細(cancel_deadline_rule_when_oos): ${l.cancelRuleWhenOos}`);
    }

    // ---------- 4. 解約金 ----------
    textParts.push('\n【4. 解約金】');
    if (m.exitFeeRule) {
      textParts.push(`・解約金ルール区分(exit_fee_rule): ${m.exitFeeRule}`);
    }
    if (l.exitFeeCalcMethod) {
      textParts.push(`・解約金計算方法(exit_fee_calc_method): ${l.exitFeeCalcMethod}`);
    }
    if (l.exitFeeAmount !== "" && l.exitFeeAmount !== null) {
      textParts.push(`・解約金額（代表値または固定値 exit_fee_amount）: ${l.exitFeeAmount}`);
    }
    if (l.exitFeeDetail) {
      textParts.push(`・解約金概要(exit_fee_detail): ${l.exitFeeDetail}`);
    }
    if (l.exitFeeCondDetail) {
      textParts.push(`・解約金発生条件の詳細(exit_fee_condition_detail): ${l.exitFeeCondDetail}`);
    }
    if (l.exitFeeWaiverCond) {
      textParts.push(`・解約金免除条件(exit_fee_waiver_condition): ${l.exitFeeWaiverCond}`);
    }
    if (l.exitFeeNoticeTemplate) {
      textParts.push(`・解約金案内トーク（CS向けテンプレ exit_fee_notice_template）:\n${l.exitFeeNoticeTemplate}`);
    }

    // ---------- 5. アップセル部分の解約 ----------
    textParts.push('\n【5. アップセル部分の解約】');
    if (m.upsellType) {
      textParts.push(`・アップセル種別(upsell_type): ${m.upsellType}`);
    }
    if (l.upsellExitFeeFlag !== "" && l.upsellExitFeeFlag !== null) {
      textParts.push(`・アップセル解約金有無(upsell_exit_fee_flag): ${l.upsellExitFeeFlag}`);
    }
    if (l.upsellExitLogicDetail) {
      textParts.push(`・アップセルの解約ロジック詳細(upsell_exit_logic_detail): ${l.upsellExitLogicDetail}`);
    }

    // ---------- 6. 返金保証・返品 ----------
    textParts.push('\n【6. 返金保証・返品】');
    if (l.refundFlag !== "" && l.refundFlag !== null) {
      textParts.push(`・返金保証の有無(refund_guarantee_flag): ${l.refundFlag}`);
    }
    if (l.refundTerm) {
      textParts.push(`・返金保証期間(refund_guarantee_term): ${l.refundTerm}`);
    }
    if (l.refundCondDetail) {
      textParts.push(`・返金保証の詳細条件(refund_guarantee_condition_detail): ${l.refundCondDetail}`);
    }

    if (l.guaranteeReturnReq !== "" && l.guaranteeReturnReq !== null) {
      textParts.push(`・返金保証利用時の返品要否(guarantee_return_required): ${l.guaranteeReturnReq}`);
    }
    if (l.guaranteeReturnDeadline) {
      textParts.push(`・返金保証で返品が必要な場合の期限(guarantee_return_deadline): ${l.guaranteeReturnDeadline}`);
    }

    if (l.exceptionReturnRule) {
      textParts.push(`・特例返品ルール(exception_return_rule): ${l.exceptionReturnRule}`);
    }
    if (l.exceptionReturnDeadline) {
      textParts.push(`・特例返品の期限(exception_return_deadline): ${l.exceptionReturnDeadline}`);
    }

    // ---------- 7. クーリングオフ ----------
    textParts.push('\n【7. クーリングオフ】');
    if (l.coFlag !== "" && l.coFlag !== null) {
      textParts.push(`・クーリングオフ可否(cooling_off_flag): ${l.coFlag}`);
    }
    if (l.coTermType) {
      textParts.push(`・クーリングオフ期間区分(cooling_off_term_type): ${l.coTermType}`);
    }
    if (l.coTerm) {
      textParts.push(`・クーリングオフの具体的な期間数値(cooling_off_term): ${l.coTerm}`);
    }
    if (l.coCondDetail) {
      textParts.push(`・クーリングオフ詳細条件(cooling_off_condition_detail): ${l.coCondDetail}`);
    }

    // ---------- 8. 発送前・発送後キャンセルと注意点 ----------
    textParts.push('\n【8. 発送前・発送後キャンセルと注意点】');
    if (l.firstCancelBeforeShip !== "" && l.firstCancelBeforeShip !== null) {
      textParts.push(`・初回発送前キャンセル可否(first_order_cancelable_before_ship): ${l.firstCancelBeforeShip}`);
    }
    if (l.firstCancelCondition) {
      textParts.push(`・初回発送前キャンセル条件(first_order_cancel_condition): ${l.firstCancelCondition}`);
    }

    if (l.recurCancelBeforeShip !== "" && l.recurCancelBeforeShip !== null) {
      textParts.push(`・継続分の発送前キャンセル可否(recurring_order_cancelable_before_ship): ${l.recurCancelBeforeShip}`);
    }
    if (l.recurCancelCondition) {
      textParts.push(`・継続分発送前キャンセル条件(recurring_order_cancel_condition): ${l.recurCancelCondition}`);
    }

    if (l.cancelAfterShip) {
      textParts.push(`・発送後キャンセル可否(cancel_after_ship): ${l.cancelAfterShip}`);
    }
    if (l.cancelExplainTemplate) {
      textParts.push(`・解約説明テンプレ（顧客案内用 cancel_explanation_template）:\n${l.cancelExplainTemplate}`);
    }

    textParts.push('・顧客が誤認しやすいポイント:');
    if (l.misunderstandingPoints) {
      textParts.push(l.misunderstandingPoints);
    } else {
      textParts.push('　特に明示されたものはありません。');
    }

    const text = textParts.join('\n');

    const docId = `contract_all:${companyId || ''}:${courseId}`;
    rows.push([
      docId,
      companyId || '',
      companyName || '',
      courseId,
      courseName || '',
      'contract_all',
      lastUpdated || '',
      text
    ]);
  });

  // --- 4) CSVを保存
  const folder = getOrCreateRagFolder_();
  deleteFileIfExists_(folder, RAG_FILE_NAME_LONGFORM);

  const csvStr = toCsvString_(rows);
  const blob   = Utilities.newBlob(csvStr, 'text/csv', RAG_FILE_NAME_LONGFORM);
  folder.createFile(blob);

  ui.alert(
    'RAG用CSV（1コース1行・要約版）を出力しました。\n' +
    `フォルダ: ${RAG_FOLDER_NAME}\nファイル: ${RAG_FILE_NAME_LONGFORM}`
  );
}


