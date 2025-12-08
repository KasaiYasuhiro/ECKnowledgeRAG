/**************************************************
 * contract_master バリデーション関連
 **************************************************/
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
/**************************************************
 * contract_master 全行バリデーション → レポート出力
 **************************************************/
function validateAllContractRows() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
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

  // 列マッピング（0始まり index）
  const COL = {
    client_company_id:        1,  // B
    client_company_name:      2,  // C
    course_name:              3,  // D
    category:                 4,  // E
    course_id:                6,  // G
    contract_type:            7,  // H
    commit_rule:              8,  // I
    exit_fee_rule:            9,  // J
    guarantee_type:          10,  // K
    sales_channels:          11,  // L
    fulfillment_rule:        12,  // M
    payment_type:            13,  // N
    payment_method_category: 14,  // O
    installment_available:   15,  // P
    first_price:             16,  // Q
    second_price:            17,  // R
    recurring_price:         18,  // S
    billing_interval:        20,  // U
    first_commit_count:      21,  // V
    total_commit_count:      22,  // W
    initial_gift_flag:       23,  // X
    is_upsell_target:        25,  // Z
    has_trigger_keyword:     26,  // AA
    use_coupon_50_off:       27,  // AB
    use_coupon_30_off:       28,  // AC
    has_point_cancel_request:29   // AD
  };

  const REPORT_SHEET_NAME = 'master_validation_report';

  // レポートシート準備
  let reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(REPORT_SHEET_NAME);
  } else {
    reportSheet.clear();
  }

  reportSheet
    .getRange(1, 1, 1, 5)
    .setValues([['sheet', 'row', 'client_company_id', 'course_id', 'message']]);

  const courseIdColIndex = 7; // G列 = course_id（レポート用に残しておく）

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

  const allValues  = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const reportRows = [];

  // 共通の「空判定」ヘルパー
  const isEmpty = (v) => v === '' || v === null;

  allValues.forEach((rowValues, idx) => {
    const rowNum = idx + 3;
    const errors = [];

    const clientCompanyId   = rowValues[COL.client_company_id];
    const clientCompanyName = rowValues[COL.client_company_name];
    const courseName        = rowValues[COL.course_name];
    const category          = rowValues[COL.category];
    const courseId          = rowValues[COL.course_id];

    // 事実上「未入力」の行はスキップ
    const isEffectivelyEmpty =
      isEmpty(clientCompanyId) &&
      isEmpty(clientCompanyName) &&
      isEmpty(courseName) &&
      isEmpty(category) &&
      isEmpty(courseId);

    if (isEffectivelyEmpty) {
      return; // この行はチェック対象外
    }

    const contractType   = rowValues[COL.contract_type];
    const guaranteeType  = rowValues[COL.guarantee_type];
    const paymentTypeRaw = rowValues[COL.payment_type];

    /** ------------------------------
     * 必須チェック
     * ------------------------------ */
    if (!clientCompanyId)   errors.push('会社ID が未入力です。');
    if (!clientCompanyName) errors.push('会社名 が未入力です。');
    if (!courseName)        errors.push('コース名 が未入力です。');
    if (!category)          errors.push('カテゴリ（category）が未入力です。');
    if (!courseId)          errors.push('コースID（course_id）が未入力です。');
    if (!contractType)      errors.push('契約種別（contract_type）が未入力です。');
    if (!guaranteeType)     errors.push('保証種別（guarantee_type）が未入力です。');
    // ★ 支払い区分は「任意」にするので、必須チェックしない
    // if (!paymentTypeRaw) errors.push('支払い区分（payment_type）が未入力です。');

    /** ------------------------------
     * 単一値チェック
     * ------------------------------ */
    validateSingleValueField_('カテゴリ（category）', category, mCategory, errors);
    validateSingleValueField_('契約種別（contract_type）', contractType, mContractType, errors);
    validateSingleValueField_('回数縛り条件（commit_rule）', rowValues[COL.commit_rule], mCommitRule, errors);
    validateSingleValueField_('解約金条件（exit_fee_rule）', rowValues[COL.exit_fee_rule], mExitFeeRule, errors);
    validateSingleValueField_('保証種別（guarantee_type）', guaranteeType, mGuarantee, errors);
    validateSingleValueField_('定期サイクル（fulfillment_rule）', rowValues[COL.fulfillment_rule], mFulfill, errors);
    validateSingleValueField_('課金間隔（billing_interval）', rowValues[COL.billing_interval], mBillingInt, errors);

    /** ------------------------------
     * 複数値チェック（; 区切り）
     * ------------------------------ */
    validateMultiValueField_('申込可能チャネル（sales_channels）', rowValues[COL.sales_channels], mSalesCh, errors);
    validateMultiValueField_('支払い区分（payment_type）', paymentTypeRaw, mPayType, errors);
    validateMultiValueField_('支払い方法カテゴリ（payment_method_category）', rowValues[COL.payment_method_category], mPayCat, errors);

    /** ------------------------------
     * TRUE/FALSE 系
     * ------------------------------ */
    const installAvailable = rowValues[COL.installment_available];
    const initialGift      = rowValues[COL.initial_gift_flag];
    const isUpsellTarget   = rowValues[COL.is_upsell_target];
    const hasTrigger       = rowValues[COL.has_trigger_keyword];
    const coupon50         = rowValues[COL.use_coupon_50_off];
    const coupon30         = rowValues[COL.use_coupon_30_off];
    const hasPointCancel   = rowValues[COL.has_point_cancel_request];

    // installment_available は code_master があればそれを優先
    if (installAvailable && mInstall.length > 0) {
      validateSingleValueField_(
        '分割払い可否（installment_available）',
        installAvailable,
        mInstall,
        errors
      );
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
      if (!isEmpty(v)) {
        const vv      = String(v).trim();
        const vvUpper = vv.toUpperCase();
        if (vv && mBoolFlag.indexOf(vvUpper) === -1) {
          errors.push(`${label} は TRUE か FALSE で入力してください（現在の値: "${vv}"）`);
        }
      }
    });

    /** ------------------------------
     * 数値チェック
     * ------------------------------ */
    validateNumberField_('初回/単品価格（first_price）', rowValues[COL.first_price], errors);
    validateNumberField_('2回目特別価格（second_price）', rowValues[COL.second_price], errors);
    validateNumberField_('定期の通常価格（recurring_price）', rowValues[COL.recurring_price], errors);
    validateNumberField_('初回縛り回数（first_commit_count）', rowValues[COL.first_commit_count], errors);
    validateNumberField_('累計縛り回数（total_commit_count）', rowValues[COL.total_commit_count], errors);

    /** ------------------------------
     * レポート行に反映
     * ------------------------------ */
    if (errors.length > 0) {
      const cid  = clientCompanyId || '';
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

  /** ------------------------------
   * レポートシートへの書き込み & ダイアログ
   * ------------------------------ */
  if (reportRows.length === 0) {
    reportSheet
      .getRange(2, 1, 1, 5)
      .setValues([[SHEET_CONTRACT_MASTER, '', '', '', 'エラーは見つかりませんでした ✅']]);

    ui.alert('全ての行をチェックしました。エラーはありませんでした。');
  } else {
    reportSheet
      .getRange(2, 1, reportRows.length, 5)
      .setValues(reportRows);

    ui.alert(
      `contract_master 全行のチェックが完了しました。\n` +
      `エラー件数: ${reportRows.length}\n` +
      `詳細はシート "${REPORT_SHEET_NAME}" をご確認ください。`
    );
  }
}
