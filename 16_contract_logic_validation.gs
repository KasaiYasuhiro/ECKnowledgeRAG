/**************************************************
 * contract_logic_rules バリデーション（選択行）
 **************************************************/
function validateSelectedLogicRow() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName(SHEET_LOGIC_RULES);
  if (!sheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== SHEET_LOGIC_RULES) {
    ui.alert('contract_logic_rules シートで、チェックしたい行を選択してから実行してください。');
    return;
  }

  const range = activeSheet.getActiveRange();
  if (!range) {
    ui.alert('contract_logic_rules シートで、チェックしたい行を選択してください。');
    return;
  }

  const row = range.getRow();
  if (row <= HEADER_ROW) { // HEADER_ROW = 2 の想定
    ui.alert('3行目以降のデータ行を選択してください。');
    return;
  }

  const lastCol = sheet.getLastColumn();
  const values  = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const errors = [];

  // 列マッピング（0始まり index）
  const COL = {
    last_updated:                          0,
    client_company_id:                     1,
    course_id:                             2,
    cancel_deadline:                       3,
    cancel_deadline_logic:                 4,
    holiday_handling:                      5,
    oos_cancel_deadline_rule:              6,
    skip_rule:                             7,
    long_pause_rule:                       8,
    cancel_deadline_rule_in_long_holiday:  9,
    exit_fee_amount:                      10,
    exit_fee_calc_method:                 11,
    exit_fee_detail:                      12,
    exit_fee_condition_detail:            13,
    exit_fee_waiver_condition:            14,
    exit_fee_notice_template:             15,
    upsell_exit_fee_flag:                 16,
    upsell_exit_logic_detail:             17,
    refund_guarantee_flag:                18,
    refund_guarantee_term:                19,
    refund_guarantee_condition_detail:    20,
    guarantee_return_required:            21,
    guarantee_return_deadline:            22,
    exception_return_rule:                23,
    exception_return_deadline:            24,
    cooling_off_flag:                     25,
    cooling_off_term_type:                26,
    cooling_off_term:                     27,
    cooling_off_condition_detail:         28,
    first_order_cancelable_before_ship:   29,
    first_order_cancel_condition:         30,
    recurring_order_cancelable_before_ship: 31,
    recurring_order_cancel_condition:     32,
    cancel_after_ship:                    33,
    cancel_explanation_template:          34,
    customer_misunderstanding_points:     35,
    cancel_deadline_rule_when_oos:        36
  };

  const isEmpty = (v) => v === '' || v === null;

  const courseId = values[COL.course_id];

  /** ------------------------------
   * 必須チェック
   * ------------------------------ */
  if (isEmpty(courseId)) {
    errors.push('コースID（course_id）が未入力です。');
  }

  // course_id が contract_master に存在するか
  const courseSet = getAllCourseIdSet_();
  if (!isEmpty(courseId) && !courseSet.has(String(courseId))) {
    errors.push(`コースID "${courseId}" は contract_master に存在しません。タイポか未登録の可能性があります。`);
  }

  /** ------------------------------
   * マスタ値（code_master）一覧取得
   * ------------------------------ */
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

  // マスタが無い場合のフォールバック用
  const mBoolFlag = ['TRUE', 'FALSE'];

  /** ------------------------------
   * 単一値チェック
   * ------------------------------ */
  validateSingleValueField_(
    '解約期限の土日祝の扱い（holiday_handling）',
    values[COL.holiday_handling],
    mHolidayHandling,
    errors
  );

  validateSingleValueField_(
    '解約金計算方法（exit_fee_calc_method）',
    values[COL.exit_fee_calc_method],
    mExitFeeCalc,
    errors
  );

  /** ------------------------------
   * TRUE/FALSE 系 & フラグ系
   *  - code_master があればそれを優先
   *  - なければ TRUE/FALSE チェック
   * ------------------------------ */
  function checkFlagWithMasterOrBool(label, value, masterList) {
    if (isEmpty(value)) return;

    const v      = String(value).trim();
    const vUpper = v.toUpperCase();

    if (masterList && masterList.length > 0) {
      validateSingleValueField_(label, v, masterList, errors);
      return;
    }

    if (mBoolFlag.indexOf(vUpper) === -1) {
      errors.push(`${label} は TRUE か FALSE で入力してください（現在の値: "${v}"）`);
    }
  }

  checkFlagWithMasterOrBool(
    'アップセル解約金有無（upsell_exit_fee_flag）',
    values[COL.upsell_exit_fee_flag],
    mUpsellExitFlag
  );

  checkFlagWithMasterOrBool(
    '返金保証フラグ（refund_guarantee_flag）',
    values[COL.refund_guarantee_flag],
    mRefundFlag
  );

  checkFlagWithMasterOrBool(
    '返金保証で返品が必要か（guarantee_return_required）',
    values[COL.guarantee_return_required],
    mGuaranteeReturn
  );

  checkFlagWithMasterOrBool(
    'クーリングオフフラグ（cooling_off_flag）',
    values[COL.cooling_off_flag],
    mCoFlag
  );

  checkFlagWithMasterOrBool(
    '初回発送前キャンセル可否（first_order_cancelable_before_ship）',
    values[COL.first_order_cancelable_before_ship],
    mFirstCancelFlag
  );

  checkFlagWithMasterOrBool(
    '継続分の発送前キャンセル可否（recurring_order_cancelable_before_ship）',
    values[COL.recurring_order_cancelable_before_ship],
    mRecurCancelFlag
  );

  validateSingleValueField_(
    'クーリングオフ期間区分（cooling_off_term_type）',
    values[COL.cooling_off_term_type],
    mCoTermType,
    errors
  );

  validateSingleValueField_(
    '発送後キャンセル可否（cancel_after_ship）',
    values[COL.cancel_after_ship],
    mCancelAfterShip,
    errors
  );

  /** ------------------------------
   * 数値チェック
   * ------------------------------ */
  validateNumberField_(
    '解約金額（exit_fee_amount）',
    values[COL.exit_fee_amount],
    errors
  );

  /** ------------------------------
   * 結果表示
   * ------------------------------ */
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
 * contract_logic_rules 全行バリデーション → レポート出力
 **************************************************/
function validateAllLogicRows() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(SHEET_LOGIC_RULES); // 'contract_logic_rules'

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

  // 列マッピング（0始まり index）
  const COL = {
    last_updated:                         0,
    client_company_id:                    1,
    course_id:                            2,
    cancel_deadline:                      3,
    cancel_deadline_logic:                4,
    holiday_handling:                     5,
    oos_cancel_deadline_rule:             6,
    skip_rule:                            7,
    long_pause_rule:                      8,
    cancel_deadline_rule_in_long_holiday: 9,
    exit_fee_amount:                     10,
    exit_fee_calc_method:                11,
    exit_fee_detail:                     12,
    exit_fee_condition_detail:           13,
    exit_fee_waiver_condition:           14,
    exit_fee_notice_template:            15,
    upsell_exit_fee_flag:                16,
    upsell_exit_logic_detail:            17,
    refund_guarantee_flag:               18,
    refund_guarantee_term:               19,
    refund_guarantee_condition_detail:   20,
    guarantee_return_required:           21,
    guarantee_return_deadline:           22,
    exception_return_rule:               23,
    exception_return_deadline:           24,
    cooling_off_flag:                    25,
    cooling_off_term_type:               26,
    cooling_off_term:                    27,
    cooling_off_condition_detail:        28,
    first_order_cancelable_before_ship:  29,
    first_order_cancel_condition:        30,
    recurring_order_cancelable_before_ship: 31,
    recurring_order_cancel_condition:    32,
    cancel_after_ship:                   33,
    cancel_explanation_template:         34,
    customer_misunderstanding_points:    35,
    cancel_deadline_rule_when_oos:       36
  };

  // 「空」判定の共通ヘルパー
  const isEmpty = (v) => v === '' || v === null;

  // レポートシート準備
  const REPORT_SHEET_NAME = 'logic_validation_report';
  let reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(REPORT_SHEET_NAME);
  } else {
    reportSheet.clear();
  }

  reportSheet
    .getRange(1, 1, 1, 5)
    .setValues([['sheet', 'row', 'client_company_id', 'course_id', 'message']]);

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

  const allValues  = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const reportRows = [];

  // TRUE/FALSE or master付きフラグをチェックする共通ヘルパー
  function checkFlagWithMasterOrBool(label, value, masterList, errors) {
    if (isEmpty(value)) return;

    const v      = String(value).trim();
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

    const clientCompanyId = rowValues[COL.client_company_id];
    const courseId        = rowValues[COL.course_id];
    const cancelDeadline  = rowValues[COL.cancel_deadline];
    const cancelLogic     = rowValues[COL.cancel_deadline_logic];

    // ★ 完全に「まだ使ってない行」はスキップ
    const isEffectivelyEmpty =
      isEmpty(clientCompanyId) &&
      isEmpty(courseId) &&
      isEmpty(cancelDeadline) &&
      isEmpty(cancelLogic);

    if (isEffectivelyEmpty) {
      return;
    }

    /** ------------------------------
     * 必須 & 整合性チェック
     * ------------------------------ */
    if (!clientCompanyId) {
      errors.push('会社ID（client_company_id）が未入力です。');
    }
    if (!courseId) {
      errors.push('コースID（course_id）が未入力です。');
    } else if (!courseSet.has(String(courseId))) {
      errors.push(`コースID "${courseId}" は contract_master に存在しません。タイポか未登録の可能性があります。`);
    }

    /** ------------------------------
     * マスタ値チェック
     * ------------------------------ */
    validateSingleValueField_(
      '解約期限の土日祝の扱い（holiday_handling）',
      rowValues[COL.holiday_handling],
      mHolidayHandling,
      errors
    );

    validateSingleValueField_(
      '解約金計算方法（exit_fee_calc_method）',
      rowValues[COL.exit_fee_calc_method],
      mExitFeeCalc,
      errors
    );

    validateSingleValueField_(
      'クーリングオフ期間区分（cooling_off_term_type）',
      rowValues[COL.cooling_off_term_type],
      mCoTermType,
      errors
    );

    validateSingleValueField_(
      '発送後キャンセル可否（cancel_after_ship）',
      rowValues[COL.cancel_after_ship],
      mCancelAfterShip,
      errors
    );

    /** ------------------------------
     * フラグ系（code_master があればそちら優先）
     * ------------------------------ */
    checkFlagWithMasterOrBool(
      'アップセル解約金有無（upsell_exit_fee_flag）',
      rowValues[COL.upsell_exit_fee_flag],
      mUpsellExitFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '返金保証の有無（refund_guarantee_flag）',
      rowValues[COL.refund_guarantee_flag],
      mRefundFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '返金保証利用時の返品要否（guarantee_return_required）',
      rowValues[COL.guarantee_return_required],
      mGuaranteeReturn,
      errors
    );

    checkFlagWithMasterOrBool(
      'クーリングオフ可否（cooling_off_flag）',
      rowValues[COL.cooling_off_flag],
      mCoFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '初回発送前キャンセル可否（first_order_cancelable_before_ship）',
      rowValues[COL.first_order_cancelable_before_ship],
      mFirstCancelFlag,
      errors
    );

    checkFlagWithMasterOrBool(
      '継続分の発送前キャンセル可否（recurring_order_cancelable_before_ship）',
      rowValues[COL.recurring_order_cancelable_before_ship],
      mRecurCancelFlag,
      errors
    );

    /** ------------------------------
     * 数値チェック
     * ------------------------------ */
    validateNumberField_(
      '解約金額（exit_fee_amount）',
      rowValues[COL.exit_fee_amount],
      errors
    );

    /** ------------------------------
     * エラーをレポート行として保存
     * ------------------------------ */
    if (errors.length > 0) {
      const cid  = clientCompanyId || '';
      const coid = courseId || '';
      errors.forEach(msg => {
        reportRows.push([
          SHEET_LOGIC_RULES,
          rowNum,
          cid,
          coid,
          msg
        ]);
      });
    }
  });

  /** ------------------------------
   * レポート書き込み & ダイアログ
   * ------------------------------ */
  if (reportRows.length === 0) {
    reportSheet
      .getRange(2, 1, 1, 5)
      .setValues([[SHEET_LOGIC_RULES, '', '', '', 'エラーは見つかりませんでした ✅']]);

    ui.alert('contract_logic_rules 全行をチェックしました。エラーはありませんでした。');
  } else {
    reportSheet
      .getRange(2, 1, reportRows.length, 5)
      .setValues(reportRows);

    ui.alert(
      `contract_logic_rules 全行のチェックが完了しました。\n` +
      `エラー件数: ${reportRows.length}\n` +
      `詳細はシート "${REPORT_SHEET_NAME}" をご確認ください。`
    );
  }
}
