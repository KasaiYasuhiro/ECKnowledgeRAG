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