
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
  removeFilesByName_(folder, RAG_FILE_NAME_LONGFORM);

  const csvStr = toCsvString(rows);
  const blob   = Utilities.newBlob(csvStr, 'text/csv', RAG_FILE_NAME_LONGFORM);
  folder.createFile(blob);

  ui.alert(
    'RAG用CSV（1コース1行・要約版）を出力しました。\n' +
    `フォルダ: ${RAG_FOLDER_NAME}\nファイル: ${RAG_FILE_NAME_LONGFORM}`
  );
}


