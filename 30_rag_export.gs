// 30_rag_export.gs
// ナレッジDB → RAG用CSVエクスポート関連

/**
 * CSVを安全に生成する関数
 * （カンマ・改行・ダブルクォートを正しくエスケープ）
 */
function toCsv_(rows) {
  return rows
    .map(row => row.map(field => {
      if (field === null || field === undefined) return '';
      const str = String(field);
      // ダブルクォート → "" に変換
      const escaped = str.replace(/"/g, '""');
      // カンマ or 改行 を含む場合は "" で囲む
      if (/[,"\n]/.test(escaped)) {
        return `"${escaped}"`;
      }
      return escaped;
    }).join(','))
    .join('\n');
}

/**
 * ナレッジDBからRAG用CSVをエクスポートする
 * - ステータスが「承認済み」の行だけを対象
 * - テキストを1フィールドにまとめて出力
 * - Driveフォルダ「05_RAG連携」にCSVを保存
 */
function exportRagCsv() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_KNOWLEDGE_DB); // ← 定数を利用（'ナレッジDB'）

  if (!sheet) {
    throw new Error('ナレッジDB シートが見つかりません');
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length <= 1) {
    throw new Error('ナレッジDBにデータ行がありません');
  }

  // 1行目をヘッダーとして想定（KNOW_HEADER_ROW = 1）
  const header = values[0];
  const dataRows = values.slice(1);

  // 出力CSVのヘッダー
  const output = [];
  output.push([
    'id',
    'title',
    'text',
    'main_category',
    'sub_categories',
    'type',
    'product',
    'client',
    'updated_at',
    'status',
    'source_url',
    'tags'
  ]);

  dataRows.forEach(row => {
    const KNOW_ID    = row[0];
    const TITLE      = row[1];
    const SUMMARY    = row[2];
    const MAIN_CAT   = row[3];
    const SUB_CATS   = row[4];
    const TYPE       = row[5];
    const PRODUCT    = row[6];
    const CLIENT     = row[7];
    const CORE       = row[8];
    const DELTA      = row[9];
    const NG         = row[10];
    const REASON     = row[11];
    const SOURCE_URL = row[12];
    const DUE        = row[13];
    const UPDATED_AT = row[14];
    const AUTHOR     = row[15];
    const STATUS     = row[16];

    // 承認済みだけ出力
    if (STATUS !== '承認済み') return;
    if (!CORE) return; // 空行はスキップ

    // 本文結合
    const textParts = [];

    if (TITLE)   textParts.push('【タイトル】\n' + TITLE);
    if (SUMMARY) textParts.push('【概要】\n' + SUMMARY);
    if (CORE)    textParts.push('【本文（共通ルール）】\n' + CORE);
    if (DELTA)   textParts.push('【差分・例外ルール】\n' + DELTA);
    if (NG)      textParts.push('【禁止事項・注意事項】\n' + NG);
    if (REASON)  textParts.push('【更新理由】\n' + REASON);
    if (DUE)     textParts.push('【希望反映期限】\n' + DUE);
    if (AUTHOR)  textParts.push('【登録者】\n' + AUTHOR);

    const text = textParts.join('\n\n');

    // タグ
    const tags = [
      MAIN_CAT,
      SUB_CATS,
      TYPE,
      PRODUCT,
      CLIENT
    ].filter(String).join(', ');

    // 日付整形
    let updatedStr = '';
    if (UPDATED_AT instanceof Date) {
      updatedStr = Utilities.formatDate(UPDATED_AT, 'Asia/Tokyo', 'yyyy-MM-dd\'T\'HH:mm:ss');
    } else if (UPDATED_AT) {
      updatedStr = UPDATED_AT;
    }

    output.push([
      KNOW_ID,
      TITLE,
      text,
      MAIN_CAT,
      SUB_CATS,
      TYPE,
      PRODUCT,
      CLIENT,
      updatedStr,
      STATUS,
      SOURCE_URL,
      tags
    ]);
  });

  if (output.length <= 1) {
    throw new Error('エクスポート対象の承認済みナレッジがありません');
  }

  const csvString = toCsv_(output);

  const folderName = '05_RAG連携';
  const folder = getOrCreateFolderByName_(folderName);

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `knowledge_rag_${timestamp}.csv`;

  const blob = Utilities.newBlob(csvString, 'text/csv', fileName);
  folder.createFile(blob);
}

/**
 * 指定名のフォルダを取得（なければ作成）
 */
function getOrCreateFolderByName_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) {
    return it.next();
  }
  return DriveApp.createFolder(name);
}

/**************************************************
 * RAG用CSV出力（30_rag_export.gs）
 **************************************************/

/**
 * 05_RAG連携 フォルダを取得（なければ作成）
 * ・スプレッドシートと同じ親フォルダ配下に作成
 * ・親フォルダが見つからなければマイドライブ直下に作成
 */
function getOrCreateRagFolder_() {
  var parentFolder = getSpreadsheetParentFolder_(); // 共通ユーティリティ
  return getOrCreateChildFolder_(parentFolder, RAG_FOLDER_NAME); // 00_constants.gs 側の定数を利用
}

/**
 * RAGエクスポート用ファイルを上書きしたい場合に使用する削除ヘルパー
 * - 実装は共通ユーティリティのラッパー
 */
function deleteRagFileIfExists_(fileName) {
  var ragFolder = getOrCreateRagFolder_();
  removeFilesByName_(ragFolder, fileName);
}

/**************************************************
 * 契約マスタ＋解約ロジック → RAG CSV出力
 **************************************************/

function exportContractsRagCsv() {
  const ss          = SpreadsheetApp.getActive();
  const masterSheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  const logicSheet  = ss.getSheetByName('contract_logic_rules');
  const ui          = SpreadsheetApp.getUi();

  if (!masterSheet) {
    ui.alert('contract_master シートが見つかりません。');
    return;
  }
  if (!logicSheet) {
    ui.alert('contract_logic_rules シートが見つかりません。');
    return;
  }

  // --- 1) master / logic の生データ取得 -----------------------
  const lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster < 3) {
    ui.alert('contract_master にデータ行がありません（3行目以降）。');
    return;
  }
  const lastColMaster = masterSheet.getLastColumn();
  const masterValues = masterSheet.getRange(3, 1, lastRowMaster - 2, lastColMaster).getValues();

  const lastRowLogic = logicSheet.getLastRow();
  const lastColLogic = logicSheet.getLastColumn();
  const logicValues =
    lastRowLogic > 2
      ? logicSheet.getRange(3, 1, lastRowLogic - 2, lastColLogic).getValues()
      : [];

  // --- 2) 行データを course_id ごとのオブジェクトに変換 -------

  /** course_id → master行のオブジェクト */
  const masterMap = {};
  /** course_id → logic行のオブジェクト */
  const logicMap = {};
  /** course_id → { companyId, companyName, courseName, lastUpdated } */
  const courseBasicMap = {};

  // contract_master 側のマップ作成
  masterValues.forEach(row => {
    const m = buildMasterRowObject_(row);
    if (!m.courseId) return;

    const key = String(m.courseId);
    masterMap[key] = m;
    courseBasicMap[key] = {
      companyId:   m.companyId,
      companyName: m.companyName,
      courseName:  m.courseName,
      lastUpdated: m.lastUpdated
    };
  });

  // contract_logic_rules 側のマップ作成
  logicValues.forEach(row => {
    const l = buildLogicRowObject_(row);
    if (!l.courseId) return;

    const key = String(l.courseId);
    logicMap[key] = l;
    // master に存在しない courseId は company 情報が空のままになるが、それは許容
  });

  // --- 3) RAG行生成 -------------------------------------------

  const masterRowsForRag = [];
  const logicRowsForRag  = [];

  // 3-1) master + logic をまとめた RAG行（source = contract_master）
  Object.keys(masterMap).forEach(courseId => {
    const m = masterMap[courseId];
    const l = logicMap[courseId] || {};
    const basic = courseBasicMap[courseId] || {};

    const text = buildContractMasterText_(m, l);

    const docId = `contract_master:${basic.companyId || ''}:${courseId}`;

    masterRowsForRag.push([
      docId,
      basic.companyId || '',
      basic.companyName || '',
      courseId,
      'contract_master',
      basic.lastUpdated || '',
      text
    ]);
  });

  // 3-2) 解約ロジックのみの RAG行（source = contract_logic_rules）
  Object.keys(logicMap).forEach(courseId => {
    const l = logicMap[courseId];
    const basic = courseBasicMap[courseId] || {};

    const text = buildLogicText_(courseId, basic, l);

    const docId = `contract_logic_rules:${basic.companyId || ''}:${courseId}`;

    logicRowsForRag.push([
      docId,
      basic.companyId || '',
      basic.companyName || '',
      courseId,
      'contract_logic_rules',
      basic.lastUpdated || '',
      text
    ]);
  });

  // --- 4) CSV を生成して Drive に保存 --------------------------

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
  const csvStr  = toCsvString(allRows);    // 90_utils.gs で定義済み想定

  const folder = getOrCreateRagFolder_();  // 30_rag_export.gs のヘルパ or 90_utils
  removeFilesByName_(folder, RAG_FILE_NAME);

  const blob = Utilities.newBlob(csvStr, 'text/csv', RAG_FILE_NAME);
  folder.createFile(blob);

  ui.alert(
    'RAG用CSVを出力しました。\n' +
    'フォルダ: ' + RAG_FOLDER_NAME + '\n' +
    'ファイル: ' + RAG_FILE_NAME
  );
}
/**************************************************
 * 行 → オブジェクト変換ヘルパー
 **************************************************/

/**
 * contract_master の 1行分をオブジェクトに変換
 */
function buildMasterRowObject_(row) {
  const [
    lastUpdated,     // A: last_updated
    companyId,       // B: client_company_id
    companyName,     // C: client_company_name
    courseName,      // D: course_name
    category,        // E: category
    subcategories,   // F: subcategories
    courseId,        // G: course_id
    contractType,    // H: contract_type
    commitRule,      // I: commit_rule
    exitFeeRule,     // J: exit_fee_rule
    guaranteeType,   // K: guarantee_type
    salesChannels,   // L: sales_channels
    fulfillRule,     // M: fulfillment_rule
    paymentType,     // N: payment_type
    paymentCat,      // O: payment_method_category
    installment,     // P: installment_available
    firstPrice,      // Q: first_price
    secondPrice,     // R: second_price
    recurringPrice,  // S: recurring_price
    productBundle,   // T: product_bundle
    billingInt,      // U: billing_interval
    firstCommit,     // V: first_commit_count
    totalCommit,     // W: total_commit_count
    initialGift,     // X: initial_gift_flag
    upsellType,      // Y: upsell_type
    isUpsellTarget,  // Z: is_upsell_target
    hasTrigger,      // AA: has_trigger_keyword
    coupon50,        // AB: use_coupon_50_off
    coupon30,        // AC: use_coupon_30_off
    hasPointCancel,  // AD: has_point_cancel_request
    warningTags,     // AE: contract_warning_tags
    termsLink,       // AF: terms_link
    remarks          // AG: remarks
  ] = row;

  return {
    lastUpdated,
    companyId,
    companyName,
    courseName,
    category,
    subcategories,
    courseId,
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
}

/**
 * contract_logic_rules の 1行分をオブジェクトに変換
 */
function buildLogicRowObject_(row) {
  const [
    courseId,              // A
    cancelDeadline,        // B
    cancelDeadlineLogic,   // C
    holidayHandling,       // D
    oosCancelRule,         // E
    skipRule,              // F
    longPauseRule,         // G
    longHolidayRule,       // H
    exitFeeAmount,         // I
    exitFeeCalcMethod,     // J
    exitFeeDetail,         // K
    exitFeeCondDetail,     // L
    exitFeeWaiverCond,     // M
    exitFeeNoticeTemplate, // N
    upsellExitFeeFlag,     // O
    upsellExitLogicDetail, // P
    refundFlag,            // Q
    refundTerm,            // R
    refundCondDetail,      // S
    guaranteeReturnReq,    // T
    guaranteeReturnDeadline,// U
    exceptionReturnRule,   // V
    exceptionReturnDeadline,// W
    coFlag,                // X
    coTermType,            // Y
    coTerm,                // Z
    coCondDetail,          // AA
    firstCancelBeforeShip, // AB
    firstCancelCondition,  // AC
    recurCancelBeforeShip, // AD
    recurCancelCondition,  // AE
    cancelAfterShip,       // AF
    cancelExplainTemplate, // AG
    misunderstandingPoints, // AH
    cancelRuleWhenOos      // AI
  ] = row;

  return {
    courseId,
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
}
/**************************************************
 * テキスト組み立てヘルパー
 **************************************************/

/**
 * contract_master をベースに、logic も含めた全文を組み立て
 */
function buildContractMasterText_(m, l) {
  const textParts = [];

  // ---------- サマリー ----------
  const summaryLines = [];
  summaryLines.push('この文書は、以下のコースに関する契約条件と解約・返金ロジックの要約です。');
  if (m.companyName) summaryLines.push(`・会社名: ${m.companyName}`);
  if (m.companyId)   summaryLines.push(`・会社ID: ${m.companyId}`);
  if (m.courseName)  summaryLines.push(`・コース名: ${m.courseName}`);
  summaryLines.push(`・コースID: ${m.courseId}`);

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
  if (m.installment !== '' && m.installment !== null) {
    textParts.push(`・分割払い可否(installment_available): ${m.installment}`);
  }

  if (m.firstPrice !== '' && m.firstPrice !== null) {
    textParts.push(`・初回/単品価格(first_price): ${m.firstPrice}`);
  }
  if (m.secondPrice !== '' && m.secondPrice !== null) {
    textParts.push(`・2回目特別価格(second_price): ${m.secondPrice}`);
  }
  if (m.recurringPrice !== '' && m.recurringPrice !== null) {
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
  if (m.firstCommit !== '' && m.firstCommit !== null) {
    textParts.push(`・初回縛り回数（最短受取回数 first_commit_count）: ${m.firstCommit}`);
  }
  if (m.totalCommit !== '' && m.totalCommit !== null) {
    textParts.push(`・累計縛り回数（total_commit_count）: ${m.totalCommit}`);
  }

  if (m.initialGift !== '' && m.initialGift !== null) {
    textParts.push(`・初回特典プレゼント同梱フラグ(initial_gift_flag): ${m.initialGift}`);
  }
  if (m.upsellType) {
    textParts.push(`・アップセル種別(upsell_type): ${m.upsellType}`);
  }
  if (m.isUpsellTarget !== '' && m.isUpsellTarget !== null) {
    textParts.push(`・アップセル対象フラグ(is_upsell_target): ${m.isUpsellTarget}`);
  }
  if (m.hasTrigger !== '' && m.hasTrigger !== null) {
    textParts.push(`・トリガーワード有無(has_trigger_keyword): ${m.hasTrigger}`);
  }

  if (m.coupon50 !== '' && m.coupon50 !== null) {
    textParts.push(`・50％OFFクーポン提案有無(use_coupon_50_off): ${m.coupon50}`);
  }
  if (m.coupon30 !== '' && m.coupon30 !== null) {
    textParts.push(`・30％OFFクーポン提案有無(use_coupon_30_off): ${m.coupon30}`);
  }
  if (m.hasPointCancel !== '' && m.hasPointCancel !== null) {
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
  if (l.exitFeeAmount !== '' && l.exitFeeAmount !== null) {
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
  if (l.upsellExitFeeFlag !== '' && l.upsellExitFeeFlag !== null) {
    textParts.push(`・アップセル解約金有無(upsell_exit_fee_flag): ${l.upsellExitFeeFlag}`);
  }
  if (l.upsellExitLogicDetail) {
    textParts.push(`・アップセルの解約ロジック詳細(upsell_exit_logic_detail): ${l.upsellExitLogicDetail}`);
  }

  // ---------- 6. 返金保証・返品 ----------
  textParts.push('\n【6. 返金保証・返品】');
  if (l.refundFlag !== '' && l.refundFlag !== null) {
    textParts.push(`・返金保証の有無(refund_guarantee_flag): ${l.refundFlag}`);
  }
  if (l.refundTerm) {
    textParts.push(`・返金保証期間(refund_guarantee_term): ${l.refundTerm}`);
  }
  if (l.refundCondDetail) {
    textParts.push(`・返金保証の詳細条件(refund_guarantee_condition_detail): ${l.refundCondDetail}`);
  }

  if (l.guaranteeReturnReq !== '' && l.guaranteeReturnReq !== null) {
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
  if (l.coFlag !== '' && l.coFlag !== null) {
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
  if (l.firstCancelBeforeShip !== '' && l.firstCancelBeforeShip !== null) {
    textParts.push(`・初回発送前キャンセル可否(first_order_cancelable_before_ship): ${l.firstCancelBeforeShip}`);
  }
  if (l.firstCancelCondition) {
    textParts.push(`・初回発送前キャンセル条件(first_order_cancel_condition): ${l.firstCancelCondition}`);
  }

  if (l.recurCancelBeforeShip !== '' && l.recurCancelBeforeShip !== null) {
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

  return textParts.join('\n');
}
/**
 * 解約ロジックのみのテキスト
 */
function buildLogicText_(courseId, basic, l) {
  const textParts = [];

  textParts.push('【解約・返金ロジック】');
  if (basic.companyName) textParts.push(`会社名: ${basic.companyName}`);
  if (basic.companyId)   textParts.push(`会社ID: ${basic.companyId}`);
  textParts.push(`コースID: ${courseId}`);

  if (l.cancelDeadline)      textParts.push(`解約受付期限: ${l.cancelDeadline}`);
  if (l.cancelDeadlineLogic) textParts.push(`受付期限判定ロジック: ${l.cancelDeadlineLogic}`);
  if (l.holidayHandling)     textParts.push(`解約期限の土日祝の扱い: ${l.holidayHandling}`);
  if (l.longHolidayRule)     textParts.push(`長期連休の解約締切ルール: ${l.longHolidayRule}`);
  if (l.oosCancelRule)       textParts.push(`欠品時の受付期限ルール: ${l.oosCancelRule}`);
  if (l.cancelRuleWhenOos)   textParts.push(`欠品時の解約締切ルール（詳細）: ${l.cancelRuleWhenOos}`);

  if (l.skipRule)      textParts.push(`スキップ反映ルール: ${l.skipRule}`);
  if (l.longPauseRule) textParts.push(`長期休止時のルール: ${l.longPauseRule}`);

  if (l.exitFeeCalcMethod)  textParts.push(`解約金計算方法: ${l.exitFeeCalcMethod}`);
  if (l.exitFeeAmount !== '' && l.exitFeeAmount !== null) textParts.push(`解約金額: ${l.exitFeeAmount}`);
  if (l.exitFeeDetail)      textParts.push(`解約金詳細: ${l.exitFeeDetail}`);
  if (l.exitFeeCondDetail)  textParts.push(`解約金発生条件詳細: ${l.exitFeeCondDetail}`);
  if (l.exitFeeWaiverCond)  textParts.push(`解約金免除条件: ${l.exitFeeWaiverCond}`);
  if (l.exitFeeNoticeTemplate) textParts.push(`解約金案内トーク: ${l.exitFeeNoticeTemplate}`);

  if (l.upsellExitFeeFlag !== '' && l.upsellExitFeeFlag !== null) {
    textParts.push(`アップセル解約金有無: ${l.upsellExitFeeFlag}`);
  }
  if (l.upsellExitLogicDetail) {
    textParts.push(`アップセルの解約ロジック詳細: ${l.upsellExitLogicDetail}`);
  }

  if (l.refundFlag !== '' && l.refundFlag !== null) {
    textParts.push(`返金保証フラグ: ${l.refundFlag}`);
  }
  if (l.refundTerm)       textParts.push(`返金保証期間: ${l.refundTerm}`);
  if (l.refundCondDetail) textParts.push(`返金保証の詳細条件: ${l.refundCondDetail}`);
  if (l.guaranteeReturnReq !== '' && l.guaranteeReturnReq !== null) {
    textParts.push(`返金保証で返品が必要か: ${l.guaranteeReturnReq}`);
  }
  if (l.guaranteeReturnDeadline) {
    textParts.push(`返品期限: ${l.guaranteeReturnDeadline}`);
  }

  if (l.exceptionReturnRule) {
    textParts.push(`特例返品ルール: ${l.exceptionReturnRule}`);
  }
  if (l.exceptionReturnDeadline) {
    textParts.push(`特例返品の期限: ${l.exceptionReturnDeadline}`);
  }

  if (l.coFlag !== '' && l.coFlag !== null) textParts.push(`クーリングオフフラグ: ${l.coFlag}`);
  if (l.coTermType) textParts.push(`クーリングオフ期間区分: ${l.coTermType}`);
  if (l.coTerm)     textParts.push(`クーリングオフ期間: ${l.coTerm}`);
  if (l.coCondDetail) textParts.push(`クーリングオフ詳細条件: ${l.coCondDetail}`);

  if (l.firstCancelBeforeShip !== '' && l.firstCancelBeforeShip !== null) {
    textParts.push(`初回発送前キャンセル可否: ${l.firstCancelBeforeShip}`);
  }
  if (l.firstCancelCondition) {
    textParts.push(`初回発送前キャンセル条件: ${l.firstCancelCondition}`);
  }

  if (l.recurCancelBeforeShip !== '' && l.recurCancelBeforeShip !== null) {
    textParts.push(`継続分の発送前キャンセル可否: ${l.recurCancelBeforeShip}`);
  }
  if (l.recurCancelCondition) {
    textParts.push(`継続分発送前キャンセル条件: ${l.recurCancelCondition}`);
  }

  if (l.cancelAfterShip)       textParts.push(`発送後キャンセル可否: ${l.cancelAfterShip}`);
  if (l.cancelExplainTemplate) textParts.push(`解約説明テンプレ: ${l.cancelExplainTemplate}`);
  if (l.misunderstandingPoints) {
    textParts.push(`顧客が誤認しやすいポイント: ${l.misunderstandingPoints}`);
  }

  return textParts.join('\n');
}
