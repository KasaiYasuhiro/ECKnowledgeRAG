
// 20_contract_master_sync.gs
// course_master_source → contract_master / contract_logic_rules / fee_table_master
    // の自動反映ロジック一式をまとめたファイル

/**************************************************
 * コア処理本体
 * mode = 'appendOnly' | 'overwrite'
 **************************************************/
function updateContractsFromCourseSource_v2_core_(mode) { 
  mode = mode || 'appendOnly';  // 念のためデフォルト

  const ss          = SpreadsheetApp.getActive();
  const srcSheet    = ss.getSheetByName(SHEET_COURSE_SOURCE);
  const masterSheet = ss.getSheetByName(SHEET_CONTRACT_MASTER);
  const logicSheet  = ss.getSheetByName(SHEET_LOGIC_RULES);
  const ui          = SpreadsheetApp.getUi();

  if (!srcSheet || !masterSheet || !logicSheet) {
    ui.alert('course_master_source / contract_master / contract_logic_rules のいずれかが見つかりません。');
    return;
  }

  const srcValues = srcSheet.getDataRange().getValues();
  const numRows   = srcValues.length;
  const numCols   = srcValues[0].length;

  if (numRows < 2 || numCols < 2) {
    ui.alert('course_master_source に有効なデータ（2列目以降）がありません。');
    return;
  }

  // ---- ラベル行（A列）から必要な行番号を取得 ----
  const rowCompanyId  = findRowIndexByLabel_(srcValues, '会社ID');
  const rowCourseName = findRowIndexByLabel_(srcValues, 'コース名');
  const rowCourseId   = findRowIndexByLabel_(srcValues, '商品コード');
  const rowMemo       = findRowIndexByLabel_(srcValues, '備考');

  const rowCycleFirst  = findRowIndexByLabel_(srcValues, '初回⇒定期2回目');
  const rowCycleSecond = findRowIndexByLabel_(srcValues, '定期2回目以降');

  const rowBundleFirst   = findRowIndexByLabel_(srcValues, '初回');        // 商品内訳ブロックの「初回」
  const rowBundleSecond  = findRowIndexByLabelAfter_(srcValues, '定期2回目', rowBundleFirst);
  const rowBundleThird   = findRowIndexByLabelAfter_(srcValues, '定期3回目以降', rowBundleFirst);

  const rowPriceHeader = findRowIndexByLabel_(srcValues, '価格関連（税込）');
  if (rowCompanyId < 0 || rowCourseName < 0 || rowCourseId < 0 || rowPriceHeader < 0) {
    ui.alert('「会社ID」「コース名」「商品コード」「価格関連（税込）」のいずれかの行が見つかりません。');
    return;
  }

  // 価格ブロック内の行（メインのみ側）を見つける
  const rowMainFirstPrice  = findRowIndexByLabelAfter_(srcValues, '初回', rowPriceHeader);
  const rowMainShip1       = findRowIndexByLabelAfter_(srcValues, '送料／後払手数料', rowMainFirstPrice);
  const rowMainSecondPrice = findRowIndexByLabelAfter_(srcValues, '定期2回目以降', rowPriceHeader);
  const rowMainShip2       = findRowIndexByLabelAfter_(srcValues, '送料／後払手数料', rowMainSecondPrice);

  const rowDiffCode   = findRowIndexByLabelAfter_(srcValues, '差額用商品コード', rowPriceHeader);
  const rowDiffAmount = findRowIndexByLabelAfter_(srcValues, '差額', rowPriceHeader);
  const rowDiffFee    = findRowIndexByLabelAfter_(srcValues, '後払手数料', rowPriceHeader);

  // CS対応関連ブロック
  const rowCoolingOff = findRowIndexByLabel_(srcValues, 'クーリングオフ');
  const rowRefund     = findRowIndexByLabel_(srcValues, '返金保証');
  const rowPause      = findRowIndexByLabel_(srcValues, '休止');
  const rowSkip       = findRowIndexByLabel_(srcValues, 'スキップ');
  const rowCancel1    = findRowIndexByLabel_(srcValues, '初回解約');
  const rowCancel2    = findRowIndexByLabel_(srcValues, '定期2回目以降の解約');

  // ---- contract_master / logic_rules 側の course_id マップを構築 ----
  const masterMap = buildCourseRowMap_(masterSheet, 6); // G列: course_id
  const logicMap  = buildCourseRowMap_(logicSheet, 2);  // C列: course_id

  const nowStr         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const lastColMaster  = masterSheet.getLastColumn();
  const lastColLogic   = logicSheet.getLastColumn();
  const masterNewRows  = [];
  const logicNewRows   = [];

  // ---- B列以降の各コースをループ ----
  for (let col = 1; col < numCols; col++) {
    const companyId  = safeStr_(srcValues[rowCompanyId]?.[col]);
    const courseName = safeStr_(srcValues[rowCourseName]?.[col]);
    const courseId   = safeStr_(srcValues[rowCourseId]?.[col]);

    if (!courseId || !companyId) {
      continue; // コースID or 会社IDがなければスキップ
    }

    // ===== contract_master 向け情報 =====
    const memoText     = (rowMemo >= 0) ? safeStr_(srcValues[rowMemo][col]) : '';
    const cycleFirst   = (rowCycleFirst  >= 0) ? safeStr_(srcValues[rowCycleFirst][col])  : '';
    const cycleSecond  = (rowCycleSecond >= 0) ? safeStr_(srcValues[rowCycleSecond][col]) : '';

    const priceCell1   = (rowMainFirstPrice  >= 0) ? srcValues[rowMainFirstPrice][col]  : '';
    const priceCell2   = (rowMainSecondPrice >= 0) ? srcValues[rowMainSecondPrice][col] : '';

    const priceInfo    = extractPriceInfoFromSource_v2_(priceCell1, priceCell2);
    const fulfillRule  = inferFulfillmentRule_(cycleFirst, cycleSecond);
    const productBundle = buildProductBundleText_v2_(
      srcValues,
      rowBundleFirst,
      rowBundleSecond,
      rowBundleThird,
      col
    );

    // 回数縛りの推定（備考から）
    const commitInfo = inferCommitRuleFromMemo_(memoText);

    // 返金保証（あり / なし）
    const refundText = (rowRefund >= 0) ? safeStr_(srcValues[rowRefund][col]) : '';

    const guaranteeType = inferGuaranteeTypeFromText_(refundText);

    // ===== contract_logic_rules 向け情報 =====
    const coText      = (rowCoolingOff >= 0) ? safeStr_(srcValues[rowCoolingOff][col]) : '';
    const pauseText   = (rowPause      >= 0) ? safeStr_(srcValues[rowPause][col])      : '';
    const skipText    = (rowSkip       >= 0) ? safeStr_(srcValues[rowSkip][col])       : '';
    const cancel1Text = (rowCancel1    >= 0) ? safeStr_(srcValues[rowCancel1][col])    : '';
    const cancel2Text = (rowCancel2    >= 0) ? safeStr_(srcValues[rowCancel2][col])    : '';

    const refundFlagAndDetail     = inferRefundFlag_(refundText);
    const coolingOffFlagAndDetail = inferCoolingOffFlag_(coText);

    // ---- contract_master 更新 or 追加 ----
    const masterKey      = companyId + '||' + courseId;
    const masterRowIndex = masterMap[masterKey];  // 既存行の行番号（3行目以降）

    if (masterRowIndex) {
      // ★ appendOnly の場合は「既存行は触らない」
      if (mode === 'appendOnly') {
        // 何もしない（このコースはスキップ）
      } else {
        // ★ overwrite モード：既存行を上書き
        const row = masterSheet.getRange(masterRowIndex, 1, 1, lastColMaster).getValues()[0];

        row[0]  = nowStr;        // last_updated (A)
        row[1]  = companyId;     // client_company_id (B)
        row[3]  = courseName;    // course_name (D)
        row[6]  = courseId;      // course_id (G)

        // fulfillment_rule（M）
        if (fulfillRule) row[12] = fulfillRule;

        // 価格（Q, R, S）
        if (priceInfo.firstPrice !== '')      row[16] = priceInfo.firstPrice;
        if (priceInfo.secondPrice !== '')     row[17] = priceInfo.secondPrice;
        if (priceInfo.recurringPrice !== '')  row[18] = priceInfo.recurringPrice;

        // 商品構成（T）
        if (productBundle) row[19] = productBundle;

        // commit_rule / first_commit_count / total_commit_count
        if (commitInfo.commitRule) {
          row[8]  = commitInfo.commitRule;             // commit_rule (I)
          row[21] = commitInfo.firstCommitCount || ''; // first_commit_count (V)
          row[22] = commitInfo.totalCommitCount || ''; // total_commit_count (W)
        }

        // guarantee_type (K)
        if (guaranteeType) {
          row[10] = guaranteeType;
        }

        // billing_interval（空なら per_order）
        if (!row[20]) {
          row[20] = 'per_order';
        }

        // remarks（備考）: すでに何か入っていれば上書きしない運用もそのまま維持
        if (!safeStr_(row[32]) && memoText) {
          row[32] = memoText;
        }

        masterSheet.getRange(masterRowIndex, 1, 1, lastColMaster).setValues([row]);
      }

    } else {
      // ★ 新規行は mode に関係なく追加
      const newRow = new Array(lastColMaster).fill('');

      newRow[0]  = nowStr;        // last_updated
      newRow[1]  = companyId;     // client_company_id
      newRow[3]  = courseName;    // course_name
      newRow[6]  = courseId;      // course_id

      if (fulfillRule) newRow[12] = fulfillRule;

      newRow[16] = priceInfo.firstPrice;
      newRow[17] = priceInfo.secondPrice;
      newRow[18] = priceInfo.recurringPrice;
      newRow[19] = productBundle;

      if (commitInfo.commitRule) {
        newRow[8]  = commitInfo.commitRule;
        newRow[21] = commitInfo.firstCommitCount || '';
        newRow[22] = commitInfo.totalCommitCount || '';
      }

      if (guaranteeType) {
        newRow[10] = guaranteeType;
      }

      newRow[20] = 'per_order';  // billing_interval
      if (memoText) {
        newRow[32] = memoText;   // remarks
      }

      masterNewRows.push(newRow);
      masterMap[masterKey] = masterSheet.getLastRow() + masterNewRows.length;
    }

    // ---- contract_logic_rules 更新 or 追加 ----
    const logicKey      = companyId + '||' + courseId;
    const logicRowIndex = logicMap[logicKey];

    if (logicRowIndex) {
      if (mode === 'appendOnly') {
        // 既存行は触らない
      } else {
        const rowL = logicSheet.getRange(logicRowIndex, 1, 1, lastColLogic).getValues()[0];

        rowL[0] = nowStr;      // last_updated
        rowL[1] = companyId;   // client_company_id
        rowL[2] = courseId;    // course_id

        // 解約受付期限（フェーズ別）
        if (cancel1Text || cancel2Text) {
          rowL[3] = 'variable_by_phase'; // cancel_deadline
          rowL[4] = `初回: ${cancel1Text} / 2回目以降: ${cancel2Text}`; // cancel_deadline_logic
        }

        if (skipText)  rowL[7] = skipText;   // skip_rule
        if (pauseText) rowL[8] = pauseText;  // long_pause_rule

        // 返金保証
        if (refundFlagAndDetail.flag !== '') {
          rowL[18] = refundFlagAndDetail.flag;   // refund_guarantee_flag (S)
        }
        if (refundFlagAndDetail.detail) {
          rowL[20] = refundFlagAndDetail.detail; // refund_guarantee_condition_detail (U)
        }

        // クーリングオフ
        if (coolingOffFlagAndDetail.flag !== '') {
          rowL[24] = coolingOffFlagAndDetail.flag; // cooling_off_flag (Z)
        }
        if (coolingOffFlagAndDetail.detail) {
          rowL[28] = coolingOffFlagAndDetail.detail; // cooling_off_condition_detail (AC)
        }

        logicSheet.getRange(logicRowIndex, 1, 1, lastColLogic).setValues([rowL]);
      }

    } else {
      // 新規行追加
      const newRowL = new Array(lastColLogic).fill('');

      newRowL[0] = nowStr;
      newRowL[1] = companyId;
      newRowL[2] = courseId;

      if (cancel1Text || cancel2Text) {
        newRowL[3] = 'variable_by_phase';
        newRowL[4] = `初回: ${cancel1Text} / 2回目以降: ${cancel2Text}`;
      }

      if (skipText)  newRowL[7] = skipText;
      if (pauseText) newRowL[8] = pauseText;

      if (refundFlagAndDetail.flag !== '') {
        newRowL[18] = refundFlagAndDetail.flag;
      }
      if (refundFlagAndDetail.detail) {
        newRowL[20] = refundFlagAndDetail.detail;
      }

      if (coolingOffFlagAndDetail.flag !== '') {
        newRowL[24] = coolingOffFlagAndDetail.flag;
      }
      if (coolingOffFlagAndDetail.detail) {
        newRowL[28] = coolingOffFlagAndDetail.detail;
      }

      logicNewRows.push(newRowL);
      logicMap[logicKey] = logicSheet.getLastRow() + logicNewRows.length;
    }
  }

  if (masterNewRows.length > 0) {
    masterSheet.getRange(masterSheet.getLastRow() + 1, 1, masterNewRows.length, lastColMaster)
      .setValues(masterNewRows);
  }
  if (logicNewRows.length > 0) {
    logicSheet.getRange(logicSheet.getLastRow() + 1, 1, logicNewRows.length, lastColLogic)
      .setValues(logicNewRows);
  }

  ui.alert(
    'course_master_source から contract_master / contract_logic_rules への反映が完了しました（v2, モード: ' +
    mode +
    '）。'
  );
}


/* ===== ヘルパー関数群 ===== */

// A列ラベル完全一致で行番号を返す（見つからなければ -1）
function findRowIndexByLabel_(values, label) {
  for (let r = 0; r < values.length; r++) {
    if (safeStr_(values[r][0]) === label) {
      return r;
    }
  }
  return -1;
}

// 指定行より下で、最初に label に一致する行を探す
function findRowIndexByLabelAfter_(values, label, startRow) {
  if (startRow < 0) return -1;
  for (let r = startRow + 1; r < values.length; r++) {
    if (safeStr_(values[r][0]) === label) {
      return r;
    }
  }
  return -1;
}

// null/undefined → '' にして trim
function safeStr_(v) {
  if (v == null) return '';
  return String(v).trim();
}

// 価格文字列から first / second / recurring を抽出
function extractPriceInfoFromSource_v2_(cellInitial, cellSecond) {
  const firstPrice = parseNumberOrEmpty_(cellInitial);

  let secondPrice    = '';
  let recurringPrice = '';

  const textSecond = safeStr_(cellSecond);
  if (!textSecond) {
    return { firstPrice, secondPrice, recurringPrice };
  }

  // "2回目：4,990円\n3回目以降：9,980円" のようなパターン
  if (textSecond.indexOf('2回目') !== -1 && textSecond.indexOf('3回目以降') !== -1) {
    const lines = textSecond.split(/\r?\n/);
    lines.forEach(line => {
      if (line.indexOf('2回目') !== -1) {
        const p = parseNumberOrEmpty_(line);
        if (p !== '') secondPrice = p;
      } else if (line.indexOf('3回目以降') !== -1) {
        const p = parseNumberOrEmpty_(line);
        if (p !== '') recurringPrice = p;
      }
    });
  } else {
    // 「定期2回目以降, 9,980円」のような単一値の場合 → recurringPrice として扱う
    const p = parseNumberOrEmpty_(textSecond);
    if (p !== '') recurringPrice = p;
  }

  return { firstPrice, secondPrice, recurringPrice };
}

// "BEAST 1袋" + "BEAST 2袋" + "BEAST 2袋" をまとめてテキスト化
function buildProductBundleText_v2_(values, rowFirst, rowSecond, rowThird, col) {
  const parts = [];
  if (rowFirst >= 0) {
    const v = safeStr_(values[rowFirst][col]);
    if (v) parts.push(`初回: ${v}`);
  }
  if (rowSecond >= 0) {
    const v = safeStr_(values[rowSecond][col]);
    if (v) parts.push(`2回目: ${v}`);
  }
  if (rowThird >= 0) {
    const v = safeStr_(values[rowThird][col]);
    if (v) parts.push(`3回目以降: ${v}`);
  }
  return parts.join(' / ');
}

// "14日" "30日" → cycle_initial14_then30 等に変換
function inferFulfillmentRule_(cycleFirstText, cycleSecondText) {
  const d1 = parseInt(cycleFirstText.replace(/[^\d]/g, ''), 10);
  const d2 = parseInt(cycleSecondText.replace(/[^\d]/g, ''), 10);
  if (isNaN(d1) || isNaN(d2)) return '';

  if (d1 === 14 && d2 === 30) {
    return 'cycle_initial14_then30';
  }
  if (d1 === d2) {
    switch (d1) {
      case 14:  return 'cycle_14days';
      case 30:  return 'cycle_30days';
      case 90:  return 'cycle_90days';
      case 180: return 'cycle_180days';
      case 360: return 'cycle_360days';
      default:  return '';
    }
  }
  return ''; // 特殊なパターンは手入力想定
}

// 備考から回数縛りを推定（「○回受取後〜」など）
function inferCommitRuleFromMemo_(memoText) {
  const result = {
    commitRule: '',
    firstCommitCount: '',
    totalCommitCount: ''
  };
  if (!memoText) return result;

  const m = memoText.match(/(\d+)回/);
  if (!m) return result;

  const n = parseInt(m[1], 10);
  if (isNaN(n)) return result;

  if (n >= 2 && n <= 7) {
    result.commitRule       = `commit_${n}`;
    result.firstCommitCount = n;
    result.totalCommitCount = n;
  } else if (n > 7) {
    result.commitRule       = 'commit_multi';
    result.firstCommitCount = n;
    result.totalCommitCount = n;
  }
  return result;
}

// "1,980円" などから数字部分だけ抜き、Number化して返す（失敗時は ''）
function parseNumberOrEmpty_(value) {
  if (value == null || value === '') return '';
  const text = String(value);
  const numStr = text.replace(/[^\d.-]/g, '');
  if (!numStr) return '';
  const num = Number(numStr);
  if (isNaN(num)) return '';
  return num;
}

// 保証種別を返金保証のテキストからざっくり推定
function inferGuaranteeTypeFromText_(text) {
  if (!text) return '';
  if (text.indexOf('なし') !== -1) return 'none';
  if (text.indexOf('あり') !== -1) return 'refund';
  return ''; // その他は手入力
}

// 返金保証フラグ＋詳細文
function inferRefundFlag_(text) {
  if (!text) return { flag: '', detail: '' };
  let flag = '';
  if (text.indexOf('なし') !== -1) flag = 'FALSE';
  if (text.indexOf('あり') !== -1) flag = 'TRUE';
  return { flag, detail: text };
}

// クーリングオフフラグ＋詳細文
function inferCoolingOffFlag_(text) {
  if (!text) return { flag: '', detail: '' };
  let flag = '';
  if (text.indexOf('なし') !== -1) flag = 'FALSE';
  if (text.indexOf('あり') !== -1) flag = 'TRUE';
  return { flag, detail: text };
}

/**
 * 指定シートの course_id 列から {companyId||courseId: rowIndex} マップを作る
 * colCourseId: 0始まりインデックス
 */
function buildCourseRowMap_(sheet, colCourseId) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const map = {};
  if (lastRow < 3) return map;

  const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const companyId = safeStr_(row[1]); // B列=client_company_id
    const courseId  = safeStr_(row[colCourseId]);
    if (!courseId || !companyId) continue;
    const key = companyId + '||' + courseId;
    map[key] = i + 3; // 実際の行番号（3行目〜）
  }
  return map;
}



/**************************************************
 * コース表（course_master_source） → fee_table_master 自動反映
 *
 * 想定シート構成:
 * - course_master_source:
 *   A列: 項目名
 *   B列以降: 各コース列
 *   行のどこかに「会社ID」行がある（例: 先頭行）
 *
 * - fee_table_master:
 *   1行目: 和名ヘッダ
 *   2行目: 英名ヘッダ
 *   3行目以降: データ
 **************************************************/





/**
 * メイン処理:
 * course_master_source から fee_table_master に行を追加する
 */
function updateFeeTableFromCourseSheet() {
  const ss = SpreadsheetApp.getActive();
  const srcSheet = ss.getSheetByName(SHEET_COURSE_SOURCE);
  const feeSheet = ss.getSheetByName(SHEET_FEE_TABLE);
  const ui = SpreadsheetApp.getUi();

  if (!srcSheet) {
    ui.alert('course_master_source シートが見つかりません。');
    return;
  }
  if (!feeSheet) {
    ui.alert('fee_table_master シートが見つかりません。');
    return;
  }

  const srcValues = srcSheet.getDataRange().getValues();
  const numRows   = srcValues.length;
  const numCols   = srcValues[0].length;

  if (numRows < 2 || numCols < 2) {
    ui.alert('course_master_source に有効なデータがありません。');
    return;
  }

  // --- 0) 会社ID行のインデックスを探す
  const rowCompanyId = findRowIndexByLabel_(srcValues, '会社ID');
  if (rowCompanyId < 0) {
    ui.alert('course_master_source 内に「会社ID」行が見つかりませんでした。');
    return;
  }

  // --- 1) 「価格関連（税込）」ブロック内の各行インデックスを特定する
  const idxMap = findPriceSectionRowIndexes_(srcValues);
  if (!idxMap) {
    ui.alert('course_master_source 内で「価格関連（税込）」ブロックが見つかりませんでした。');
    return;
  }

  const rowCourseName      = idxMap.rowCourseName;
  const rowCourseId        = idxMap.rowCourseId;
  const rowShip1           = idxMap.rowShipInitial;
  const rowShip2           = idxMap.rowShipSecond;
  const rowDiffItem        = idxMap.rowDiffItem;
  const rowDiffAmount      = idxMap.rowDiffAmount;
  const rowDiffFee         = idxMap.rowDiffFee;

  // --- 2) fee_table_master の既存行を読み込み、キーセットを作成
  const lastRowFee = feeSheet.getLastRow();
  const lastColFee = feeSheet.getLastColumn();

  let existingValues = [];
  let existingKeySet = new Set();

  if (lastRowFee >= 3) {
    existingValues = feeSheet.getRange(3, 1, lastRowFee - 2, lastColFee).getValues();
    existingValues.forEach(row => {
      const companyId   = row[0]; // client_company_id
      const courseId    = row[1]; // course_id
      const paymentType = row[2]; // payment_type
      const orderCount  = row[3]; // order_count
      const region      = row[4]; // region

      if (!courseId) return;
      const key = buildFeeKey_(companyId, courseId, paymentType, orderCount, region);
      existingKeySet.add(key);
    });
  }

  const newRows = [];

  // --- 3) コース表の各コース列を走査（B列以降）
  for (let col = 1; col < numCols; col++) {
    const companyId = String(srcValues[rowCompanyId][col] || '').trim();
    const courseName = String(srcValues[rowCourseName][col] || '').trim();
    const courseId   = String(srcValues[rowCourseId][col]   || '').trim();

    // コースIDがない列はスキップ
    if (!courseId) continue;

    // （会社IDが空のコースは一旦スキップする設計にします）
    if (!companyId) continue;

    // 初回「送料／後払手数料」
    const shipCell1 = srcValues[rowShip1][col];
    const feePair1  = parseShippingAndFee_(shipCell1);

    // 定期2回目以降「送料／後払手数料」
    const shipCell2 = srcValues[rowShip2][col];
    const feePair2  = parseShippingAndFee_(shipCell2);

    // 差額用商品コード & 差額 & 差額用決済手数料
    const diffItemCell   = srcValues[rowDiffItem][col];
    const diffAmountCell = srcValues[rowDiffAmount][col];
    const diffFeeCell    = srcValues[rowDiffFee][col];

    const diffMap        = parseDiffInfo_(diffItemCell, diffAmountCell); // {1: {code, amount}, 2: {...}}
    const diffPaymentFee = parseNumberOrEmpty_(diffFeeCell);

    // --- order_count = 1（初回）行を作成
    {
      const orderCount = 1;
      const shippingFee = feePair1.shipping;
      const paymentFee  = feePair1.fee;

      const diffInfo = diffMap[1] || null;
      const diffItem   = diffInfo ? diffInfo.code   : '';
      const diffAmount = diffInfo ? diffInfo.amount : '';

      const row = buildFeeRow_({
        companyId,
        courseId,
        paymentType: '',        // 今はまとめて空運用（将来、payment_type 別に展開してもOK）
        orderCount,
        region: DEFAULT_REGION,
        shippingFee,
        paymentFee,
        discountAmount: '',
        diffItemCode: diffItem,
        diffAmount,
        diffPaymentFee: diffAmount !== '' ? diffPaymentFee : '',
        remarks: ''
      });

      const key = buildFeeKey_(
        row[0], row[1], row[2], row[3], row[4]
      );

      if (!existingKeySet.has(key)) {
        newRows.push(row);
        existingKeySet.add(key);
      }
    }

    // --- order_count = 2（定期2回目以降）行を作成
    {
      const orderCount = 2;
      const shippingFee = feePair2.shipping;
      const paymentFee  = feePair2.fee;

      const diffInfo = diffMap[2] || null;
      const diffItem   = diffInfo ? diffInfo.code   : '';
      const diffAmount = diffInfo ? diffInfo.amount : '';

      const row = buildFeeRow_({
        companyId,
        courseId,
        paymentType: '',
        orderCount,
        region: DEFAULT_REGION,
        shippingFee,
        paymentFee,
        discountAmount: '',
        diffItemCode: diffItem,
        diffAmount,
        diffPaymentFee: diffAmount !== '' ? diffPaymentFee : '',
        remarks: ''
      });

      const key = buildFeeKey_(
        row[0], row[1], row[2], row[3], row[4]
      );

      if (!existingKeySet.has(key)) {
        newRows.push(row);
        existingKeySet.add(key);
      }
    }
  }

  // --- 4) 追加行を書き込む
  if (newRows.length > 0) {
    feeSheet.getRange(feeSheet.getLastRow() + 1, 1, newRows.length, lastColFee).setValues(newRows);
    ui.alert('fee_table_master に ' + newRows.length + ' 行を追加しました。');
  } else {
    ui.alert('追加対象となる新しい行はありませんでした。');
  }
}


/**
 * 「価格関連（税込）」ブロック内の各行インデックスを特定する
 * 返り値:
 * {
 *   rowCourseName,
 *   rowCourseId,
 *   rowShipInitial,
 *   rowShipSecond,
 *   rowDiffItem,
 *   rowDiffAmount,
 *   rowDiffFee
 * }
 */
function findPriceSectionRowIndexes_(values) {
  const numRows = values.length;

  let rowCourseName   = -1;
  let rowCourseId     = -1;
  let rowPriceHeader  = -1;

  for (let r = 0; r < numRows; r++) {
    const label = String(values[r][0] || '').trim();
    if (label === 'コース名') {
      rowCourseName = r;
    } else if (label === '商品コード') {
      rowCourseId = r;
    } else if (label === '価格関連（税込）') {
      rowPriceHeader = r;
      break;
    }
  }

  if (rowCourseName < 0 || rowCourseId < 0 || rowPriceHeader < 0) {
    return null;
  }

  let rowShipInitial = -1;
  let rowShipSecond  = -1;
  let rowDiffItem    = -1;
  let rowDiffAmount  = -1;
  let rowDiffFee     = -1;

  for (let r = rowPriceHeader + 1; r < numRows; r++) {
    const label = String(values[r][0] || '').trim();

    if (label === '送料／後払手数料') {
      if (rowShipInitial < 0) {
        rowShipInitial = r;  // 初回の直後の送料／後払手数料
      } else if (rowShipSecond < 0) {
        rowShipSecond = r;   // 定期2回目以降の送料／後払手数料
      }
    } else if (label === '差額用商品コード') {
      rowDiffItem = r;
    } else if (label === '差額') {
      rowDiffAmount = r;
    } else if (label === '後払手数料') {
      rowDiffFee = r;
    }
  }

  if (
    rowShipInitial < 0 ||
    rowShipSecond  < 0 ||
    rowDiffItem    < 0 ||
    rowDiffAmount  < 0 ||
    rowDiffFee     < 0
  ) {
    return null;
  }

  return {
    rowCourseName,
    rowCourseId,
    rowShipInitial,
    rowShipSecond,
    rowDiffItem,
    rowDiffAmount,
    rowDiffFee
  };
}

/**
 * "500円／330円" のような文字列から送料・決済手数料をパース
 * @return {{shipping: string|number, fee: string|number}}
 */
function parseShippingAndFee_(cellValue) {
  if (cellValue === null || cellValue === '') {
    return { shipping: '', fee: '' };
  }
  const text = String(cellValue);

  // "500円／330円" or "500/330" などを想定
  const parts = text.split(/[／\/]/);
  const shipping = parts[0] != null ? parseNumberOrEmpty_(parts[0]) : '';
  const fee      = parts[1] != null ? parseNumberOrEmpty_(parts[1]) : '';

  return { shipping, fee };
}

/**
 * 差額用商品コード & 差額 を order_count ごとのマップに変換
 */
function parseDiffInfo_(diffItemCell, diffAmountCell) {
  const result = {};

  const codes = String(diffItemCell || '')
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(s => s !== '');

  const amountLines = String(diffAmountCell || '')
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(s => s !== '');

  if (codes.length === 0 || amountLines.length === 0) {
    return result;
  }

  // ラベルなし1行だけ → order_count=1 に紐付け
  const hasLabelWord = amountLines.some(line =>
    line.indexOf('初回') !== -1 ||
    line.indexOf('1回目') !== -1 ||
    line.indexOf('2回目') !== -1 ||
    line.indexOf('3回目') !== -1
  );

  if (!hasLabelWord && amountLines.length === 1) {
    const amt = parseNumberOrEmpty_(amountLines[0]);
    if (amt !== '') {
      result[1] = {
        code: codes[0] || '',
        amount: amt
      };
    }
    return result;
  }

  // "初回：8,000円" / "2回目：4,990円" のようなパターン
  amountLines.forEach((line, i) => {
    let orderCount = null;
    if (line.indexOf('初回') !== -1 || line.indexOf('1回目') !== -1) {
      orderCount = 1;
    } else if (line.indexOf('2回目') !== -1) {
      orderCount = 2;
    } else if (line.indexOf('3回目') !== -1) {
      orderCount = 3;
    }

    if (orderCount == null) return;

    const amt = parseNumberOrEmpty_(line);
    if (amt === '') return;

    const code = codes[i] || codes[0] || '';

    result[orderCount] = {
      code,
      amount: amt
    };
  });

  return result;
}


/**
 * fee_table_master の1行分を配列で構築
 */
function buildFeeRow_(opts) {
  return [
    opts.companyId || '',
    opts.courseId || '',
    opts.paymentType || '',
    opts.orderCount || '',
    opts.region || '',
    opts.shippingFee !== undefined ? opts.shippingFee : '',
    opts.paymentFee !== undefined ? opts.paymentFee : '',
    opts.discountAmount !== undefined ? opts.discountAmount : '',
    opts.diffItemCode !== undefined ? opts.diffItemCode : '',
    opts.diffAmount !== undefined ? opts.diffAmount : '',
    opts.diffPaymentFee !== undefined ? opts.diffPaymentFee : '',
    opts.remarks !== undefined ? opts.remarks : ''
  ];
}

/**
 * 重複判定用のキーを生成
 */
function buildFeeKey_(companyId, courseId, paymentType, orderCount, region) {
  return [
    String(companyId || ''),
    String(courseId || ''),
    String(paymentType || ''),
    String(orderCount || ''),
    String(region || '')
  ].join('||');
}

/**************************************************
 * fee_table_master の「差額用〜」列を
 * course_master_source から補完する
 *
 * - 既存の updateFeeTableFromCourseSheet で
 *   ベース行を作ったあとに実行する想定
 **************************************************/



function fillFeeTableDiffFromCourseSource() {
  const ss         = SpreadsheetApp.getActive();
  const srcSheet   = ss.getSheetByName(SHEET_COURSE_SOURCE_FEE);
  const feeSheet   = ss.getSheetByName(SHEET_FEE_TABLE);
  const ui         = SpreadsheetApp.getUi();

  if (!srcSheet || !feeSheet) {
    ui.alert('course_master_source または fee_table_master シートが見つかりません。');
    return;
  }

  const srcValues = srcSheet.getDataRange().getValues();
  const numRows   = srcValues.length;
  const numCols   = srcValues[0].length;

  if (numRows < 2 || numCols < 2) {
    ui.alert('course_master_source に有効なデータがありません。');
    return;
  }

  // ==== course_master_source 側：必要な行インデックスを取得 ====
  const rowCompanyId   = findRowIndexByLabel_(srcValues, '会社ID');
  const rowCourseId    = findRowIndexByLabel_(srcValues, '商品コード');
  const rowMemo        = findRowIndexByLabel_(srcValues, '備考');

  const rowDiffCode    = findRowIndexByLabel_(srcValues, '差額用商品コード');
  const rowDiffAmount  = findRowIndexByLabel_(srcValues, '差額');
  // 「差額」の直後の「後払手数料」を差額用手数料とみなす
  const rowDiffFee     = findRowIndexByLabelAfter_(srcValues, '後払手数料', rowDiffAmount);

  if (rowCompanyId < 0 || rowCourseId < 0 || rowDiffCode < 0 || rowDiffAmount < 0) {
    ui.alert('course_master_source 内に「会社ID」「商品コード」「差額用商品コード」「差額」行が見つかりません。');
    return;
  }

  // ==== course_master_source -> diffInfoMap を作成 ====
  /**
   * diffInfoMap[key] = {
   *   diffProductCode: "sagaku_1;sagaku_6",
   *   diffAmountByOrderCount: {1:8000, 2:4990},
   *   diffPaymentFee: 330,
   *   memo: "備考テキスト"
   * }
   */
  const diffInfoMap = {};

  for (let col = 1; col < numCols; col++) {
    const companyId = safeStr_(srcValues[rowCompanyId][col]);
    const courseId  = safeStr_(srcValues[rowCourseId][col]);
    if (!companyId || !courseId) continue;

    const key = companyId + '||' + courseId;

    // 差額用商品コード（改行で複数 → セミコロン区切り）
    const diffCodeRaw = safeStr_(srcValues[rowDiffCode][col]);
    let diffProductCode = '';
    if (diffCodeRaw) {
      const codes = diffCodeRaw.split(/\r?\n/).map(c => safeStr_(c)).filter(c => c);
      diffProductCode = codes.join(';');
    }

    // 差額金額（「初回:8,000円\n2回目:4,990円」など）
    const diffAmountRaw = safeStr_(srcValues[rowDiffAmount][col]);
    const diffAmountByOrderCount = {};  // {1: 8000, 2: 4990 ...}
    if (diffAmountRaw) {
      const lines = diffAmountRaw.split(/\r?\n/);
      lines.forEach(line => {
        const t = safeStr_(line);
        if (!t) return;
        let oc = null;
        if (t.indexOf('初回') !== -1) {
          oc = 1;
        } else {
          const m = t.match(/(\d+)回目/);
          if (m) {
            oc = parseInt(m[1], 10);
          }
        }
        if (oc != null && !isNaN(oc)) {
          const amount = parseNumberOrEmpty_(t);
          if (amount !== '') {
            diffAmountByOrderCount[oc] = amount;
          }
        }
      });
    }

    // 差額用決済手数料
    let diffPaymentFee = '';
    if (rowDiffFee >= 0) {
      diffPaymentFee = parseNumberOrEmpty_(srcValues[rowDiffFee][col]);
    }

    // 備考
    const memoText = (rowMemo >= 0) ? safeStr_(srcValues[rowMemo][col]) : '';

    diffInfoMap[key] = {
      diffProductCode,
      diffAmountByOrderCount,
      diffPaymentFee,
      memo: memoText
    };
  }

  // ==== fee_table_master 側を更新 ====
  const lastRow = feeSheet.getLastRow();
  const lastCol = feeSheet.getLastColumn();
  if (lastRow < 3) {
    ui.alert('fee_table_master にデータ行がありません（3行目以降）。');
    return;
  }

  const feeValues = feeSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  // 列インデックス（0始まり）を明示
  const COL_COMPANY_ID      = 0;  // A: client_company_id
  const COL_COURSE_ID       = 1;  // B: course_id
  const COL_ORDER_COUNT     = 3;  // D: order_count
  const COL_DIFF_PRODUCT    = 8;  // I: diff_product_code
  const COL_DIFF_AMOUNT     = 9;  // J: diff_amount
  const COL_DIFF_PAYMENT    = 10; // K: diff_payment_fee
  const COL_REMARKS         = 11; // L: remarks

  for (let i = 0; i < feeValues.length; i++) {
    const row = feeValues[i];
    const companyId = safeStr_(row[COL_COMPANY_ID]);
    const courseId  = safeStr_(row[COL_COURSE_ID]);
    if (!companyId || !courseId) continue;

    const key = companyId + '||' + courseId;
    const info = diffInfoMap[key];
    if (!info) continue;

    const orderCountVal = row[COL_ORDER_COUNT];
    const orderCount = parseInt(orderCountVal, 10);

    // 差額用商品コード（全行共通で同じ値でOKとする）
    if (info.diffProductCode) {
      row[COL_DIFF_PRODUCT] = info.diffProductCode;
    }

    // 差額金額（受取回数ごとに変える）
    if (!isNaN(orderCount) && info.diffAmountByOrderCount[orderCount] != null) {
      row[COL_DIFF_AMOUNT] = info.diffAmountByOrderCount[orderCount];
    }

    // 差額用決済手数料（差額が発生する受取回だけ入れる運用でも、全受取回に入れる運用でもOK）
    if (info.diffPaymentFee !== '') {
      // ここでは「差額金額がある行」にだけ入れる方針にする
      if (!isNaN(orderCount) && info.diffAmountByOrderCount[orderCount] != null) {
        row[COL_DIFF_PAYMENT] = info.diffPaymentFee;
      }
    }

    // 備考：fee_table_master 側が空欄なら course_master_source の備考を入れる
    if (!safeStr_(row[COL_REMARKS]) && info.memo) {
      row[COL_REMARKS] = info.memo;
    }

    feeValues[i] = row;
  }

  // 反映
  feeSheet.getRange(3, 1, feeValues.length, lastCol).setValues(feeValues);

  ui.alert('fee_table_master の差額用項目（商品コード / 金額 / 決済手数料 / 備考）を course_master_source から補完しました。');
}


