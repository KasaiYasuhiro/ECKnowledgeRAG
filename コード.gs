

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ナレッジDB');
  if (!sheet) {
    throw new Error('ナレッジDB シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    throw new Error('ナレッジDB にデータがありません');
  }

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

    if (TITLE) textParts.push('【タイトル】\n' + TITLE);
    if (SUMMARY) textParts.push('【概要】\n' + SUMMARY);
    if (CORE) textParts.push('【本文（共通ルール）】\n' + CORE);
    if (DELTA) textParts.push('【差分・例外ルール】\n' + DELTA);
    if (NG) textParts.push('【禁止事項・注意事項】\n' + NG);
    if (REASON) textParts.push('【更新理由】\n' + REASON);
    if (DUE) textParts.push('【希望反映期限】\n' + DUE);
    if (AUTHOR) textParts.push('【登録者】\n' + AUTHOR);

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

  // ← ★ここを修正：formatCsv ではなく toCsv_ を使う
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





/**
 * 2行目の英名ヘッダーを検索し、該当列番号を返す。
 * もし存在しなければ新規列として追加する。
 */
function getOrCreateHeaderColumn_(sheet, headerName) {
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
  const headerValues = headerRange.getValues()[0];

  for (let i = 0; i < headerValues.length; i++) {
    if (headerValues[i] === headerName) {
      return i + 1; // 列番号（1始まり）
    }
  }

  // 無ければヘッダー右側に追加
  const newCol = headerValues.length + 1;
  sheet.getRange(HEADER_ROW, newCol).setValue(headerName);     // ★英名ヘッダーに追加
  sheet.getRange(HEADER_ROW - 1, newCol).setValue(headerName); // ★和名ヘッダーにも仮値を入れておく（任意）
  return newCol;
}


/**
 * last_updated に現在時刻を入れる
 */
function updateLastUpdated_(sheet, row, col) {
  const cell = sheet.getRange(row, col);
  cell.setValue(new Date());
  cell.setNumberFormat('yyyy-mm-dd hh:mm:ss');
}


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
 * course_master_source → contract_master / contract_logic_rules
 * 旧スクリプト方針を踏襲した簡易版
 *
 * - A列：項目名
 * - B列以降：コース列
 **************************************************/


/**
 * メイン：course_master_source から各マスタに自動反映
 * 既存行は触らず「新規コースのみ追加」
 */
function updateContractsFromCourseSource_AppendOnly() {
  updateContractsFromCourseSource_v2_core_('appendOnly');
}

/**
 * メイン：course_master_source から各マスタに自動反映
 * 既存行も含めて上書き
 */
function updateContractsFromCourseSource_Overwrite() {
  updateContractsFromCourseSource_v2_core_('overwrite');
}

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
 * ↓ 旧スクリプト（現在は使用しない） ここから
 *   - メニュー・トリガーからは呼ばれていません
 *   - 仕様参照用に残しています
 **************************************************/
/**************************************************
 * メイン：course_master_source → contract_master / logic_rules
 **************************************************/
function oldimportCoursesFromSource() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('course_master_source');
  const masterSheet = ss.getSheetByName('contract_master');
  const logicSheet  = ss.getSheetByName('contract_logic_rules');

  if (!src) {
    SpreadsheetApp.getUi().alert('course_master_source シートが見つかりません。');
    return;
  }
  if (!masterSheet || !logicSheet) {
    SpreadsheetApp.getUi().alert('contract_master または contract_logic_rules シートが見つかりません。');
    return;
  }

  const values = src.getDataRange().getValues();
  const numRows = values.length;
  const numCols = values[0].length;

  // A列のラベル → 行インデックス マップ
  const rowIndexByLabel = {};
  for (let r = 0; r < numRows; r++) {
    const label = (values[r][0] || '').toString().trim();
    if (label) {
      rowIndexByLabel[label] = r;
    }
  }

  function get(label) {
    const r = rowIndexByLabel[label];
    if (r == null) return null;
    return { row: r };
  }

  // 必要なラベル行の存在チェック
  const requiredLabels = [
    'コース名',
    '商品コード',
    '初回⇒定期2回目',
    '定期2回目以降',
    '初回',
    '定期2回目',
    '定期3回目以降',
    '備考',
    'クーリングオフ',
    '返金保証',
    '休止',
    'スキップ',
    '初回解約',
    '定期2回目以降の解約'
  ];
  const missing = requiredLabels.filter(lbl => rowIndexByLabel[lbl] == null);
  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert(
      'course_master_source に次のラベル行が見つかりませんでした:\n' +
      missing.join(', ')
    );
    return;
  }

  // 各ラベル行のインデックス取得
  const rowCourseName   = rowIndexByLabel['コース名'];
  const rowCourseId     = rowIndexByLabel['商品コード'];
  const rowCycle1       = rowIndexByLabel['初回⇒定期2回目'];
  const rowCycle2       = rowIndexByLabel['定期2回目以降'];
  const rowItemFirst    = rowIndexByLabel['初回'];
  const rowItemSecond   = rowIndexByLabel['定期2回目'];
  const rowItemThird    = rowIndexByLabel['定期3回目以降'];
  const rowRemark       = rowIndexByLabel['備考'];
  const rowCoolingOff   = rowIndexByLabel['クーリングオフ'];
  const rowRefund       = rowIndexByLabel['返金保証'];
  const rowPause        = rowIndexByLabel['休止'];
  const rowSkip         = rowIndexByLabel['スキップ'];
  const rowFirstCancel  = rowIndexByLabel['初回解約'];
  const rowRecurCancel  = rowIndexByLabel['定期2回目以降の解約'];
  const rowPriceInitial = rowIndexByLabel['初回'];               // 価格関連のブロックと名前が被るので注意
  const rowFeeInitial   = rowIndexByLabel['送料／後払手数料'];   // 最初の出現を初回として扱う

  // 簡易的に、「価格関連（税込）」以降の最初の「初回」「送料／後払手数料」を価格ブロックとみなす
  // （もし構造を変えたらここを調整）
  const rowPriceBlockStart = rowIndexByLabel['価格関連（税込）'] || rowPriceInitial;

  // 価格部分の行は、一番最初に現れる「価格関連（税込）」以降の同名ラベルを優先的に使う
  function findRowAfter(label, startRow) {
    let found = null;
    for (let r = startRow; r < numRows; r++) {
      if ((values[r][0] || '').toString().trim() === label) {
        found = r;
        break;
      }
    }
    return found != null ? found : rowIndexByLabel[label];
  }

  const rowPriceInit   = findRowAfter('初回', rowPriceBlockStart);
  const rowPriceInitFee = findRowAfter('送料／後払手数料', rowPriceInit + 1);
  const rowPriceRecur  = findRowAfter('定期2回目以降', rowPriceInitFee + 1);
  const rowPriceRecurFee = findRowAfter('送料／後払手数料', rowPriceRecur + 1);

  // 日付フォーマット
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 列B以降を1コースずつ処理
  for (let col = 1; col < numCols; col++) {
    const courseId = (values[rowCourseId][col] || '').toString().trim();
    const courseName = (values[rowCourseName][col] || '').toString().trim();
    if (!courseId && !courseName) {
      continue; // 完全空コース列はスキップ
    }
    if (!courseId) {
      // course_id がない列はスキップ（メッセージだけ出す）
      Logger.log(`列 ${col + 1} は course_id が空のためスキップしました。コース名: ${courseName}`);
      continue;
    }

    // --- 1) コース表から情報を解析してオブジェクト化 ---
    const course = parseCourseColumn_(
      values,
      {
        col,
        rowCourseName,
        rowCourseId,
        rowCycle1,
        rowCycle2,
        rowItemFirst,
        rowItemSecond,
        rowItemThird,
        rowRemark,
        rowCoolingOff,
        rowRefund,
        rowPause,
        rowSkip,
        rowFirstCancel,
        rowRecurCancel,
        rowPriceInit,
        rowPriceInitFee,
        rowPriceRecur,
        rowPriceRecurFee
      }
    );

    // --- 2) contract_master に反映 ---
    upsertContractMasterFromCourse_(masterSheet, course, nowStr);

    // --- 3) contract_logic_rules に反映 ---
    upsertContractLogicFromCourse_(logicSheet, course, nowStr);

    // --- 4) fee_table_master 用のデータは後続で（別関数で拡張予定） ---
    // upsertFeeTableFromCourse_(feeTableSheet, course, nowStr); // TODO: あとで列構成に合わせて実装
  }

  SpreadsheetApp.getUi().alert('course_master_source から contract_master / contract_logic_rules への取り込みが完了しました。');
}
/**************************************************
 * ↑ 旧スクリプト（現在は使用しない） ここまで
 **************************************************/
/**************************************************
 * 1コース分の列を解析して、汎用オブジェクトにする
 **************************************************/
function parseCourseColumn_(values, ctx) {
  const col = ctx.col;

  function getCell(rowIndex) {
    if (rowIndex == null || rowIndex < 0 || rowIndex >= values.length) return '';
    return (values[rowIndex][col] || '').toString().trim();
  }

  const courseName = getCell(ctx.rowCourseName);
  const courseId   = getCell(ctx.rowCourseId);

  // サイクル（例：初回14日 / 2回目以降30日）
  const cycle1Str = getCell(ctx.rowCycle1); // "14日"
  const cycle2Str = getCell(ctx.rowCycle2); // "30日"
  const cycle1Days = parseInt(cycle1Str.replace('日', ''), 10);
  const cycle2Days = parseInt(cycle2Str.replace('日', ''), 10);
  let fulfillmentRule = '';
  if (cycle1Days === 14 && cycle2Days === 30) {
    fulfillmentRule = 'cycle_initial14_then30';
  } else if (cycle1Days === cycle2Days) {
    // code_master に登録のあるサイクルだけ簡易対応
    const m = {
      14: 'cycle_14days',
      30: 'cycle_30days',
      90: 'cycle_90days',
      180: 'cycle_180days',
      360: 'cycle_360days'
    };
    fulfillmentRule = m[cycle1Days] || 'none';
  } else {
    fulfillmentRule = 'none';
  }

  // 商品構成（初回 / 2回目 / 3回目以降）
  const itemFirst  = getCell(ctx.rowItemFirst);
  const itemSecond = getCell(ctx.rowItemSecond);
  const itemThird  = getCell(ctx.rowItemThird);
  let productBundleParts = [];
  if (itemFirst)  productBundleParts.push(`初回${itemFirst}`);
  if (itemSecond) productBundleParts.push(`2回目${itemSecond}`);
  if (itemThird)  productBundleParts.push(`3回目以降${itemThird}`);
  const productBundle = productBundleParts.join('、');

  // 備考 → commit_rule / first_commit_count / total_commit_count 推定
  const remark = getCell(ctx.rowRemark);
  let commitRule = 'none';
  let firstCommitCount = 0;
  let totalCommitCount = 0;
  if (remark) {
    const m = remark.match(/(\d+)回受[取け取り]/);
    if (m) {
      const n = parseInt(m[1], 10);
      firstCommitCount = n;
      totalCommitCount = n;
      if (n >= 2 && n <= 7) {
        commitRule = `commit_${n}`;
      } else if (n > 7) {
        commitRule = 'commit_multi';
      }
    }
  }

  // 価格（初回 / 2回目特別 / 3回目以降）
  const priceInitialStr = getCell(ctx.rowPriceInit);   // 例: "1,980円"
  const priceRecurStr   = getCell(ctx.rowPriceRecur);  // 例: "9,980円" または "2回目：4,990円\n3回目以降：9,980円"

  function parsePrice(raw) {
    if (!raw) return '';
    const m = raw.match(/([\d,]+)\s*円?/);
    return m ? m[1] : raw;
  }

  let firstPrice = parsePrice(priceInitialStr);
  let secondPrice = '';
  let recurringPrice = '';

  if (priceRecurStr.indexOf('\n') >= 0 || /2回目/.test(priceRecurStr)) {
    // 2回目と3回目以降が分かれているパターン
    const lines = priceRecurStr.split(/\n/);
    lines.forEach(line => {
      if (/2回目/.test(line)) {
        secondPrice = parsePrice(line);
      } else if (/3回目以降/.test(line)) {
        recurringPrice = parsePrice(line);
      }
    });
    if (!recurringPrice && secondPrice) {
      recurringPrice = secondPrice;
    }
  } else {
    // 単一の価格 → 3回目以降価格とみなす
    recurringPrice = parsePrice(priceRecurStr);
  }

  // クーリングオフ / 返金保証 / 休止 / スキップ
  const coolingOffStr = getCell(ctx.rowCoolingOff); // 例: "なし" / "なし　※補足メモあり"
  const refundStr     = getCell(ctx.rowRefund);     // 例: "あり" / "なし"
  const pauseStr      = getCell(ctx.rowPause);
  const skipStr       = getCell(ctx.rowSkip);

  let coolingOffFlag = '';
  let coolingOffDetail = '';
  if (coolingOffStr) {
    if (coolingOffStr.indexOf('なし') >= 0) {
      coolingOffFlag = 'FALSE';
    } else if (coolingOffStr.indexOf('あり') >= 0) {
      coolingOffFlag = 'TRUE';
    }
    coolingOffDetail = coolingOffStr;
  }

  let refundFlag = '';
  if (refundStr) {
    if (refundStr.indexOf('なし') >= 0) {
      refundFlag = 'FALSE';
    } else if (refundStr.indexOf('あり') >= 0) {
      refundFlag = 'TRUE';
    }
  }

  // 解約連絡期限
  const firstCancelText = getCell(ctx.rowFirstCancel); // "定期2回目発送予定日の5日前"
  const recurCancelText = getCell(ctx.rowRecurCancel); // "次回発送予定日の15日前"

  let cancelDeadline = '';
  let cancelDeadlineLogic = '';

  if (firstCancelText && recurCancelText && firstCancelText === recurCancelText) {
    cancelDeadline = firstCancelText;
    cancelDeadlineLogic = 'same_for_all_phases';
  } else if (firstCancelText || recurCancelText) {
    cancelDeadline = 'variable_by_phase';
    const parts = [];
    if (firstCancelText) {
      parts.push(`初回解約: ${firstCancelText}`);
    }
    if (recurCancelText) {
      parts.push(`2回目以降の解約: ${recurCancelText}`);
    }
    cancelDeadlineLogic = parts.join(' / ');
  }

  return {
    courseId,
    courseName,
    fulfillmentRule,
    productBundle,
    commitRule,
    firstCommitCount,
    totalCommitCount,
    firstPrice,
    secondPrice,
    recurringPrice,
    remark,
    coolingOffFlag,
    coolingOffDetail,
    refundFlag,
    pauseStr,
    skipStr,
    cancelDeadline,
    cancelDeadlineLogic
  };
}

/**************************************************
 * contract_master に 1コース分を upsert
 * 既存 course_id があれば上書き、なければ新規追加
 **************************************************/
function upsertContractMasterFromCourse_(sheet, course, nowStr) {
  const lastRow = sheet.getLastRow();
  const startRow = 3; // 1行目:和名, 2行目:英名, 3行目以降データ

  let targetRow = null;
  for (let r = startRow; r <= lastRow; r++) {
    const val = sheet.getRange(r, 7).getValue(); // G列: course_id
    if (val && String(val) === course.courseId) {
      targetRow = r;
      break;
    }
  }
  if (!targetRow) {
    targetRow = lastRow >= startRow ? lastRow + 1 : startRow;
  }

  // 既存行を読み込んで更新（触る列だけ上書き）
  const lastCol = sheet.getLastColumn();
  let rowValues = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];

  function set(colIndex1based, value) {
    const idx = colIndex1based - 1;
    if (idx < rowValues.length) {
      rowValues[idx] = value;
    }
  }

  // A: last_updated
  set(1, nowStr);

  // B: client_company_id → 自動では入れない（手入力前提）
  // C: client_company_name → 自動では入れない

  // D: course_name
  if (course.courseName) set(4, course.courseName);

  // G: course_id
  set(7, course.courseId);

  // I: commit_rule
  if (course.commitRule) set(9, course.commitRule);

  // L: guarantee_type → 返金保証あり/なしから簡易変換（必要に応じて上書き可）
  if (course.refundFlag === 'TRUE') {
    set(11, 'refund');
  } else if (course.refundFlag === 'FALSE') {
    set(11, 'none');
  }

  // M: sales_channels → 自動では入れない（web固定などを避ける）

  // N: fulfillment_rule
  if (course.fulfillmentRule) set(13, course.fulfillmentRule);

  // O: payment_type → 自動では入れない（NO と指定あり）
  // P: payment_method_category → payment_type から別ロジックで自動（既存スクリプトを使用）

  // Q: installment_available → 自動では入れない

  // R: first_price
  if (course.firstPrice) set(17, course.firstPrice);

  // S: second_price
  if (course.secondPrice) set(18, course.secondPrice);

  // T: recurring_price
  if (course.recurringPrice) set(19, course.recurringPrice);

  // U: product_bundle
  if (course.productBundle) set(20, course.productBundle);

  // V: billing_interval → per_order をデフォルト（ほぼ全コース同じなので）
  set(21, 'per_order');

  // W: first_commit_count
  set(22, course.firstCommitCount);

  // X: total_commit_count
  set(23, course.totalCommitCount);

  // AG: remarks
  if (course.remark) set(33, course.remark);

  // 反映
  sheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
}

/**************************************************
 * contract_logic_rules に 1コース分を upsert
 **************************************************/
function upsertContractLogicFromCourse_(sheet, course, nowStr) {
  const lastRow = sheet.getLastRow();
  const startRow = 3; // 1行目:和名, 2行目:英名

  let targetRow = null;
  for (let r = startRow; r <= lastRow; r++) {
    const val = sheet.getRange(r, 3).getValue(); // C列: course_id
    if (val && String(val) === course.courseId) {
      targetRow = r;
      break;
    }
  }
  if (!targetRow) {
    targetRow = lastRow >= startRow ? lastRow + 1 : startRow;
  }

  const lastCol = sheet.getLastColumn();
  let rowValues = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];

  function set(colIndex1based, value) {
    const idx = colIndex1based - 1;
    if (idx < rowValues.length) {
      rowValues[idx] = value;
    }
  }

  // A: last_updated
  set(1, nowStr);

  // B: client_company_id → 自動では入れない
  // C: course_id
  set(3, course.courseId);

  // D: cancel_deadline
  if (course.cancelDeadline) set(4, course.cancelDeadline);

  // E: cancel_deadline_logic
  if (course.cancelDeadlineLogic) set(5, course.cancelDeadlineLogic);

  // F: holiday_handling → コース表からは不明なので空欄のまま（必要なら手入力）
  // G: oos_cancel_deadline_rule → 手入力
  // H: skip_rule
  if (course.skipStr) {
    set(8, `スキップ: ${course.skipStr}`);
  }
  // I: long_pause_rule
  if (course.pauseStr) {
    set(9, `休止: ${course.pauseStr}`);
  }
  // J: cancel_deadline_rule_in_long_holiday → 手入力

  // K〜O: 解約金関連 → コース表だけでは判断できないので自動では入れない
  // P: exit_fee_notice_template なども同様

  // S: refund_guarantee_flag
  if (course.refundFlag) {
    set(19, course.refundFlag);
  }

  // Z: cooling_off_flag
  if (course.coolingOffFlag) {
    set(26, course.coolingOffFlag);
  }

  // AC: cooling_off_condition_detail
  if (course.coolingOffDetail) {
    set(29, course.coolingOffDetail);
  }

  // その他の列（返品期限・発送前キャンセル等）はコース表だけでは定まらないので空欄のまま

  sheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
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
 * 指定ラベルと一致する行番号を返す（0始まり）
 */
function findRowIndexByLabel_(values, label) {
  const numRows = values.length;
  for (let r = 0; r < numRows; r++) {
    const v = String(values[r][0] || '').trim();
    if (v === label) return r;
  }
  return -1;
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
 * "1,980円" などから数値部分だけ抜き出して Number にする
 * 数字がなければ '' を返す
 */
function parseNumberOrEmpty_(value) {
  if (value === null || value === '') return '';
  const text = String(value);
  const numStr = text.replace(/[^\d.-]/g, '');
  if (!numStr) return '';
  const num = Number(numStr);
  if (isNaN(num)) return '';
  return num;
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

/* ==== ここから下は、既に同名ヘルパーがあれば省略してOK ==== */

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

// "1,980円" などから数字部分だけ抜いて Number にする（失敗時は ''）
function parseNumberOrEmpty_(value) {
  if (value == null || value === '') return '';
  const text = String(value);
  const numStr = text.replace(/[^\d.-]/g, '');
  if (!numStr) return '';
  const num = Number(numStr);
  if (isNaN(num)) return '';
  return num;
}




/**************************************************
 * ここから 注意タグ（contract_warning_tags）関連
 **************************************************/

/**************************************
 * 注意タグ サイドバーを開く
 **************************************/
function openWarningTagSidebar() {
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
  // 1行目：和名, 2行目：英名 という前提
  if (row <= 2) {
    ui.alert('データ行（3行目以降）を選択してください。');
    return;
  }

  const companyId = sheet.getRange(row, COL_COMPANY_ID).getValue(); // client_company_id
  const courseId  = sheet.getRange(row, COL_COURSE_ID).getValue();  // course_id

  if (!companyId || !courseId) {
    ui.alert('選択した行に 会社ID または コースID がありません。');
    return;
  }

  // タグマスタ取得（code_master 内の item = "contract_warning_tags"）
  const tags = getAllWarningTagMaster_();
  if (tags.length === 0) {
    ui.alert('code_master に item="contract_warning_tags" のデータがありません。');
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
  if (lastRow < 2) return [];

  const data = codeSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  // [ [item, value, desc], ... ]
  const tags = data
    .filter(r => String(r[0]) === 'contract_warning_tags' && r[1])
    .map(r => ({
      value: String(r[1]),
      desc:  String(r[2] || '')
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
  if (lastRow < 3) return [];

  const data = flagSheet.getRange(3, 1, lastRow - 2, 3).getValues();
  // [ [client_company_id, course_id, warning_tag], ... ]

  const tags = data
    .filter(r =>
      String(r[0]) === String(companyId) &&
      String(r[1]) === String(courseId) &&
      r[2]
    )
    .map(r => String(r[2]));

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

  const lastRow = flagSheet.getLastRow();
  if (lastRow >= 3) {
    // 既存行を削除（後ろから削除）
    const data = flagSheet.getRange(3, 1, lastRow - 2, 3).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const rowIndex = i + 3;
      const row = data[i];
      if (String(row[0]) === String(companyId) && String(row[1]) === String(courseId)) {
        flagSheet.deleteRow(rowIndex);
      }
    }
  }

  // 新しいタグを追加
  if (Array.isArray(selectedTags) && selectedTags.length > 0) {
    const rowsToInsert = selectedTags.map(tag => [
      companyId,
      courseId,
      tag,
      '' // remarks
    ]);
    flagSheet.getRange(flagSheet.getLastRow() + 1, 1, rowsToInsert.length, 4).setValues(rowsToInsert);
  }
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

/**************************************************
 * ナレッジDB 承認フロー
 **************************************************/

/**
 * （Step3で使う予定）ID指定でステータス変更する汎用関数
 * - knowId: "N0001" など
 */
function setKnowledgeStatusById_(knowId, newStatus, comment) {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName(SHEET_KNOWLEDGE_DB);
  if (!sh) throw new Error('ナレッジDB シートが見つかりません');

  const lastRow = sh.getLastRow();
  if (lastRow <= KNOW_HEADER_ROW) return;

  const idRange = sh.getRange(KNOW_HEADER_ROW + 1, 1, lastRow - KNOW_HEADER_ROW, 1); // A列
  const ids     = idRange.getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(knowId)) {
      const row = KNOW_HEADER_ROW + 1 + i;
      changeKnowledgeStatus_(sh, row, newStatus, comment);
      return;
    }
  }
  throw new Error(`KNOW_ID "${knowId}" が見つかりません`);
}

/**************************************************
 * 契約マスタ／解約ロジック スナップショット履歴
 **************************************************/


/**
 * ヘルパー: 履歴シートを取得（なければ作成）
 * 1・2行目に元シートのヘッダーをコピーしておく
 */
function getOrCreateHistorySheet_(sourceSheetName, historySheetName) {
  const ss = SpreadsheetApp.getActive();
  let hist = ss.getSheetByName(historySheetName);
  if (!hist) {
    const src = ss.getSheetByName(sourceSheetName);
    if (!src) throw new Error(`${sourceSheetName} シートが見つかりません`);

    hist = ss.insertSheet(historySheetName);

    // 1〜2行目のヘッダーをコピー
    const lastCol = src.getLastColumn();
    const header  = src.getRange(1, 1, 2, lastCol).getValues();
    // 履歴シートの1〜2行目はヘッダーではなく「メタ情報」を置きたいので、
    // 3行目から元シートのヘッダーを置く。
    hist.getRange(3, 1, 2, lastCol).setValues(header);

    // 1行目に履歴用のメタヘッダー
    hist.getRange(1, 1, 1, 4).setValues([[
      'snapshot_ts',
      'source_sheet',
      'course_id',
      'version_note'
    ]]);
  }
  return hist;
}

/**
 * 共通ヘルパー：指定シートの指定行を履歴シートにコピー
 */
function snapshotRowToHistory_(sourceSheetName, historySheetName, row, versionNote) {
  const ss   = SpreadsheetApp.getActive();
  const src  = ss.getSheetByName(sourceSheetName);
  if (!src) throw new Error(`${sourceSheetName} シートが見つかりません`);

  const hist = getOrCreateHistorySheet_(sourceSheetName, historySheetName);

  const lastCol = src.getLastColumn();
  const values  = src.getRange(row, 1, 1, lastCol).getValues()[0];

  // course_id の列位置（既にCOL_COURSE_ID定数があるので流用可）
  const courseId = src.getRange(row, COL_COURSE_ID).getValue();

  const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  const lastHistRow = hist.getLastRow();
  const writeRow    = lastHistRow + 1;

  // 1〜4列：スナップショットメタ情報
  hist.getRange(writeRow, 1, 1, 4).setValues([[
    ts,
    sourceSheetName,
    courseId,
    versionNote || ''
  ]]);

  // 5列目以降に元行を丸コピー
  hist.getRange(writeRow, 5, 1, lastCol).setValues([values]);
}

/**
 * 選択行の contract_master をスナップショット保存
 */
function snapshotSelectedContractRowToHistory() {
  const ss   = SpreadsheetApp.getActive();
  const ui   = SpreadsheetApp.getUi();
  const sh   = ss.getSheetByName(SHEET_CM);
  if (!sh) {
    ui.alert('contract_master シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('スナップショットしたい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= 2) {
    ui.alert('3行目以降のデータ行を選択してください');
    return;
  }

  const res = ui.prompt(
    'バージョンメモ（任意）',
    '今回のスナップショットの理由や変更内容をメモしておくと後で便利です',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const note = res.getResponseText();
  snapshotRowToHistory_(SHEET_CM, SHEET_CM_HISTORY, row, note);
  ui.alert(`contract_master 行 ${row} のスナップショットを保存しました`);
}

/**
 * 選択行の contract_logic_rules をスナップショット保存
 */
function snapshotSelectedLogicRowToHistory() {
  const ss   = SpreadsheetApp.getActive();
  const ui   = SpreadsheetApp.getUi();
  const sh   = ss.getSheetByName(SHEET_CL);
  if (!sh) {
    ui.alert('contract_logic_rules シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('スナップショットしたい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= 2) {
    ui.alert('3行目以降のデータ行を選択してください');
    return;
  }

  const res = ui.prompt(
    'バージョンメモ（任意）',
    '今回のスナップショットの理由や変更内容をメモしておくと後で便利です',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const note = res.getResponseText();
  snapshotRowToHistory_(SHEET_CL, SHEET_CL_HISTORY, row, note);
  ui.alert(`contract_logic_rules 行 ${row} のスナップショットを保存しました`);
}

/**************************************************
 * 更新承認フロー表 と ナレッジDB の連携（ナレッジ版）
 **************************************************/


/**
 * 更新承認フロー表で選択行を「承認（ナレッジ）」にする
 * ・更新承認フロー表のステータスを更新
 * ・ナレッジDB側も承認済みにする
 */
function approveFromApprovalSheet() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sh  = ss.getSheetByName(SHEET_APPROVAL_FLOW);
  if (!sh) {
    ui.alert('更新承認フロー シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('承認したい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= AF_HEADER_ROW) {
    ui.alert('データ行（2行目以降）を選択してください');
    return;
  }

  const type   = sh.getRange(row, COL_AF_TYPE).getValue();
  const knowId = sh.getRange(row, COL_AF_TARGET).getValue();

  if (type !== 'knowledge') {
    ui.alert('現在は「種別 = knowledge」の行のみ自動承認対象としています');
    return;
  }
  if (!knowId) {
    ui.alert('対象キー（KNOW_ID）が空です');
    return;
  }

  // ナレッジDB側を承認済みに
  setKnowledgeStatusById_(knowId, '承認済み', '承認フローから承認');

  // 承認フロー表のステータス更新
  const user = Session.getActiveUser().getEmail() || 'unknown_user';
  const ts   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  sh.getRange(row, COL_AF_STATUS).setValue('承認済み');
  sh.getRange(row, COL_AF_APPROVER).setValue(user);
  sh.getRange(row, COL_AF_UPDATED_AT).setValue(ts);
}

/**
 * 更新承認フロー表で選択行を「差し戻し（ナレッジ）」にする
 */
function rejectFromApprovalSheet() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const sh  = ss.getSheetByName(SHEET_APPROVAL_FLOW);
  if (!sh) {
    ui.alert('更新承認フロー シートが見つかりません');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) {
    ui.alert('差し戻したい行を選択してください');
    return;
  }

  const row = range.getRow();
  if (row <= AF_HEADER_ROW) {
    ui.alert('データ行（2行目以降）を選択してください');
    return;
  }

  const type   = sh.getRange(row, COL_AF_TYPE).getValue();
  const knowId = sh.getRange(row, COL_AF_TARGET).getValue();

  if (type !== 'knowledge') {
    ui.alert('現在は「種別 = knowledge」の行のみ自動差し戻し対象としています');
    return;
  }
  if (!knowId) {
    ui.alert('対象キー（KNOW_ID）が空です');
    return;
  }

  const res = ui.prompt(
    '差し戻し理由を入力',
    'ナレッジ作成者に伝えたい修正ポイントなどを入力してください',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const reason = res.getResponseText() || '';

  // ナレッジDB側を差し戻しに
  setKnowledgeStatusById_(knowId, '差し戻し', reason);

  // 承認フロー表のステータス更新
  const user = Session.getActiveUser().getEmail() || 'unknown_user';
  const ts   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  sh.getRange(row, COL_AF_STATUS).setValue('差し戻し');
  sh.getRange(row, COL_AF_APPROVER).setValue(user);
  sh.getRange(row, COL_AF_UPDATED_AT).setValue(ts);

  const oldNote = sh.getRange(row, COL_AF_NOTE).getValue() || '';
  const line    = `[${ts}] ${user} 差し戻し理由: ${reason}`;
  sh.getRange(row, COL_AF_NOTE).setValue(oldNote ? (oldNote + '\n' + line) : line);
}
