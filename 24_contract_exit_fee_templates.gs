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
  const ss   = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_CONTRACT_LOGIC); // ← 定数を利用
  const ui   = SpreadsheetApp.getUi();

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
  const values  = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  // 列番号（1始まり）をわかりやすく定義
  const COL_COURSE_ID            = 3;  // C: course_id
  const COL_EXIT_FEE_AMOUNT      = 11; // K: exit_fee_amount
  const COL_EXIT_FEE_CALC_METHOD = 12; // L: exit_fee_calc_method
  const COL_EXIT_FEE_COND_DETAIL = 14; // N: exit_fee_condition_detail

  let updateCount = 0;

  values.forEach((row) => {
    const courseId    = row[COL_COURSE_ID - 1];
    const exitFeeAmt  = row[COL_EXIT_FEE_AMOUNT - 1];
    const method      = row[COL_EXIT_FEE_CALC_METHOD - 1];
    const currentCond = row[COL_EXIT_FEE_COND_DETAIL - 1];

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
      const hasExitFeeAmt = exitFeeAmt !== '' && exitFeeAmt != null;
      const amtStr = hasExitFeeAmt
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
    row[COL_EXIT_FEE_COND_DETAIL - 1] = templateText;
    updateCount++;
  });

  // 変更があった行だけまとめて書き戻す
  if (updateCount > 0) {
    sheet.getRange(3, 1, values.length, lastCol).setValues(values);
    ui.alert('解約金条件テンプレを ' + updateCount + ' 行に反映しました。');
  } else {
    ui.alert(
      '自動反映対象の行がありませんでした。\n' +
      '（course_id 空行、または exit_fee_condition_detail が既に入力済みの行のみでした。）'
    );
  }
}
