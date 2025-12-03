/**
 * プロジェクト全体の .gs ファイルから
 * const 定義を自動抽出するツール
 */
function exportAllConstDefinitions() {
  const files = DriveApp.getFileById(SpreadsheetApp.getActive().getId())
    .getParents()
    .next()
    .getFiles();

  let result = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!name.endsWith('.gs')) continue;

    const content = file.getBlob().getDataAsString();
    const lines = content.split('\n');

    lines.forEach((line, idx) => {
      const trimmed = line.trim();
      // 「const ** =」の行だけ抽出
      if (trimmed.startsWith('const ')) {
        result.push({
          file: name,
          line: idx + 1,
          code: trimmed
        });
      }
    });
  }

  // 結果をログへ出力（必要ならシートに書き出せる）
  console.log(JSON.stringify(result, null, 2));

  // 利便性のためにシートにも書き出し
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'const_export';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();

  sheet.getRange(1, 1, 1, 3).setValues([['ファイル名', '行番号', 'コード']]);

  const rows = result.map(r => [r.file, r.line, r.code]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  SpreadsheetApp.getUi().alert('抽出が完了しました。\n「const_export」シートをご確認ください。');
}
