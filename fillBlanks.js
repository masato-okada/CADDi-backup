function fillEmptyCells() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート2");
  const targetColumns = [2, 3, 4]; // 処理対象の列番号（例: B列, D列, E列）
  const lastRow = sheet.getLastRow();
  
  targetColumns.forEach(col => { // target columnsに含まれる各列に対して処理を行う
    let values = sheet.getRange(1, col, lastRow).getValues(); // 各列の値を1行目から配列で取得
    
    for (let row = 1; row < values.length; row++) { // 空白セルに値を埋める処理
      if (values[row][0] === '' || values[row][0] === null) {
        values[row][0] = values[row - 1][0];
      }
    }
    
    sheet.getRange(1, col, lastRow).setValues(values); // 処理結果をシートに書き込み
  });
}
