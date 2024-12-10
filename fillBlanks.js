function fillEmptyCells() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート2");
  const targetColumns = [2, 3, 4]; // 処理対象の列番号（例: B列, D列, E列）
  const lastRow = sheet.getLastRow();
  
  targetColumns.forEach(col => {
    let values = sheet.getRange(1, col, lastRow).getValues();
    
    for (let row = 1; row < values.length; row++) {
      if (values[row][0] === '' || values[row][0] === null) {
        values[row][0] = values[row - 1][0];
      }
    }
    
    sheet.getRange(1, col, lastRow).setValues(values);
  });
}
