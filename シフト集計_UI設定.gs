// ------------------------------------------------------------------------------------
// UI操作・書式設定 関数群
// ------------------------------------------------------------------------------------
function clearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('データの削除と書式リセットを開始します...', '処理開始', 2);
    const sourceSheet = ss.getSheetByName('貼付用');
    const targetSheet = ss.getSheetByName('確認用');

    if (sourceSheet) {
      const lastRowSource = sourceSheet.getLastRow();
      if (lastRowSource >= 3) sourceSheet.getRange(3, 1, lastRowSource - 2, 49).clear();
    }
    if (targetSheet) {
      const lastRowTarget = targetSheet.getLastRow();
      if (lastRowTarget >= 2) targetSheet.getRange(2, 1, lastRowTarget - 1, 10).clear();
    }
    ss.toast('削除とリセットが完了しました。', '完了', 3);
  } catch (e) {
    ss.toast(`データ削除中にエラーが発生しました: ${e.message}`, 'エラー', 5);
  }
}

function applyConditionalFormatting_CellSpecific() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName('確認用');
  if (!targetSheet) return;

  const startRow = 2;
  const lastRow = targetSheet.getLastRow();
  if (lastRow < startRow) return;

  const numColumns = 10;
  const dataRange = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, numColumns);
  const values = dataRange.getValues();
  const backgroundColors = [];
  const redColor = '#FFCCCC';
  const whiteColor = '#FFFFFF';

  for (let i = 0; i < values.length; i++) {
    const rowBackgrounds = [];
    for (let j = 0; j < numColumns; j++) {
      let cellColor = whiteColor;
      const cellValue = values[i][j];
      switch (j) {
        case 3: case 4: case 5:
          if (cellValue === '' || cellValue == null || Number(cellValue) === 0) cellColor = redColor;
          break;
        case 7: case 8: case 9:
          if (cellValue === '' || cellValue == null) cellColor = redColor;
          break;
      }
      rowBackgrounds.push(cellColor);
    }
    backgroundColors.push(rowBackgrounds);
  }
  dataRange.setBackgrounds(backgroundColors);
}

function setupDateSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('文章自動作成');
  const sourceSheet = ss.getSheetByName('確認用');
  if (!sheet || !sourceSheet) return;

  const baseDate = new Date();
  if (baseDate.getHours() >= 15) baseDate.setDate(baseDate.getDate() + 1);
  const endDate = new Date(baseDate);
  endDate.setDate(baseDate.getDate() + 6); 

  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"];
  const formattedStart = fastFormatDate(baseDate) + `（${weekdaysJP[baseDate.getDay()]}）`; 
  const formattedEnd = fastFormatDate(endDate) + `（${weekdaysJP[endDate.getDay()]}）`; 

  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  const lastRow = sourceSheet.getLastRow();
  let uniqueValues = [];

  if (lastRow >= 2) {
    const bColumnValues = sourceSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    uniqueValues = [...new Set(bColumnValues.filter(Boolean))]
      .map(dateStr => {
        const dateObj = parseDateToSafeDateObj(dateStr); 
        if (dateObj && dateObj < todayStart) return null;
        return !dateObj ? dateStr : fastFormatDate(dateObj) + `（${weekdaysJP[dateObj.getDay()]}）`;
      }).filter(Boolean);
  }

  if (!uniqueValues.includes(formattedStart)) uniqueValues.unshift(formattedStart);
  if (!uniqueValues.includes(formattedEnd)) uniqueValues.push(formattedEnd);

  if (uniqueValues.length === 0) return;

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(uniqueValues, true).setAllowInvalid(true).build();

  const cellB2 = sheet.getRange('B2');
  const cellB4 = sheet.getRange('B4');
  
  cellB2.setNumberFormat('@').setDataValidation(null);
  cellB4.setNumberFormat('@').setDataValidation(null);
  
  // ★変更点：前回実行時の古い日付が残っていても、強制的に新しい日付をセットする
  cellB2.setValue(formattedStart);
  cellB4.setValue(formattedEnd);
  
  cellB2.setDataValidation(rule);
  cellB4.setDataValidation(rule);
}