function setupDateSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('文章自動作成');
  const sourceSheet = ss.getSheetByName('確認用');

  if (!sheet || !sourceSheet) {
    Logger.log('「文章自動作成」または「確認用」シートが見つかりません。');
    return;
  }

  const today = new Date();
  const sevenDaysLater = new Date();
  sevenDaysLater.setDate(today.getDate() + 7);

  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"];
  const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd') + `（${weekdaysJP[today.getDay()]}）`;
  const formattedSevenDaysLater = Utilities.formatDate(sevenDaysLater, Session.getScriptTimeZone(), 'yyyy/MM/dd') + `（${weekdaysJP[sevenDaysLater.getDay()]}）`;

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('「確認用」シートにデータがありません。');
    return;
  }

  const bColumnValues = sourceSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();

  const uniqueValues = [...new Set(bColumnValues.filter(value => value !== ''))]
    .map(dateStr => {
      const dateObj = new Date(dateStr);
      if (!isNaN(dateObj.getTime())) {
        return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd') + `（${weekdaysJP[dateObj.getDay()]}）`;
      }
      return dateStr;
    });

  if (uniqueValues.length === 0) {
    Logger.log('プルダウンに設定する有効な日付データがありません。');
    return;
  }

  sheet.getRange('B2').setDataValidation(null);
  sheet.getRange('B4').setDataValidation(null);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueValues, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange('B2').setValue(formattedToday).setDataValidation(rule);
  sheet.getRange('B4').setValue(formattedSevenDaysLater).setDataValidation(rule);

  Logger.log('B2に今日の日付、B4に7日後の日付を設定し、プルダウンを適用しました。');
}