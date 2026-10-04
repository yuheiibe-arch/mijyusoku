function setupDateSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('文章自動作成'); // シートを取得
  const sourceSheet = ss.getSheetByName('確認用'); // 日付データの取得元

  if (!sheet || !sourceSheet) {
    Logger.log('「文章自動作成」または「確認用」シートが見つかりません。');
    return;
  }

  const today = new Date();
  const sevenDaysLater = new Date();
  sevenDaysLater.setDate(today.getDate() + 7); // **7日後に変更**

  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"]; // 日本語の曜日
  const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd') + `（${weekdaysJP[today.getDay()]}）`;
  const formattedSevenDaysLater = Utilities.formatDate(sevenDaysLater, Session.getScriptTimeZone(), 'yyyy/MM/dd') + `（${weekdaysJP[sevenDaysLater.getDay()]}）`;

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('「確認用」シートにデータがありません。');
    return;
  }

  // **B列のデータを取得**
  const bColumnValues = sourceSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat(); // B列（2列目）

  // **重複を削除し、日本語の曜日を追加**
  const uniqueValues = [...new Set(bColumnValues.filter(value => value !== ''))]
    .map(dateStr => {
      const dateObj = new Date(dateStr);
      if (!isNaN(dateObj.getTime())) {
        return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd') + `（${weekdaysJP[dateObj.getDay()]}）`;
      }
      return dateStr; // 変換できない場合はそのまま
    });

  if (uniqueValues.length === 0) {
    Logger.log('プルダウンに設定する有効な日付データがありません。');
    return;
  }

  // **B2・B4 のデータ検証を削除**
  sheet.getRange('B2').setDataValidation(null);
  sheet.getRange('B4').setDataValidation(null);

  // **データ検証（プルダウン）を設定**
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueValues, true) // B列のユニークな日付をプルダウンに設定
    .setAllowInvalid(false) // 無効な値を許可しない
    .build();

  // **B2セルに今日の日付を設定**
  sheet.getRange('B2').setValue(formattedToday).setDataValidation(rule);

  // **B4セルに7日後の日付を設定**
  sheet.getRange('B4').setValue(formattedSevenDaysLater).setDataValidation(rule); // **7日後に変更**

  Logger.log('B2に今日の日付、B4に7日後の日付を設定し、プルダウンを適用しました。');
}