function setupDateSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('文章自動作成');
  const sourceSheet = ss.getSheetByName('確認用');

  if (!sheet || !sourceSheet) {
    Logger.log('「文章自動作成」または「確認用」シートが見つかりません。');
    return;
  }

  // ★15:00以降は基準日（開始期日）を「明日」、終了期日も1日足す（明日+6日）
  const baseDate = new Date();
  if (baseDate.getHours() >= 15) {
    baseDate.setDate(baseDate.getDate() + 1);
  }
  const baseDateZero = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());

  const endDate = new Date(baseDateZero);
  endDate.setDate(baseDateZero.getDate() + 6);

  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"];
  const formattedStart = fastFormatDate(baseDateZero) + `（${weekdaysJP[baseDateZero.getDay()]}）`;
  const formattedEnd = fastFormatDate(endDate) + `（${weekdaysJP[endDate.getDay()]}）`;

  const lastRow = sourceSheet.getLastRow();
  let uniqueValues = [];

  if (lastRow >= 2) {
    const bColumnValues = sourceSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    uniqueValues = [...new Set(bColumnValues.filter(Boolean))]
      .map(dateStr => {
        const dateObj = parseDateToSafeDateObj(dateStr);
        // 基準日（15時以降は明日）より前の日付はプルダウン候補から除外
        if (dateObj && dateObj < baseDateZero) return null;
        return !dateObj ? dateStr : fastFormatDate(dateObj) + `（${weekdaysJP[dateObj.getDay()]}）`;
      })
      .filter(Boolean);
  }

  if (!uniqueValues.includes(formattedStart)) uniqueValues.unshift(formattedStart);
  if (!uniqueValues.includes(formattedEnd)) uniqueValues.push(formattedEnd);

  if (uniqueValues.length === 0) {
    Logger.log('プルダウンに設定する有効な日付データがありません。');
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueValues, true)
    .setAllowInvalid(true)
    .build();

  const cellB2 = sheet.getRange('B2');
  const cellB4 = sheet.getRange('B4');

  cellB2.setNumberFormat('@').setDataValidation(null);
  cellB4.setNumberFormat('@').setDataValidation(null);

  // 強制的に新しい基準日（15時以降は明日〜明日+6日）をセットする
  cellB2.setValue(formattedStart);
  cellB4.setValue(formattedEnd);

  cellB2.setDataValidation(rule);
  cellB4.setDataValidation(rule);

  Logger.log(`B2に ${formattedStart}、B4に ${formattedEnd} を設定し、プルダウンを適用しました。`);
}