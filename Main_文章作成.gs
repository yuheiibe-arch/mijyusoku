/**
 * 1. メイン処理: Chatwork用メッセージ生成
 */
function generateChatworkMessage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptTimeZone = Session.getScriptTimeZone(); 
  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"]; 

  ss.toast("文章自動作成を開始します...", "処理開始", 3);

  // --- シート取得 ---
  const sheets = {
    source: ss.getSheetByName('確認用'),
    target: ss.getSheetByName('文章自動作成'),
    mention: ss.getSheetByName('メンション先選択'),
    ishiFuzai: ss.getSheetByName('医師不在拠点'),
    closed: ss.getSheetByName('休館日'), 
    irregular: ss.getSheetByName('変則営業')
  };

  if (!sheets.source || !sheets.target || !sheets.mention || !sheets.ishiFuzai) { 
    ss.toast('エラー: 必要なシートが見つかりません。', 'エラー', 5); 
    return; 
  }

  // --- 🌟UI日付の取得とデフォルト設定ロジック ---
  const baseDate = new Date();
  if (baseDate.getHours() >= 15) baseDate.setDate(baseDate.getDate() + 1);
  const endDateCalc = new Date(baseDate);
  endDateCalc.setDate(baseDate.getDate() + 6); 

  const formattedStart = fastFormatDate(baseDate) + `（${weekdaysJP[baseDate.getDay()]}）`;
  const formattedEnd = fastFormatDate(endDateCalc) + `（${weekdaysJP[endDateCalc.getDay()]}）`;

  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  let uniqueValues = [];
  if (sheets.source.getLastRow() >= 2) {
    const bColumnValues = sheets.source.getRange(2, 2, sheets.source.getLastRow() - 1, 1).getValues().flat();
    uniqueValues = [...new Set(bColumnValues.filter(Boolean))]
      .map(dateStr => {
        const dateObj = parseDateToSafeDateObj(dateStr);
        if (dateObj && dateObj < todayStart) return null;
        return !dateObj ? dateStr : fastFormatDate(dateObj) + `（${weekdaysJP[dateObj.getDay()]}）`;
      })
      .filter(Boolean);
  }
  if (!uniqueValues.includes(formattedStart)) uniqueValues.unshift(formattedStart);
  if (!uniqueValues.includes(formattedEnd)) uniqueValues.push(formattedEnd);
  
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(uniqueValues, true).setAllowInvalid(true).build();
  
  const cellB2 = sheets.target.getRange('B2');
  const cellB4 = sheets.target.getRange('B4');
  
  cellB2.setNumberFormat('@').setDataValidation(null);
  cellB4.setNumberFormat('@').setDataValidation(null);
  
  // ★変更点：セルに値が入っている場合はそれを優先する（手動での任意日付の実行を可能にする）
  let startDateRaw = cellB2.getValue();
  let endDateRaw = cellB4.getValue();
  
  if (!startDateRaw) {
    cellB2.setValue(formattedStart);
    startDateRaw = formattedStart;
  }
  if (!endDateRaw) {
    cellB4.setValue(formattedEnd);
    endDateRaw = formattedEnd;
  }
  
  cellB2.setDataValidation(rule);
  cellB4.setDataValidation(rule);

  const startDate = parseDateToSafeDateObj(startDateRaw);
  const endDate = parseDateToSafeDateObj(endDateRaw);

  if (!startDate || !endDate || startDate > endDate) {
    ss.toast(`日付指定が無効です。\n開始: ${startDateRaw}\n終了: ${endDateRaw}`, 'エラー', 5); 
    return; 
  }

  // --- データの読み込み (分割した関数を呼び出し) ---
  const masterData = loadMasterAndExternalData(ss, sheets);
  const shiftData = loadShiftAndAbsenceData(sheets.source, sheets.ishiFuzai);

  // --- レポート生成 (分割した関数を呼び出し) ---
  const report = buildDailyAndSummaryReport(startDate, endDate, masterData, shiftData, weekdaysJP);

  // --- テキスト出力 ---
  const now = new Date();
  let hours = now.getHours();
  let minutes = now.getMinutes();
  if (minutes <= 19) minutes = "00"; else if (minutes <= 49) minutes = "30"; else { minutes = "00"; hours = (hours + 1) % 24; }
  const formattedHours = String(hours).padStart(2, '0');
  const formattedReportDate = `${now.getMonth() + 1}月${now.getDate()}日（${weekdaysJP[now.getDay()]}）`;
  
  let initialText = `【未充足報告】${formattedReportDate} ${formattedHours}:${minutes}時点\n\n`;

  let mentionsArray = [], ccArray = [];
  const mentionData = sheets.mention.getDataRange().getValues();
  for (let i = 1; i < mentionData.length; i++) {
    if (mentionData[i][0]) mentionsArray.push(String(mentionData[i][0]).trim());
    if (mentionData[i][1]) ccArray.push(String(mentionData[i][1]).trim());
  }
  if (mentionsArray.length > 0) initialText += mentionsArray.join('') + "\n";
  if (ccArray.length > 0) initialText += "CC:" + ccArray.join('') + "\n";
  initialText += "\n";

  sheets.target.getRange('A6').clearContent();
  sheets.target.getRange('A6').setValue(initialText + report.summaryText + report.dailyText);
  ss.toast('文章自動作成が完了しました。', '完了', 3); 
}