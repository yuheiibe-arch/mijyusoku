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

  // --- 🌟UI日付の取得とデフォルト設定ロジック（15:00以降は明日基準＆終了日も+1日） ---
  const baseDate = new Date();
  if (baseDate.getHours() >= 15) {
    baseDate.setDate(baseDate.getDate() + 1);
  }
  const baseDateZero = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());

  const endDateCalc = new Date(baseDateZero);
  endDateCalc.setDate(baseDateZero.getDate() + 6); 

  const formattedStart = fastFormatDate(baseDateZero) + `（${weekdaysJP[baseDateZero.getDay()]}）`;
  const formattedEnd = fastFormatDate(endDateCalc) + `（${weekdaysJP[endDateCalc.getDay()]}）`;

  let uniqueValues = [];
  if (sheets.source.getLastRow() >= 2) {
    const bColumnValues = sheets.source.getRange(2, 2, sheets.source.getLastRow() - 1, 1).getValues().flat();
    uniqueValues = [...new Set(bColumnValues.filter(Boolean))]
      .map(dateStr => {
        const dateObj = parseDateToSafeDateObj(dateStr);
        // 15時以降は「今日」以前の選択肢を除外し、明日以降のみにする
        if (dateObj && dateObj < baseDateZero) return null;
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
  
  let startDateRaw = cellB2.getValue();
  let endDateRaw = cellB4.getValue();
  
  const parsedCellStart = parseDateToSafeDateObj(startDateRaw);
  const parsedCellEnd = parseDateToSafeDateObj(endDateRaw);

  // ★15:00以降に実行した際、B2セルが空欄または「今日以前」になっていたら、強制的に「明日〜明日+6日」へ更新する
  if (!parsedCellStart || parsedCellStart < baseDateZero) {
    cellB2.setValue(formattedStart);
    startDateRaw = formattedStart;
    cellB4.setValue(formattedEnd);
    endDateRaw = formattedEnd;
  } else if (!parsedCellEnd || parsedCellEnd < parsedCellStart) {
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

  // --- データの読み込み（下の専用関数を直接呼び出すので他ファイルに邪魔されません！） ---
  const masterData = loadMasterAndExternalData(ss, sheets);
  const shiftData = loadShiftAndAbsenceData_New_(sheets.source, sheets.ishiFuzai);

  // --- レポート生成 ---
  const report = buildDailyAndSummaryReport(startDate, endDate, masterData, shiftData, weekdaysJP);

  // ★ バックアップが0人で空欄になった日（[/title][hr] が直結している日）に「バックアップ医師ご応募なし」を挿入
  const finalDailyText = report.dailyText
    .split('【バックアップ】').join('【関東バックアップ】')
    .split('（全拠点）').join('（関東全域）')
    .replace(/\[\/title\]\s*[hr]/g, '[/title]バックアップ医師ご応募なし[hr]');

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

  const fullMessage = initialText + report.summaryText + finalDailyText;

  sheets.target.getRange('A6').clearContent();
  sheets.target.getRange('A6').setValue(fullMessage);
  SpreadsheetApp.flush(); // ★即座にシートへ反映させる
  ss.toast('文章自動作成が完了しました（関東バックアップ対応済）。', '完了', 3); 
}

/**
 * ★他ファイルの古い関数と名前が被らないよう、専用の新しい関数名（_New_）にしました
 */
function loadShiftAndAbsenceData_New_(sourceSheet, ishiFuzaiSheet) {
  const EXCLUDED_KEYWORDS = ["有給", "欠勤", "院外勤務", "バックアップ", "医師会", "嘱託医", "出張インフルエンザワクチン"];
  const allClinics = new Set(); 
  const backupInfoMap = {}; 
  const rawNameMap = {}; 
  const workingClinics = {}; 
  
  const srcData = sourceSheet.getDataRange().getValues();
  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const name = row[0] ? String(row[0]).trim() : "";
    
    if (name === "【関東】バックアップシフト") {
        const dateObj = parseDateToSafeDateObj(row[1]);
        if (dateObj) {
            const dateKey = fastFormatDate(dateObj);
            const slots = ["09:00~13:00", "15:00~18:00", "18:00~21:00"];
            let backups = [];
            for (let j = 0; j < slots.length; j++) {
                const doc = row[7 + j]; 
                if (doc && String(doc).trim()) backups.push(`${slots[j]}：${doc}先生（関東全域）`);
            }
            if (backups.length > 0) {
                backupInfoMap[dateKey] = `【関東バックアップ】${backups.join('、')}`;
            } else {
                backupInfoMap[dateKey] = `バックアップ医師ご応募なし`;
            }
        }
        continue; 
    }

    if (!name || EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;
    const normName = normalizeClinicName(name);
    allClinics.add(normName);
    rawNameMap[normName] = name; 

    const dateObj = parseDateToSafeDateObj(row[1]);
    if (dateObj) {
        const dateKey = fastFormatDate(dateObj);
        if (!workingClinics[dateKey]) workingClinics[dateKey] = new Set();
        workingClinics[dateKey].add(normName);
    }
  }
  
  const ishiFuzaiRawData = ishiFuzaiSheet.getDataRange().getValues();
  const fuzaiRecords = []; 

  for (let i = 1; i < ishiFuzaiRawData.length; i++) { 
    const row = ishiFuzaiRawData[i];
    const dateValue = row[0];        
    const name = row[1] ? row[1].toString().trim() : ""; 
    const time = row[2] ? row[2].toString().trim() : "";

    if (!dateValue || !name || !time) continue; 
    if (EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;

    const normName = normalizeClinicName(name);
    allClinics.add(normName); 
    rawNameMap[normName] = name; 

    const dObj = parseDateToSafeDateObj(dateValue);
    if (dObj) {
      fuzaiRecords.push({
        dateKey: fastFormatDate(dObj),
        name: name,      
        normName: normName, 
        timeRaw: time,
        isInternalMedicine: name.includes("内科")
      });
    }
  }

  return { sortedClinicList: Array.from(allClinics).sort(), backupInfoMap, rawNameMap, workingClinics, fuzaiRecords };
}