// ------------------------------------------------------------------------------------
// グローバル定数（必須設定）
// ------------------------------------------------------------------------------------
const GLOBAL_STANDARD_SHIFT_ORDER = ["A", "B", "C"];
const GLOBAL_SHIFT_TIMES_MIN = {
  "北葛西小児科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 20 * 60] },
  "北葛西内科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 20 * 60] },
  "亀有小児科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] },
  "亀有内科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] },
  "その他": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] }
};
const GLOBAL_EXCLUDED_LOCATIONS = [
  "有給", "欠勤", "院外勤務（小児科）", "院外勤務（内科）",
  "【関東】バックアップシフト", "医師会・嘱託医業務（小児科）", "医師会・嘱託医業務（内科）", "医師会", "医師会業務", "嘱託医業務"
];
const GLOBAL_EXCLUDED_DEPARTMENTS = [
  "小児科ワクチン専任(対象：小児～成人)", "内科ワクチン専任(対象：小児～成人)"
];
const GLOBAL_TARGET_CLINICS_FOR_DEPT_INFO = ["北葛西", "亀有"];

// ------------------------------------------------------------------------------------
// メインの更新関数
// ------------------------------------------------------------------------------------
function updateSheetRowAdjusted_CallingCellSpecificFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('貼付用');
  const targetSheet = ss.getSheetByName('確認用');
  
  // 空行の大量ログ出力を防ぎ、処理を高速化
  const DETAILED_LOGGING = false; 

  if (!sourceSheet || !targetSheet) {
    ss.toast('貼付用 または 確認用 シートが見つかりません。', 'エラー', 5);
    return;
  }

  ss.toast('シフト集計処理を開始します...', '処理開始', 3);

  const localExcludedLocations = GLOBAL_EXCLUDED_LOCATIONS.filter(item => item !== "【関東】バックアップシフト");

  const data = sourceSheet.getDataRange().getValues();
  const rows = data.slice(2);

  if (rows.length === 0) {
    ss.toast('貼付用シートの3行目以降にデータが見つかりません。', 'エラー', 5);
    return;
  }

  const results = {};
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];

  rows.forEach((row, index) => {
    const rowIndex = index + 3;
    
    // 空行スキップ
    if (!row[0] && !row[12] && !row[14]) return;

    const doctorName = row[0] ? row[0].toString().trim() : '未設定';
    const originalClinicName = row[12] ? row[12].toString().trim() : null;
    const originalDepartment = row[13] ? row[13].toString().trim() : null;
    const shiftDateRaw = row[14];
    
    const aShiftValue = row[63]; // BL列
    const bShiftValue = row[64]; // BM列
    const cShiftValue = row[65]; // BN列

    if ((originalClinicName && localExcludedLocations.includes(originalClinicName)) ||
        (originalDepartment && GLOBAL_EXCLUDED_DEPARTMENTS.includes(originalDepartment))) {
      return;
    }
    
    if (!originalClinicName || !shiftDateRaw || !originalDepartment) return;
    
    let shiftDateObj, shiftDateKey;
    try {
      shiftDateObj = parseDateToSafeDateObj(shiftDateRaw); 
      if (!shiftDateObj || isNaN(shiftDateObj.getTime())) return;
      shiftDateKey = fastFormatDate(shiftDateObj).replace(/\//g, '-'); 
    } catch (e) {
      return;
    }

    let displayClinicName = originalClinicName;
    if (GLOBAL_TARGET_CLINICS_FOR_DEPT_INFO.includes(originalClinicName) &&
        (originalDepartment === "小児科" || originalDepartment === "内科")) {
      displayClinicName = `${originalClinicName}（${originalDepartment}）`;
    }
    const key = `${shiftDateKey}-${displayClinicName}-${originalDepartment}`;

    if (!results[key]) {
      results[key] = {
        shiftDate: shiftDateObj, department: originalDepartment, clinicName: displayClinicName,
        aShiftSum: 0, bShiftSum: 0, cShiftSum: 0,
        doctorsA: [], doctorsB: [], doctorsC: []
      };
    }
    results[key].aShiftSum += Number(aShiftValue) || 0;
    results[key].bShiftSum += Number(bShiftValue) || 0;
    results[key].cShiftSum += Number(cShiftValue) || 0;
    if (Number(aShiftValue) == 1 && doctorName !== '未設定' && !results[key].doctorsA.includes(doctorName)) { results[key].doctorsA.push(doctorName); }
    if (Number(bShiftValue) == 1 && doctorName !== '未設定' && !results[key].doctorsB.includes(doctorName)) { results[key].doctorsB.push(doctorName); }
    if (Number(cShiftValue) == 1 && doctorName !== '未設定' && !results[key].doctorsC.includes(doctorName)) { results[key].doctorsC.push(doctorName); }
  });

  const outputHeader = ['拠点名', '勤務日', '診療科', '09:00~13:00', '15:00~18:00', '18:00~21:00', '', 'Aシフト医師 09:00-13:00', 'Bシフト医師 15:00-18:00', 'Cシフト医師 18:00-21:00'];
  const outputRows = [];
  
  for (const key in results) {
    const record = results[key];
    let formattedShiftDate = '日付エラー';
    try {
      if (record.shiftDate instanceof Date && !isNaN(record.shiftDate.getTime())) {
        const yyyymmdd = fastFormatDate(record.shiftDate);
        const weekday = weekdays[record.shiftDate.getDay()];
        formattedShiftDate = `${yyyymmdd}（${weekday}）`;
      }
    } catch (e) { formattedShiftDate = 'フォーマットエラー'; }
    
    outputRows.push([
      record.clinicName, formattedShiftDate, record.department,
      record.aShiftSum, record.bShiftSum, record.cShiftSum, '',
      record.doctorsA.join(', '), record.doctorsB.join(', '), record.doctorsC.join(', ')
    ]);
  }
  
  const sortedData = outputRows.sort((a, b) => {
    const clinicA = a[0], clinicB = b[0];
    let dateA = new Date("invalid"), dateB = new Date("invalid");
    try { dateA = parseDateToSafeDateObj(a[1].split('（')[0].trim()) || new Date("invalid"); } catch(e){}
    try { dateB = parseDateToSafeDateObj(b[1].split('（')[0].trim()) || new Date("invalid"); } catch(e){}
    if (clinicA < clinicB) return -1; if (clinicA > clinicB) return 1;
    if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) return dateA.getTime() - dateB.getTime();
    if (!isNaN(dateA.getTime())) return -1; if (!isNaN(dateB.getTime())) return 1;
    return a[1].localeCompare(b[1]);
  });
  
  const lastRowOutput = targetSheet.getLastRow();
  if (lastRowOutput >= 1) {
    targetSheet.getRange(1, 1, lastRowOutput, targetSheet.getMaxColumns()).clearContent();
  }
  targetSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]);
  if (sortedData.length > 0) {
    targetSheet.getRange(2, 1, sortedData.length, outputHeader.length).setValues(sortedData);
  }
  SpreadsheetApp.flush();
  
  // ▼ 別のファイルに分けた機能を呼び出す ▼
  try { applyConditionalFormatting_CellSpecific(); } catch (e) {}
  try {
    if (typeof generateDoctorAbsenceReportWithContext === 'function') {
        generateDoctorAbsenceReportWithContext();
    }
  } catch (e) {
    ss.toast(`医師不在拠点書き出しエラー: ${e.message}`, 'エラー', 5);
  }
  try { setupDateSelection(); } catch (e) {}
  try {
    if (typeof generateChatworkMessage === 'function') {
        generateChatworkMessage();
    }
  } catch (e) {}
  
  // 🔽🔽 アプローチシート更新処理を呼び出す 🔽🔽
  try {
    if (typeof updateApproachSheet === 'function') {
        updateApproachSheet();
    }
  } catch (e) {
    ss.toast(`アプローチシート更新エラー: ${e.message}`, 'エラー', 5);
  }
  // 🔼🔼 ここまで 🔼🔼
  
  ss.toast('本番シフト集計が完了しました。', '完了', 3);
}