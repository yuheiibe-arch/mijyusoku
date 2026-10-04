/**
 * 時間文字列を分数に変換するヘルパー関数
 */
function parseTimeToMins(tStr) {
  if (!tStr) return 0;
  if (tStr instanceof Date) return tStr.getHours() * 60 + tStr.getMinutes();
  const p = String(tStr).trim().split(':');
  if (p.length >= 2) return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  if (typeof tStr === 'number' && tStr >= 0 && tStr < 1) return Math.round(tStr * 24 * 60);
  return 0;
}

/**
 * =========================================================
 * アプローチシート自動更新バッチ処理
 * =========================================================
 */
function updateApproachSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const absenceSheet = ss.getSheetByName('医師不在拠点');
  const approachSheet = ss.getSheetByName('アプローチシート');
  const pasteSheet = ss.getSheetByName('貼付用');
  
  if (!absenceSheet || !approachSheet || !pasteSheet) return;

  // 内閣府の公式CSVから動的に日本の祝日を取得する（半永久対応）
  const holidayCache = new Map();
  let isHolidayDataLoaded = false;
  
  const loadOfficialHolidays = () => {
    if (isHolidayDataLoaded) return;
    const url = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
    try {
      const response = UrlFetchApp.fetch(url);
      const csvText = response.getContentText('Shift_JIS');
      const lines = csvText.split(/\r\n|\n/);
      for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        const parts = line.split(',');
        if (parts.length >= 2) {
          const dateObj = new Date(parts[0]);
          if (!isNaN(dateObj.getTime())) {
            const formatted = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy/MM/dd");
            holidayCache.set(formatted, true); 
          }
        }
      }
      isHolidayDataLoaded = true;
    } catch (e) {
      Logger.log("祝日データ取得失敗: " + e.message);
    }
  };

  const isHoliday = (dateObj) => {
    loadOfficialHolidays();
    const dStr = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy/MM/dd");
    return holidayCache.has(dStr);
  };
  
  // 1. 貼付用シートから [拠点名_科目] -> [クリニックID] のマッピング
  const pasteData = pasteSheet.getDataRange().getValues();
  let clinicNoIdx = -1;
  for (let c = 0; c < pasteData[0].length; c++) {
    if (pasteData[0][c] === "クリニックNo" || pasteData[1][c] === "クリニックNo") {
      clinicNoIdx = c; break;
    }
  }
  const clinicIdMap = new Map();
  if (clinicNoIdx !== -1) {
    for (let i = 2; i < pasteData.length; i++) {
      const idRaw = pasteData[i][clinicNoIdx];
      const name = String(pasteData[i][12] || "").trim();
      const dept = String(pasteData[i][13] || "").trim();
      if (idRaw && name) {
        clinicIdMap.set(`${name}_${dept}`, parseInt(idRaw, 10));
      }
    }
  }

  // 2. マスタシート（時給表）の動的キャッシュ関数
  const MASTER_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  let masterSs = null;
  const masterCache = new Map();
  
  const getMasterData = (sheetName) => {
    if (masterCache.has(sheetName)) return masterCache.get(sheetName);
    if (!masterSs) {
      try { masterSs = SpreadsheetApp.openById(MASTER_SS_ID); } catch (e) { return null; }
    }
    const sheet = masterSs.getSheetByName(sheetName);
    if (!sheet) { masterCache.set(sheetName, null); return null; }
    
    const data = sheet.getDataRange().getValues();
    const sheetMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const idRaw = data[i][0];
      const dept = String(data[i][2]).trim();
      if (idRaw) {
        const idNum = parseInt(idRaw, 10);
        sheetMap.set(`${idNum}_${dept}`, data[i]); 
      }
    }
    masterCache.set(sheetName, sheetMap);
    return sheetMap;
  };

  // 3. 募集シフトの読み込みとマップ化
  const BOSYU_SS_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  let bosyuMap = new Map();
  let hasBosyuSheet = false;
  try {
    const extSs = SpreadsheetApp.openById(BOSYU_SS_ID);
    const bosyuSheet = extSs.getSheetByName('募集シフト');
    if (bosyuSheet) {
      hasBosyuSheet = true;
      const bData = bosyuSheet.getDataRange().getValues();
      if (bData.length > 0) {
        const bHeaders = bData[0];
        let hIdx = { clinicNo: -1, date: -1, start: -1, end: -1, s1: -1, s2: -1, s3: -1, s4: -1 };
        
        for (let c = 0; c < bHeaders.length; c++) {
          const hStr = String(bHeaders[c]).trim();
          if (hStr === "クリニックNo") hIdx.clinicNo = c;
          if (hStr === "勤務日") hIdx.date = c;
          if (hStr === "勤務開始時間") hIdx.start = c;
          if (hStr === "勤務終了時間") hIdx.end = c;
          if (hStr === "時給1") hIdx.s1 = c;
          if (hStr === "時給2") hIdx.s2 = c;
          if (hStr === "時給3") hIdx.s3 = c;
          if (hStr === "時給4") hIdx.s4 = c;
        }
        
        for (let i = 1; i < bData.length; i++) {
          const row = bData[i];
          if (hIdx.clinicNo !== -1 && hIdx.date !== -1) {
             const cId = parseInt(row[hIdx.clinicNo], 10);
             let bDate = row[hIdx.date];
             if (bDate instanceof Date) {
               bDate = Utilities.formatDate(bDate, "Asia/Tokyo", "yyyy/MM/dd");
             } else {
               const dObj = new Date(bDate);
               if (!isNaN(dObj.getTime())) bDate = Utilities.formatDate(dObj, "Asia/Tokyo", "yyyy/MM/dd");
             }
             
             const key = `${bDate}_${cId}`;
             const sTime = hIdx.start !== -1 ? parseTimeToMins(row[hIdx.start]) : 0;
             const eTime = hIdx.end !== -1 ? parseTimeToMins(row[hIdx.end]) : 0;
             
             let maxBosyuSal = 0;
             [hIdx.s1, hIdx.s2, hIdx.s3, hIdx.s4].forEach(idx => {
               if (idx !== -1) {
                 const salNum = parseInt(String(row[idx]).replace(/,/g, ''), 10) || 0;
                 if (salNum > maxBosyuSal) maxBosyuSal = salNum;
               }
             });
             
             if (!bosyuMap.has(key)) bosyuMap.set(key, []);
             bosyuMap.get(key).push({ start: sTime, end: eTime, salary: maxBosyuSal });
          }
        }
      }
    }
  } catch (e) {}

  // 4. 既存データの読み込み＆ログ（完了済み）データのキャッシュ
  const absenceData = absenceSheet.getDataRange().getValues();
  if (absenceData.length <= 1) return; 
  
  const existingData = approachSheet.getDataRange().getValues();
  const headers = existingData[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);
  
  const existingRowsMap = new Map();
  for (let i = 2; i < existingData.length; i++) {
    const row = existingData[i];
    let dateStr = String(row[colIdx['日付']] || "").trim();
    const dObj = new Date(dateStr);
    if (!isNaN(dObj.getTime())) dateStr = Utilities.formatDate(dObj, "Asia/Tokyo", "yyyy/MM/dd");
    
    const clinicStr = String(row[colIdx['拠点名']] || "").trim();
    const timeStr = String(row[colIdx['不在時間']] || "").trim();
    if (dateStr && clinicStr) {
      existingRowsMap.set(`${dateStr}_${clinicStr}_${timeStr}`, row);
    }
  }
  
  // 【追加】ログシートの完了済み案件をキャッシュ（復活防止用）
  const logSheet = ss.getSheetByName('ログ');
  const logKeys = new Set();
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    if (logData.length > 0) {
      const logHeaders = logData[0];
      const lColIdx = {};
      logHeaders.forEach((h, i) => lColIdx[h] = i);
      
      for (let i = 1; i < logData.length; i++) {
        const lRow = logData[i];
        let dStr = String(lRow[lColIdx['日付']] || "").trim();
        const dObj = new Date(dStr);
        if (!isNaN(dObj.getTime())) dStr = Utilities.formatDate(dObj, "Asia/Tokyo", "yyyy/MM/dd");
        const cStr = String(lRow[lColIdx['拠点名']] || "").trim();
        const tStr = String(lRow[lColIdx['不在時間']] || "").trim();
        if (dStr && cStr && tStr) {
          logKeys.add(`${dStr}_${cStr}_${tStr}`);
        }
      }
    }
  }
  
  const todayDate = new Date();
  todayDate.setHours(0,0,0,0);
  const rowsToWrite = [];
  const backgroundsToWrite = [];
  
  // 5. 不在拠点ごとにメイン処理
  for (let i = 1; i < absenceData.length; i++) {
    const aRow = absenceData[i];
    if (!aRow[0] || !aRow[1]) continue;
    
    let shiftDate = null;
    let dateStr = "";
    const rawDateObj = new Date(aRow[0]);
    if (!isNaN(rawDateObj.getTime())) {
      shiftDate = rawDateObj;
      dateStr = Utilities.formatDate(shiftDate, "Asia/Tokyo", "yyyy/MM/dd");
    } else {
      dateStr = String(aRow[0]).split('（')[0].trim();
      shiftDate = parseDateToSafeDateObj(dateStr);
    }
    if (!shiftDate) continue;
    
    // 【追加】今日、および過去の日付を除外する
    const diffDays = Math.ceil((shiftDate.getTime() - todayDate.getTime()) / (1000 * 60 * 60 * 24));
    if (diffDays <= 0) continue; 
    
    const clinicNameRaw = String(aRow[1]).trim();
    const absenceTime = String(aRow[2]).trim();
    
    // 【追加】ログに存在する（＝完了済み転記済み）場合はスキップして復活を防ぐ
    const key = `${dateStr}_${clinicNameRaw}_${absenceTime}`;
    if (logKeys.has(key)) continue;

    let prevDoc = String(aRow[3]).trim();
    let nextDoc = String(aRow[4]).trim();
    if (prevDoc === "未充足") prevDoc = "";
    if (nextDoc === "未充足") nextDoc = "";
    
    let clinicName = clinicNameRaw;
    let dept = "小児科"; 
    const match = clinicNameRaw.match(/(.+?)\s*[\(（](.+)[\)）]/);
    if (match) {
      clinicName = match[1].trim();
      dept = match[2].trim();
    }
    
    const timeParts = absenceTime.split('-');
    const startMin = parseTimeToMins(timeParts[0]);
    const endMin = parseTimeToMins(timeParts[1]);
    
    const y = shiftDate.getFullYear();
    const m = shiftDate.getMonth() + 1;
    let termSheetName = "";
    if (m >= 4 && m <= 9) termSheetName = `${y}上期時給`;
    else if (m >= 10 && m <= 12) termSheetName = `${y}下期時給`;
    else if (m >= 1 && m <= 3) termSheetName = `${y - 1}下期時給`;
    
    const cId = clinicIdMap.get(`${clinicName}_${dept}`);
    let sArr = [];
    const masterSheetData = getMasterData(termSheetName);
    
    if (masterSheetData && cId) {
      const mRow = masterSheetData.get(`${cId}_${dept}`);
      if (mRow) {
        const day = shiftDate.getDay();
        const isHol = isHoliday(shiftDate); 
        
        let baseIdx = 3; 
        if (day === 0 || isHol) baseIdx = 9; // 日・祝
        else if (day === 6) baseIdx = 6;     // 土
        
        const sAm = mRow[baseIdx], sPm = mRow[baseIdx + 1], sNight = mRow[baseIdx + 2];
        
        if (startMin < 780 && endMin > 540) sArr.push(sAm); 
        if (startMin < 1080 && endMin > 900) sArr.push(sPm); 
        if (startMin < 1260 && endMin > 1080) {              
          if (!sArr.includes(sNight)) sArr.push(sNight);
        }
      }
    }
    
    let salaryStr = sArr.join('/');
    const masterSalaryMax = sArr.length > 0 ? Math.max(...sArr.map(s => parseInt(String(s).replace(/,/g, ''), 10)).filter(s => !isNaN(s))) : 0;
    
    let isFoundInBosyu = false;
    let bosyuSalaryMax = 0;
    let exactMatchSalary = 0;
    let enclosedSalary = 0;
    
    if (hasBosyuSheet && cId) {
       const bList = bosyuMap.get(`${dateStr}_${cId}`) || [];
       for (const b of bList) {
           if (startMin === b.start && endMin === b.end) {
               isFoundInBosyu = true;
               if (b.salary > exactMatchSalary) exactMatchSalary = b.salary;
           }
           else if (b.start <= startMin && b.end >= endMin) {
               isFoundInBosyu = true;
               if (b.salary > enclosedSalary) enclosedSalary = b.salary;
           }
       }
       if (exactMatchSalary > 0) bosyuSalaryMax = exactMatchSalary;
       else if (enclosedSalary > 0) bosyuSalaryMax = enclosedSalary;
    }
    
    let phase = "";
    if (diffDays >= 7) phase = "Phaze.1";
    else if (diffDays >= 4 && diffDays <= 6) phase = "Phaze.2";
    else if (diffDays === 3) phase = "Phaze.3";
    else if (diffDays >= 0 && diffDays <= 2) phase = "Phaze.4";

    let isSpecialSalary = false;
    let memoText = "";
    if (hasBosyuSheet) {
      if (isFoundInBosyu) {
         if (bosyuSalaryMax > masterSalaryMax) {
             isSpecialSalary = true;
             salaryStr = String(bosyuSalaryMax);
         }
      } else {
         if (diffDays <= 3) {
             memoText = "要確認（直前キャンセル等）";
         }
      }
    }
    
    let preExt = "未実施", postEarly = "未実施";
    if (startMin <= 540) preExt = "調整余地なし"; 
    if (endMin >= 1260 || (endMin >= 1200 && clinicName.includes("北葛西"))) postEarly = "調整余地なし";
    
    if (!prevDoc) preExt = "調整余地なし";
    if (!nextDoc) postEarly = "調整余地なし";
    
    let newRow = new Array(headers.length).fill("");
    newRow[colIdx['日付']] = dateStr;
    newRow[colIdx['拠点名']] = clinicNameRaw;
    newRow[colIdx['不在時間']] = absenceTime;
    if(colIdx['掲載時給'] !== undefined) newRow[colIdx['掲載時給']] = salaryStr;
    newRow[colIdx['前の時間枠の医師']] = prevDoc;
    newRow[colIdx['後ろの時間枠の医師']] = nextDoc;
    if(colIdx['Phaze'] !== undefined) newRow[colIdx['Phaze']] = phase;
    
    if (existingRowsMap.has(key)) {
      const ext = existingRowsMap.get(key);
      const keepCols = ['担当', '対応状況', '個別交渉', '時給UP', 'お知らせ配信', '配信日']; 
      keepCols.forEach(cName => {
        if (colIdx[cName] !== undefined && ext[colIdx[cName]] !== "") {
          newRow[colIdx[cName]] = ext[colIdx[cName]];
        }
      });
      
      if (colIdx['作業メモ'] !== undefined) {
         let oldMemo = String(ext[colIdx['作業メモ']] || "");
         if (oldMemo === "要確認（直前キャンセル等）" && !memoText) {
             oldMemo = "";
         }
         newRow[colIdx['作業メモ']] = oldMemo || memoText;
      }

      if (colIdx['前医延長'] !== undefined) newRow[colIdx['前医延長']] = ext[colIdx['前医延長']] || preExt;
      if (colIdx['後医早出'] !== undefined) newRow[colIdx['後医早出']] = ext[colIdx['後医早出']] || postEarly;
      if (colIdx['特別時給'] !== undefined) newRow[colIdx['特別時給']] = isSpecialSalary ? "特別時給" : (ext[colIdx['特別時給']] || "通常時給");
      
      existingRowsMap.delete(key);
    } else {
      if (colIdx['前医延長'] !== undefined) newRow[colIdx['前医延長']] = preExt;
      if (colIdx['後医早出'] !== undefined) newRow[colIdx['後医早出']] = postEarly;
      if (colIdx['対応状況'] !== undefined) newRow[colIdx['対応状況']] = "未着手";
      if (colIdx['特別時給'] !== undefined) newRow[colIdx['特別時給']] = isSpecialSalary ? "特別時給" : "通常時給";
      if (colIdx['作業メモ'] !== undefined && memoText) newRow[colIdx['作業メモ']] = memoText;
    }
    
    rowsToWrite.push(newRow);
    
    let rowBg = new Array(headers.length).fill(null);
    if (!prevDoc && colIdx['前の時間枠の医師'] !== undefined) rowBg[colIdx['前の時間枠の医師']] = '#efefef';
    if (!nextDoc && colIdx['後ろの時間枠の医師'] !== undefined) rowBg[colIdx['後ろの時間枠の医師']] = '#efefef';
    backgroundsToWrite.push(rowBg);
  }
  
  // 6. アプローチシートのフルリセット（完全消去）とシートサイズの自動拡張
  const maxRows = approachSheet.getMaxRows();
  const maxCols = approachSheet.getMaxColumns();
  
  if (maxRows >= 3) {
    approachSheet.getRange(3, 1, maxRows - 2, maxCols).clear();
  }
  
  if (rowsToWrite.length > 0) {
    const requiredRows = rowsToWrite.length + 2; 
    if (requiredRows > maxRows) {
      approachSheet.insertRowsAfter(maxRows, requiredRows - maxRows);
    }
    
    const range = approachSheet.getRange(3, 1, rowsToWrite.length, headers.length);
    const templateRange = approachSheet.getRange(2, 1, 1, headers.length);
    
    templateRange.copyTo(range);
    range.setValues(rowsToWrite);
    
    const currentBgs = range.getBackgrounds();
    for (let r = 0; r < currentBgs.length; r++) {
      for (let c = 0; c < currentBgs[r].length; c++) {
        if (backgroundsToWrite[r][c] !== null) {
          currentBgs[r][c] = backgroundsToWrite[r][c];
        }
      }
    }
    range.setBackgrounds(currentBgs);
  }
}

/**
 * =========================================================
 * リアルタイム連動（シンプルトリガー）
 * =========================================================
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'アプローチシート') return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const value = e.value;
  
  if (row <= 2) return; 
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colName = headers[col - 1];
  
  // 担当者が入力されたら「対応中」
  if (colName === '担当') {
    const statusColIdx = headers.indexOf('対応状況') + 1;
    if (statusColIdx > 0 && value) {
      sheet.getRange(row, statusColIdx).setValue('対応中');
    }
  }
  
  // お知らせ配信が「済」ならタイムスタンプ
  if (colName === 'お知らせ配信') {
    const dateColIdx = headers.indexOf('配信日') + 1;
    if (dateColIdx > 0 && value === '済') {
      const now = new Date();
      const formatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      sheet.getRange(row, dateColIdx).setValue(formatted);
    }
  }
  
  // 【追加】対応状況が「完了」になった場合、即座に「ログ」シートへ転記し削除
  if (colName === '対応状況' && value === '完了') {
    const logSheet = e.source.getSheetByName('ログ');
    if (logSheet) {
      const lastCol = sheet.getLastColumn();
      const sourceRange = sheet.getRange(row, 1, 1, lastCol);
      const rowData = sourceRange.getValues()[0];
      const memoIdx = headers.indexOf('作業メモ');
      
      // ログシートの最終行の次へ、書式（プルダウン色等）ごと完全コピー
      const logLastRow = Math.max(logSheet.getLastRow(), 1);
      const targetRange = logSheet.getRange(logLastRow + 1, 1, 1, lastCol);
      sourceRange.copyTo(targetRange); 
      
      // コピー先の「作業メモ」末尾に完了タイムスタンプを追記
      if (memoIdx !== -1) {
        const now = new Date();
        const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
        let memo = rowData[memoIdx] ? String(rowData[memoIdx]) + "\n" : "";
        memo += `[完了: ${ts}]`;
        targetRange.getCell(1, memoIdx + 1).setValue(memo);
      }
      
      // アプローチシートから該当行を消去
      sheet.deleteRow(row);
    }
  }
}