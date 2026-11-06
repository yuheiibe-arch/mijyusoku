/**
 * 外部データ・マスタデータの読み込み
 */
function loadMasterAndExternalData(ss, sheets) {
  // ヘッダー名から列インデックスを動的に検索する内部ヘルパー（ハードコード防止）
  const findColIdx = (rows, candidates, fallbackIdx = -1) => {
    if (!rows || rows.length === 0) return fallbackIdx;
    const maxRow = Math.min(2, rows.length);
    for (let r = 0; r < maxRow; r++) {
      const headers = rows[r].map(c => String(c || "").trim());
      for (const name of candidates) {
        const idx = headers.indexOf(name);
        if (idx !== -1) return idx;
      }
      for (const name of candidates) {
        const idx = headers.findIndex(h => h && h.includes(name));
        if (idx !== -1) return idx;
      }
    }
    return fallbackIdx;
  };

  // 条件に一致する複数列のインデックス配列を動的に取得する内部ヘルパー
  const findColIndices = (rows, matcherFn, fallbackIndices = []) => {
    if (!rows || rows.length === 0) return fallbackIndices;
    const maxRow = Math.min(2, rows.length);
    for (let r = 0; r < maxRow; r++) {
      const matched = [];
      rows[r].forEach((cell, idx) => {
        const h = String(cell || "").trim();
        if (h && matcherFn(h, idx)) matched.push(idx);
      });
      if (matched.length > 0) return matched;
    }
    return fallbackIndices;
  };

  // 追加データソース
  const EXT_SS_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  const extDataByDate = {}; 
  let isExtSsLoaded = false;
  const initExtDate = (dStr) => {
    if (!extDataByDate[dStr]) extDataByDate[dStr] = { paste: [], oubo: [], bosyu: [] };
  };

  // 動的特定した列インデックス（Main_レポート生成.gs へも引き渡す）
  const colMap = {
    paste: { doc: 0, clinicNo: 11, clinic: 12, dept: 13, date: 14, times: [15, 16, 17, 18, 19] },
    oubo: { doc: 0, clinicNo: -1, clinic: 3, dept: 4, date: 5, start: 6, end: 7 }
  };

  // 貼付用シートから「拠点名 ⇔ クリニックNo」の対応表を作成
  const clinicNameToNoMap = new Map();

  try {
    const extSs = SpreadsheetApp.openById(EXT_SS_ID);
    isExtSsLoaded = true;
    
    const pasteSheet = ss.getSheetByName('貼付用');
    if (pasteSheet && pasteSheet.getLastRow() > 2) {
      const pasteData = pasteSheet.getDataRange().getValues();
      colMap.paste.doc = findColIdx(pasteData, ['名前', '医師名', '氏名'], 0);
      colMap.paste.clinicNo = findColIdx(pasteData, ['クリニックNo', '拠点No', 'クリニックID'], 11);
      colMap.paste.clinic = findColIdx(pasteData, ['クリニック名', '拠点名'], 12);
      colMap.paste.dept = findColIdx(pasteData, ['診療科', '診療科目'], 13);
      colMap.paste.date = findColIdx(pasteData, ['勤務日', '日付'], 14);

      const detectedTimes = findColIndices(
        pasteData,
        (h) => /^(勤務開始|勤務終了|開始時間|終了時間|休憩開始|休憩終了)/.test(h) && !h.includes('元のシフト'),
        [15, 16, 17, 18, 19]
      );
      if (detectedTimes.length >= 2) colMap.paste.times = detectedTimes.slice(0, 5);

      for (let i = 2; i < pasteData.length; i++) {
        const row = pasteData[i];
        const rawDate = row[colMap.paste.date];
        const dStr = fastFormatDate(parseDateToSafeDateObj(rawDate) || rawDate);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].paste.push(row); }

        const cNo = colMap.paste.clinicNo !== -1 ? String(row[colMap.paste.clinicNo] || "").trim() : "";
        const cName = colMap.paste.clinic !== -1 ? String(row[colMap.paste.clinic] || "").trim() : "";
        const cDept = colMap.paste.dept !== -1 ? String(row[colMap.paste.dept] || "").trim() : "";
        if (cNo && cName && Number(cNo) < 1000) {
          const normC = normalizeClinicName(cName);
          clinicNameToNoMap.set(cName, cNo);
          clinicNameToNoMap.set(normC, cNo);
          if (cDept) {
            clinicNameToNoMap.set(`${cName}_${cDept}`, cNo);
            clinicNameToNoMap.set(`${normC}_${cDept}`, cNo);
            clinicNameToNoMap.set(`${normC}（${cDept}）`, cNo);
            clinicNameToNoMap.set(`${normC} (${cDept})`, cNo);
          }
        }
      }
    }

    const ouboSheet = extSs.getSheetByName('応募シフト');
    if (ouboSheet && ouboSheet.getLastRow() > 1) {
      const ouboData = ouboSheet.getDataRange().getValues();
      colMap.oubo.doc = findColIdx(ouboData, ['名前', '医師名', '氏名'], 0);
      colMap.oubo.clinicNo = findColIdx(ouboData, ['クリニックNo', '拠点No', 'クリニックID'], -1);
      colMap.oubo.clinic = findColIdx(ouboData, ['クリニック名', '拠点名'], 3);
      colMap.oubo.dept = findColIdx(ouboData, ['診療科', '診療科目'], 4);
      colMap.oubo.date = findColIdx(ouboData, ['勤務日', '日付'], 5);
      colMap.oubo.start = findColIdx(ouboData, ['勤務開始時間', '開始時間'], 6);
      colMap.oubo.end = findColIdx(ouboData, ['勤務終了時間', '終了時間'], 7);

      for (let i = 1; i < ouboData.length; i++) {
        const rawDate = ouboData[i][colMap.oubo.date];
        const dStr = fastFormatDate(parseDateToSafeDateObj(rawDate) || rawDate);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].oubo.push(ouboData[i]); }
      }
    }

    const bosyuSheet = extSs.getSheetByName('募集シフト');
    if (bosyuSheet && bosyuSheet.getLastRow() > 1) {
      const bosyuData = bosyuSheet.getDataRange().getValues();
      const idxBosyuDate = findColIdx(bosyuData, ['勤務日', '日付'], 3);
      for (let i = 1; i < bosyuData.length; i++) {
        const rawDate = bosyuData[i][idxBosyuDate];
        const dStr = fastFormatDate(parseDateToSafeDateObj(rawDate) || rawDate);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].bosyu.push(bosyuData[i]); }
      }
    }
  } catch (e) {
    Logger.log('追加データソースの読み込みに失敗しました: ' + e.message);
  }

  // COO室依頼２診要望データ（ヘッダー動的検索）
  const COO_SS_ID = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';
  const cooDataByDate = {};
  try {
    const cooSS = SpreadsheetApp.openById(COO_SS_ID);
    const cooSheet = cooSS.getSheetByName("２診要望一覧");
    if (cooSheet && cooSheet.getLastRow() > 1) {
      const cooAll = cooSheet.getDataRange().getDisplayValues();
      const idxClinic = findColIdx(cooAll, ['拠点名', 'クリニック名', '拠点'], 0);
      const idxDate = findColIdx(cooAll, ['日付', '勤務日'], 1);
      const idxStart = findColIdx(cooAll, ['開始時間', '勤務開始時間', '開始'], 2);
      const idxEnd = findColIdx(cooAll, ['終了時間', '勤務終了時間', '終了'], 3);
      const idxFlag = findColIdx(cooAll, ['判定', 'f判定', '要望', '備考', 'ステータス'], 4);

      for (let i = 1; i < cooAll.length; i++) {
        const row = cooAll[i];
        const dObj = parseDateToSafeDateObj(row[idxDate]);
        if (dObj) {
          const dStr = fastFormatDate(dObj);
          if (!cooDataByDate[dStr]) cooDataByDate[dStr] = [];
          cooDataByDate[dStr].push([row[idxClinic], row[idxDate], row[idxStart], row[idxEnd], row[idxFlag]]);
        }
      }
    }
  } catch (e) {
    Logger.log('COO要望一覧の読み込みに失敗しました: ' + e.message);
  }

  // 開院日マスタ・エリアマスタ（ヘッダー名からの動的検索 ＆ クリニックNo突合）
  const openDateMap = new Map();
  const openDateByNoMap = new Map();
  const areaMap = {}; 
  const areaByNoMap = new Map();

  try {
    const masterSS = SpreadsheetApp.openById('14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs');
    const extData = masterSS.getSheetByName('拠点名').getDataRange().getValues();
    
    const idxGroup = findColIdx(extData, ['拠点グループ', 'グループ'], 5);
    const idxDept = findColIdx(extData, ['診療科', '診療科目'], 6);
    const idxAreaCol = findColIdx(extData, ['エリア', '地域'], 7);
    const idxOpenDate = findColIdx(extData, ['開院日', 'オープン日'], 8);
    const idxClinicNo = findColIdx(extData, ['クリニックNo', '拠点No', 'クリニックID'], 10);

    const nameColIndices = findColIndices(
      extData,
      (h) => h.includes('正規記載') || h.includes('表記揺れ') || h === '拠点名' || h === 'クリニック名',
      [0, 1, 2, 3, 4]
    );

    for (let i = 1; i < extData.length; i++) {
      const row = extData[i];
      const groupName = idxGroup !== -1 && row[idxGroup] ? String(row[idxGroup]).trim() : "";
      const areaColStr = idxAreaCol !== -1 && row[idxAreaCol] ? String(row[idxAreaCol]).trim() : "";

      let area = "その他";
      if (groupName.includes("関東第一") || groupName.includes("関東第二") || groupName.includes("東京第一") || groupName.includes("東京第二") || groupName.includes("東京")) area = "東京";
      else if (groupName.includes("神奈川")) area = "神奈川";
      else if (groupName.includes("埼玉")) area = "埼玉";
      else if (groupName.includes("千葉")) area = "千葉";
      else if (groupName.includes("茨城")) area = "茨城";
      else if (groupName.includes("大阪") || groupName.includes("関西") || areaColStr.includes("関西")) area = "大阪";

      const openDateObj = idxOpenDate !== -1 ? parseDateToSafeDateObj(row[idxOpenDate]) : null;
      const cNo = idxClinicNo !== -1 && row[idxClinicNo] !== "" && row[idxClinicNo] !== null
        ? String(row[idxClinicNo]).trim()
        : "";
      const deptStr = idxDept !== -1 && row[idxDept] ? String(row[idxDept]).trim() : "";

      if (cNo) {
        if (openDateObj) {
          openDateByNoMap.set(cNo, openDateObj);
          if (deptStr) openDateByNoMap.set(`${cNo}_${deptStr}`, openDateObj);
        }
        if (area !== "その他" || !areaByNoMap.has(cNo)) {
          areaByNoMap.set(cNo, area);
        }
      }

      nameColIndices.forEach(colIdx => {
        const cleanName = row[colIdx] ? String(row[colIdx]).trim() : "";
        if (cleanName !== "") {
          const normName = normalizeClinicName(cleanName);
          [cleanName, normName].forEach(k => {
            if (!k) return;
            if (openDateObj) openDateMap.set(k, openDateObj);
            areaMap[k] = area;
            if (cNo && !clinicNameToNoMap.has(k)) clinicNameToNoMap.set(k, cNo);
          });
        }
      });
    }

    // 貼付用の「拠点名 ⇔ クリニックNo」を使って、クリニックNo基準の開院日・エリアをopenDateMap/areaMapへ完全反映
    clinicNameToNoMap.forEach((cNo, cName) => {
      if (openDateByNoMap.has(cNo)) {
        openDateMap.set(cName, openDateByNoMap.get(cNo));
      }
      if (areaByNoMap.has(cNo) && (!areaMap[cName] || areaMap[cName] === "その他")) {
        areaMap[cName] = areaByNoMap.get(cNo);
      }
    });
  } catch (e) {
    Logger.log('開院日マスタの読み込みに失敗: ' + e.message);
  }

  // 変則営業データ（ヘッダー動的検索）
  const irregularMap = {}; 
  if (sheets.irregular && sheets.irregular.getLastRow() > 1) {
    const iData = sheets.irregular.getDataRange().getValues();
    const idxIrrDate = findColIdx(iData, ['日付', '勤務日', '対象日'], 0);
    const idxIrrTime = findColIdx(iData, ['営業時間', '時間', '変則時間', '時間帯'], 1);
    const idxIrrClinic = findColIdx(iData, ['拠点名', 'クリニック名', '対象拠点', '拠点'], 2);

    for (let i = 1; i < iData.length; i++) {
      const row = iData[i];
      const dateObj = parseDateToSafeDateObj(row[idxIrrDate]); 
      const timeRangeStr = row[idxIrrTime];
      const clinicName = row[idxIrrClinic] ? String(row[idxIrrClinic]).trim() : "";

      if (dateObj && timeRangeStr && clinicName) {
        const dateKey = fastFormatDate(dateObj);
        const rangesStr = String(timeRangeStr).split(/[,/、]/);
        const validRanges = [];
        for (let rStr of rangesStr) {
          const parts = rStr.split('-');
          if (parts.length === 2) {
            const openMin = safeParseTime(parts[0]);
            const closeMin = safeParseTime(parts[1]);
            if (!isNaN(openMin) && !isNaN(closeMin) && openMin < closeMin) {
              validRanges.push({ open: openMin, close: closeMin });
            }
          }
        }
        
        if (validRanges.length > 0) {
           const normName = normalizeClinicName(clinicName);
           if (normName === "全拠点") {
               irregularMap[`${dateKey}_全拠点`] = validRanges;
           } else {
               irregularMap[`${dateKey}_${normName}`] = validRanges;
               irregularMap[`${dateKey}_${clinicName}`] = validRanges;
           }
        }
      }
    }
  }

  // 休館日データ（ヘッダー動的検索）
  const closedDataMap = new Map();
  if (sheets.closed && sheets.closed.getLastRow() > 1) {
    const cData = sheets.closed.getDataRange().getValues();
    const idxClosedDate = findColIdx(cData, ['日付', '休館日', '対象日'], 0);
    const idxClosedDept = findColIdx(cData, ['診療科', '科目'], 2);
    const idxClosedLoc = findColIdx(cData, ['拠点名', 'クリニック名', '拠点', '対象拠点'], 3);
    const idxClosedTime = findColIdx(cData, ['休館時間', '時間帯', '時間', '区分'], 4);

    for (let i = 1; i < cData.length; i++) {
      const row = cData[i];
      const dateObj = parseDateToSafeDateObj(row[idxClosedDate]);
      const cDept = idxClosedDept !== -1 && row[idxClosedDept] ? String(row[idxClosedDept]).trim() : "";
      const cLoc = idxClosedLoc !== -1 && row[idxClosedLoc] ? String(row[idxClosedLoc]).trim() : "";
      const cTime = idxClosedTime !== -1 && row[idxClosedTime] ? String(row[idxClosedTime]).trim() : "全日"; 
      
      if (dateObj && cLoc) {
        const dateKey = fastFormatDate(dateObj);
        if (cLoc === "全拠点") {
             closedDataMap.set(`${dateKey}_全拠点`, cTime);
        } else if (cDept === "内科") {
             closedDataMap.set(`${dateKey}_${normalizeClinicName(cLoc)}_内科`, cTime);
             closedDataMap.set(`${dateKey}_${cLoc}_内科`, cTime);
        } else {
             closedDataMap.set(`${dateKey}_${normalizeClinicName(cLoc)}`, cTime);
        }
      }
    }
  }

  return {
    extDataByDate,
    isExtSsLoaded,
    cooDataByDate,
    openDateMap,
    openDateByNoMap,
    clinicNameToNoMap,
    areaMap,
    areaByNoMap,
    irregularMap,
    closedDataMap,
    colMap
  };
}

/**
 * シフトと不在データの読み込み・整形
 */
function loadShiftAndAbsenceData(sourceSheet, ishiFuzaiSheet) {
  const findColIdx = (rows, candidates, fallbackIdx = -1) => {
    if (!rows || rows.length === 0) return fallbackIdx;
    const headers = rows[0].map(c => String(c || "").trim());
    for (const name of candidates) {
      const idx = headers.indexOf(name);
      if (idx !== -1) return idx;
    }
    for (const name of candidates) {
      const idx = headers.findIndex(h => h && h.includes(name));
      if (idx !== -1) return idx;
    }
    return fallbackIdx;
  };

  const EXCLUDED_KEYWORDS = ["有給", "欠勤", "院外勤務", "バックアップ", "医師会", "嘱託医", "出張インフルエンザワクチン"];
  const allClinics = new Set(); 
  const backupInfoMap = {}; 
  const rawNameMap = {}; 
  const workingClinics = {}; 
  
  const srcData = sourceSheet.getDataRange().getValues();
  const idxClinic = findColIdx(srcData, ['拠点名', 'クリニック名'], 0);
  const idxDate = findColIdx(srcData, ['勤務日', '日付'], 1);
  const idxDocA = findColIdx(srcData, ['Aシフト医師', 'Aシフト'], 7);
  const idxDocB = findColIdx(srcData, ['Bシフト医師', 'Bシフト'], 8);
  const idxDocC = findColIdx(srcData, ['Cシフト医師', 'Cシフト'], 9);
  const docCols = [idxDocA, idxDocB, idxDocC];

  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const name = row[idxClinic] ? String(row[idxClinic]).trim() : "";
    
    if (name === "【関東】バックアップシフト") {
        const dateObj = parseDateToSafeDateObj(row[idxDate]);
        if (dateObj) {
            const dateKey = fastFormatDate(dateObj);
            const slots = ["09:00~13:00", "15:00~18:00", "18:00~21:00"];
            let backups = [];
            for (let j = 0; j < slots.length; j++) {
                const doc = row[docCols[j]]; 
                if (doc && String(doc).trim()) backups.push(`${slots[j]}：${doc}先生（全拠点）`);
            }
            if (backups.length > 0) backupInfoMap[dateKey] = `【バックアップ】${backups.join('、')}`;
        }
        continue; 
    }

    if (!name || EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;
    const normName = normalizeClinicName(name);
    allClinics.add(normName);
    rawNameMap[normName] = name; 

    const dateObj = parseDateToSafeDateObj(row[idxDate]);
    if (dateObj) {
        const dateKey = fastFormatDate(dateObj);
        if (!workingClinics[dateKey]) workingClinics[dateKey] = new Set();
        workingClinics[dateKey].add(normName);
    }
  }
  
  const ishiFuzaiRawData = ishiFuzaiSheet.getDataRange().getValues();
  const idxFDate = findColIdx(ishiFuzaiRawData, ['日付', '勤務日'], 0);
  const idxFName = findColIdx(ishiFuzaiRawData, ['拠点名', 'クリニック名'], 1);
  const idxFTime = findColIdx(ishiFuzaiRawData, ['不在時間', '時間帯', '時間'], 2);
  const fuzaiRecords = []; 

  for (let i = 1; i < ishiFuzaiRawData.length; i++) { 
    const row = ishiFuzaiRawData[i];
    const dateValue = row[idxFDate];        
    const name = row[idxFName] ? row[idxFName].toString().trim() : ""; 
    const time = row[idxFTime] ? row[idxFTime].toString().trim() : "";

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