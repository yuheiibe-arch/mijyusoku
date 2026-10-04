/**
 * 外部データ・マスタデータの読み込み
 */
function loadMasterAndExternalData(ss, sheets) {
  // 追加データソース
  const EXT_SS_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  const extDataByDate = {}; 
  let isExtSsLoaded = false;
  const initExtDate = (dStr) => {
    if (!extDataByDate[dStr]) extDataByDate[dStr] = { paste: [], oubo: [], bosyu: [] };
  };

  try {
    const extSs = SpreadsheetApp.openById(EXT_SS_ID);
    isExtSsLoaded = true;
    
    const pasteSheet = ss.getSheetByName('貼付用');
    if (pasteSheet && pasteSheet.getLastRow() > 2) {
      const pasteData = pasteSheet.getDataRange().getValues();
      for (let i = 2; i < pasteData.length; i++) {
        const dStr = fastFormatDate(parseDateToSafeDateObj(pasteData[i][14]) || pasteData[i][14]);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].paste.push(pasteData[i]); }
      }
    }

    const ouboSheet = extSs.getSheetByName('応募シフト');
    if (ouboSheet && ouboSheet.getLastRow() > 1) {
      const ouboData = ouboSheet.getDataRange().getValues();
      for (let i = 1; i < ouboData.length; i++) {
        const dStr = fastFormatDate(ouboData[i][5]);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].oubo.push(ouboData[i]); }
      }
    }

    const bosyuSheet = extSs.getSheetByName('募集シフト');
    if (bosyuSheet && bosyuSheet.getLastRow() > 1) {
      const bosyuData = bosyuSheet.getDataRange().getValues();
      for (let i = 1; i < bosyuData.length; i++) {
        const dStr = fastFormatDate(bosyuData[i][3]);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].bosyu.push(bosyuData[i]); }
      }
    }
  } catch (e) {
    Logger.log('追加データソースの読み込みに失敗しました: ' + e.message);
  }

  // COO室依頼２診要望データ
  const COO_SS_ID = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';
  const cooDataByDate = {};
  try {
    const cooSS = SpreadsheetApp.openById(COO_SS_ID);
    const cooSheet = cooSS.getSheetByName("２診要望一覧");
    if (cooSheet && cooSheet.getLastRow() > 1) {
      const cooRaw = cooSheet.getDataRange().getDisplayValues().slice(1);
      cooRaw.forEach(row => {
        const clinic = String(row[0] || "").replace(/[（(]小児科[）)]/, "").trim();
        const dStrRaw = String(row[1] || "").trim().split(/[（(]/)[0].trim().replace(/-/g, '/');
        const parts = dStrRaw.split('/');
        if (parts.length >= 3) {
          const dObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          if (!isNaN(dObj.getTime())) {
            const dStr = fastFormatDate(dObj);
            if (!cooDataByDate[dStr]) cooDataByDate[dStr] = [];
            cooDataByDate[dStr].push(row);
          }
        }
      });
    }
  } catch (e) {
    Logger.log('COO要望一覧の読み込みに失敗しました: ' + e.message);
  }

  // 開院日マスタ・エリアマスタ
  const openDateMap = new Map();
  const areaMap = {}; 
  try {
    const masterSS = SpreadsheetApp.openById('14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs');
    const extData = masterSS.getSheetByName('拠点名').getDataRange().getValues();
    
    for (let i = 1; i < extData.length; i++) {
      const groupName = extData[i][5] ? String(extData[i][5]).trim() : "";
      let area = "その他";
      if (groupName.includes("関東第一") || groupName.includes("関東第二") || groupName.includes("東京第一") || groupName.includes("東京第二")) area = "東京";
      else if (groupName.includes("神奈川")) area = "神奈川";
      else if (groupName.includes("埼玉")) area = "埼玉";
      else if (groupName.includes("千葉")) area = "千葉";
      else if (groupName.includes("茨城")) area = "茨城";
      else if (groupName.includes("大阪") || groupName.includes("関西")) area = "大阪";

      const openDate = extData[i][7]; 
      const namesToRegister = [extData[i][0], extData[i][1], extData[i][2], extData[i][3], extData[i][4]];
      
      namesToRegister.forEach(name => {
        const cleanName = name ? String(name).trim() : "";
        if (cleanName !== "") {
          if (openDate instanceof Date) openDateMap.set(cleanName, openDate);
          areaMap[cleanName] = area;
        }
      });
    }
  } catch (e) {
    Logger.log('開院日マスタの読み込みに失敗: ' + e.message);
  }

  // 変則営業データ
  const irregularMap = {}; 
  if (sheets.irregular) {
    const iData = sheets.irregular.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      const row = iData[i];
      const dateObj = parseDateToSafeDateObj(row[0]); 
      const timeRangeStr = row[1];
      const clinicName = row[2] ? String(row[2]).trim() : "";

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

  // 休館日データ
  const closedDataMap = new Map();
  if (sheets.closed) {
    const cData = sheets.closed.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const row = cData[i];
      const dateObj = parseDateToSafeDateObj(row[0]);
      const cDept = row[2] ? String(row[2]).trim() : "";
      const cLoc = row[3] ? String(row[3]).trim() : "";
      const cTime = row[4] ? String(row[4]).trim() : "全日"; 
      
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

  return { extDataByDate, isExtSsLoaded, cooDataByDate, openDateMap, areaMap, irregularMap, closedDataMap };
}

/**
 * シフトと不在データの読み込み・整形
 */
function loadShiftAndAbsenceData(sourceSheet, ishiFuzaiSheet) {
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