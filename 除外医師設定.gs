/**
 * ============================================================================
 * 除外医師設定.gs
 * 特定の医師が確定シフト（貼付用・外部シート）に存在していても、
 * システム全体で「最初からいないもの（未充足）」として扱うための専用モジュール。
 *
 * 【将来不要になった場合の消し方】
 * 1. 一時停止したい場合：下の EXCLUDED_DOCTORS_CONFIG の配列を空 `[]` にする
 * 2. 完全撤去したい場合：このファイル（除外医師設定.gs）を削除するだけ！
 * ============================================================================
 */

const EXCLUDED_DOCTORS_CONFIG = [
  {
    name: '蛯原慎太郎', // スペースなしで記載（全角・半角スペースや「先生」は自動除去して判定します）
    licenseNo: '474101' // 医籍番号
  }
];

/**
 * 医師名または医籍番号が除外対象かどうかを判定する関数
 */
function isExcludedDoctor_(name, licenseNo) {
  if (!EXCLUDED_DOCTORS_CONFIG || EXCLUDED_DOCTORS_CONFIG.length === 0) return false;

  const cleanId = String(licenseNo || '').trim();
  const cleanName = String(name || '')
    .replace(/[\(（].*?[\)）]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();

  for (let i = 0; i < EXCLUDED_DOCTORS_CONFIG.length; i++) {
    const target = EXCLUDED_DOCTORS_CONFIG[i];
    if (target.licenseNo && cleanId && cleanId === target.licenseNo) return true;
    if (target.name && cleanName && (cleanName === target.name || cleanName.includes(target.name))) return true;
  }
  return false;
}

/**
 * 『貼付用』シートの1行データ（配列）から除外対象医師かを判定するヘルパー
 * （A列:0 = 名前, I列:8 = 医籍番号）
 */
function isExcludedDoctorInPasteRow_(row) {
  if (!row) return false;
  return isExcludedDoctor_(row[0], row[8]);
}

// ============================================================================
// 以下、他ファイルを直接改造せずに除外を適用するための自動中継（オーバーライド）処理
// ============================================================================

/**
 * 1. Main_データ読込.gs の拡張：
 * 貼付用・応募シフトの読み込み時に、除外医師の行を自動的にフィルタリングする
 */
function loadMasterAndExternalData(ss, sheets) {
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

  const EXT_SS_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  const extDataByDate = {};
  let isExtSsLoaded = false;
  const initExtDate = (dStr) => {
    if (!extDataByDate[dStr]) extDataByDate[dStr] = { paste: [], oubo: [], bosyu: [] };
  };

  const colMap = {
    paste: { doc: 0, clinicNo: 11, clinic: 12, dept: 13, date: 14, times: [15, 16, 17, 18, 19] },
    oubo: { doc: 0, clinicNo: -1, clinic: 3, dept: 4, date: 5, start: 6, end: 7 }
  };

  const clinicNameToNoMap = new Map();

  try {
    const extSs = SpreadsheetApp.openById(EXT_SS_ID);
    isExtSsLoaded = true;

    const pasteSheet = ss.getSheetByName('貼付用');
    if (pasteSheet && pasteSheet.getLastRow() > 2) {
      const pasteData = pasteSheet.getDataRange().getValues();
      colMap.paste.doc = findColIdx(pasteData, ['名前', '医師名', '氏名'], 0);
      const idxPasteLic = findColIdx(pasteData, ['医籍番号', '医籍登録番号'], 8);
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
        // ★除外対象医師は読み込まない
        if (isExcludedDoctor_(row[colMap.paste.doc], idxPasteLic !== -1 ? row[idxPasteLic] : '')) continue;

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
      const idxOuboLic = findColIdx(ouboData, ['医籍番号', '医籍登録番号'], -1);
      colMap.oubo.clinicNo = findColIdx(ouboData, ['クリニックNo', '拠点No', 'クリニックID'], -1);
      colMap.oubo.clinic = findColIdx(ouboData, ['クリニック名', '拠点名'], 3);
      colMap.oubo.dept = findColIdx(ouboData, ['診療科', '診療科目'], 4);
      colMap.oubo.date = findColIdx(ouboData, ['勤務日', '日付'], 5);
      colMap.oubo.start = findColIdx(ouboData, ['勤務開始時間', '開始時間'], 6);
      colMap.oubo.end = findColIdx(ouboData, ['勤務終了時間', '終了時間'], 7);

      for (let i = 1; i < ouboData.length; i++) {
        const row = ouboData[i];
        // ★除外対象医師は応募シフトからも除外
        if (isExcludedDoctor_(row[colMap.oubo.doc], idxOuboLic !== -1 ? row[idxOuboLic] : '')) continue;

        const rawDate = row[colMap.oubo.date];
        const dStr = fastFormatDate(parseDateToSafeDateObj(rawDate) || rawDate);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].oubo.push(row); }
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
 * 2. AI連携_3_連勤判定.gs の拡張：
 * 連勤カレンダー作成時に除外医師をスキップする
 */
function buildDoctorWorkCalendarFromSheets_(pasteSheet) {
  const cfg = getAiBridgeConfig_();
  const clinicNoMap = new Map();
  const docLicenseMap = new Map();
  const docDailyStatusMap = new Map();
  const docWorkCalendarMap = new Map();

  if (pasteSheet && pasteSheet.getLastRow() >= 3) {
    const pData = pasteSheet.getDataRange().getDisplayValues();
    for (let i = 2; i < pData.length; i++) {
      const row = pData[i];
      const rawDocName = String(row[0] || '').trim();
      const licNo = String(row[8] || '').trim();
      if (!rawDocName || isExcludedDoctor_(rawDocName, licNo)) continue;

      const docKey = normalizeDocKey_(rawDocName);
      const cNo = String(row[11] || '').trim();
      const cName = String(row[12] || '').trim();
      const baseC = cleanBaseClinicName_(cName);
      const dept = String(row[13] || '').trim();
      const normDate = normalizeBridgeDate_(row[14]);

      if (docKey && licNo && !docLicenseMap.has(docKey)) docLicenseMap.set(docKey, licNo);
      if (cNo && baseC) {
        if (dept) clinicNoMap.set(`${baseC}_${dept}`, cNo);
        if (!clinicNoMap.has(baseC)) clinicNoMap.set(baseC, cNo);
      }

      if (!normDate || !docKey || isExcludedFromStreakCount_(cNo, cName)) continue;

      const wageSumStr = String(row[29] || '').replace(/,/g, '').trim();
      const isFullTime = (wageSumStr !== '' && Number(wageSumStr) === 0);
      const staffComments = [row[44], row[45], row[46], row[47], row[48]].map(v => String(v || '')).join(' ');
      const isShoteiKyushutsu = staffComments.includes('所定休出');

      const statusKey = `${normDate}_${docKey}`;
      const prevStatus = docDailyStatusMap.get(statusKey) || {};
      docDailyStatusMap.set(statusKey, {
        isFullTime: Boolean(prevStatus.isFullTime || isFullTime),
        isShoteiKyushutsu: Boolean(prevStatus.isShoteiKyushutsu || isShoteiKyushutsu)
      });

      if (!docWorkCalendarMap.has(docKey)) {
        docWorkCalendarMap.set(docKey, {
          rawName: rawDocName.replace(/[\s\u3000]+/g, ''),
          licenseNo: licNo,
          isFullTimeEver: false,
          hasShoteiEver: false,
          workDates: new Set()
        });
      }
      const calEntry = docWorkCalendarMap.get(docKey);
      if (isFullTime) calEntry.isFullTimeEver = true;
      if (isShoteiKyushutsu) calEntry.hasShoteiEver = true;
      if (licNo && !calEntry.licenseNo) calEntry.licenseNo = licNo;
      calEntry.workDates.add(normDate);
    }
  }

  try {
    const extSs = SpreadsheetApp.openById(cfg.EXT_SHIFT_SS_ID);
    const fixedSheet = extSs.getSheetByName('確定シフト');
    if (fixedSheet && fixedSheet.getLastRow() >= 2) {
      const fData = fixedSheet.getDataRange().getDisplayValues();
      const headers = fData[0].map(h => String(h || '').trim());
      const iName = headers.indexOf('名前') !== -1 ? headers.indexOf('名前') : 0;
      const iLic = headers.indexOf('医籍番号');
      const iClinicNo = headers.indexOf('クリニックNo');
      const iClinic = headers.indexOf('クリニック名') !== -1 ? headers.indexOf('クリニック名') : 3;
      const iDate = headers.indexOf('勤務日') !== -1 ? headers.indexOf('勤務日') : 5;

      const now = new Date();
      const scanStartStr = Utilities.formatDate(new Date(now.getTime() - 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
      const scanEndStr = Utilities.formatDate(new Date(now.getTime() + 21 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

      for (let i = 1; i < fData.length; i++) {
        const r = fData[i];
        const dName = String(r[iName] || '').trim();
        const licNo = iLic >= 0 ? String(r[iLic] || '').trim() : '';
        if (!dName || isExcludedDoctor_(dName, licNo)) continue;

        const cNo = iClinicNo >= 0 ? String(r[iClinicNo] || '').trim() : '';
        const cName = String(r[iClinic] || '').trim();
        if (isExcludedFromStreakCount_(cNo, cName)) continue;

        const normDate = normalizeBridgeDate_(r[iDate]);
        if (!normDate || normDate < scanStartStr || normDate > scanEndStr) continue;

        const docKey = normalizeDocKey_(dName);
        if (docWorkCalendarMap.has(docKey)) {
          docWorkCalendarMap.get(docKey).workDates.add(normDate);
        }
      }
    }
  } catch (e) {
    Logger.log('外部確定シフトからの前月連勤補完スキップ: ' + e.message);
  }

  return { clinicNoMap, docLicenseMap, docDailyStatusMap, docWorkCalendarMap };
}

/**
 * 3. AI連携_暗号化送信2.gs（BQパック）の拡張：
 * 確定シフト・応募シフト抽出結果から除外医師を取り除く
 */
function buildBqShiftPayload_() {
  const ss = SpreadsheetApp.openById(BQ_SYNC_CONFIG.BQ_SHIFT_SS_ID);
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");

  const monthStartObj = new Date(now.getFullYear(), now.getMonth(), 1);
  const lookbackObj = new Date(now.getFullYear(), now.getMonth(), now.getDate() - BQ_SYNC_CONFIG.LOOKBACK_DAYS);
  const startDateObj = monthStartObj < lookbackObj ? monthStartObj : lookbackObj;
  const startDateStr = Utilities.formatDate(startDateObj, "Asia/Tokyo", "yyyy/MM/dd");

  const maxDateObj = new Date(now.getFullYear(), now.getMonth(), now.getDate() + BQ_SYNC_CONFIG.TARGET_DAYS);
  const maxDateStr = Utilities.formatDate(maxDateObj, "Asia/Tokyo", "yyyy/MM/dd");
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  const rawKakutei = extractShiftRows_(ss.getSheetByName("確定シフト"), startDateStr, maxDateStr, "kakutei");
  const rawObo     = extractShiftRows_(ss.getSheetByName("応募シフト"), todayStr, maxDateStr, "obo");
  const rawBosyu   = extractShiftRows_(ss.getSheetByName("募集シフト"), todayStr, maxDateStr, "bosyu");

  return {
    schemaVersion: 2,
    generatedAt: nowStr,
    baseDate: todayStr,
    startDate: startDateStr,
    maxDate: maxDateStr,
    kakutei: rawKakutei.filter(r => !isExcludedDoctor_(r.n, r.id)),
    obo:     rawObo.filter(r => !isExcludedDoctor_(r.n, r.id)),
    bosyu:   rawBosyu
  };
}

/**
 * 4. 勤務くん連携.gs の拡張：
 * 過去履歴スキャン時に除外医師を候補から外す
 */
function scanKinmukunHistoryOnce_(startDate, endDate, todayYmd, targetClinics, targetDates, agencyIds) {
  const statsByClinicDoc = new Map();
  const bookedMap = new Map();
  const startMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const endMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

  const scanBook = (ssId, isNew) => {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      ss.getSheets().forEach(sh => {
        const sName = sh.getName();
        if (!isNew) {
          if (!/^\d{4}\/\d{2}$/.test(sName)) return;
          const sd = new Date(sName + '/01');
          if (sd < startMonth || sd > endMonth) return;
        } else if (sName !== '確定シフト' && sName !== '応募シフト') return;

        const isObo = (sName === '応募シフト');
        const vals = sh.getDataRange().getValues();
        if (vals.length < 2) return;

        const h = vals[0];
        const iName = h.indexOf('名前'), iId = h.indexOf('医籍番号'), iKyoten = h.indexOf('クリニック名');
        const iShinryo = h.indexOf('診療科'), iDate = h.indexOf('勤務日');
        let iStart = h.indexOf('(元のシフト）勤務開始時間'), iEnd = h.indexOf('(元のシフト)勤務終了時間');
        if (iStart === -1) iStart = h.indexOf('勤務開始時間');
        if (iEnd === -1) iEnd = h.indexOf('勤務終了時間');
        const iWage = h.indexOf('時給合計'), iRem = h.indexOf('業務内容（備考）');
        if (iName === -1 || iId === -1 || iKyoten === -1 || iDate === -1) return;

        for (let i = 1; i < vals.length; i++) {
          const r = vals[i];
          const wd = new Date(r[iDate]);
          if (isNaN(wd.getTime()) || wd < startDate || wd > endDate) continue;

          const id = String(r[iId] || '').trim();
          const name = String(r[iName] || '').trim();
          if (!id || !name || isExcludedDoctor_(name, id)) continue;

          const kName = String(r[iKyoten] || '').trim();
          const shin = iShinryo !== -1 ? String(r[iShinryo] || '').trim() : '';
          const dispK = shin && (kName === '北葛西' || kName === '亀有') ? `${kName}（${shin}）` : kName;
          const dateStr = Utilities.formatDate(wd, 'Asia/Tokyo', 'yyyy/MM/dd');
          const fc = calcKinmukunFrames_(r[iStart], r[iEnd], kName === '北葛西');

          if (!isObo && targetDates.has(dateStr)) {
            const bKey = `${dateStr}_${id}`;
            if (!bookedMap.has(bKey)) bookedMap.set(bKey, []);
            bookedMap.get(bKey).push({ clinic: dispK, frames: fc });
          }

          if (!targetClinics.has(dispK)) continue;
          const wage = iWage !== -1 ? Number(r[iWage] || 0) : 0;
          const rem = iRem !== -1 ? String(r[iRem] || '') : '';
          if (wage === 0) continue;
          if (!isObo && /\d{4}\/\d{2}\/\d{2}/.test(rem) && rem.includes('【') && rem.includes('曜日')) continue;

          const shiftSum = fc.am + fc.pm1 + fc.pm2;
          if (shiftSum <= 0) continue;

          if (!statsByClinicDoc.has(dispK)) statsByClinicDoc.set(dispK, new Map());
          const dMap = statsByClinicDoc.get(dispK);
          if (!dMap.has(id)) {
            dMap.set(id, {
              id, name,
              attribute: agencyIds.has(id) ? '紹介会社' : 'スポット',
              totalCount: 0, lastWork: '-',
              totalFrames: { am: 0, pm1: 0, pm2: 0 },
              dayFrames: {
                0: { am: 0, pm1: 0, pm2: 0 }, 1: { am: 0, pm1: 0, pm2: 0 }, 2: { am: 0, pm1: 0, pm2: 0 },
                3: { am: 0, pm1: 0, pm2: 0 }, 4: { am: 0, pm1: 0, pm2: 0 }, 5: { am: 0, pm1: 0, pm2: 0 },
                6: { am: 0, pm1: 0, pm2: 0 }
              }
            });
          }
          const st = dMap.get(id);
          const dow = wd.getDay();
          st.totalCount += shiftSum;
          st.totalFrames.am += fc.am; st.totalFrames.pm1 += fc.pm1; st.totalFrames.pm2 += fc.pm2;
          st.dayFrames[dow].am += fc.am; st.dayFrames[dow].pm1 += fc.pm1; st.dayFrames[dow].pm2 += fc.pm2;
          if (dateStr < todayYmd && (st.lastWork === '-' || dateStr > st.lastWork)) st.lastWork = dateStr;
        }
      });
    } catch (e) {
      Logger.log(`シート読込スキップ (${ssId}): ${e.message}`);
    }
  };

  scanBook(KINMUKUN_BRIDGE_CFG.MONTHLY_2_SS_ID, false);
  scanBook(KINMUKUN_BRIDGE_CFG.SHIFT_NEW_SS_ID, true);

  return { statsByClinicDoc, bookedMap };
}