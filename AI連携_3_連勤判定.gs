/**
 * ============================================================================
 * AI連携_3_連勤判定.gs（勤務カレンダー構築・連続勤務NG判定）
 * ============================================================================
 */

function isExcludedFromStreakCount_(clinicNo, clinicName) {
  const cfg = getAiBridgeConfig_();
  const cNo = String(clinicNo || '').trim();
  const cName = String(clinicName || '').trim();
  if (!cName) return true;

  if (cNo && (cfg.EXCLUDED_CLINIC_NOS.has(cNo) || Number(cNo) >= 1000)) {
    return true;
  }
  for (const exKeyword of cfg.EXCLUDED_CLINIC_NAMES) {
    if (cName.includes(exKeyword)) return true;
  }
  return false;
}

function buildDoctorWorkCalendarFromSheets_(pasteSheet) {
  const cfg = getAiBridgeConfig_();
  const clinicNoMap = new Map();
  const docLicenseMap = new Map();
  const docDailyStatusMap = new Map();
  const docWorkCalendarMap = new Map();

  if (pasteSheet && pasteSheet.getLastRow() >= 3) {
    const pData = pasteSheet.getDataRange().getDisplayValues();
    const colDocName = 0;     // 名前
    const colLicenseNo = 8;   // 医籍番号
    const colClinicNo = 11;   // クリニックNo
    const colClinicName = 12; // クリニック名
    const colDept = 13;       // 診療科
    const colDate = 14;       // 勤務日
    const colWageSum = 29;    // 時給合計

    for (let i = 2; i < pData.length; i++) {
      const row = pData[i];
      const rawDocName = String(row[colDocName] || '').trim();
      if (!rawDocName) continue;

      const docKey = normalizeDocKey_(rawDocName);
      const licNo = String(row[colLicenseNo] || '').trim();
      const cNo = String(row[colClinicNo] || '').trim();
      const cName = String(row[colClinicName] || '').trim();
      const baseC = cleanBaseClinicName_(cName);
      const dept = String(row[colDept] || '').trim();
      const normDate = normalizeBridgeDate_(row[colDate]);

      if (docKey && licNo && !docLicenseMap.has(docKey)) {
        docLicenseMap.set(docKey, licNo);
      }
      if (cNo && baseC) {
        if (dept) clinicNoMap.set(`${baseC}_${dept}`, cNo);
        if (!clinicNoMap.has(baseC)) clinicNoMap.set(baseC, cNo);
      }

      if (!normDate || !docKey) continue;
      if (isExcludedFromStreakCount_(cNo, cName)) continue;

      const wageSumStr = String(row[colWageSum] || '').replace(/,/g, '').trim();
      const wageSumVal = wageSumStr !== '' ? Number(wageSumStr) : -1;
      const isFullTime = (wageSumVal === 0);

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

  // 外部『確定シフト』からの前月後半補完
  try {
    const extSs = SpreadsheetApp.openById(cfg.EXT_SHIFT_SS_ID);
    const fixedSheet = extSs.getSheetByName('確定シフト');
    if (fixedSheet && fixedSheet.getLastRow() >= 2) {
      const fData = fixedSheet.getDataRange().getDisplayValues();
      const headers = fData[0].map(h => String(h || '').trim());
      const iName = headers.indexOf('名前') !== -1 ? headers.indexOf('名前') : 0;
      const iClinicNo = headers.indexOf('クリニックNo');
      const iClinic = headers.indexOf('クリニック名') !== -1 ? headers.indexOf('クリニック名') : 3;
      const iDate = headers.indexOf('勤務日') !== -1 ? headers.indexOf('勤務日') : 5;

      const now = new Date();
      const scanStartStr = Utilities.formatDate(new Date(now.getTime() - 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
      const scanEndStr = Utilities.formatDate(new Date(now.getTime() + 21 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

      for (let i = 1; i < fData.length; i++) {
        const r = fData[i];
        const dName = String(r[iName] || '').trim();
        if (!dName) continue;
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

function buildConsecutiveWorkWarnings_(docWorkCalendarMap) {
  const warnings = [];
  const now = new Date();
  const todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
  const limitStr = Utilities.formatDate(new Date(now.getTime() + 16 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

  const toDateMs = (ymd) => {
    const p = ymd.split('/').map(Number);
    return new Date(p[0], p[1] - 1, p[2]).getTime();
  };
  const toYmd = (ms) => Utilities.formatDate(new Date(ms), 'Asia/Tokyo', 'yyyy/MM/dd');

  for (const [, info] of docWorkCalendarMap.entries()) {
    const sortedDates = Array.from(info.workDates).sort();
    if (sortedDates.length < 6) continue;

    const blocks = [];
    let curStart = sortedDates[0], curEnd = sortedDates[0], curLen = 1;

    for (let i = 1; i < sortedDates.length; i++) {
      const diffDays = Math.round((toDateMs(sortedDates[i]) - toDateMs(curEnd)) / 86400000);
      if (diffDays === 1) {
        curEnd = sortedDates[i];
        curLen++;
      } else {
        blocks.push({ start: curStart, end: curEnd, len: curLen });
        curStart = sortedDates[i];
        curEnd = sortedDates[i];
        curLen = 1;
      }
    }
    blocks.push({ start: curStart, end: curEnd, len: curLen });

    for (let b = 0; b < blocks.length; b++) {
      const blk = blocks[b];
      if (blk.end < todayStr || blk.start > limitStr) continue;

      const prevDayYmd = toYmd(toDateMs(blk.start) - 86400000);
      const nextDayYmd = toYmd(toDateMs(blk.end) + 86400000);

      if (blk.len >= 11) {
        warnings.push({
          licenseNo: info.licenseNo || '',
          doctor: info.rawName,
          isFullTime: info.isFullTimeEver,
          hasShotei: info.hasShoteiEver,
          type: 'STREAK',
          streakDays: blk.len,
          period: `${blk.start}〜${blk.end}`,
          ngDates: `${prevDayYmd} / ${nextDayYmd}`,
          note: `🚫前日(${prevDayYmd})・翌日(${nextDayYmd})は12連勤到達のため打診NG`
        });
      }

      if (b + 1 < blocks.length) {
        const nextBlk = blocks[b + 1];
        const gapDays = Math.round((toDateMs(nextBlk.start) - toDateMs(blk.end)) / 86400000) - 1;
        if (gapDays === 1 && (blk.len + 1 + nextBlk.len) >= 12) {
          const bridgeDate = toYmd(toDateMs(blk.end) + 86400000);
          if (bridgeDate >= todayStr && bridgeDate <= limitStr) {
            warnings.push({
              licenseNo: info.licenseNo || '',
              doctor: info.rawName,
              isFullTime: info.isFullTimeEver,
              hasShotei: info.hasShoteiEver,
              type: 'BRIDGE',
              streakDays: blk.len + 1 + nextBlk.len,
              period: `${blk.start}〜${blk.end} [中1日:${bridgeDate}] ${nextBlk.start}〜${nextBlk.end}`,
              ngDates: bridgeDate,
              note: `🚫${bridgeDate} に勤務を入れると前後が繋がり ${blk.len + 1 + nextBlk.len}連勤 になるため打診NG`
            });
          }
        }
      }
    }
  }
  return warnings;
}