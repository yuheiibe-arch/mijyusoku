// ------------------------------------------------------------------------------------
// 「医師不在拠点」シートへの出力処理関数群
// ------------------------------------------------------------------------------------
function getAdjacentShiftDoctors(currentClinicWorkData, currentShiftKey, clinicShiftTimesForLookup, shiftOrder) {
  let prevDoctorsStr = "";
  let nextDoctorsStr = "";
  const currentShiftIndex = shiftOrder.indexOf(currentShiftKey);

  if (currentShiftIndex > 0) {
    const prevShiftKey = shiftOrder[currentShiftIndex - 1];
    if (clinicShiftTimesForLookup && clinicShiftTimesForLookup[prevShiftKey]) {
      const [prevStart, prevEnd] = clinicShiftTimesForLookup[prevShiftKey];
      const doctorsInPrev = [...new Set(currentClinicWorkData
        .filter(work => work.start < prevEnd && work.end > prevStart)
        .map(work => work.doctor))];
      prevDoctorsStr = doctorsInPrev.length > 0 ? `${doctorsInPrev.join(',')}：${formatMinutesToHHMM(prevStart)}-${formatMinutesToHHMM(prevEnd)}` : "未充足";
    }
  }

  if (currentShiftIndex < shiftOrder.length - 1) {
    const nextShiftKey = shiftOrder[currentShiftIndex + 1];
    if (clinicShiftTimesForLookup && clinicShiftTimesForLookup[nextShiftKey]) {
      const [nextStart, nextEnd] = clinicShiftTimesForLookup[nextShiftKey];
       const doctorsInNext = [...new Set(currentClinicWorkData
        .filter(work => work.start < nextEnd && work.end > nextStart)
        .map(work => work.doctor))];
      nextDoctorsStr = doctorsInNext.length > 0 ? `${doctorsInNext.join(',')}：${formatMinutesToHHMM(nextStart)}-${formatMinutesToHHMM(nextEnd)}` : "未充足";
    }
  }
  return { prevDoctorsStr, nextDoctorsStr };
}

function generateDoctorAbsenceReportWithContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  const targetSheet = ss.getSheetByName("医師不在拠点");

  if (!sourceSheet || !targetSheet) return;

  const header = [["日付", "拠点名", "不在時間", "前の時間枠の医師", "後ろの時間枠の医師"]];
  targetSheet.getRange(1, 1, 1, header[0].length).setValues(header);
  const lastTargetRow = targetSheet.getLastRow();
  if (lastTargetRow > 1) {
    targetSheet.getRange(2, 1, lastTargetRow - 1, targetSheet.getMaxColumns()).clearContent();
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const sourceData = sourceSheet.getDataRange().getValues();
  const processedWorkData = {};

  // 爆速版：中抜けインターバル抽出
  const getIntervalsFast = (row) => {
    const p = (t) => {
      if (!t) return NaN;
      if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
      if (typeof t === 'string') {
        const pts = t.trim().split(':');
        if (pts.length >= 2) return parseInt(pts[0], 10) * 60 + parseInt(pts[1], 10);
      }
      if (typeof t === 'number' && t >= 0 && t < 1) return Math.round(t * 24 * 60);
      return NaN;
    };
    const times = [];
    for (let i = 15; i <= 19; i++) {
      const val = p(row[i]);
      if (!isNaN(val)) times.push(val);
    }
    const len = times.length;
    if (len === 4) {
      const res = [];
      if (times[0] < times[1]) res.push({ start: times[0], end: times[1] });
      if (times[2] < times[3]) res.push({ start: times[2], end: times[3] });
      return res;
    } else if (len === 2) {
      if (times[0] < times[1]) return [{ start: times[0], end: times[1] }];
    } else if (len > 0) {
      const start = times[0], end = times[len - 1];
      if (start < end) return [{ start: start, end: end }];
    }
    return [];
  };

  for (let i = 2; i < sourceData.length; i++) {
    const row = sourceData[i];
    const doctorName = row[0] ? String(row[0]).trim() : "不明医師";
    const clinicName = row[12] ? String(row[12]).trim() : "";
    const department = row[13] ? String(row[13]).trim() : "";
    if (!row[14] || !clinicName) continue;
    if (GLOBAL_EXCLUDED_LOCATIONS.includes(clinicName) || GLOBAL_EXCLUDED_DEPARTMENTS.includes(department)) continue;

    const dateObj = parseDateToSafeDateObj(row[14]); 
    if (!dateObj) continue;
    
    const dateKey = fastFormatDate(dateObj); 

    const aggregationKey = `${clinicName}_${department || ""}`;
    
    if (!processedWorkData[dateKey]) processedWorkData[dateKey] = {};
    if (!processedWorkData[dateKey][aggregationKey]) processedWorkData[dateKey][aggregationKey] = [];

    const intervals = getIntervalsFast(row);
    
    intervals.forEach(inv => {
        processedWorkData[dateKey][aggregationKey].push({ doctor: doctorName, start: inv.start, end: inv.end });
    });
  }

  const unfulfilledShiftsList = []; 
  const sortedDates = Object.keys(processedWorkData).sort((a, b) => new Date(a) - new Date(b));

  for (const dateKey of sortedDates) {
    const currentDateObj = parseDateToSafeDateObj(dateKey); 
    if (!currentDateObj || currentDateObj < today) continue;

    const dailyWorkDataByAggregationKey = processedWorkData[dateKey];
    for (const aggregationKey in dailyWorkDataByAggregationKey) {
      const [currentClinicName, currentDepartment] = aggregationKey.split('_');
      const workIntervalsForClinic = dailyWorkDataByAggregationKey[aggregationKey];

      let shiftLookupKey = "その他";
      const specificKeyWithDept = `${currentClinicName}${currentDepartment}`;
      if (GLOBAL_SHIFT_TIMES_MIN.hasOwnProperty(specificKeyWithDept)) shiftLookupKey = specificKeyWithDept;
      else if (GLOBAL_SHIFT_TIMES_MIN.hasOwnProperty(currentClinicName)) shiftLookupKey = currentClinicName;
      
      const currentClinicShifts = GLOBAL_SHIFT_TIMES_MIN[shiftLookupKey] || GLOBAL_SHIFT_TIMES_MIN["その他"];

      for (const shiftKey of GLOBAL_STANDARD_SHIFT_ORDER) {
        if (!currentClinicShifts[shiftKey]) continue;
        const [shiftStartTarget, shiftEndTarget] = currentClinicShifts[shiftKey];
        
        const mergedIntervals = mergeIntervals(workIntervalsForClinic
          .map(interval => ({ start: Math.max(interval.start, shiftStartTarget), end: Math.min(interval.end, shiftEndTarget) }))
          .filter(interval => interval.start < interval.end));
        
        let currentTimePointer = shiftStartTarget;
        let gaps = [];

        for (const merged of mergedIntervals) {
            if (currentTimePointer < merged.start) {
                gaps.push(`${formatMinutesToHHMM(currentTimePointer)}-${formatMinutesToHHMM(merged.start)}`);
            }
            currentTimePointer = Math.max(currentTimePointer, merged.end);
        }
        if (currentTimePointer < shiftEndTarget) {
            gaps.push(`${formatMinutesToHHMM(currentTimePointer)}-${formatMinutesToHHMM(shiftEndTarget)}`);
        }

        if (gaps.length > 0) {
          unfulfilledShiftsList.push({ 
              dateKey, aggregationKey, standardShiftKey: shiftKey, preciseAbsenceStr: gaps.join(", ")
          });
        }
      }
    }
  }

  const outputDataRows = [];
  for (const {dateKey, aggregationKey, standardShiftKey, preciseAbsenceStr} of unfulfilledShiftsList) {
    const [clinicName, department] = aggregationKey.split('_');
    let outputClinicName = GLOBAL_TARGET_CLINICS_FOR_DEPT_INFO.includes(clinicName) && department ? `${clinicName} (${department})` : clinicName;
    
    let shiftLookupKey = "その他";
    if (GLOBAL_SHIFT_TIMES_MIN.hasOwnProperty(`${clinicName}${department}`)) shiftLookupKey = `${clinicName}${department}`;
    else if (GLOBAL_SHIFT_TIMES_MIN.hasOwnProperty(clinicName)) shiftLookupKey = clinicName;
    
    const { prevDoctorsStr, nextDoctorsStr } = getAdjacentShiftDoctors(
      processedWorkData[dateKey]?.[aggregationKey] || [], standardShiftKey,
      GLOBAL_SHIFT_TIMES_MIN[shiftLookupKey] || GLOBAL_SHIFT_TIMES_MIN["その他"], GLOBAL_STANDARD_SHIFT_ORDER
    );

    outputDataRows.push([dateKey, outputClinicName, preciseAbsenceStr, prevDoctorsStr, nextDoctorsStr]);
  }
  
  if (outputDataRows.length > 0) {
    outputDataRows.sort((a, b) => new Date(a[0]) - new Date(b[0]) || a[1].localeCompare(b[1]));
    targetSheet.getRange(2, 1, outputDataRows.length, outputDataRows[0].length).setValues(outputDataRows).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } else {
    targetSheet.getRange(2,1).setValue("該当する不在情報はありませんでした。(今日以降)");
  }
  ss.toast("「医師不在拠点」シートの更新が完了しました。", "完了", 3);
}

// ------------------------------------------------------------------------------------
// 安全対策用の必須ヘルパー関数
// ------------------------------------------------------------------------------------
function mergeIntervals(intervals) {
  if (!intervals || intervals.length === 0) return [];
  intervals.sort((a, b) => a.start - b.start);
  const merged = [];
  let currentMerge = { ...intervals[0] };
  for (let i = 1; i < intervals.length; i++) {
    const nextInterval = intervals[i];
    if (nextInterval.start <= currentMerge.end) {
      currentMerge.end = Math.max(currentMerge.end, nextInterval.end);
    } else {
      merged.push(currentMerge);
      currentMerge = { ...nextInterval };
    }
  }
  merged.push(currentMerge);
  return merged;
}

function parseTimeToMinutes(timeInput) {
  if (!timeInput) return NaN;
  if (timeInput instanceof Date) return timeInput.getHours() * 60 + timeInput.getMinutes();
  if (typeof timeInput === 'string') {
    const parts = timeInput.trim().split(':');
    if (parts.length >= 2) return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  }
  if (typeof timeInput === 'number' && timeInput >= 0 && timeInput < 1) return Math.round(timeInput * 24 * 60);
  return NaN;
}

function formatMinutesToHHMM(totalMinutes) {
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}