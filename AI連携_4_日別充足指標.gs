/**
 * ============================================================================
 * AI連携_4_日別充足指標.gs（要望2：当月1日〜月末・2週間先までの日別充足指標算出）
 * ============================================================================
 */

function buildStructuredDailyMetrics_(ss) {
  const sheets = {
    source: ss.getSheetByName('確認用'),
    ishiFuzai: ss.getSheetByName('医師不在拠点'),
    closed: ss.getSheetByName('休館日'),
    irregular: ss.getSheetByName('変則営業')
  };
  if (!sheets.source || !sheets.ishiFuzai) return {};
  if (typeof loadMasterAndExternalData !== 'function' || typeof calculateAdvancedMetricsFast !== 'function') {
    return {};
  }

  const masterData = loadMasterAndExternalData(ss, sheets);
  const shiftData = (typeof loadShiftAndAbsenceData_New_ === 'function')
    ? loadShiftAndAbsenceData_New_(sheets.source, sheets.ishiFuzai)
    : loadShiftAndAbsenceData(sheets.source, sheets.ishiFuzai);

  const now = new Date();
  const loopStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const twoWeeksAhead = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 14);
  const loopEnd = monthEnd > twoWeeksAhead ? monthEnd : twoWeeksAhead;

  const pasteCols = (masterData.colMap && masterData.colMap.paste)
    ? masterData.colMap.paste
    : { doc: 0, clinicNo: 11, clinic: 12, dept: 13, date: 14, times: [15, 16, 17, 18, 19] };
  const ouboCols = (masterData.colMap && masterData.colMap.oubo)
    ? masterData.colMap.oubo
    : { doc: 0, clinicNo: -1, clinic: 3, dept: 4, date: 5, start: 6, end: 7 };

  const isUnopenedClinic = (clinicName, targetDate, explicitClinicNo = '') => {
    const cleanName = String(clinicName || '').trim();
    const normName = normalizeClinicName(cleanName);
    const baseName = cleanName.replace(/[\(（].*?[\)）]/g, '').trim();

    let cNo = String(explicitClinicNo || '').trim();
    if (!cNo && masterData.clinicNameToNoMap) {
      cNo = masterData.clinicNameToNoMap.get(cleanName)
         || masterData.clinicNameToNoMap.get(normName)
         || masterData.clinicNameToNoMap.get(baseName)
         || '';
    }

    let openDate = null;
    if (cNo && masterData.openDateByNoMap && masterData.openDateByNoMap.has(cNo)) {
      openDate = masterData.openDateByNoMap.get(cNo);
    } else if (masterData.openDateMap) {
      openDate = masterData.openDateMap.get(cleanName)
              || masterData.openDateMap.get(normName)
              || masterData.openDateMap.get(baseName)
              || null;
    }
    if (!openDate) return false;

    const openZero = new Date(openDate.getFullYear(), openDate.getMonth(), openDate.getDate());
    const checkZero = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    return checkZero < openZero;
  };

  const getIntervalsFast = (row) => {
    const times = [];
    for (const i of pasteCols.times) {
      const val = safeParseTime(row[i]);
      if (!isNaN(val)) times.push(val);
    }
    if (times.length === 4) {
      const res = [];
      if (times[0] < times[1]) res.push({ start: times[0], end: times[1] });
      if (times[2] < times[3]) res.push({ start: times[2], end: times[3] });
      return res;
    } else if (times.length >= 2) {
      const start = times[0], end = times[times.length - 1];
      if (start < end) return [{ start, end }];
    }
    return [];
  };

  const metricsMap = {};

  for (let d = new Date(loopStart); d <= loopEnd; d.setDate(d.getDate() + 1)) {
    const dateKey = fastFormatDate(d);
    const isSpecialDay = (d.getMonth() + 1 === 12 && d.getDate() === 31) || (d.getMonth() + 1 === 1 && d.getDate() <= 3);

    // 1診目必要時間の算出（小児科）
    let totalRequiredMinutes = 0;
    shiftData.sortedClinicList.forEach(clinic => {
      if (isUnopenedClinic(clinic, d) || clinic.includes('内科')) return;
      const closedTime = masterData.closedDataMap.get(`${dateKey}_${clinic}`) || masterData.closedDataMap.get(`${dateKey}_全拠点`);
      if (closedTime === '全日') return;

      let reqMin = isSpecialDay ? 420 : (clinic.includes('北葛西') ? 540 : 600);
      if (closedTime === '午前') reqMin -= 240;
      else if (closedTime === '午後' || closedTime === '夜間') reqMin -= 180;
      else if (closedTime === '午後夜間') reqMin -= 360;

      const irregularRule = masterData.irregularMap[`${dateKey}_${clinic}`] || masterData.irregularMap[`${dateKey}_全拠点`];
      if (irregularRule) {
        reqMin = 0;
        irregularRule.forEach(r => { reqMin += (r.close - r.open); });
      }
      totalRequiredMinutes += Math.max(0, reqMin);
    });

    // 不在時間の算出
    let totalGapMinutes = 0;
    const absenceDetails = [];
    const dailyRecords = shiftData.fuzaiRecords.filter(r => r.dateKey === dateKey);
    const groupedFuzai = {};

    dailyRecords.forEach(r => {
      if (isUnopenedClinic(r.normName, d)) return;
      if (!groupedFuzai[r.normName]) groupedFuzai[r.normName] = [];
      groupedFuzai[r.normName].push(r.timeRaw);
    });

    shiftData.sortedClinicList.forEach(clinic => {
      if (isUnopenedClinic(clinic, d) || clinic.includes('内科')) return;
      const closedTime = masterData.closedDataMap.get(`${dateKey}_${clinic}`) || masterData.closedDataMap.get(`${dateKey}_全拠点`);
      if (closedTime === '全日') return;

      const hasFuzai = groupedFuzai[clinic] && groupedFuzai[clinic].length > 0;
      const isWorking = shiftData.workingClinics[dateKey] && shiftData.workingClinics[dateKey].has(clinic);
      if (!hasFuzai && !isWorking) {
        groupedFuzai[clinic] = isSpecialDay
          ? ['10:00-19:00']
          : (clinic.includes('北葛西') ? ['09:00-13:00', '15:00-20:00'] : ['09:00-13:00', '15:00-21:00']);
      }
    });

    for (const normName in groupedFuzai) {
      if (isUnopenedClinic(normName, d)) continue;
      const closedTime = normName.includes('内科')
        ? (masterData.closedDataMap.get(`${dateKey}_${normName}`) || masterData.closedDataMap.get(`${dateKey}_${normName}_内科`) || masterData.closedDataMap.get(`${dateKey}_全拠点`))
        : (masterData.closedDataMap.get(`${dateKey}_${normName}`) || masterData.closedDataMap.get(`${dateKey}_全拠点`));
      if (closedTime === '全日') continue;

      const mergedTimeStr = safeNormalizeAndMerge(groupedFuzai[normName], isSpecialDay);
      if (!mergedTimeStr) continue;

      let finalTimeStr = closedTime ? safeRemoveClosed(mergedTimeStr, closedTime) : mergedTimeStr;
      if (!finalTimeStr) continue;
      if (!isSpecialDay && normName.includes('北葛西')) finalTimeStr = finalTimeStr.replace('21:00', '20:00');

      const irregularRule = masterData.irregularMap[`${dateKey}_${normName}`] || masterData.irregularMap[`${dateKey}_全拠点`];
      if (irregularRule) {
        const validRanges = [];
        for (const range of finalTimeStr.split('/')) {
          const parts = range.split('-');
          if (parts.length === 2) {
            const s = safeParseTime(parts[0]), e = safeParseTime(parts[1]);
            for (const r of irregularRule) {
              const adjS = Math.max(s, r.open), adjE = Math.min(e, r.close);
              if (adjS < adjE) validRanges.push(`${safeFormatTime(adjS)}-${safeFormatTime(adjE)}`);
            }
          }
        }
        if (validRanges.length === 0) continue;
        finalTimeStr = validRanges.join('/');
      }

      const gapMin = safeCalcTotalMin(finalTimeStr);
      const isInternal = normName.includes('内科');
      if (!isInternal) totalGapMinutes += gapMin;

      absenceDetails.push({ clinic: normName, dept: isInternal ? '内科' : '小児科', gapMin, timeStr: finalTimeStr });
    }

    const totalFilledMinutes = Math.max(0, totalRequiredMinutes - totalGapMinutes);
    const rate1st = totalRequiredMinutes > 0 ? Math.floor((totalFilledMinutes / totalRequiredMinutes) * 100) : 100;

    // 2診目およびCOO室依頼2診の算出
    const dailyExtData = masterData.extDataByDate[dateKey] || { paste: [], oubo: [], bosyu: [] };
    const advMetrics = calculateAdvancedMetricsFast(dailyExtData.paste, dailyExtData.oubo, dailyExtData.bosyu);
    const cooMetrics = calcDailyCooMetrics_(d, dateKey, masterData, dailyExtData, pasteCols, ouboCols, isUnopenedClinic, getIntervalsFast);

    metricsMap[dateKey] = {
      req1stMin: totalRequiredMinutes,
      filled1stMin: totalFilledMinutes,
      gap1stMin: totalGapMinutes,
      rate1st: rate1st,
      req2ndMin: advMetrics.secondReqMin || 0,
      filled2ndMin: advMetrics.secondActualMin || 0,
      rate2nd: advMetrics.secondRate || 100,
      cooReqMin: cooMetrics.reqMin,
      cooFilledMin: cooMetrics.filledMin,
      cooFMin: cooMetrics.fMin,
      absences: absenceDetails
    };
  }
  return metricsMap;
}

function calcDailyCooMetrics_(d, dateKey, masterData, dailyExtData, pasteCols, ouboCols, isUnopenedClinic, getIntervalsFast) {
  let reqMin = 0, filledMin = 0, fMin = 0;
  if (!masterData.cooDataByDate[dateKey]) return { reqMin, filledMin, fMin };

  const covMap = {};
  const getCov = (c) => { if (!covMap[c]) covMap[c] = new Array(1440).fill(0); return covMap[c]; };

  dailyExtData.paste.forEach(row => {
    const doc = String(row[pasteCols.doc] || '').trim();
    const clinic = String(row[pasteCols.clinic] || '').replace(/[（(]小児科[）)]/, '').trim();
    const cNo = pasteCols.clinicNo !== -1 ? String(row[pasteCols.clinicNo] || '').trim() : '';
    if (!doc || !clinic || /バックアップ|有給|欠勤/.test(doc) || isUnopenedClinic(clinic, d, cNo)) return;
    getIntervalsFast(row).forEach(inv => {
      const cov = getCov(clinic);
      for (let m = inv.start; m < inv.end; m++) cov[m]++;
    });
  });

  dailyExtData.oubo.forEach(row => {
    const doc = String(row[ouboCols.doc] || '').trim();
    const clinic = String(row[ouboCols.clinic] || '').replace(/[（(]小児科[）)]/, '').trim();
    const cNo = ouboCols.clinicNo !== -1 ? String(row[ouboCols.clinicNo] || '').trim() : '';
    if (!doc || !clinic || clinic.includes('バックアップ') || doc.replace(/\s+/g, '') === '橋本浩' || isUnopenedClinic(clinic, d, cNo)) return;
    const s = safeParseTime(row[ouboCols.start]), e = safeParseTime(row[ouboCols.end]);
    if (!isNaN(s) && !isNaN(e) && s < e) {
      const cov = getCov(clinic);
      for (let m = s; m < e; m++) cov[m]++;
    }
  });

  masterData.cooDataByDate[dateKey].forEach(row => {
    const clinic = String(row[0] || '').replace(/[（(]小児科[）)]/, '').trim();
    if (isUnopenedClinic(clinic, d)) return;
    const s = safeParseTime(row[2]), e = safeParseTime(row[3]);
    const flag = String(row[4] || '').toLowerCase();
    if (!isNaN(s) && !isNaN(e) && s < e) {
      reqMin += (e - s);
      if (flag.includes('f')) {
        filledMin += (e - s);
        fMin += (e - s);
      } else if (covMap[clinic]) {
        const cov = covMap[clinic];
        for (let m = s; m < e; m++) {
          if (cov[m] >= 2) filledMin++;
        }
      }
    }
  });

  return { reqMin, filledMin, fMin };
}