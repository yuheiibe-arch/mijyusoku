/**
 * 集計およびテキスト組み立て処理
 */
function buildDailyAndSummaryReport(startDate, endDate, masterData, shiftData, weekdaysJP) {
  let hasAnyContent = false;
  let dailyText = "";

  let weeklyReq1st = 0, weeklyFilled1st = 0, weeklyGapMin = 0;
  let weeklyReq2nd = 0, weeklyFilled2nd = 0;
  let weeklyCooReqMin = 0, weeklyCooFilledMin = 0, weeklyCooFMin = 0;
  let weeklyAbsenceClinics = []; 

  let monthlyReq1st = 0, monthlyFilled1st = 0, monthlyGapMin = 0;
  let monthlyReq2nd = 0, monthlyFilled2nd = 0;
  let monthlyCooReqMin = 0, monthlyCooFilledMin = 0, monthlyCooFMin = 0;
  let monthlyAbsenceCount = 0;

  const loopStart = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const loopEnd = new Date(Math.max(endDate.getTime(), new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0).getTime()));

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

  for (let d = new Date(loopStart); d <= loopEnd; d.setDate(d.getDate() + 1)) {
    const isWeekly = (d.getTime() >= startDate.getTime() && d.getTime() <= endDate.getTime());
    const isMonthly = (d.getFullYear() === startDate.getFullYear() && d.getMonth() === startDate.getMonth());

    if (!isWeekly && !isMonthly) continue;

    const dateKey = fastFormatDate(d);
    const dateTitle = `${d.getMonth() + 1}月${d.getDate()}日（${weekdaysJP[d.getDay()]}）`;
    const isSpecialDay = (d.getMonth() + 1 === 12 && d.getDate() === 31) || (d.getMonth() + 1 === 1 && d.getDate() <= 3);

    let totalRequiredMinutes = 0;
    
    shiftData.sortedClinicList.forEach(clinic => {
        if (masterData.openDateMap.has(clinic)) {
            const openDate = masterData.openDateMap.get(clinic);
            openDate.setHours(0, 0, 0, 0);
            const checkDate = new Date(d);
            checkDate.setHours(0, 0, 0, 0);
            if (checkDate < openDate) return; 
        }

        if (clinic.includes("内科")) return; 
        
        let closedTime = masterData.closedDataMap.get(`${dateKey}_${clinic}`) || masterData.closedDataMap.get(`${dateKey}_全拠点`);
        if (closedTime === "全日") return; 

        let reqMin = 0;
        if (isSpecialDay) {
            reqMin = 420; 
        } else {
            if (clinic.includes("北葛西")) reqMin = 540; 
            else reqMin = 600; 
        }

        if (closedTime === "午前") reqMin -= 4 * 60;
        else if (closedTime === "午後") reqMin -= 3 * 60;
        else if (closedTime === "夜間") reqMin -= 3 * 60;
        else if (closedTime === "午後夜間") reqMin -= 6 * 60;

        const irregularRule = masterData.irregularMap[`${dateKey}_${clinic}`] || masterData.irregularMap[`${dateKey}_全拠点`];
        if (irregularRule) {
            reqMin = 0;
            irregularRule.forEach(r => { reqMin += (r.close - r.open); });
        }

        totalRequiredMinutes += Math.max(0, reqMin);
    });

    let totalGapMinutes = 0;
    const dailyOutputLines = [];
    
    const dailyRecords = shiftData.fuzaiRecords.filter(r => r.dateKey === dateKey);
    const groupedFuzai = {};
    dailyRecords.forEach(r => {
        if (!groupedFuzai[r.normName]) groupedFuzai[r.normName] = [];
        groupedFuzai[r.normName].push(r.timeRaw);
    });

    shiftData.sortedClinicList.forEach(clinic => {
        if (masterData.openDateMap.has(clinic)) {
            const openDate = masterData.openDateMap.get(clinic);
            openDate.setHours(0, 0, 0, 0);
            const checkDate = new Date(d);
            checkDate.setHours(0, 0, 0, 0);
            if (checkDate < openDate) return;
        }

        if (clinic.includes("内科")) return; 

        let closedTime = masterData.closedDataMap.get(`${dateKey}_${clinic}`) || masterData.closedDataMap.get(`${dateKey}_全拠点`);
        if (closedTime === "全日") return;

        const hasFuzai = groupedFuzai[clinic] && groupedFuzai[clinic].length > 0;
        const isWorking = shiftData.workingClinics[dateKey] && shiftData.workingClinics[dateKey].has(clinic);

        if (!hasFuzai && !isWorking) {
            let pureAbsenceArr = isSpecialDay ? ["10:00-19:00"] : (clinic.includes("北葛西") ? ["09:00-13:00", "15:00-20:00"] : ["09:00-13:00", "15:00-21:00"]);
            groupedFuzai[clinic] = pureAbsenceArr;
        }
    });

    for (const normName in groupedFuzai) {
        if (masterData.openDateMap.has(normName)) {
            const openDate = masterData.openDateMap.get(normName);
            openDate.setHours(0, 0, 0, 0);
            const checkDate = new Date(d);
            checkDate.setHours(0, 0, 0, 0);
            if (checkDate < openDate) continue; 
        }

        let closedTime = null;
        if (normName.includes("内科")) {
             closedTime = masterData.closedDataMap.get(`${dateKey}_${normName}`) || masterData.closedDataMap.get(`${dateKey}_${normName}_内科`) || masterData.closedDataMap.get(`${dateKey}_全拠点`);
        } else {
             closedTime = masterData.closedDataMap.get(`${dateKey}_${normName}`) || masterData.closedDataMap.get(`${dateKey}_全拠点`);
        }
        if (closedTime === "全日") continue; 

        let mergedTimeStr = safeNormalizeAndMerge(groupedFuzai[normName], isSpecialDay);
        
        if (mergedTimeStr) {
            let finalTimeStr = mergedTimeStr;

            if (closedTime) finalTimeStr = safeRemoveClosed(finalTimeStr, closedTime);
            if (!finalTimeStr) continue;

            if (!isSpecialDay && normName.includes("北葛西")) {
                 finalTimeStr = finalTimeStr.replace("21:00", "20:00");
            }
            
            const irregularRule = masterData.irregularMap[`${dateKey}_${normName}`] || masterData.irregularMap[`${dateKey}_全拠点`];
            if (irregularRule) {
                 const ranges = finalTimeStr.split('/');
                 const validRanges = [];
                 for (const range of ranges) {
                     const parts = range.split('-');
                     if (parts.length === 2) {
                         const s = safeParseTime(parts[0]);
                         const e = safeParseTime(parts[1]);
                         for (const r of irregularRule) {
                             const adjustedS = Math.max(s, r.open);
                             const adjustedE = Math.min(e, r.close);
                             if (adjustedS < adjustedE) {
                                 validRanges.push(`${safeFormatTime(adjustedS)}-${safeFormatTime(adjustedE)}`);
                             }
                         }
                     }
                 }
                 if (validRanges.length === 0) continue; 
                 finalTimeStr = validRanges.join('/');
            }

            if (!normName.includes("内科")) {
                totalGapMinutes += safeCalcTotalMin(finalTimeStr);
            }

            let displayTimeStr = finalTimeStr;
            if (displayTimeStr.startsWith("09:00-13:00/15:00-")) {
                displayTimeStr = displayTimeStr.replace("09:00-13:00/15:00-", "09:00-");
            } else if (displayTimeStr.startsWith("10:00-13:00/15:00-")) { 
                displayTimeStr = displayTimeStr.replace("10:00-13:00/15:00-", "10:00-");
            }

            let record = dailyRecords.find(r => r.normName === normName);
            let displayName = record ? record.name : (shiftData.rawNameMap[normName] || normName);
            
            const currentDepartment = record && record.isInternalMedicine ? "内科" : (displayName.includes("内科") ? "内科" : "小児科");
            
            if (isWeekly) {
                dailyOutputLines.push(`【${displayName}】${displayTimeStr}`);
                weeklyAbsenceClinics.push({ clinic: normName, dept: currentDepartment });
            }
            if (isMonthly) {
                monthlyAbsenceCount++;
            }
        }
    }

    const totalFilledMinutes = Math.max(0, totalRequiredMinutes - totalGapMinutes);
    const rate = totalRequiredMinutes > 0 ? Math.floor((totalFilledMinutes / totalRequiredMinutes) * 100) : 100;
    const requiredHours = Math.round(totalRequiredMinutes / 60);
    const filledHours = Math.round(totalFilledMinutes / 60);

    const dailyExtData = masterData.extDataByDate[dateKey] || { paste: [], oubo: [], bosyu: [] };
    const advMetrics = calculateAdvancedMetricsFast(dailyExtData.paste, dailyExtData.oubo, dailyExtData.bosyu);

    let dailyCooReqMin = 0;
    let dailyCooFilledMin = 0;
    let dailyCooFMin = 0;

    if (masterData.cooDataByDate[dateKey]) {
      const covMap = {};
      const getCov = (c) => { if (!covMap[c]) covMap[c] = new Array(1440).fill(0); return covMap[c]; };

      dailyExtData.paste.forEach(row => {
        const doc = String(row[0] || "").trim();
        const clinic = String(row[12] || "").replace(/[（(]小児科[）)]/, "").trim();
        if (!doc || !clinic || doc.includes("バックアップ") || doc.includes("有給") || doc.includes("欠勤")) return;
        
        const intervals = getIntervalsFast(row);
        intervals.forEach(inv => {
            const cov = getCov(clinic);
            for (let m = inv.start; m < inv.end; m++) cov[m]++;
        });
      });

      dailyExtData.oubo.forEach(row => {
        const doc = String(row[0] || "").trim();
        const clinic = String(row[3] || "").replace(/[（(]小児科[）)]/, "").trim();
        if (!doc || !clinic || clinic.includes("バックアップ")) return;
        if (doc.replace(/\s+/g, '') === "橋本浩") return;
        const startMin = safeParseTime(row[6]);
        const endMin = safeParseTime(row[7]);
        if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
          const cov = getCov(clinic);
          for (let m = startMin; m < endMin; m++) cov[m]++;
        }
      });

      masterData.cooDataByDate[dateKey].forEach(row => {
        const clinic = String(row[0] || "").replace(/[（(]小児科[）)]/, "").trim();
        const startMin = safeParseTime(row[2]);
        const endMin = safeParseTime(row[3]);
        const reqStr = String(row[4] || "").toLowerCase(); 

        if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
          dailyCooReqMin += (endMin - startMin); 
          if (reqStr.includes('f')) {
            dailyCooFilledMin += (endMin - startMin);
            dailyCooFMin += (endMin - startMin); 
          } else {
            if (covMap[clinic]) {
              const cov = covMap[clinic];
              for (let m = startMin; m < endMin; m++) {
                if (cov[m] >= 2) dailyCooFilledMin++;
              }
            }
          }
        }
      });
    }

    if (isWeekly) {
        weeklyReq1st += totalRequiredMinutes;
        weeklyFilled1st += totalFilledMinutes;
        weeklyGapMin += totalGapMinutes;
        weeklyReq2nd += advMetrics.secondReqMin;
        weeklyFilled2nd += advMetrics.secondActualMin;
        weeklyCooReqMin += dailyCooReqMin;
        weeklyCooFilledMin += dailyCooFilledMin;
        weeklyCooFMin += dailyCooFMin;
        
        if (totalRequiredMinutes > 0) {
            let entry = `[info][title]${dateTitle}[/title]`;
            if (shiftData.backupInfoMap[dateKey]) entry += shiftData.backupInfoMap[dateKey];
            entry += `[hr]\n`;
            entry += `小児科１診目充足率：${rate}%（充足済み：${filledHours}h/募集：${requiredHours}h）\n`;
            
            if (masterData.isExtSsLoaded) {
                entry += `募集全体充足率：${advMetrics.overallRate}%（充足済み：${advMetrics.overallActualH}h/募集：${advMetrics.overallReqH}h）\n`;
                entry += `２診目充足率（全体）：${advMetrics.secondRate}%（充足済み：${advMetrics.secondActualH}h/募集：${advMetrics.secondReqH}h）\n`;
                if (dailyCooReqMin > 0) {
                    const dCooRate = Math.floor((dailyCooFilledMin / dailyCooReqMin) * 100);
                    const dCooReqH = Math.round(dailyCooReqMin / 60);
                    const dCooFilledH = Math.round(dailyCooFilledMin / 60);
                    const dCooFH = Math.round(dailyCooFMin / 60);
                    const dCooOuboH = dCooFilledH - dCooFH; 
                    
                    entry += `└COO室依頼２診：${dCooRate}%（充足済み：${dCooFilledH}h/募集：${dCooReqH}h）\n`;
                    entry += `（f判定含む依頼総数）\n`;
                    entry += `f判定：${dCooFH}h\n`;
                    entry += `応募時間数：${dCooOuboH}h\n`;
                }
            }
            entry += `\n`;
            if (dailyOutputLines.length > 0) {
                dailyOutputLines.sort();
                entry += `＜医師不在拠点＞\n` + dailyOutputLines.join('\n') + '\n';
            } else {
                entry += `充足\n`;
            }
            entry += `[/info]\n`;
            dailyText += entry;
            hasAnyContent = true;
        }
    }

    if (isMonthly) {
        monthlyReq1st += totalRequiredMinutes;
        monthlyFilled1st += totalFilledMinutes;
        monthlyGapMin += totalGapMinutes;
        monthlyReq2nd += advMetrics.secondReqMin;
        monthlyFilled2nd += advMetrics.secondActualMin;
        monthlyCooReqMin += dailyCooReqMin;
        monthlyCooFilledMin += dailyCooFilledMin;
        monthlyCooFMin += dailyCooFMin; 
    }
  }

  const summaryParams = {
    startDate: startDate,
    endDate: endDate,
    isExtSsLoaded: masterData.isExtSsLoaded,
    hasCooData: (weeklyCooReqMin > 0 || Object.keys(masterData.cooDataByDate).length > 0),
    weekly: {
      req1st: weeklyReq1st, filled1st: weeklyFilled1st, gapMin: weeklyGapMin,
      req2nd: weeklyReq2nd, filled2nd: weeklyFilled2nd,
      cooReq: weeklyCooReqMin, cooFilled: weeklyCooFilledMin, cooFMin: weeklyCooFMin, 
      absenceClinics: weeklyAbsenceClinics
    },
    monthly: {
      req1st: monthlyReq1st, filled1st: monthlyFilled1st, gapMin: monthlyGapMin,
      req2nd: monthlyReq2nd, filled2nd: monthlyFilled2nd,
      cooReq: monthlyCooReqMin, cooFilled: monthlyCooFilledMin, cooFMin: monthlyCooFMin, 
      absenceCount: monthlyAbsenceCount
    },
    areaMap: masterData.areaMap,
    TARGET_SPLIT_CLINICS: ["北葛西", "西葛西"]
  };

  const summaryText = buildWeeklyMonthlySummaryText(summaryParams);

  if (!hasAnyContent) {
    dailyText = "対象期間内に報告すべきデータ（小児科）はありませんでした。\n";
  }

  return { summaryText, dailyText };
}