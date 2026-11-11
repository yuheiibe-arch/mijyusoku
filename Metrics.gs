/**
 * 区間のマージヘルパー
 */
function mergeIntervalsForMetrics(intervals) {
  if (!intervals || intervals.length === 0) return [];
  intervals.sort((a, b) => a.start - b.start);
  const merged = [];
  let current = { ...intervals[0] };
  for (let i = 1; i < intervals.length; i++) {
    if (intervals[i].start <= current.end) {
      current.end = Math.max(current.end, intervals[i].end);
    } else {
      merged.push(current);
      current = { ...intervals[i] };
    }
  }
  merged.push(current);
  return merged;
}

/**
 * 全体・2診目充足率の算出ロジック（爆速版）
 */
function calculateAdvancedMetricsFast(pasteList, ouboList, bosyuList) {
  const parseTimeToMinutes = (t) => {
    if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
    if (typeof t === 'string') {
      const p = t.split(':');
      if (p.length >= 2) return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
    }
    if (typeof t === 'number' && t >= 0 && t < 1) return Math.round(t * 24 * 60);
    return 0;
  };

  const actualShiftsMap = new Map();

  // ① 貼付用（TimeHelpers.gs の extractWorkIntervals で中抜け・単価切替を正確に判別）
  pasteList.forEach(row => {
    const doctor = row[0] ? String(row[0]).trim() : "";
    const clinicRaw = row[12] ? String(row[12]).trim() : "";

    if (!doctor || doctor.includes("バックアップ") || doctor.includes("有給") || doctor.includes("欠勤") || !clinicRaw) return;
    const clinic = clinicRaw.replace(/[（(]小児科[）)]/, "").trim();

    const intervals = extractWorkIntervals(row);
    intervals.forEach(inv => {
      actualShiftsMap.set(`${clinic}_${inv.start}_${inv.end}_${doctor}`, { clinic, start: inv.start, end: inv.end });
    });
  });

  // ② 応募シフト
  ouboList.forEach(row => {
    const doctor = row[0] ? String(row[0]).trim() : "";
    const clinicRaw = row[3] ? String(row[3]).trim() : "";
    const startTime = row[6];
    const endTime = row[7];

    if (!doctor || clinicRaw.includes("バックアップ")) return;
    if (doctor.replace(/\s+/g, '') === "橋本浩") return;

    const startMin = parseTimeToMinutes(startTime);
    const endMin = parseTimeToMinutes(endTime);
    if (startMin < endMin) {
      const clinic = clinicRaw.replace(/[（(]小児科[）)]/, "").trim();
      const key = `${clinic}_${startMin}_${endMin}_${doctor}`;
      if (!actualShiftsMap.has(key)) {
        actualShiftsMap.set(key, { clinic, start: startMin, end: endMin });
      }
    }
  });

  let totalActualMin = 0;
  const rawIntervalsByClinic = {};
  actualShiftsMap.forEach(shift => {
    totalActualMin += (shift.end - shift.start);
    if (!rawIntervalsByClinic[shift.clinic]) rawIntervalsByClinic[shift.clinic] = [];
    rawIntervalsByClinic[shift.clinic].push({ start: shift.start, end: shift.end });
  });

  let firstActualMin = 0;
  const coverageByClinic = {};
  for (const clinic in rawIntervalsByClinic) {
    const merged = mergeIntervalsForMetrics(rawIntervalsByClinic[clinic]);
    coverageByClinic[clinic] = merged;
    merged.forEach(m => { firstActualMin += (m.end - m.start); });
  }

  const secondActualMin = totalActualMin - firstActualMin;
  let firstUnfilledMin = 0;
  let secondUnfilledMin = 0;

  // ③ 募集シフト
  bosyuList.forEach(row => {
    const clinicRaw = row[1] ? String(row[1]).trim() : "";
    const startTime = row[4];
    const endTime = row[5];
    const status = row[11] ? String(row[11]).trim() : "";
    const secondFlag = row[12] ? String(row[12]).trim() : "";

    if (clinicRaw.includes("バックアップ") || status !== '掲載') return;

    const startMin = parseTimeToMinutes(startTime);
    const endMin = parseTimeToMinutes(endTime);
    if (startMin >= endMin) return;

    const clinic = clinicRaw.replace(/[（(]小児科[）)]/, "").trim();

    if (secondFlag.includes("二診目")) {
      secondUnfilledMin += (endMin - startMin);
    } else {
      const coverage = coverageByClinic[clinic] || [];
      const hasOverlap = coverage.some(c => Math.max(startMin, c.start) < Math.min(endMin, c.end));
      if (!hasOverlap) {
        firstUnfilledMin += (endMin - startMin);
      }
    }
  });

  const totalRequiredMin = totalActualMin + firstUnfilledMin + secondUnfilledMin;
  const secondRequiredMin = secondActualMin + secondUnfilledMin;

  return {
    overallRate: totalRequiredMin > 0 ? Math.floor((totalActualMin / totalRequiredMin) * 100) : 100,
    overallActualH: Math.round(totalActualMin / 60),
    overallReqH: Math.round(totalRequiredMin / 60),
    secondRate: secondRequiredMin > 0 ? Math.floor((secondActualMin / secondRequiredMin) * 100) : 100,
    secondActualH: Math.round(secondActualMin / 60),
    secondReqH: Math.round(secondRequiredMin / 60),
    secondActualMin: secondActualMin,
    secondReqMin: secondRequiredMin
  };
}