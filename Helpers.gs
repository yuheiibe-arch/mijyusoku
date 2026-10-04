/**
 * 爆速の日付フォーマット関数
 */
function fastFormatDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return "";
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

/**
 * 日付パース関数（曜日カッコを強制的に切り捨てる確実版）
 */
function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) {
    if (isNaN(dateInput.getTime())) return null;
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }

  let dateStr = String(dateInput).trim();

  // 1. "2026年9月1日" などの表記に対応
  const jpMatch = dateStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (jpMatch) {
    return new Date(parseInt(jpMatch[1], 10), parseInt(jpMatch[2], 10) - 1, parseInt(jpMatch[3], 10));
  }

  // 2. カッコ（全角・半角）で分割し、左側の日付部分だけを取り出す
  dateStr = dateStr.split(/[（(]/)[0].trim();
  
  // 3. ハイフン表記をスラッシュに統一
  dateStr = dateStr.replace(/-/g, '/');
  
  // 4. "YYYY/MM/DD" に分解してDateオブジェクト化
  const parts = dateStr.split('/');
  if (parts.length >= 3) {
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      return new Date(year, month, day);
    }
  }

  // 5. 念のため標準の Date パーサーも試行
  const fallbackDate = new Date(dateStr);
  if (!isNaN(fallbackDate.getTime())) {
    return new Date(fallbackDate.getFullYear(), fallbackDate.getMonth(), fallbackDate.getDate());
  }

  return null;
}

/**
 * 時間文字列("09:00") -> 分
 */
function parseTimeToMinutes(timeStr) {
  if (!timeStr || typeof timeStr !== 'string') return NaN;
  const parts = timeStr.trim().split(':');
  if (parts.length < 2) return NaN;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

/**
 * 分 -> 時間文字列
 */
function formatMinutesToHHMM(totalMinutes) {
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

/**
 * 時間結合・補正ロジック
 */
function normalizeAndMergeTimes(timeStrings, isSpecialDay) {
  const OPEN = isSpecialDay ? 10 * 60 : 9 * 60;   
  const CLOSE = isSpecialDay ? 19 * 60 : 21 * 60; 

  let intervals = [];
  timeStrings.forEach(str => {
    const parts = str.split(/[-~]/);
    if (parts.length !== 2) return;
    let start = parseTimeToMinutes(parts[0]);
    let end = parseTimeToMinutes(parts[1]);

    if (isNaN(start) || isNaN(end)) return;
    if (end <= OPEN || start >= CLOSE) return; 

    start = Math.max(start, OPEN);
    end = Math.min(end, CLOSE);

    if (start < end) intervals.push({ start, end });
  });

  if (intervals.length === 0) return "";

  intervals.sort((a, b) => a.start - b.start);
  let merged = [];
  let current = intervals[0];
  for (let i = 1; i < intervals.length; i++) {
    if (current.end >= intervals[i].start) {
      current.end = Math.max(current.end, intervals[i].end);
    } else {
      merged.push(current);
      current = intervals[i];
    }
  }
  merged.push(current);
  return merged.map(m => `${formatMinutesToHHMM(m.start)}-${formatMinutesToHHMM(m.end)}`).join('/');
}

/**
 * 結合文字列から合計分数を算出
 */
function calculateTotalMinutesFromStr(mergedStr) {
    let total = 0;
    const parts = mergedStr.split('/');
    parts.forEach(p => {
        const rng = p.split('-');
        if (rng.length === 2) {
            const s = parseTimeToMinutes(rng[0]);
            const e = parseTimeToMinutes(rng[1]);
            if (!isNaN(s) && !isNaN(e)) total += (e - s);
        }
    });
    return total;
}

/**
 * 拠点名正規化
 */
function normalizeClinicName(name) {
    return name.replace(/[（(]小児科[）)]/, "").trim();
}

/**
 * 休館時間（プルダウン）をカットするヘルパー
 */
function removeClosedTime(mergedTimeStr, closedTime) {
    if (!closedTime || closedTime === "全日") return "";
    let closedRanges = [];
    if (closedTime === "午前") closedRanges.push({s: 9*60, e: 13*60});
    else if (closedTime === "午後") closedRanges.push({s: 15*60, e: 18*60});
    else if (closedTime === "夜間") closedRanges.push({s: 18*60, e: 21*60});
    else if (closedTime === "午後夜間") closedRanges.push({s: 15*60, e: 21*60});

    if (closedRanges.length === 0) return mergedTimeStr;

    const ranges = mergedTimeStr.split('/');
    const validRanges = [];

    for (const range of ranges) {
        const parts = range.split('-');
        if (parts.length !== 2) continue;
        let start = parseTimeToMinutes(parts[0]);
        let end = parseTimeToMinutes(parts[1]);

        for (const c of closedRanges) {
            if (start >= end) break;
            if (c.s <= start && c.e >= end) {
                start = end; 
            } else if (c.s <= start && c.e > start) {
                start = c.e; 
            } else if (c.s < end && c.e >= end) {
                end = c.s; 
            } else if (c.s > start && c.e < end) {
                validRanges.push(`${formatMinutesToHHMM(start)}-${formatMinutesToHHMM(c.s)}`);
                start = c.e; 
            }
        }
        if (start < end) {
            validRanges.push(`${formatMinutesToHHMM(start)}-${formatMinutesToHHMM(end)}`);
        }
    }
    return validRanges.join('/');
}