/**
 * 時間のパース、フォーマット、マージを行う安全なヘルパー関数群
 * 完全に独立しているため、他の処理に影響を与えません。
 */

function safeParseTime(t) {
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  if (typeof t === 'string') {
    const p = t.trim().split(':');
    if (p.length >= 2) return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  }
  if (typeof t === 'number' && t >= 0 && t < 1) return Math.round(t * 24 * 60);
  return NaN;
}

function safeFormatTime(min) {
  return `${String(Math.floor(min / 60)).padStart(2, '0')}:${String(min % 60).padStart(2, '0')}`;
}

function safeNormalizeAndMerge(timeStrings, isSpecialDay) {
  const OPEN = isSpecialDay ? 10 * 60 : 9 * 60;
  const CLOSE = isSpecialDay ? 19 * 60 : 21 * 60;
  let intervals = [];

  timeStrings.forEach(str => {
    // カンマ(,)やスラッシュ(/)で区切られた「複数の時間帯」を分解して処理する
    const subStrings = str.split(/[,/、]/);

    subStrings.forEach(subStr => {
      const parts = subStr.trim().split(/[-~]/);
      if (parts.length !== 2) return;

      let start = safeParseTime(parts[0]);
      let end = safeParseTime(parts[1]);
      if (isNaN(start) || isNaN(end)) return;
      if (end <= OPEN || start >= CLOSE) return;

      start = Math.max(start, OPEN);
      end = Math.min(end, CLOSE);

      if (start < end) intervals.push({ start, end });
    });
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

  // チャット出力用にスラッシュ区切りでフォーマット（例：15:00-15:30/17:30-18:00）
  return merged.map(m => `${safeFormatTime(m.start)}-${safeFormatTime(m.end)}`).join('/');
}

function safeCalcTotalMin(mergedStr) {
  let total = 0;
  mergedStr.split('/').forEach(p => {
    const rng = p.split('-');
    if (rng.length === 2) {
      const s = safeParseTime(rng[0]);
      const e = safeParseTime(rng[1]);
      if (!isNaN(s) && !isNaN(e)) total += (e - s);
    }
  });
  return total;
}

function safeRemoveClosed(mergedTimeStr, closedTime) {
  if (!closedTime || closedTime === "全日") return "";
  let cRanges = [];
  if (closedTime === "午前") cRanges.push({ s: 9 * 60, e: 13 * 60 });
  else if (closedTime === "午後") cRanges.push({ s: 15 * 60, e: 18 * 60 });
  else if (closedTime === "夜間") cRanges.push({ s: 18 * 60, e: 21 * 60 });
  else if (closedTime === "午後夜間") cRanges.push({ s: 15 * 60, e: 21 * 60 });
  if (cRanges.length === 0) return mergedTimeStr;

  const validRanges = [];
  mergedTimeStr.split('/').forEach(range => {
    const parts = range.split('-');
    if (parts.length !== 2) return;
    let start = safeParseTime(parts[0]);
    let end = safeParseTime(parts[1]);

    for (const c of cRanges) {
      if (start >= end) break;
      if (c.s <= start && c.e >= end) start = end;
      else if (c.s <= start && c.e > start) start = c.e;
      else if (c.s < end && c.e >= end) end = c.s;
      else if (c.s > start && c.e < end) {
        validRanges.push(`${safeFormatTime(start)}-${safeFormatTime(c.s)}`);
        start = c.e;
      }
    }
    if (start < end) validRanges.push(`${safeFormatTime(start)}-${safeFormatTime(end)}`);
  });

  return validRanges.join('/');
}

/**
 * 開始〜終了の5列（P〜T列等）と、時給1〜4・受入可能人数1〜4を照合し、
 * 本物の中抜け（休憩）と時給単価の切替境界を判別して実働区間配列 [{start, end}, ...] を返す関数
 */
function extractWorkIntervals(row, colOpts) {
  const tCols = (colOpts && colOpts.times) ? colOpts.times : [15, 16, 17, 18, 19];
  const wCols = (colOpts && colOpts.wages) ? colOpts.wages : [25, 26, 27, 28];
  const wageSumCol = (colOpts && colOpts.wageSum !== undefined) ? colOpts.wageSum : 29;
  const capCols = (colOpts && colOpts.caps) ? colOpts.caps : [54, 55, 56, 57];

  const parseMin = (v) => {
    if (v === '' || v === null || v === undefined || v === 0) return NaN;
    return safeParseTime(v);
  };

  // 5つの時間列から値が入っている境界点（order: 0〜4, min: 分）を抽出
  const points = [];
  for (let k = 0; k < tCols.length && k < 5; k++) {
    const colIdx = tCols[k];
    const min = colIdx !== -1 ? parseMin(row[colIdx]) : NaN;
    if (!isNaN(min)) points.push({ order: k, min: min });
  }

  if (points.length < 2) return [];
  if (points.length === 2) {
    return points[0].min < points[1].min ? [{ start: points[0].min, end: points[1].min }] : [];
  }

  const toNum = (v) => {
    if (v === '' || v === null || v === undefined) return NaN;
    const n = Number(String(v).replace(/,/g, '').trim());
    return isNaN(n) ? NaN : n;
  };

  const wages = wCols.map(c => (c !== -1 && c < row.length) ? toNum(row[c]) : NaN);
  const caps = capCols.map(c => (c !== -1 && c < row.length) ? toNum(row[c]) : NaN);
  const wageSum = (wageSumCol !== -1 && wageSumCol < row.length) ? (toNum(row[wageSumCol]) || 0) : 0;

  const hasAnyPositiveWage = wages.some(w => !isNaN(w) && w > 0) || wageSum > 0;
  const hasAnyPositiveCap = caps.some(c => !isNaN(c) && c > 0);

  const segCount = points.length - 1;
  const rawIntervals = [];

  for (let p = 0; p < segCount; p++) {
    const segStart = points[p].min;
    const segEnd = points[p + 1].min;
    if (segStart >= segEnd) continue;

    // 終了列の位置(1:終了1, 2:終了2, 3:終了3, 4:勤務終了)がそのままスロット(0〜3)に対応
    const slotIdx = Math.max(0, Math.min(3, points[p + 1].order - 1));
    const segWage = wages[slotIdx];
    const segCap = caps[slotIdx];

    // 先頭区間(p===0)と最終区間(p===segCount-1)は必ず実働。中間区間のみ「休憩」か「単価切替」かを判定
    const isInterior = (p > 0 && p < segCount - 1);
    let isWork = true;

    if (isInterior) {
      if (!isNaN(segWage) && segWage > 0) {
        isWork = true; // 単価が入っている＝実働（単価切替）
      } else if (!isNaN(segCap) && segCap > 0) {
        isWork = true; // 受入可能人数が入っている＝実働
      } else if (hasAnyPositiveWage && (isNaN(segWage) || segWage === 0)) {
        isWork = false; // 非常勤で該当枠だけ時給0/空欄＝休憩
      } else if (hasAnyPositiveCap && (isNaN(segCap) || segCap === 0)) {
        isWork = false; // 他枠に受入人数があり該当枠だけ空欄/0＝休憩
      } else {
        // 時給・受入とも全枠0/空欄の行（外部シートの常勤医師等）のフォールバック
        if (points.length === 4 && p === 1) {
          isWork = false;
        } else if (points.length === 5 && segStart >= 11 * 60 && segEnd <= 16 * 60) {
          isWork = false;
        }
      }
    }

    if (isWork) {
      rawIntervals.push({ start: segStart, end: segEnd });
    }
  }

  // 連続する実働区間をマージ（例: 15:00-17:00 + 17:00-18:00 + 18:00-20:00 ➔ 15:00-20:00）
  const merged = [];
  for (let i = 0; i < rawIntervals.length; i++) {
    const cur = rawIntervals[i];
    if (merged.length > 0 && merged[merged.length - 1].end >= cur.start) {
      merged[merged.length - 1].end = Math.max(merged[merged.length - 1].end, cur.end);
    } else {
      merged.push({ start: cur.start, end: cur.end });
    }
  }
  return merged;
}