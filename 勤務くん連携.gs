/**
 * ============================================================================
 * 勤務くん連携.gs（mnys.jp 側：タイムアウト防止・高速抽出エディション）
 * ============================================================================
 */

const KINMUKUN_BRIDGE_CFG = {
  AGENCY_SS_ID: '1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE',
  MONTHLY_1_SS_ID: '1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc',
  MONTHLY_2_SS_ID: '10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw',
  SHIFT_NEW_SS_ID: '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co',
  MIN_TOTAL_COUNT: 2, // 過去6ヶ月〜1年で2回以上勤務している医師を抽出
  TOP_N: 5,
  LOOKBACK_MONTHS: 6, // タイムアウト防止のため直近6ヶ月分の月次シート＋確定シフトを高速スキャン
  TARGET_DAYS_AHEAD: 35 // 会議対象となる直近35日以内の不在枠に対して候補を抽出
};

/**
 * ★手動テスト用：確実に「勤務くん候補付き」で暗号化メールを送る専用関数
 * （Apps Script上部のプルダウンからこの関数を選んで「実行」してください）
 */
function runManualSendWithKinmukun() {
  sendEncryptedShiftContextToCaps(true);
}

/**
 * AI連携_暗号化送信.gs から呼ばれるメイン関数
 */
function buildKinmukunCandidatesFromUnfilledList_(unfilledList) {
  const t0 = Date.now();
  const today = new Date();
  const todayYmd = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  const limitYmd = Utilities.formatDate(
    new Date(today.getTime() + KINMUKUN_BRIDGE_CFG.TARGET_DAYS_AHEAD * 86400000),
    'Asia/Tokyo',
    'yyyy/MM/dd'
  );

  // タイムアウト防止のため、過去LOOKBACK_MONTHSヶ月前〜未来を対象にする
  const startDate = new Date(today.getFullYear(), today.getMonth() - KINMUKUN_BRIDGE_CFG.LOOKBACK_MONTHS, 1);
  const endDate = new Date(today.getFullYear(), today.getMonth() + 3, 1);

  const targetClinics = new Set();
  const targetDates = new Set();
  const slots = [];

  unfilledList.forEach(u => {
    if (!u.date || u.date < todayYmd || u.date > limitYmd || !u.clinic) return;
    const { kyotenName, displayKyoten } = parseKinmukunClinic_(u.clinic);
    const frameKey = detectKinmukunFrameKey_(u.absenceTime);
    const dow = new Date(u.date).getDay();

    targetClinics.add(displayKyoten);
    targetDates.add(u.date);

    slots.push({
      dateStr: u.date,
      dow: dow,
      dayLabel: ['日', '月', '火', '水', '木', '金', '土'][dow],
      kyotenName: kyotenName,
      displayKyoten: displayKyoten,
      absTime: u.absenceTime,
      frameKey: frameKey
    });
  });

  Logger.log(`🔍 勤務くん候補抽出開始: 対象不在枠 ${slots.length}件（直近${KINMUKUN_BRIDGE_CFG.TARGET_DAYS_AHEAD}日分 / 対象拠点数: ${targetClinics.size}）`);

  const agencyIds = loadKinmukunAgencyIds_();
  const { statsByClinicDoc, bookedMap } = scanKinmukunHistoryOnce_(startDate, endDate, todayYmd, targetClinics, targetDates, agencyIds);

  const spotCandidateRows = [[
    '日付', '曜日', '拠点名', '不在時間帯', '候補順位',
    '医師名', '医籍番号', '属性', '同曜日×同時間帯実績', '同曜日実績', '拠点総回数', '最終勤務日', '当日他拠点シフト'
  ]];

  const bundleMap = new Map();

  slots.forEach(s => {
    const docMap = statsByClinicDoc.get(s.displayKyoten) || new Map();

    const cands = Array.from(docMap.values())
      .filter(d => d.totalCount >= KINMUKUN_BRIDGE_CFG.MIN_TOTAL_COUNT)
      .map(d => {
        const sameDowFrame = d.dayFrames[s.dow][s.frameKey] || 0;
        const sameDowTotal = d.dayFrames[s.dow].am + d.dayFrames[s.dow].pm1 + d.dayFrames[s.dow].pm2;
        const sameFrameTotal = d.totalFrames[s.frameKey] || 0;

        const bList = bookedMap.get(`${s.dateStr}_${d.id}`) || [];
        const hasOverlap = bList.some(b => b.frames[s.frameKey] > 0);
        const bookedNote = bList.length === 0 ? '当日空き' : `同日別枠あり(${bList.map(b => b.clinic).join(',')})`;

        return { ...d, sameDowFrame, sameDowTotal, sameFrameTotal, hasOverlap, bookedNote };
      })
      .filter(d => !d.hasOverlap)
      .sort((a, b) =>
        (b.sameDowFrame - a.sameDowFrame) ||
        (b.sameDowTotal - a.sameDowTotal) ||
        (b.sameFrameTotal - a.sameFrameTotal) ||
        (b.totalCount - a.totalCount)
      )
      .slice(0, KINMUKUN_BRIDGE_CFG.TOP_N);

    cands.forEach((c, idx) => {
      spotCandidateRows.push([
        s.dateStr, s.dayLabel, s.displayKyoten, s.absTime, `候補${idx + 1}`,
        c.name, c.id, c.attribute, `${c.sameDowFrame}回`, `${c.sameDowTotal}回`, `${c.totalCount}回`, c.lastWork, c.bookedNote
      ]);

      if (!bundleMap.has(c.id)) {
        bundleMap.set(c.id, { name: c.name, id: c.id, attribute: c.attribute, items: [] });
      }
      bundleMap.get(c.id).items.push(`${s.dateStr}(${s.dayLabel})【${s.displayKyoten}】${s.absTime}[候補${idx + 1}/同枠${c.sameDowFrame}回]`);
    });
  });

  const bundledCandidateRows = [['医師名', '医籍番号', '属性', '対象枠数', 'まとめて打診する日程・拠点一覧']];
  Array.from(bundleMap.values())
    .filter(b => b.items.length >= 2)
    .sort((a, b) => b.items.length - a.items.length)
    .forEach(b => {
      bundledCandidateRows.push([b.name, b.id, b.attribute, `${b.items.length}枠`, b.items.join(' ／ ')]);
    });

  Logger.log(`✅ 勤務くん抽出完了（所要時間: ${((Date.now() - t0) / 1000).toFixed(1)}秒 / 候補行数: ${spotCandidateRows.length - 1}件 / 複数日程まとめ: ${bundledCandidateRows.length - 1}名）`);
  return { spotCandidateRows, bundledCandidateRows };
}

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
          if (!id || !name) continue;

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

function parseKinmukunClinic_(rawClinic) {
  const c = String(rawClinic || '').trim();
  if (c.includes('北葛西')) {
    return { kyotenName: '北葛西', displayKyoten: c.includes('内科') ? '北葛西（内科）' : '北葛西（小児科）' };
  }
  if (c.includes('亀有')) {
    return { kyotenName: '亀有', displayKyoten: c.includes('内科') ? '亀有（内科）' : '亀有（小児科）' };
  }
  const clean = c.replace(/小児科|内科|[\(（].*?[\)）]/g, '').trim();
  return { kyotenName: clean, displayKyoten: clean };
}

function detectKinmukunFrameKey_(absTime) {
  const s = String(absTime || '');
  if (s.startsWith('09') || s.startsWith('9')) return 'am';
  if (s.startsWith('18') || s.startsWith('17')) return 'pm2';
  return 'pm1';
}

function calcKinmukunFrames_(sVal, eVal, isKitakasai) {
  const f = { am: 0, pm1: 0, pm2: 0 };
  const toH = v => {
    if (!v) return null;
    if (v instanceof Date) return v.getHours() + v.getMinutes() / 60;
    const m = String(v).match(/(\d{1,2}):(\d{2})/);
    return m ? Number(m[1]) + Number(m[2]) / 60 : null;
  };
  const s = toH(sVal), e = toH(eVal);
  if (s === null || e === null) return f;
  if (s < 13 && e > 9) f.am = 1;
  if (s < 18 && e > 15) f.pm1 = 1;
  if (s < (isKitakasai ? 20 : 21) && e > 18) f.pm2 = 1;
  return f;
}

function loadKinmukunAgencyIds_() {
  const set = new Set();
  try {
    const ss = SpreadsheetApp.openById(KINMUKUN_BRIDGE_CFG.AGENCY_SS_ID);
    const sh = ss.getSheetByName('紹介会社応募表') || ss.getSheets()[0];
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return set;
    let idx = vals[0].indexOf('医籍番号');
    if (idx === -1) idx = 1;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][idx]) set.add(String(vals[i][idx]).trim());
    }
  } catch (e) {}
  return set;
}