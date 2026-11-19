/**
 * ============================================================================
 * 【mijyusoku 新規追加】AI連携_5_前日差分_要因検知.gs
 * - 『貼付用』シートの日次ベースラインを非表示シート『_シフト日次基準』に保持
 * - 「昨日いたが今日いない＝シフトキャンセル」「勤務➔欠勤＝定期/常勤キャンセル」「移動玉突き」を検知
 * ============================================================================
 */

function detectAndLogMijyusokuDailyDiffs_(ss) {
  const pasteSheet = ss.getSheetByName('貼付用');
  if (!pasteSheet || pasteSheet.getLastRow() < 3) return [];

  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const limitStr = Utilities.formatDate(new Date(Date.now() + 35 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');

  let snapSheet = ss.getSheetByName('_シフト日次基準');
  if (!snapSheet) {
    snapSheet = ss.insertSheet('_シフト日次基準');
    snapSheet.getRange(1, 1, 1, 9).setValues([['基準日', '勤務日', '拠点名', '診療科', '医師名', '医籍番号', '実働分数', '時間帯', 'ステータス']]);
    snapSheet.hideSheet();
  }

  let logSheet = ss.getSheetByName('_シフト変動検知ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('_シフト変動検知ログ');
    logSheet.getRange(1, 1, 1, 11).setValues([['検知日', '勤務日', '残日数', '拠点名', '診療科', '医師名', '医籍番号', '変動区分', '喪失分数', '詳細', '一意キー']]);
    logSheet.hideSheet();
  }

  // 1. 現在の『貼付用』シートを医師×勤務日単位で構造化
  const pVals = pasteSheet.getDataRange().getValues();
  const currByDocDate = new Map();
  const currSnapshotRows = [];

  for (let i = 2; i < pVals.length; i++) {
    const r = pVals[i];
    const docName = String(r[0] || '').replace(/\s+/g, '').trim();
    const licNo = String(r[8] || '').replace(/\.0+$/, '').trim();
    const cNo = String(r[11] || '').trim();
    const cRaw = String(r[12] || '').trim();
    const dept = String(r[13] || '小児科').trim();
    const dObj = parseDateToSafeDateObj(r[14]);
    if (!docName || !cRaw || !dObj) continue;

    const dStr = fastFormatDate(dObj);
    if (dStr < todayStr || dStr > limitStr) continue;

    const wageSum = Number(String(r[29] || '0').replace(/,/g, '')) || 0;
    const rmStr = String(r[53] || r[38] || '');
    const isRegular = wageSum === 0 || (/【.+?】/.test(rmStr) && /曜日|毎週|第\d/.test(rmStr));
    const docKey = licNo || docName;
    const key = `${dStr}_${docKey}`;

    if (!currByDocDate.has(key)) {
      currByDocDate.set(key, { dStr, docKey, docName, licNo, isFullTime: wageSum === 0, isRegular, hasYukyu: false, hasKekkin: false, kekkinMin: 0, clinics: new Map() });
    }
    const entry = currByDocDate.get(key);
    if (wageSum === 0) entry.isFullTime = true;

    const ivs = typeof extractWorkIntervals === 'function' ? extractWorkIntervals(r) : [];
    const workMin = ivs.reduce((s, iv) => s + Math.max(0, iv.end - iv.start), 0);
    const ivText = ivs.map(iv => `${safeFormatTime(iv.start)}-${safeFormatTime(iv.end)}`).join('/');

    if (cRaw === '有給' || cRaw === '有休' || cNo === '1017') {
      entry.hasYukyu = true;
      currSnapshotRows.push([todayStr, dStr, '有給', dept, docName, licNo, workMin, ivText, 'YUKYU']);
      continue;
    }
    if (cRaw === '欠勤' || cNo === '1018') {
      entry.hasKekkin = true;
      entry.kekkinMin += workMin;
      currSnapshotRows.push([todayStr, dStr, '欠勤', dept, docName, licNo, workMin, ivText, 'KEKKIN']);
      continue;
    }
    if (/バックアップ|院外|ワクチン|医師会|嘱託医|休館/.test(cRaw) || (cNo && Number(cNo) >= 1000)) continue;

    const cleanC = normalizeClinicName(cRaw);
    const cKey = `${cleanC}_${dept}`;
    if (!entry.clinics.has(cKey)) entry.clinics.set(cKey, { clinic: cleanC, dept, min: 0, ivs: [] });
    const ce = entry.clinics.get(cKey);
    ce.min += workMin;
    if (ivText) ce.ivs.push(ivText);

    currSnapshotRows.push([todayStr, dStr, cleanC, dept, docName, licNo, workMin, ivText, wageSum === 0 ? 'FT' : (isRegular ? 'REG' : 'SPOT')]);
  }

  // 2. 前回基準と比較して差分を抽出
  if (snapSheet.getLastRow() >= 2) {
    const sVals = snapSheet.getRange(2, 1, snapSheet.getLastRow() - 1, 9).getDisplayValues();
    const prevBaseDate = String(sVals[0][0] || '').trim();
    const prevByDocDate = new Map();

    sVals.forEach(r => {
      const dStr = normalizeBridgeDate_(r[1]);
      if (!dStr || dStr < todayStr) return;
      const docName = r[4], licNo = r[5], docKey = licNo || docName;
      const key = `${dStr}_${docKey}`;
      if (!prevByDocDate.has(key)) {
        prevByDocDate.set(key, { dStr, docKey, docName, licNo, status: r[8], hasYukyu: false, hasKekkin: false, clinics: new Map() });
      }
      const pe = prevByDocDate.get(key);
      if (r[8] === 'YUKYU') pe.hasYukyu = true;
      else if (r[8] === 'KEKKIN') pe.hasKekkin = true;
      else {
        const cKey = `${r[2]}_${r[3]}`;
        if (!pe.clinics.has(cKey)) pe.clinics.set(cKey, { clinic: r[2], dept: r[3], min: 0, ivs: [], status: r[8] });
        const ce = pe.clinics.get(cKey);
        ce.min += Number(r[6] || 0);
        if (r[7]) ce.ivs.push(r[7]);
      }
    });

    const existingKeys = new Set();
    if (logSheet.getLastRow() >= 2) {
      logSheet.getRange(2, 11, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(r => { if (r[0]) existingKeys.add(r[0]); });
    }

    const newDiffs = [];
    const calcDays = wStr => Math.max(0, Math.round((new Date(wStr).getTime() - new Date(todayStr).getTime()) / 86400000));

    for (const [key, prev] of prevByDocDate.entries()) {
      const curr = currByDocDate.get(key);
      const daysBefore = calcDays(prev.dStr);

      for (const [cKey, pc] of prev.clinics.entries()) {
        const cc = curr ? curr.clinics.get(cKey) : null;
        if (cc) {
          if (cc.min + 15 <= pc.min) {
            const lost = pc.min - cc.min;
            const uKey = `${todayStr}_${prev.dStr}_${prev.docKey}_${cKey}_PARTIAL_${lost}`;
            if (!existingKeys.has(uKey)) {
              existingKeys.add(uKey);
              newDiffs.push([todayStr, prev.dStr, daysBefore, pc.clinic, pc.dept, prev.docName, prev.licNo, '部分キャンセル', lost, `${pc.ivs.join(',')}➔${cc.ivs.join(',')}`, uKey]);
            }
          }
        } else {
          // 勤務シフトが消えた場合：欠勤変更 or 有給変更 or 他拠点移動 or シフトキャンセル を厳密分岐
          if (curr && curr.hasKekkin) {
            const typeLabel = pc.status === 'FT' ? '常勤欠勤' : '定期キャンセル(欠勤)';
            const uKey = `${todayStr}_${prev.dStr}_${prev.docKey}_${cKey}_KEKKIN`;
            if (!existingKeys.has(uKey)) {
              existingKeys.add(uKey);
              newDiffs.push([todayStr, prev.dStr, daysBefore, pc.clinic, pc.dept, prev.docName, prev.licNo, typeLabel, pc.min, `${pc.ivs.join(',')}➔欠勤シフト変更`, uKey]);
            }
          } else if (curr && curr.hasYukyu) {
            const uKey = `${todayStr}_${prev.dStr}_${prev.docKey}_${cKey}_YUKYU`;
            if (!existingKeys.has(uKey)) {
              existingKeys.add(uKey);
              newDiffs.push([todayStr, prev.dStr, daysBefore, pc.clinic, pc.dept, prev.docName, prev.licNo, '直前有給変更', pc.min, `${pc.ivs.join(',')}➔有給変更`, uKey]);
            }
          } else if (curr && curr.clinics.size > 0 && Array.from(curr.clinics.keys()).some(k => !prev.clinics.has(k))) {
            const dest = Array.from(curr.clinics.values()).map(x => x.clinic).join('・');
            const uKey = `${todayStr}_${prev.dStr}_${prev.docKey}_${cKey}_MOVE`;
            if (!existingKeys.has(uKey)) {
              existingKeys.add(uKey);
              newDiffs.push([todayStr, prev.dStr, daysBefore, pc.clinic, pc.dept, prev.docName, prev.licNo, '移動', pc.min, `${pc.clinic}➔${dest}`, uKey]);
            }
          } else if (pc.status !== 'FT') {
            const uKey = `${todayStr}_${prev.dStr}_${prev.docKey}_${cKey}_CANCEL`;
            if (!existingKeys.has(uKey)) {
              existingKeys.add(uKey);
              newDiffs.push([todayStr, prev.dStr, daysBefore, pc.clinic, pc.dept, prev.docName, prev.licNo, 'シフトキャンセル', pc.min, `${pc.ivs.join(',')}が消滅`, uKey]);
            }
          }
        }
      }
    }

    if (newDiffs.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, newDiffs.length, 11).setValues(newDiffs);
    }

    // 日付が変わった最初の実行時に前日基準スナップショットを更新
    if (prevBaseDate !== todayStr && currSnapshotRows.length > 0) {
      snapSheet.getRange(2, 1, snapSheet.getLastRow() - 1, 9).clearContent();
      snapSheet.getRange(2, 1, currSnapshotRows.length, 9).setValues(currSnapshotRows);
    }
  } else if (currSnapshotRows.length > 0) {
    snapSheet.getRange(2, 1, currSnapshotRows.length, 9).setValues(currSnapshotRows);
  }

  // 直近14日以内に検知された変動イベントを配列で返却（AI連携パック送信用）
  const cutoffDetect = Utilities.formatDate(new Date(Date.now() - 14 * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  if (logSheet.getLastRow() < 2) return [];
  return logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 11).getDisplayValues()
    .filter(r => normalizeBridgeDate_(r[0]) >= cutoffDetect)
    .map(r => ({
      detectD: normalizeBridgeDate_(r[0]),
      workD: normalizeBridgeDate_(r[1]),
      daysBefore: Number(r[2] || 0),
      clinic: r[3], dept: r[4], doc: r[5], id: r[6],
      type: r[7], lostMin: Number(r[8] || 0), note: r[9], uKey: r[10]
    }));
}