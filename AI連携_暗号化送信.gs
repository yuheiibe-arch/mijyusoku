/**
 * ============================================================================
 * AI連携_暗号化送信.gs（mnys.jp 側：指定クリニックNo・非拠点シフト完全除外版）
 * 宛先: capsdoctor365@caps365.jp
 * ============================================================================
 */

const AI_BRIDGE_CONFIG = {
  RECIPIENT_EMAIL: 'capsdoctor365@caps365.jp',
  SUBJECT_PREFIX: '【AI連携暗号化JSON】医師不在・シフト状況パック',
  EXT_SHIFT_SS_ID: '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co',
  DEFAULT_SECRET_KEY: 'CAPS_MNYS_AI_AGENT_SECRET_2026_KEY_#9876',

  PENDING_FLAG_KEY: 'AI_BRIDGE_PENDING_SEND',
  LAST_SENT_HASH_KEY: 'AI_BRIDGE_LAST_SENT_HASH',
  LAST_SENT_TIME_KEY: 'AI_BRIDGE_LAST_SENT_TIME_MS',

  MIN_INTERVAL_MS: 30 * 60 * 1000,

  // ★連勤カウントから除外するクリニックNoリスト
  EXCLUDED_CLINIC_NOS: new Set(['1000', '1002', '1004', '1005', '1017', '1018']),

  // ★連勤カウントから除外するシフト名称リスト
  EXCLUDED_CLINIC_NAMES: [
    '出張インフルエンザワクチン',
    '【埼玉】ワクチンバックアップシフト',
    '職域新型コロナワクチン接種',
    '【千葉】ワクチンバックアップシフト',
    '【東京】ワクチンバックアップシフト',
    '院外勤務（内科）',
    '院外勤務（小児科）',
    '【関東】バックアップシフト',
    '欠勤',
    '有給',
    '嘱託医業務',
    '医師会業務',
    '休館'
  ]
};

function triggerAiBridgeExport() {
  const props = PropertiesService.getScriptProperties();
  try {
    props.deleteProperty(AI_BRIDGE_CONFIG.PENDING_FLAG_KEY);
    sendEncryptedShiftContextToCaps(false);
  } catch (e) {
    props.setProperty(AI_BRIDGE_CONFIG.PENDING_FLAG_KEY, String(Date.now()));
    Logger.log('直接送信できなかったため、1回限りの代理送信予約をセットしました: ' + e.message);
  }
}

function installAiBridgeFallbackTrigger() {
  const funcName = 'processPendingAiBridgeExport';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(funcName).timeBased().everyMinutes(10).create();
  Logger.log('✅ AI連携のバックグラウンド代理送信トリガー（10分間隔）を設定しました。');
}

function processPendingAiBridgeExport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const props = PropertiesService.getScriptProperties();
    const pendingTimestamp = props.getProperty(AI_BRIDGE_CONFIG.PENDING_FLAG_KEY);
    if (!pendingTimestamp) return;

    props.deleteProperty(AI_BRIDGE_CONFIG.PENDING_FLAG_KEY);

    if (Date.now() - Number(pendingTimestamp) > 2 * 60 * 60 * 1000) {
      Logger.log('古い代理送信フラグのため破棄しました。');
      return;
    }

    sendEncryptedShiftContextToCaps(false);
  } catch (e) {
    Logger.log('代理送信時にエラーが発生しました（暴走防止のため再予約はしません）: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function testSendEncryptedShiftContextManual() {
  sendEncryptedShiftContextToCaps(true);
}

function sendEncryptedShiftContextToCaps(isForceManual) {
  const props = PropertiesService.getScriptProperties();
  const nowMs = Date.now();

  const lastSentTimeMs = Number(props.getProperty(AI_BRIDGE_CONFIG.LAST_SENT_TIME_KEY) || 0);
  if (!isForceManual && lastSentTimeMs > 0 && (nowMs - lastSentTimeMs) < AI_BRIDGE_CONFIG.MIN_INTERVAL_MS) {
    Logger.log('🛑 暴走防止ストッパー作動：直近30分以内に送信済みのためメール送信をスキップしました。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payloadObj = buildAiShiftContextPayload_(ss);

  if (!payloadObj || payloadObj.unfilledList.length === 0) {
    Logger.log('AI連携送信スキップ：対象の不在枠データがありません。');
    return;
  }

  const contentForHash = JSON.stringify({
    u: payloadObj.unfilledList,
    a: payloadObj.activeShiftsByDate,
    o: payloadObj.ouboShiftsByDate,
    s: payloadObj.streakWarnings,
    k: payloadObj.spotCandidateRows
  });
  const currentHash = computeSha256Hex_(contentForHash);
  const lastSentHash = props.getProperty(AI_BRIDGE_CONFIG.LAST_SENT_HASH_KEY) || '';

  if (!isForceManual && currentHash === lastSentHash) {
    Logger.log('🛑 2重送信防止ストッパー作動：前回送信したデータと内容が同一のため送信をスキップしました。');
    return;
  }

  props.setProperty(AI_BRIDGE_CONFIG.LAST_SENT_TIME_KEY, String(nowMs));
  props.setProperty(AI_BRIDGE_CONFIG.LAST_SENT_HASH_KEY, currentHash);

  const secretKey = props.getProperty('AI_BRIDGE_SECRET_KEY')
    || AI_BRIDGE_CONFIG.DEFAULT_SECRET_KEY;

  const jsonString = JSON.stringify(payloadObj);
  const encryptedPackage = encryptAndSignPayload_(jsonString, secretKey);

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const subject = `${AI_BRIDGE_CONFIG.SUBJECT_PREFIX} (${nowStr})`;

  const body = [
    '※このメールは医療法人側システム（mijyusoku）からAIエージェント会議用に自動送信された暗号化データです。',
    `生成日時: ${nowStr}`,
    `不在枠件数: ${payloadObj.unfilledList.length}件`,
    `対象日数: ${Object.keys(payloadObj.activeShiftsByDate).length}日分`,
    `連勤注意候補件数: ${(payloadObj.streakWarnings || []).length}件`,
    `個別打診候補件数: ${Math.max(0, (payloadObj.spotCandidateRows || []).length - 1)}件`,
    `複数日程おまとめ医師数: ${Math.max(0, (payloadObj.bundledCandidateRows || []).length - 1)}名`,
    '',
    '---BEGIN_ENCRYPTED_AI_PAYLOAD---',
    JSON.stringify(encryptedPackage),
    '---END_ENCRYPTED_AI_PAYLOAD---'
  ].join('\n');

  MailApp.sendEmail({
    to: AI_BRIDGE_CONFIG.RECIPIENT_EMAIL,
    subject: subject,
    body: body,
    noReply: true
  });

  Logger.log(`✅ AI連携暗号化JSONを ${AI_BRIDGE_CONFIG.RECIPIENT_EMAIL} へ送信しました（不在枠: ${payloadObj.unfilledList.length}件 / 連勤警告: ${(payloadObj.streakWarnings || []).length}件）`);
}

function computeSha256Hex_(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function cleanBaseClinicName_(rawName) {
  return String(rawName || '')
    .replace(/[\(（].*?[\)）]/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

function normalizeDocKey_(name) {
  return String(name || '')
    .replace(/[\(（].*?[\)）]/g, '')
    .replace(/先生$/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

/**
 * ★ご指定の除外クリニックNo（1000, 1002, 1004, 1005, 1017, 1018）および除外名称リストに該当するかを判定
 */
function isExcludedFromStreakCount_(clinicNo, clinicName) {
  const cNo = String(clinicNo || '').trim();
  const cName = String(clinicName || '').trim();
  if (!cName) return true;

  // 1. クリニックNoが指定の除外番号（または1000番台）なら除外
  if (cNo && (AI_BRIDGE_CONFIG.EXCLUDED_CLINIC_NOS.has(cNo) || Number(cNo) >= 1000)) {
    return true;
  }

  // 2. 名称が除外リストに含まれる場合は除外
  for (const exKeyword of AI_BRIDGE_CONFIG.EXCLUDED_CLINIC_NAMES) {
    if (cName.includes(exKeyword)) return true;
  }

  return false;
}

function buildAiShiftContextPayload_(ss) {
  const approachSheet = ss.getSheetByName('アプローチシート');
  const confirmSheet = ss.getSheetByName('確認用');
  const pasteSheet = ss.getSheetByName('貼付用');
  const reportSheet = ss.getSheetByName('文章自動作成');

  if (!approachSheet || !confirmSheet) {
    throw new Error('「アプローチシート」または「確認用」シートが見つかりません。');
  }

  const clinicNoMap = new Map();
  const docLicenseMap = new Map();
  const docDailyStatusMap = new Map();
  const docWorkCalendarMap = new Map();

  if (pasteSheet && pasteSheet.getLastRow() >= 3) {
    const pData = pasteSheet.getDataRange().getDisplayValues();
    const colDocName = 0;     // 1列目: 名前
    const colLicenseNo = 8;   // 9列目: 医籍番号
    const colClinicNo = 11;   // 12列目: クリニックNo
    const colClinicName = 12; // 13列目: クリニック名
    const colDept = 13;       // 14列目: 診療科
    const colDate = 14;       // 15列目: 勤務日
    const colWageSum = 29;    // 30列目: 時給合計

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

      // ★ご指定の除外リスト（1000, 1002, 1004, 1005, 1017, 1018 やワクチン・院外業務等）は連勤カウントから完全除外
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
          workDates: new Set() // 1日単位で格納（同日に複数枠あっても1日としてカウント）
        });
      }
      const calEntry = docWorkCalendarMap.get(docKey);
      if (isFullTime) calEntry.isFullTimeEver = true;
      if (isShoteiKyushutsu) calEntry.hasShoteiEver = true;
      if (licNo && !calEntry.licenseNo) calEntry.licenseNo = licNo;
      calEntry.workDates.add(normDate);
    }
  }

  // 外部『確定シフト』からの前月後半補完（ここでも除外リストを厳密適用）
  try {
    const extSs = SpreadsheetApp.openById(AI_BRIDGE_CONFIG.EXT_SHIFT_SS_ID);
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

  const attachBaseFullTimeTag = (rawNamesStr, dateStr) => {
    if (!rawNamesStr) return '';
    return String(rawNamesStr).split(/([,、，\n]+)/).map(token => {
      if (/^[,、，\n]+$/.test(token) || !token.trim()) return token;

      const parts = token.split(/([：:])/);
      const namePart = parts[0].trim();
      const restPart = parts.slice(1).join('');
      const docKey = normalizeDocKey_(namePart);
      if (!docKey || docKey === '未充足') return token;

      const st = docDailyStatusMap.get(`${dateStr}_${docKey}`);
      const cal = docWorkCalendarMap.get(docKey);
      const isFull = (st && st.isFullTime) || (cal && cal.isFullTimeEver);
      const isShotei = Boolean(st && st.isShoteiKyushutsu);

      let tag = '';
      if (isFull) {
        tag = isShotei ? '(常勤・所定休出)' : '(常勤)';
      }
      return `${namePart}${tag}${restPart}`;
    }).join('');
  };

  const resolveClinicNo = (clinicRaw, deptRaw) => {
    const base = cleanBaseClinicName_(clinicRaw);
    const dept = String(deptRaw || '').trim();
    if (dept && clinicNoMap.has(`${base}_${dept}`)) return clinicNoMap.get(`${base}_${dept}`);
    if (clinicNoMap.has(base)) return clinicNoMap.get(base);
    return '';
  };

  const targetDateSet = new Set();
  const unfilledList = [];

  // 1. 『アプローチシート』
  const apLastRow = approachSheet.getLastRow();
  const apLastCol = approachSheet.getLastColumn();
  if (apLastRow >= 3) {
    const headers = approachSheet.getRange(1, 1, 1, apLastCol).getDisplayValues()[0].map(h => String(h || '').trim());
    const apRows = approachSheet.getRange(3, 1, apLastRow - 2, apLastCol).getDisplayValues();
    const idx = name => headers.indexOf(name);

    for (let i = 0; i < apRows.length; i++) {
      const r = apRows[i];
      const rawDate = String(r[idx('日付')] || '').trim();
      const clinic = String(r[idx('拠点名')] || '').trim();
      const absenceTime = String(r[idx('不在時間')] || '').trim();
      if (!rawDate || !clinic || !absenceTime) continue;

      const normDate = normalizeBridgeDate_(rawDate);
      if (!normDate) continue;

      const status = idx('対応状況') >= 0 ? String(r[idx('対応状況')] || '').trim() : '';
      if (status === '完了') continue;

      targetDateSet.add(normDate);

      const rawPrevDocs = idx('前の時間枠の医師') >= 0 ? String(r[idx('前の時間枠の医師')] || '').trim() : '';
      const rawNextDocs = idx('後ろの時間枠の医師') >= 0 ? String(r[idx('後ろの時間枠の医師')] || '').trim() : '';

      unfilledList.push({
        date: normDate,
        clinic: clinic,
        absenceTime: absenceTime,
        listedWage: idx('掲載時給') >= 0 ? String(r[idx('掲載時給')] || '').trim() : '',
        prevDocs: attachBaseFullTimeTag(rawPrevDocs, normDate),
        nextDocs: attachBaseFullTimeTag(rawNextDocs, normDate),
        phase: idx('Phaze') >= 0 ? String(r[idx('Phaze')] || '').trim() : (idx('Phase') >= 0 ? String(r[idx('Phase')] || '').trim() : ''),
        prevExtStatus: idx('前医延長') >= 0 ? String(r[idx('前医延長')] || '').trim() : '',
        nextEarlyStatus: idx('後医早出') >= 0 ? String(r[idx('後医早出')] || '').trim() : '',
        assignee: idx('担当') >= 0 ? String(r[idx('担当')] || '').trim() : '',
        status: status,
        specialWage: idx('特別時給') >= 0 ? String(r[idx('特別時給')] || '').trim() : '',
        memo: idx('作業メモ') >= 0 ? String(r[idx('作業メモ')] || '').trim() : ''
      });
    }
  }

  // 2. 『確認用』シート
  const activeShiftsByDate = {};
  const cfLastRow = confirmSheet.getLastRow();
  if (cfLastRow >= 2 && targetDateSet.size > 0) {
    const cfRows = confirmSheet.getRange(2, 1, cfLastRow - 1, 10).getDisplayValues();
    for (let i = 0; i < cfRows.length; i++) {
      const r = cfRows[i];
      const clinicName = String(r[0] || '').trim();
      const normDate = normalizeBridgeDate_(r[1]);
      const dept = String(r[2] || '').trim();
      if (!clinicName || !normDate || !targetDateSet.has(normDate)) continue;

      const clinicNo = resolveClinicNo(clinicName, dept);

      if (!activeShiftsByDate[normDate]) activeShiftsByDate[normDate] = [];
      activeShiftsByDate[normDate].push({
        clinicNo: clinicNo,
        clinic: clinicName,
        dept: dept,
        countA: Number(r[3] || 0),
        countB: Number(r[4] || 0),
        countC: Number(r[5] || 0),
        docsA: attachBaseFullTimeTag(String(r[7] || '').trim(), normDate),
        docsB: attachBaseFullTimeTag(String(r[8] || '').trim(), normDate),
        docsC: attachBaseFullTimeTag(String(r[9] || '').trim(), normDate)
      });
    }
  }

  // 3. 外部『応募シフト』
  const ouboShiftsByDate = {};
  if (targetDateSet.size > 0) {
    try {
      const extSs = SpreadsheetApp.openById(AI_BRIDGE_CONFIG.EXT_SHIFT_SS_ID);
      const ouboSheet = extSs.getSheetByName('応募シフト');
      if (ouboSheet && ouboSheet.getLastRow() >= 2) {
        const ouboRows = ouboSheet.getDataRange().getDisplayValues();
        const oHeaders = ouboRows[0].map(h => String(h || '').trim());

        const findCol = (names, fallbackIdx) => {
          for (const n of names) {
            const idx = oHeaders.indexOf(n);
            if (idx !== -1) return idx;
          }
          return fallbackIdx;
        };

        const idxDoc = findCol(['名前', '医師名', '氏名'], 0);
        const idxLic = findCol(['医籍番号', '医籍登録番号'], -1);
        const idxClinicNo = findCol(['クリニックNo', '拠点No', 'クリニックID'], -1);
        const idxClinic = findCol(['クリニック名', '拠点名'], 3);
        const idxDept = findCol(['診療科', '診療科目'], 4);
        const idxDate = findCol(['勤務日', '日付'], 5);
        const idxStart = findCol(['勤務開始時間', '開始時間'], 6);
        const idxEnd = findCol(['勤務終了時間', '終了時間'], 7);

        for (let i = 1; i < ouboRows.length; i++) {
          const r = ouboRows[i];
          const docName = String(r[idxDoc] || '').trim();
          const clinic = String(r[idxClinic] || '').trim();
          const dept = String(r[idxDept] || '').trim();
          const normDate = normalizeBridgeDate_(r[idxDate]);
          const startT = String(r[idxStart] || '').trim();
          const endT = String(r[idxEnd] || '').trim();
          if (!docName || !normDate || !targetDateSet.has(normDate)) continue;

          const cleanDocKey = normalizeDocKey_(docName);
          const licenseNo = (idxLic !== -1 && r[idxLic])
            ? String(r[idxLic]).trim()
            : (docLicenseMap.get(cleanDocKey) || '');

          const clinicNo = (idxClinicNo !== -1 && r[idxClinicNo])
            ? String(r[idxClinicNo]).trim()
            : resolveClinicNo(clinic, dept);

          if (!ouboShiftsByDate[normDate]) ouboShiftsByDate[normDate] = [];
          ouboShiftsByDate[normDate].push({
            licenseNo: licenseNo,
            doctor: attachBaseFullTimeTag(docName, normDate),
            clinicNo: clinicNo,
            clinic: clinic,
            dept: dept,
            time: `${startT}-${endT}`
          });
        }
      }
    } catch (e) {
      Logger.log('応募シフト参照スキップ: ' + e.message);
    }
  }

  // 4. 連続勤務ブロックデータ（純粋な拠点勤務日数のみで判定）
  const streakWarnings = buildConsecutiveWorkWarnings_(docWorkCalendarMap);

  // 5. 未充足報告テキスト
  const summaryText = reportSheet ? String(reportSheet.getRange('A6').getDisplayValue() || '').trim() : '';

  // 6. 勤務くん連携
  let spotCandidateRows = [];
  let bundledCandidateRows = [];
  if (typeof buildKinmukunCandidatesFromUnfilledList_ === 'function' && unfilledList.length > 0) {
    try {
      const kinmukunRes = buildKinmukunCandidatesFromUnfilledList_(unfilledList);
      spotCandidateRows = kinmukunRes.spotCandidateRows || [];
      bundledCandidateRows = kinmukunRes.bundledCandidateRows || [];
    } catch (e) {
      Logger.log('勤務くん候補抽出スキップ: ' + e.message);
    }
  }

  return {
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    unfilledList: unfilledList,
    activeShiftsByDate: activeShiftsByDate,
    ouboShiftsByDate: ouboShiftsByDate,
    streakWarnings: streakWarnings,
    spotCandidateRows: spotCandidateRows,
    bundledCandidateRows: bundledCandidateRows,
    summaryText: summaryText
  };
}

/**
 * パターンA：対象期間付近で実際に「追加打診NG日（11連勤以上または橋渡し12連勤以上）」が発生するケースのみ抽出
 */
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

  for (const [docKey, info] of docWorkCalendarMap.entries()) {
    const sortedDates = Array.from(info.workDates).sort();
    if (sortedDates.length < 6) continue;

    const blocks = [];
    let curStart = sortedDates[0];
    let curEnd = sortedDates[0];
    let curLen = 1;

    for (let i = 1; i < sortedDates.length; i++) {
      const prevMs = toDateMs(curEnd);
      const currMs = toDateMs(sortedDates[i]);
      const diffDays = Math.round((currMs - prevMs) / 86400000);

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

      // パターンA：実際にNG日が発生する11連勤以上のみ
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

function normalizeBridgeDate_(rawVal) {
  if (!rawVal) return '';
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    return Utilities.formatDate(rawVal, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  const clean = String(rawVal)
    .split(/[（(]/)[0]
    .trim()
    .replace(/[年月\.\-]/g, '/')
    .replace(/日/g, '');
  const parts = clean.split('/');
  if (parts.length >= 3) {
    const y = parts[0];
    const m = ('0' + parts[1]).slice(-2);
    const d = ('0' + parts[2]).slice(-2);
    return `${y}/${m}/${d}`;
  }
  return '';
}

function encryptAndSignPayload_(plainText, secretKey) {
  const textBlob = Utilities.newBlob(plainText, 'application/json', 'payload.json');
  const gzipBytes = Utilities.gzip(textBlob).getBytes();
  const ivStr = Utilities.getUuid().replace(/-/g, '').slice(0, 16);

  const cipherBytes = xorStreamCipher_(gzipBytes, secretKey, ivStr);
  const cipherBase64 = Utilities.base64Encode(cipherBytes);

  const sigBytes = Utilities.computeHmacSha256Signature(ivStr + ':' + cipherBase64, secretKey);
  const sigBase64 = Utilities.base64Encode(sigBytes);

  return {
    v: 1,
    iv: ivStr,
    data: cipherBase64,
    sig: sigBase64
  };
}

function xorStreamCipher_(inputBytes, secretKey, ivStr) {
  const output = new Array(inputBytes.length);
  let keyStream = [];
  let blockIndex = 0;

  for (let i = 0; i < inputBytes.length; i++) {
    if (i % 32 === 0) {
      keyStream = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        `${secretKey}|${ivStr}|${blockIndex}`,
        Utilities.Charset.UTF_8
      );
      blockIndex++;
    }
    output[i] = (inputBytes[i] ^ keyStream[i % 32]) & 0xff;
    if (output[i] > 127) output[i] -= 256;
  }
  return output;
}