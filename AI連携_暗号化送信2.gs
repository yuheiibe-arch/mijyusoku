/**
 * ============================================================================
 * 【mnys.jp 側】AI連携_暗号化送信2.gs（2大保管ブック横断検索・前月/前々月自動合体・全欠損月一括補完版）
 * - sendEncryptedBqShiftsToCaps()：通常送信時、前月・前々月（例: 2026/09, 2026/10）を2つの保管ブックから自動横断検索して合体送信
 * - sendAllMissingMonthlyArchivesToCaps()：caps側で199行切れ・未格納だった過去9ヶ月分を一括補完送信
 * ============================================================================
 */

const BQ_SYNC_CONFIG = {
  BQ_SHIFT_SS_ID: "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co",
  MONTHLY_ARCHIVE_SS_IDS: [
    "10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw",
    "1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc"
  ],
  TO_ADDRESS: "doctor-support@caps365.jp",
  SUBJECT_PREFIX: "【AI連携暗号化JSON】BQシフト状況パック",
  ARCHIVE_SUBJECT_PREFIX: "【AI連携暗号化JSON】月次確定シフト補完パック",
  ATTACHMENT_FILENAME: "bq_shift_pack.enc",
  DEFAULT_SECRET_KEY: "CAPS_AI_BRIDGE_SECRET_2026",
  TARGET_DAYS: 62
};

/**
 * ⏰ 本番・手動送信関数（前月・前々月の確定シフトも月次保管ブックから自動横断補完して送信）
 */
function sendEncryptedBqShiftsToCaps() {
  const props = PropertiesService.getScriptProperties();
  const secretKey = props.getProperty("BQ_SYNC_SECRET_KEY") || BQ_SYNC_CONFIG.DEFAULT_SECRET_KEY;

  const payload = buildBqShiftPayload_();
  const jsonString = JSON.stringify(payload);
  const encryptedText = encryptPayloadForMail_(jsonString, secretKey);

  const attachmentBlob = Utilities.newBlob(
    encryptedText,
    "application/octet-stream",
    BQ_SYNC_CONFIG.ATTACHMENT_FILENAME
  );

  const subject = `${BQ_SYNC_CONFIG.SUBJECT_PREFIX} (${payload.generatedAt})`;
  const bodySummary = [
    `BQシフト状況パック（自動送信・過去2ヶ月アーカイブ補完対応）`,
    `抽出期間: ${payload.startDate} 〜 ${payload.maxDate} (基準日: ${payload.baseDate})`,
    `確定シフト: ${payload.kakutei.length}件 (補完内訳: ${payload.archiveSummary || "なし"})`,
    `応募シフト: ${payload.obo.length}件`,
    `募集シフト: ${payload.bosyu.length}件`,
    `暗号化サイズ: ${encryptedText.length.toLocaleString()} bytes`,
    `※データは添付ファイル（${BQ_SYNC_CONFIG.ATTACHMENT_FILENAME}）に格納されています。`
  ].join("\n");

  GmailApp.sendEmail(BQ_SYNC_CONFIG.TO_ADDRESS, subject, bodySummary, {
    attachments: [attachmentBlob]
  });

  console.log(`✅ 添付送信完了！ 期間:${payload.startDate}〜${payload.maxDate} | 確定=${payload.kakutei.length}件 (補完:${payload.archiveSummary || "なし"}), 応募=${payload.obo.length}件, 募集=${payload.bosyu.length}件`);
}

/**
 * 📚 ワンショット補完用：caps側で199行切れ・未格納だった過去月を2つの月次保管ブックから横断検索して一括送信
 */
function sendAllMissingMonthlyArchivesToCaps() {
  const targetYms = [
    "2025/08", "2025/09", "2025/12",
    "2026/04", "2026/06", "2026/07", "2026/08", "2026/09", "2026/10"
  ];

  const props = PropertiesService.getScriptProperties();
  const secretKey = props.getProperty("BQ_SYNC_SECRET_KEY") || BQ_SYNC_CONFIG.DEFAULT_SECRET_KEY;
  const archiveBooks = openAllArchiveBooks_();

  const allRecords = [];
  const summaryParts = [];
  const missingYms = [];

  targetYms.forEach(ym => {
    const [y, m] = ym.split("/").map(Number);
    const mStart = `${ym}/01`;
    const mEnd = Utilities.formatDate(new Date(y, m, 0), "Asia/Tokyo", "yyyy/MM/dd");

    const bestRows = extractBestArchiveMonthAcrossBooks_(archiveBooks, ym, mStart, mEnd);
    if (bestRows.length >= 500) {
      bestRows.forEach(r => allRecords.push(r));
      summaryParts.push(`${ym}:${bestRows.length}件`);
    } else {
      missingYms.push(`${ym}(最大${bestRows.length}件)`);
    }
  });

  if (missingYms.length > 0) {
    console.log(`ℹ️ 500件未満のためスキップした月: ${missingYms.join(", ")}`);
  }
  if (allRecords.length === 0) {
    console.log("⚠️ 送信可能な月次アーカイブ（500件以上）が見つかりませんでした。");
    return;
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const payload = {
    schemaVersion: 2,
    isMonthlyArchivePack: true,
    generatedAt: nowStr,
    months: summaryParts,
    kakutei: allRecords
  };

  const encryptedText = encryptPayloadForMail_(JSON.stringify(payload), secretKey);
  const blob = Utilities.newBlob(encryptedText, "application/octet-stream", "monthly_archive_pack.enc");

  GmailApp.sendEmail(
    BQ_SYNC_CONFIG.TO_ADDRESS,
    `${BQ_SYNC_CONFIG.ARCHIVE_SUBJECT_PREFIX} (${nowStr})`,
    `月次確定シフト一括補完パック\n内訳: ${summaryParts.join(", ")}\n総件数: ${allRecords.length}件`,
    { attachments: [blob] }
  );
  console.log(`✅ 過去月アーカイブ一括送信完了！ 総計 ${allRecords.length}件 (${summaryParts.join(", ")})`);
}

/**
 * 2つの月次保管ブックを安全に開いて配列で返す
 */
function openAllArchiveBooks_() {
  const books = [];
  BQ_SYNC_CONFIG.MONTHLY_ARCHIVE_SS_IDS.forEach(id => {
    try {
      books.push(SpreadsheetApp.openById(id));
    } catch (e) {
      console.log(`保管ブック読込スキップ(${id}): ${e.message}`);
    }
  });
  return books;
}

/**
 * 複数ブックを横断して指定年月(ymStr)に該当する最も件数の多い確定シフト配列を取得する
 */
function extractBestArchiveMonthAcrossBooks_(books, ymStr, mStart, mEnd) {
  const [y, m] = ymStr.split("/").map(Number);
  const shortYr = String(y).slice(-2);
  const candidateSheets = [];

  books.forEach(book => {
    book.getSheets().forEach(sh => {
      const rowsCnt = sh.getLastRow();
      if (rowsCnt < 300) return; // 199行切れやサマリー表は最初から除外

      const name = sh.getName().normalize("NFKC").trim();
      if (/グラフ|集計|マスタ|テンプレ|ピボット|削除|ダミー/i.test(name)) return;

      // パターン1: 年・月の両方がタブ名に含まれている（例: "2026/09", "2026年9月", "202609", "26/09"）
      const hasFullYearMatch = new RegExp(`(^|\\D)${y}[\\/._\\-年\\s]*0?${m}(月|度|確定|$|\\D)`).test(name);
      const hasShortYearMatch = new RegExp(`^${shortYr}[\\/._\\-]0?${m}(月|$|\\D)`).test(name);
      if (hasFullYearMatch || hasShortYearMatch) {
        candidateSheets.push({ sh, priority: 0, rowsCnt });
        return;
      }

      // パターン2: 月のみのタブ名（例: "9月", "09月", "09", "9"）➔ 先頭数行の日付を覗いて年が一致するか検証
      const hasMonthOnlyMatch = new RegExp(`^0?${m}(月|月度|月確定)?$`).test(name);
      if (hasMonthOnlyMatch) {
        try {
          const sampleVals = sh.getRange(1, 1, Math.min(12, rowsCnt), Math.min(20, sh.getLastColumn())).getDisplayValues();
          const joinedSample = sampleVals.map(r => r.join(" ")).join(" ");
          if (joinedSample.includes(`${y}/0${m}`) || joinedSample.includes(`${y}/${m}/`) || joinedSample.includes(`${y}-${('0' + m).slice(-2)}`) || joinedSample.includes(`${y}年${m}月`)) {
            candidateSheets.push({ sh, priority: 1, rowsCnt });
          }
        } catch (e) {}
      }
    });
  });

  candidateSheets.sort((a, b) => (a.priority - b.priority) || (b.rowsCnt - a.rowsCnt));

  let bestRows = [];
  for (let i = 0; i < candidateSheets.length; i++) {
    const extracted = extractShiftRows_(candidateSheets[i].sh, mStart, mEnd, "kakutei");
    if (extracted.length > bestRows.length) {
      bestRows = extracted;
    }
    if (bestRows.length >= 1000) break; // 完全データが取れた時点で確定
  }
  return bestRows;
}

/**
 * 通常BQパック構築：前々月1日（月次レポートの前月＆当月を完全カバー）〜62日後を抽出
 */
function buildBqShiftPayload_() {
  const ss = SpreadsheetApp.openById(BQ_SYNC_CONFIG.BQ_SHIFT_SS_ID);
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");

  // 前々月1日を開始日とする（例: 11月実行時は 2026/09/01〜 をカバー）
  const twoMonthsAgoStart = new Date(now.getFullYear(), now.getMonth() - 2, 1);
  const startDateStr = Utilities.formatDate(twoMonthsAgoStart, "Asia/Tokyo", "yyyy/MM/dd");
  const maxDateObj = new Date(now.getFullYear(), now.getMonth(), now.getDate() + BQ_SYNC_CONFIG.TARGET_DAYS);
  const maxDateStr = Utilities.formatDate(maxDateObj, "Asia/Tokyo", "yyyy/MM/dd");
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  let kakuteiList = extractShiftRows_(ss.getSheetByName("確定シフト"), startDateStr, maxDateStr, "kakutei");
  const archiveLogs = [];
  let archiveBooks = null;

  // 前々月(-2)・前月(-1)について、確定シフト側に1,000件未満しか無ければ2つの保管ブックを横断検索して補完
  [-2, -1].forEach(offset => {
    const mStartObj = new Date(now.getFullYear(), now.getMonth() + offset, 1);
    const mEndObj = new Date(now.getFullYear(), now.getMonth() + offset + 1, 0);
    const mStart = Utilities.formatDate(mStartObj, "Asia/Tokyo", "yyyy/MM/dd");
    const mEnd = Utilities.formatDate(mEndObj, "Asia/Tokyo", "yyyy/MM/dd");
    const ym = mStart.slice(0, 7);

    const existingCnt = kakuteiList.filter(r => r.d >= mStart && r.d <= mEnd).length;
    if (existingCnt >= 1000) return;

    if (!archiveBooks) archiveBooks = openAllArchiveBooks_();
    const archRows = extractBestArchiveMonthAcrossBooks_(archiveBooks, ym, mStart, mEnd);
    if (archRows.length > existingCnt) {
      kakuteiList = kakuteiList.filter(r => r.d < mStart || r.d > mEnd).concat(archRows);
      archiveLogs.push(`${ym}:${archRows.length}件`);
    }
  });

  kakuteiList.sort((a, b) => a.d.localeCompare(b.d));

  return {
    schemaVersion: 2,
    generatedAt: nowStr,
    baseDate: todayStr,
    startDate: startDateStr,
    maxDate: maxDateStr,
    archiveSummary: archiveLogs.join(", "),
    kakutei: kakuteiList,
    obo:     extractShiftRows_(ss.getSheetByName("応募シフト"), todayStr, maxDateStr, "obo"),
    bosyu:   extractShiftRows_(ss.getSheetByName("募集シフト"), todayStr, maxDateStr, "bosyu")
  };
}

function extractShiftRows_(sheet, startDateStr, maxDateStr, sheetType) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const isBosyuSheet = (sheetType === "bosyu");
  const isKakuteiSheet = (sheetType === "kakutei");

  const lastCol = Math.min(65, sheet.getLastColumn());
  const vals = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getDisplayValues();
  let hRow = 0;
  for (let r = 0; r < Math.min(8, vals.length); r++) {
    const joined = vals[r].join("|");
    if (joined.includes("勤務日") || (joined.includes("クリニック") && joined.includes("開始"))) {
      hRow = r;
      break;
    }
  }
  const headers = vals[hRow].map(h => String(h || "").normalize("NFKC").replace(/\s+/g, "").trim());

  const findIdx = (cands, exRe) => {
    for (let i = 0; i < cands.length; i++) {
      const hit = headers.findIndex(h => h === cands[i] && (!exRe || !exRe.test(h)));
      if (hit !== -1) return hit;
    }
    for (let i = 0; i < cands.length; i++) {
      const hit = headers.findIndex(h => h.includes(cands[i]) && (!exRe || !exRe.test(h)));
      if (hit !== -1) return hit;
    }
    return -1;
  };

  const iDate = findIdx(["勤務日", "日付"], /応募|作成|元/);
  const iClinicNo = findIdx(["クリニックNo", "拠点No"], null);
  const iClinic = findIdx(["クリニック名", "拠点名", "勤務先"], /No|コード/i);
  const iDept = findIdx(["診療科", "科目"], null);
  const iStart = findIdx(["勤務開始時間", "開始時間"], /元のシフト/);
  const iEnd1 = findIdx(["終了時間1"], /元のシフト/);
  const iEnd2 = findIdx(["終了時間2"], /元のシフト/);
  const iEnd3 = findIdx(["終了時間3"], /元のシフト/);
  const iEnd = findIdx(["勤務終了時間", "終了時間"], /元のシフト|終了時間[123]/);
  const iName = findIdx(["名前", "医師名", "氏名", "医師氏名"], /カナ|カタカナ|ふりがな|口座|名義|担当|承認|変更|クリニック|拠点/);
  const iId = findIdx(["医籍番号", "医籍"], null);
  const iKana = findIdx(["名前(カタカナ)", "カタカナ", "フリガナ"], /口座/);
  const iWage = findIdx(["時給合計", "支払合計", "時給"], /時給[1-4]|特別時給/);
  const iRemarks = findIdx(["業務内容(備考)", "業務内容", "備考"], /入職経緯/);
  const iApplyDate = findIdx(["応募日", "医師シフト応募日時"], null);
  const iCreateDate = findIdx(["シフト作成日", "作成日"], null);
  const iStatus = findIdx(["掲載ステータス"], null);
  const iSecondFlag = findIdx(["二診目表示"], null);
  const iSpecialSetting = findIdx(["特別時給設定"], null);

  const wageSlotCols = ["時給1", "時給2", "時給3", "時給4"].map(n => headers.indexOf(n));
  const unitWageCols = wageSlotCols.filter(c => c !== -1);
  const capSlotCols = ["受入可能人数1", "受入可能人数2", "受入可能人数3", "受入可能人数4"].map(n => headers.indexOf(n));

  const allowanceCols = [], staffCommentCols = [];
  headers.forEach((h, c) => {
    if (h.includes("追加支給項目") || h.includes("手当")) allowanceCols.push(c);
    if (h.includes("スタッフコメント")) staffCommentCols.push(c);
  });

  if (iDate === -1 || iClinic === -1) return [];
  const fbYear = startDateStr.slice(0, 4);

  const list = [];
  for (let r = hRow + 1; r < vals.length; r++) {
    const row = vals[r];
    const dateStr = normalizeYmd_(row[iDate], fbYear);
    if (!dateStr || dateStr < startDateStr || dateStr > maxDateStr) continue;

    const clinic = String(row[iClinic] || "").trim();
    if (!clinic) continue;

    if (isKakuteiSheet) {
      if (/休暇|休診|休館|研修|会議|面談|事務/.test(clinic)) continue;
    } else {
      if (/有給|有休|休暇|欠勤|休診|休館|研修|会議|面談|事務/.test(clinic)) continue;
    }

    const docName = iName !== -1 ? String(row[iName] || "").replace(/\s+/g, "").trim() : "";
    if (!isBosyuSheet && (!docName || /休館|勤務なし|医師不在/.test(docName))) continue;

    const cNo = iClinicNo !== -1 ? String(row[iClinicNo] || "").replace(/^[="']+|[="']+$/g, "").replace(/\.0+$/, "").trim() : "";
    const start = iStart !== -1 ? padHhMm_(row[iStart]) : "";
    const end = iEnd !== -1 ? padHhMm_(row[iEnd]) : "";
    const dept = iDept !== -1 ? String(row[iDept] || "").trim() : "";
    const wage = iWage !== -1 ? String(row[iWage] || "").replace(/,/g, "").trim() : "";
    const allowances = allowanceCols.map(c => String(row[c] || "").trim()).filter(Boolean).join("/");

    let maxUnitWage = 0;
    unitWageCols.forEach(c => {
      const v = parseInt(String(row[c] || "").replace(/,/g, "").trim(), 10);
      if (!isNaN(v) && v > maxUnitWage) maxUnitWage = v;
    });

    const rawRm = iRemarks !== -1 ? String(row[iRemarks] || "").trim() : "";
    const rawSc = staffCommentCols.map(c => String(row[c] || "").trim()).filter(Boolean).join(" / ");

    let compactRm = "";
    if (rawRm) {
      if (isBosyuSheet) {
        if (rawRm.includes("特別時給")) compactRm = "特別時給";
        else if (rawRm.includes("3,000円") || rawRm.includes("3000円")) compactRm = "加算3000";
        else if (rawRm.includes("2,000円") || rawRm.includes("2000円")) compactRm = "加算2000";
      } else {
        compactRm = rawRm.replace(/\r?\n/g, " ").trim().substring(0, 120);
      }
    }

    if (isBosyuSheet) {
      const obj = { d: dateStr, c: clinic, dp: dept, t: `${start}-${end}`, w: wage };
      if (cNo) obj.cNo = cNo;
      if (maxUnitWage > 0) obj.uMax = maxUnitWage;
      if (compactRm) obj.rm = compactRm;
      if (iStatus !== -1 && row[iStatus]) obj.st = String(row[iStatus]).trim();
      if (iSecondFlag !== -1 && row[iSecondFlag]) obj.f2 = String(row[iSecondFlag]).trim();
      if (iSpecialSetting !== -1 && row[iSpecialSetting]) obj.sp = String(row[iSpecialSetting]).trim();
      list.push(obj);
    } else {
      const ivStr = buildWorkIntervalStr_(row, [iStart, iEnd1, iEnd2, iEnd3, iEnd], wageSlotCols, iWage, capSlotCols);
      const combinedNotes = `${rawRm} ${rawSc}`;
      let shiftType = "";
      if (clinic === "有給" || clinic === "有休" || cNo === "1017") shiftType = "有給";
      else if (clinic === "欠勤" || cNo === "1018") shiftType = "欠勤";
      else if (/所定休出|所定外休出|所定外勤務/.test(combinedNotes)) shiftType = "所定休出";
      else if ((/【.+?】/.test(rawRm) && /毎週|第\d|隔週|曜日/.test(rawRm)) || /定期非常勤|定期勤務|定期/.test(combinedNotes)) shiftType = "定期";

      const obj = {
        d: dateStr, cNo, c: clinic, dp: dept,
        t: `${start}-${end}`, iv: ivStr || `${start}-${end}`,
        n: docName, id: iId !== -1 ? String(row[iId] || "").replace(/^[="']+|[="']+$/g, "").replace(/\.0+$/, "").trim() : "", w: wage
      };
      if (maxUnitWage > 0) obj.uMax = maxUnitWage;
      if (shiftType) obj.tp = shiftType;

      const kana = iKana !== -1 ? String(row[iKana] || "").replace(/\s+/g, "").trim() : "";
      if (kana) obj.kn = kana;
      if (allowances) obj.al = allowances;
      if (compactRm) obj.rm = compactRm;
      if (rawSc) obj.sc = rawSc.substring(0, 150);

      const applyDateStr = iApplyDate !== -1 ? normalizeYmd_(row[iApplyDate], null) : "";
      const createDateStr = iCreateDate !== -1 ? normalizeYmd_(row[iCreateDate], null) : "";
      if (applyDateStr) obj.ad = applyDateStr;
      if (createDateStr) obj.cd = createDateStr;

      list.push(obj);
    }
  }
  return list;
}

function buildWorkIntervalStr_(row, colIndices, wageSlotCols, iWage, capSlotCols) {
  if (typeof extractWorkIntervals === "function") {
    const intervals = extractWorkIntervals(row, {
      times: colIndices,
      wages: wageSlotCols || [-1, -1, -1, -1],
      wageSum: iWage !== undefined ? iWage : -1,
      caps: capSlotCols || [-1, -1, -1, -1]
    });
    if (intervals.length > 0) {
      return intervals.map(iv => `${safeFormatTime(iv.start)}-${safeFormatTime(iv.end)}`).join("/");
    }
  }
  const times = colIndices.filter(idx => idx !== -1).map(idx => padHhMm_(row[idx])).filter(val => /^\d{2}:\d{2}$/.test(val));
  if (times.length === 4) return `${times[0]}-${times[1]}/${times[2]}-${times[3]}`;
  if (times.length >= 2) return `${times[0]}-${times[times.length - 1]}`;
  return "";
}

function encryptPayloadForMail_(plainText, secretKey) {
  const blob = Utilities.newBlob(plainText, "application/json", "data.json");
  const gzippedBytes = Utilities.gzip(blob).getBytes();
  const keyBytes = Utilities.newBlob(secretKey).getBytes();
  const outBytes = new Array(gzippedBytes.length);
  for (let i = 0; i < gzippedBytes.length; i++) {
    outBytes[i] = gzippedBytes[i] ^ keyBytes[i % keyBytes.length];
  }
  return "ENC_V1:" + Utilities.base64Encode(outBytes);
}

function normalizeYmd_(val, fallbackYear) {
  if (!val) return "";
  const s = String(val).normalize("NFKC").replace(/^[="']+|[="']+$/g, "").split(/[（(\s]/)[0].trim().replace(/[年月.\-]/g, "/").replace(/日.*$/, "");
  // YYYY/MM/DD
  const mFull = s.match(/^(20\d{2})\/(\d{1,2})\/(\d{1,2})$/);
  if (mFull) return `${mFull[1]}/${("0" + mFull[2]).slice(-2)}/${("0" + mFull[3]).slice(-2)}`;
  // MM/DD/YYYY (US形式)
  const mUs = s.match(/^(\d{1,2})\/(\d{1,2})\/(20\d{2})$/);
  if (mUs) return `${mUs[3]}/${("0" + mUs[1]).slice(-2)}/${("0" + mUs[2]).slice(-2)}`;
  // YY/MM/DD
  const m2Yr = s.match(/^(2\d)\/(\d{1,2})\/(\d{1,2})$/);
  if (m2Yr) return `20${m2Yr[1]}/${("0" + m2Yr[2]).slice(-2)}/${("0" + m2Yr[3]).slice(-2)}`;
  // MM/DD
  const mShort = s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (mShort && fallbackYear) return `${fallbackYear}/${("0" + mShort[1]).slice(-2)}/${("0" + mShort[2]).slice(-2)}`;
  return "";
}

function padHhMm_(val) {
  const m = String(val || "").normalize("NFKC").trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? `${("0" + m[1]).slice(-2)}:${m[2]}` : String(val || "").trim();
}