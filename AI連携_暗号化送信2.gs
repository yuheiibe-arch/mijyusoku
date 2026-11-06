/**
 * ============================================================================
 * 【mnys.jp 側（週次・月次報告エージェント対応版）】
 * BQシフト暗号化・添付ファイル送信スクリプト (AI連携_暗号化送信2.gs)
 * - 既存のcaps側受信処理との後方互換性を100%維持しつつ、報告用フィールドを追加
 * - 抽出期間：(当月1日 または 14日前の早い方) 〜 今日から62日後
 * - 「有給」「欠勤」を除外せず区分(tp)を付与して送信
 * ============================================================================
 */

const BQ_SYNC_CONFIG = {
  BQ_SHIFT_SS_ID: "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co",
  TO_ADDRESS: "doctor-support@caps365.jp",
  SUBJECT_PREFIX: "【AI連携暗号化JSON】BQシフト状況パック",
  ATTACHMENT_FILENAME: "bq_shift_pack.enc",
  DEFAULT_SECRET_KEY: "CAPS_AI_BRIDGE_SECRET_2026",
  TARGET_DAYS: 62, // 今日から約2ヶ月分（62日間）
  LOOKBACK_DAYS: 14 // 月初1日と14日前のどちらか早い方まで遡って抽出
};

/**
 * ⏰ 本番・テスト送信関数（暗号化データを添付ファイルとして送信）
 */
function sendEncryptedBqShiftsToCaps() {
  const props = PropertiesService.getScriptProperties();
  const secretKey = props.getProperty("BQ_SYNC_SECRET_KEY") || BQ_SYNC_CONFIG.DEFAULT_SECRET_KEY;

  const payload = buildBqShiftPayload_();
  const jsonString = JSON.stringify(payload);
  const encryptedText = encryptPayloadForMail_(jsonString, secretKey);

  // メール本文のサイズ上限（約200KB）を回避するため、添付ファイルBlobにする
  const attachmentBlob = Utilities.newBlob(
    encryptedText,
    "application/octet-stream",
    BQ_SYNC_CONFIG.ATTACHMENT_FILENAME
  );

  const subject = `${BQ_SYNC_CONFIG.SUBJECT_PREFIX} (${payload.generatedAt})`;
  const bodySummary = [
    `BQシフト状況パック（自動送信）`,
    `抽出期間: ${payload.startDate} 〜 ${payload.maxDate} (基準日: ${payload.baseDate})`,
    `確定シフト: ${payload.kakutei.length}件`,
    `応募シフト: ${payload.obo.length}件`,
    `募集シフト: ${payload.bosyu.length}件`,
    `暗号化サイズ: ${encryptedText.length.toLocaleString()} bytes`,
    `※データは添付ファイル（${BQ_SYNC_CONFIG.ATTACHMENT_FILENAME}）に格納されています。`
  ].join("\n");

  GmailApp.sendEmail(BQ_SYNC_CONFIG.TO_ADDRESS, subject, bodySummary, {
    attachments: [attachmentBlob]
  });

  console.log(`✅ 添付送信完了！ 期間:${payload.startDate}〜${payload.maxDate} | 確定=${payload.kakutei.length}件, 応募=${payload.obo.length}件, 募集=${payload.bosyu.length}件 (暗号化サイズ:${encryptedText.length.toLocaleString()}文字)`);
}

/**
 * 🔍 デバッグ確認用関数（メール送信なし）
 */
function debugTestMnysSourceData() {
  console.log("=== [mnys.jp側] BQパック抽出＆軽量化デバッグ開始 ===");
  const props = PropertiesService.getScriptProperties();
  const secretKey = props.getProperty("BQ_SYNC_SECRET_KEY") || BQ_SYNC_CONFIG.DEFAULT_SECRET_KEY;

  const payload = buildBqShiftPayload_();
  const jsonStr = JSON.stringify(payload);
  const encrypted = encryptPayloadForMail_(jsonStr, secretKey);

  const yukyuCount = payload.kakutei.filter(r => r.tp === "有給").length;
  const kekkinCount = payload.kakutei.filter(r => r.tp === "欠勤").length;
  const shoteiCount = payload.kakutei.filter(r => r.tp === "所定休出").length;
  const teikiCount = payload.kakutei.filter(r => r.tp === "定期").length;

  console.log(`📅 抽出対象期間: ${payload.startDate} 〜 ${payload.maxDate} (今日:${payload.baseDate})`);
  console.log(`✅ 確定シフト抽出件数: ${payload.kakutei.length} 件 (内訳: 有給=${yukyuCount}, 欠勤=${kekkinCount}, 所定休出=${shoteiCount}, 定期=${teikiCount})`);
  console.log(`✅ 応募シフト抽出件数: ${payload.obo.length} 件`);
  console.log(`✅ 募集シフト抽出件数: ${payload.bosyu.length} 件`);
  console.log(`📦 元JSON文字数: ${jsonStr.length.toLocaleString()} 文字`);
  console.log(`🔒 圧縮＆暗号化後の文字数: ${encrypted.length.toLocaleString()} 文字`);
  if (payload.kakutei.length > 0) {
    console.log(`🔍 確定シフトサンプル: ${JSON.stringify(payload.kakutei.slice(0, 3))}`);
  }
  console.log("=== ✅ デバッグ完了！続けて sendEncryptedBqShiftsToCaps を実行してください ===");
}

/**
 * シフトデータを抽出してオブジェクトを作成する内部関数
 */
function buildBqShiftPayload_() {
  const ss = SpreadsheetApp.openById(BQ_SYNC_CONFIG.BQ_SHIFT_SS_ID);
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");

  // 当月1日 と 14日前 のどちらか早い方を開始日とする（週次・月初の過去分集計用）
  const monthStartObj = new Date(now.getFullYear(), now.getMonth(), 1);
  const lookbackObj = new Date(now.getFullYear(), now.getMonth(), now.getDate() - BQ_SYNC_CONFIG.LOOKBACK_DAYS);
  const startDateObj = monthStartObj < lookbackObj ? monthStartObj : lookbackObj;
  const startDateStr = Utilities.formatDate(startDateObj, "Asia/Tokyo", "yyyy/MM/dd");

  const maxDateObj = new Date(now.getFullYear(), now.getMonth(), now.getDate() + BQ_SYNC_CONFIG.TARGET_DAYS);
  const maxDateStr = Utilities.formatDate(maxDateObj, "Asia/Tokyo", "yyyy/MM/dd");
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  return {
    schemaVersion: 2,
    generatedAt: nowStr,
    baseDate: todayStr,    // 既存互換用（今日）
    startDate: startDateStr, // 実際の抽出開始日
    maxDate: maxDateStr,
    kakutei: extractShiftRows_(ss.getSheetByName("確定シフト"), startDateStr, maxDateStr, "kakutei"),
    obo:     extractShiftRows_(ss.getSheetByName("応募シフト"), todayStr, maxDateStr, "obo"),
    bosyu:   extractShiftRows_(ss.getSheetByName("募集シフト"), todayStr, maxDateStr, "bosyu")
  };
}

/**
 * シートから指定期間のシフトを抽出する
 */
function extractShiftRows_(sheet, startDateStr, maxDateStr, sheetType) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const isBosyuSheet = (sheetType === "bosyu");
  const isKakuteiSheet = (sheetType === "kakutei");

  const vals = sheet.getDataRange().getDisplayValues();
  const headers = vals[0].map(h => String(h || "").trim());

  const idx = name => headers.indexOf(name);
  const iDate = idx("勤務日") !== -1 ? idx("勤務日") : idx("日付");
  const iClinicNo = idx("クリニックNo") !== -1 ? idx("クリニックNo") : idx("拠点No");
  const iClinic = idx("クリニック名") !== -1 ? idx("クリニック名") : idx("拠点名");
  const iDept = idx("診療科");
  const iStart = idx("勤務開始時間") !== -1 ? idx("勤務開始時間") : idx("開始時間");
  const iEnd1 = idx("終了時間1");
  const iEnd2 = idx("終了時間2");
  const iEnd3 = idx("終了時間3");
  const iEnd = idx("勤務終了時間") !== -1 ? idx("勤務終了時間") : idx("終了時間");
  const iName = idx("名前") !== -1 ? idx("名前") : idx("医師名");
  const iId = idx("医籍番号");
  const iKana = idx("名前（カタカナ）");
  const iWage = idx("時給合計") !== -1 ? idx("時給合計") : idx("時給");
  const iRemarks = idx("業務内容（備考）");
  const iApplyDate = idx("応募日") !== -1 ? idx("応募日") : idx("医師シフト応募日時");
  const iCreateDate = idx("シフト作成日");
  const iStatus = idx("掲載ステータス");
  const iSecondFlag = idx("二診目表示");
  const iSpecialSetting = idx("特別時給設定");

  // 時給1〜4（時間単価の最大値取得用）
  const unitWageCols = ["時給1", "時給2", "時給3", "時給4"].map(n => idx(n)).filter(c => c !== -1);

  // 追加支給項目（手当）
  const allowanceCols = [];
  // スタッフコメント1〜5
  const staffCommentCols = [];
  headers.forEach((h, c) => {
    if (h.includes("追加支給項目") || h.includes("手当")) allowanceCols.push(c);
    if (h.includes("スタッフコメント")) staffCommentCols.push(c);
  });

  if (iDate === -1 || iClinic === -1) return [];

  const list = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const dateStr = normalizeYmd_(row[iDate]);
    if (!dateStr || dateStr < startDateStr || dateStr > maxDateStr) continue;

    const clinic = String(row[iClinic] || "").trim();
    if (!clinic) continue;

    // 確定シフト以外は従来通り「有給・欠勤」等も除外。確定シフトは「有給・欠勤」を残す（要望1）
    if (isKakuteiSheet) {
      if (/休暇|休診|休館|研修|会議|面談|事務/.test(clinic)) continue;
    } else {
      if (/有給|有休|休暇|欠勤|休診|休館|研修|会議|面談|事務/.test(clinic)) continue;
    }

    const docName = iName !== -1 ? String(row[iName] || "").replace(/\s+/g, "").trim() : "";
    if (!isBosyuSheet && (!docName || docName.includes("休館"))) continue;

    const cNo = iClinicNo !== -1 ? String(row[iClinicNo] || "").trim() : "";
    const start = iStart !== -1 ? padHhMm_(row[iStart]) : "";
    const end = iEnd !== -1 ? padHhMm_(row[iEnd]) : "";
    const dept = iDept !== -1 ? String(row[iDept] || "").trim() : "";
    const wage = iWage !== -1 ? String(row[iWage] || "").replace(/,/g, "").trim() : "";
    const allowances = allowanceCols.map(c => String(row[c] || "").trim()).filter(Boolean).join("/");

    // 時間単価の最大値（特別時給判定用）
    let maxUnitWage = 0;
    unitWageCols.forEach(c => {
      const v = parseInt(String(row[c] || "").replace(/,/g, "").trim(), 10);
      if (!isNaN(v) && v > maxUnitWage) maxUnitWage = v;
    });

    // 備考とスタッフコメントの抽出
    const rawRm = iRemarks !== -1 ? String(row[iRemarks] || "").trim() : "";
    const rawSc = staffCommentCols.map(c => String(row[c] || "").trim()).filter(Boolean).join(" / ");

    // 備考の軽量化（判定に必要なキーワード・定期パターンを保持しつつ120文字以内に圧縮）
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
      // 中抜け休憩を考慮した実働区間（例: "09:00-13:00/15:00-18:00"）
      const ivStr = buildWorkIntervalStr_(row, [iStart, iEnd1, iEnd2, iEnd3, iEnd]);

      // 行ごとの区分（有給・欠勤・所定休出・定期）の判定（要望1）
      const combinedNotes = `${rawRm} ${rawSc}`;
      let shiftType = "";
      if (clinic === "有給" || clinic === "有休" || cNo === "1017") {
        shiftType = "有給";
      } else if (clinic === "欠勤" || cNo === "1018") {
        shiftType = "欠勤";
      } else if (/所定休出|所定外休出|所定外勤務/.test(combinedNotes)) {
        shiftType = "所定休出";
      } else if (
        (/【.+?】/.test(rawRm) && /毎週|第\d|隔週|曜日/.test(rawRm)) ||
        /定期|所定/.test(combinedNotes)
      ) {
        shiftType = "定期";
      }

      const obj = {
        d: dateStr,
        cNo: cNo,
        c: clinic,
        dp: dept,
        t: `${start}-${end}`,
        iv: ivStr || `${start}-${end}`,
        n: docName,
        id: iId !== -1 ? String(row[iId] || "").trim() : "",
        w: wage
      };
      if (maxUnitWage > 0) obj.uMax = maxUnitWage;
      if (shiftType) obj.tp = shiftType;

      const kana = iKana !== -1 ? String(row[iKana] || "").replace(/\s+/g, "").trim() : "";
      if (kana) obj.kn = kana;
      if (allowances) obj.al = allowances;
      if (compactRm) obj.rm = compactRm;
      if (rawSc) obj.sc = rawSc.substring(0, 150);

      const applyDateStr = iApplyDate !== -1 ? normalizeYmd_(row[iApplyDate]) : "";
      const createDateStr = iCreateDate !== -1 ? normalizeYmd_(row[iCreateDate]) : "";
      if (applyDateStr) obj.ad = applyDateStr;
      if (createDateStr) obj.cd = createDateStr;

      list.push(obj);
    }
  }
  return list;
}

/**
 * 開始〜終了（中抜け含む5列）から実働時間帯文字列（"09:00-13:00/15:00-21:00"）を組み立てる
 */
function buildWorkIntervalStr_(row, colIndices) {
  const times = colIndices
    .filter(idx => idx !== -1)
    .map(idx => padHhMm_(row[idx]))
    .filter(val => /^\d{2}:\d{2}$/.test(val));

  if (times.length === 4) {
    return `${times[0]}-${times[1]}/${times[2]}-${times[3]}`;
  } else if (times.length >= 2) {
    return `${times[0]}-${times[times.length - 1]}`;
  }
  return "";
}

function encryptPayloadForMail_(plainText, secretKey) {
  const blob = Utilities.newBlob(plainText, "application/json", "data.json");
  const gzippedBytes = Utilities.gzip(blob).getBytes();
  const keyBytes = Utilities.newBlob(secretKey).getBytes();
  const outBytes = [];
  for (let i = 0; i < gzippedBytes.length; i++) {
    outBytes.push(gzippedBytes[i] ^ keyBytes[i % keyBytes.length]);
  }
  return "ENC_V1:" + Utilities.base64Encode(outBytes);
}

function normalizeYmd_(val) {
  const s = String(val || "").trim().replace(/[年月\.\-]/g, "/").replace(/日.*$/, "");
  const m = s.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  return m ? `${m[1]}/${("0" + m[2]).slice(-2)}/${("0" + m[3]).slice(-2)}` : "";
}

function padHhMm_(val) {
  const m = String(val || "").trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? `${("0" + m[1]).slice(-2)}:${m[2]}` : String(val || "").trim();
}