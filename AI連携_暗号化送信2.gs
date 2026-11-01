/**
 * ============================================================================
 * 【mnys.jp 側（改良版）】BQシフト（今日から2ヶ月分）暗号化・添付ファイル送信スクリプト
 * - メール本文サイズ上限(200KB)を回避するため、暗号化データを添付ファイル(最大25MB)で送信
 * - 今日から2ヶ月分のみ対象、「有給・休館・休暇」や長文の備考ノイズを除外して超軽量化
 * ============================================================================
 */

const BQ_SYNC_CONFIG = {
  BQ_SHIFT_SS_ID: "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co",
  TO_ADDRESS: "doctor-support@caps365.jp",
  SUBJECT_PREFIX: "【AI連携暗号化JSON】BQシフト状況パック",
  ATTACHMENT_FILENAME: "bq_shift_pack.enc",
  SECRET_KEY: "CAPS_AI_BRIDGE_SECRET_2026",
  TARGET_DAYS: 62 // 今日から約2ヶ月分（62日間）
};

/**
 * ⏰ 本番・テスト送信関数（暗号化データを添付ファイルとして送信）
 */
function sendEncryptedBqShiftsToCaps() {
  const payload = buildBqShiftPayload_();
  const jsonString = JSON.stringify(payload);
  const encryptedText = encryptPayloadForMail_(jsonString, BQ_SYNC_CONFIG.SECRET_KEY);

  // ★メール本文のサイズ上限（約200KB）を回避するため、添付ファイルBlobにする
  const attachmentBlob = Utilities.newBlob(
    encryptedText,
    "application/octet-stream",
    BQ_SYNC_CONFIG.ATTACHMENT_FILENAME
  );

  const subject = `${BQ_SYNC_CONFIG.SUBJECT_PREFIX} (${payload.generatedAt})`;
  const bodySummary = [
    `BQシフト状況パック（自動送信）`,
    `抽出期間: ${payload.baseDate} 〜 ${payload.maxDate}`,
    `確定シフト: ${payload.kakutei.length}件`,
    `応募シフト: ${payload.obo.length}件`,
    `募集シフト: ${payload.bosyu.length}件`,
    `暗号化サイズ: ${encryptedText.length.toLocaleString()} bytes`,
    `※データは添付ファイル（${BQ_SYNC_CONFIG.ATTACHMENT_FILENAME}）に格納されています。`
  ].join("\n");

  GmailApp.sendEmail(BQ_SYNC_CONFIG.TO_ADDRESS, subject, bodySummary, {
    attachments: [attachmentBlob]
  });

  console.log(`✅ 添付送信完了！ 期間:${payload.baseDate}〜${payload.maxDate} | 確定=${payload.kakutei.length}件, 応募=${payload.obo.length}件, 募集=${payload.bosyu.length}件 (暗号化サイズ:${encryptedText.length.toLocaleString()}文字)`);
}

/**
 * 🔍 デバッグ確認用関数（メール送信なし）
 */
function debugTestMnysSourceData() {
  console.log("=== [mnys.jp側] 2ヶ月分抽出＆軽量化デバッグ開始 ===");
  const payload = buildBqShiftPayload_();
  const jsonStr = JSON.stringify(payload);
  const encrypted = encryptPayloadForMail_(jsonStr, BQ_SYNC_CONFIG.SECRET_KEY);

  console.log(`📅 抽出対象期間: ${payload.baseDate} 〜 ${payload.maxDate}`);
  console.log(`✅ 確定シフト抽出件数: ${payload.kakutei.length} 件`);
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
 * 2ヶ月分のシフトデータを抽出してオブジェクトを作成する内部関数
 */
function buildBqShiftPayload_() {
  const ss = SpreadsheetApp.openById(BQ_SYNC_CONFIG.BQ_SHIFT_SS_ID);
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
  const maxDateObj = new Date(now.getFullYear(), now.getMonth(), now.getDate() + BQ_SYNC_CONFIG.TARGET_DAYS);
  const maxDateStr = Utilities.formatDate(maxDateObj, "Asia/Tokyo", "yyyy/MM/dd");
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  return {
    generatedAt: nowStr,
    baseDate: todayStr,
    maxDate: maxDateStr,
    kakutei: extractShiftRows_(ss.getSheetByName("確定シフト"), todayStr, maxDateStr, false),
    obo:     extractShiftRows_(ss.getSheetByName("応募シフト"), todayStr, maxDateStr, false),
    bosyu:   extractShiftRows_(ss.getSheetByName("募集シフト"), todayStr, maxDateStr, true)
  };
}

/**
 * シートから指定期間（todayStr 〜 maxDateStr）の有効シフトのみを軽量抽出する
 */
function extractShiftRows_(sheet, todayStr, maxDateStr, isBosyuSheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const vals = sheet.getDataRange().getDisplayValues();
  const headers = vals[0].map(h => String(h || "").trim());

  const idx = name => headers.indexOf(name);
  const iDate = idx("勤務日") !== -1 ? idx("勤務日") : idx("日付");
  const iClinic = idx("クリニック名") !== -1 ? idx("クリニック名") : idx("拠点名");
  const iDept = idx("診療科");
  const iStart = idx("勤務開始時間") !== -1 ? idx("勤務開始時間") : idx("開始時間");
  const iEnd = idx("勤務終了時間") !== -1 ? idx("勤務終了時間") : idx("終了時間");
  const iName = idx("名前") !== -1 ? idx("名前") : idx("医師名");
  const iId = idx("医籍番号");
  const iKana = idx("名前（カタカナ）");
  const iWage = idx("時給合計") !== -1 ? idx("時給合計") : idx("時給");
  const iRemarks = idx("業務内容（備考）");

  const allowanceCols = [];
  headers.forEach((h, c) => {
    if (h.includes("追加支給項目") || h.includes("手当")) allowanceCols.push(c);
  });

  if (iDate === -1 || iClinic === -1) return [];

  const list = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const dateStr = normalizeYmd_(row[iDate]);
    // ★今日〜2ヶ月後までのデータに限定
    if (!dateStr || dateStr < todayStr || dateStr > maxDateStr) continue;

    const clinic = String(row[iClinic] || "").trim();
    // ★「有給」「休暇」「休館」「欠勤」「研修」「会議」などの非診療枠を除外
    if (!clinic || /有給|有休|休暇|欠勤|休診|休館|研修|会議|面談|事務/.test(clinic)) continue;

    const docName = iName !== -1 ? String(row[iName] || "").replace(/\s+/g, "").trim() : "";
    if (!isBosyuSheet && (!docName || docName.includes("休館"))) continue;

    const start = iStart !== -1 ? padHhMm_(row[iStart]) : "";
    const end = iEnd !== -1 ? padHhMm_(row[iEnd]) : "";
    const dept = iDept !== -1 ? String(row[iDept] || "").trim() : "";
    const wage = iWage !== -1 ? String(row[iWage] || "").replace(/,/g, "").trim() : "";
    const allowances = allowanceCols.map(c => String(row[c] || "").trim()).filter(Boolean).join("/");

    // ★備考(rm)は長文の契約メモや定型文を捨て、「特別時給」や「バックアップ」等の重要フラグだけ残す
    const rawRm = iRemarks !== -1 ? String(row[iRemarks] || "") : "";
    let shortRm = "";
    if (rawRm.includes("特別時給")) shortRm = "特別時給";
    else if (rawRm.includes("3,000円") || rawRm.includes("3000円")) shortRm = "加算3000";
    else if (rawRm.includes("2,000円") || rawRm.includes("2000円")) shortRm = "加算2000";

    if (isBosyuSheet) {
      const obj = { d: dateStr, c: clinic, dp: dept, t: `${start}-${end}`, w: wage };
      if (shortRm) obj.rm = shortRm;
      list.push(obj);
    } else {
      const obj = {
        d: dateStr,
        c: clinic,
        dp: dept,
        t: `${start}-${end}`,
        n: docName,
        id: iId !== -1 ? String(row[iId] || "").trim() : "",
        w: wage
      };
      const kana = iKana !== -1 ? String(row[iKana] || "").replace(/\s+/g, "").trim() : "";
      if (kana) obj.kn = kana;
      if (allowances) obj.al = allowances;
      if (shortRm) obj.rm = shortRm;
      list.push(obj);
    }
  }
  return list;
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