/**
 * 🔍 正しい列（I列:開院日、K列:クリニックNo、A〜E列:拠点名）でのマッピング検証スクリプト
 */
function debugVerifyCorrectMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MASTER_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';

  const masterSS = SpreadsheetApp.openById(MASTER_SS_ID);
  const masterSheet = masterSS.getSheetByName('拠点名');
  const mValues = masterSheet.getDataRange().getValues();
  const mDisp = masterSheet.getDataRange().getDisplayValues();
  const headers = mDisp[0].map(h => String(h || "").trim());

  // ★列番号を固定値ではなくヘッダー名から動的に特定
  const idxGroup = headers.indexOf("拠点グループ"); // 5 (F列)
  const idxDept = headers.indexOf("診療科");        // 6 (G列)
  const idxOpenDate = headers.indexOf("開院日");    // 8 (I列)
  const idxClinicNo = headers.indexOf("クリニックNo"); // 10 (K列)

  console.log(`📌 特定した列インデックス -> 拠点グループ:${idxGroup}, 診療科:${idxDept}, 開院日:${idxOpenDate}(I列), クリニックNo:${idxClinicNo}(K列)`);

  const openDateByNo = new Map();
  const openDateByName = new Map();
  const clinicNoToMasterName = new Map();

  console.log("\n==================================================");
  console.log("🔍 1. マスタの『クリニックNo(K列)』と『開院日(I列)』一覧（行45以降＋北葛西・亀有）");
  console.log("==================================================");

  for (let i = 1; i < mValues.length; i++) {
    const rawOpen = mValues[i][idxOpenDate];
    const dispOpen = mDisp[i][idxOpenDate];
    const parsedOpen = parseDateToSafeDateObj(rawOpen);
    const cNo = String(mDisp[i][idxClinicNo] || "").trim();
    const dept = String(mDisp[i][idxDept] || "").trim();
    const mainName = String(mDisp[i][0] || "").trim();

    // クリニックNoでマッピング登録
    if (cNo && parsedOpen) {
      openDateByNo.set(cNo, parsedOpen);
      openDateByNo.set(`${cNo}_${dept}`, parsedOpen);
    }
    if (cNo && mainName) {
      clinicNoToMasterName.set(cNo, mainName);
    }

    // 名前（A〜E列の正規記載＋表記揺れ）でもマッピング登録
    for (let c = 0; c <= 4; c++) {
      const alias = String(mDisp[i][c] || "").trim();
      if (alias && parsedOpen) {
        openDateByName.set(alias, parsedOpen);
        openDateByName.set(normalizeClinicName(alias), parsedOpen);
      }
    }

    if (i >= 44 || mainName === "北葛西" || mainName === "亀有") {
      console.log(
        `[行${i + 1}] クリニックNo(K列): "${cNo.padEnd(3, ' ')}" | 正規記載(A列): "${mainName}" (${dept}) | ` +
        `開院日(I列表示): "${dispOpen}" -> パース結果: ${parsedOpen ? fastFormatDate(parsedOpen) : "⚠️未設定/null"}`
      );
    }
  }

  console.log("\n==================================================");
  console.log("🔍 2. 『貼付用』シートの全拠点 ⇔ クリニックNo ⇔ 開院日 突合結果");
  console.log("==================================================");

  const pasteSheet = ss.getSheetByName('貼付用');
  const pasteData = pasteSheet.getDataRange().getDisplayValues();
  const clinicMapInPaste = new Map(); // normName -> { cNos: Set, rawNames: Set, minDate, maxDate }

  for (let i = 2; i < pasteData.length; i++) {
    const r = pasteData[i];
    const cNo = String(r[11] || "").trim();
    const cName = String(r[12] || "").trim();
    const dObj = parseDateToSafeDateObj(r[14]);
    if (!cName || !dObj) continue;

    // 非拠点（1000番台や除外ワード）はスキップ
    if (Number(cNo) >= 1000 || /有給|欠勤|院外勤務|バックアップ|医師会|嘱託医|ワクチン/.test(cName)) continue;

    const normName = normalizeClinicName(cName);
    const dStr = fastFormatDate(dObj);
    if (!clinicMapInPaste.has(normName)) {
      clinicMapInPaste.set(normName, { cNos: new Set(), minDate: dStr, maxDate: dStr, count: 0 });
    }
    const entry = clinicMapInPaste.get(normName);
    if (cNo) entry.cNos.add(cNo);
    if (dStr < entry.minDate) entry.minDate = dStr;
    if (dStr > entry.maxDate) entry.maxDate = dStr;
    entry.count++;
  }

  const testDate = new Date(2026, 10, 2); // 2026/11/02
  let unmappedCount = 0;

  Array.from(clinicMapInPaste.keys()).sort().forEach(normName => {
    const entry = clinicMapInPaste.get(normName);
    const cNoList = Array.from(entry.cNos);
    const cNo = cNoList[0] || "";

    const openByNo = openDateByNo.get(cNo);
    const openByName = openDateByName.get(normName);
    const finalOpenDate = openByNo || openByName || null;

    const isUnopenedOnNov2 = finalOpenDate && finalOpenDate > testDate;

    if (!finalOpenDate) unmappedCount++;

    // 開院日が2026年以降の拠点、または未マッピングの拠点を表示
    if (!finalOpenDate || finalOpenDate >= new Date(2026, 0, 1) || isUnopenedOnNov2) {
      console.log(
        `🏥 【${normName}】 (貼付用No: ${cNo || "なし"}, シフト期間: ${entry.minDate}〜${entry.maxDate})\n` +
        `   ├─ No突合(K列=${cNo}) : ${openByNo ? "✅ " + fastFormatDate(openByNo) : "❌ なし"}\n` +
        `   ├─ 名前突合(A〜E列)  : ${openByName ? "✅ " + fastFormatDate(openByName) : "❌ なし"}\n` +
        `   └─ 2026/11/02判定    : ${isUnopenedOnNov2 ? "🚫 未開院として正しく除外される！（開院日: " + fastFormatDate(finalOpenDate) + "）" : "✅ 開院済み（集計対象）"}`
      );
    }
  });

  console.log(`\n📊 突合サマリー: 貼付用シートの有効拠点数=${clinicMapInPaste.size}拠点 / 開院日が取れなかった拠点=${unmappedCount}拠点`);
}