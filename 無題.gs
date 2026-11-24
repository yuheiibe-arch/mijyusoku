/**
 * 🔍【mnys側 読取専用デバッグ】同名関数の上書き競合 ＆ 2026/09抽出テスト（送信なし）
 */
function debugDiagnoseMnysArchiveSend() {
  console.log("=== 🔍 mnys側 補完ロジック診断開始 ===");

  // 1. 別ファイルによる同名関数の上書きがないか、実際に動く関数の先頭コードを出力
  console.log("① 実行中の buildBqShiftPayload_ 定義先頭:\n" + String(buildBqShiftPayload_).slice(0, 260));
  console.log("② 実行中の normalizeYmd_ 定義先頭:\n" + String(normalizeYmd_).slice(0, 220));

  // 2. 2026確定シフトの [2026/09] シート単体で抽出テスト
  const ss2026 = SpreadsheetApp.openById("10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw");
  const sh09 = ss2026.getSheetByName("2026/09");
  if (!sh09) {
    console.log("❌ 2026確定シフト内に [2026/09] シートが見つかりません");
  } else {
    const sampleRows = sh09.getRange(1, 1, 3, Math.min(35, sh09.getLastColumn())).getDisplayValues();
    console.log(`③ [2026/09] 1行目見出し(抜粋): ${sampleRows[0].map((h, i) => `${i + 1}:${h}`).filter(s => !s.endsWith(":")).join(" | ")}`);
    console.log(`   [2026/09] 2行目データ(抜粋): 13列=${sampleRows[1][12]} | 14列=${sampleRows[1][13]} | 15列=${sampleRows[1][14]} | 16列=${sampleRows[1][15]} | 20列=${sampleRows[1][19]}`);
    console.log(`   normalizeYmd_テスト("${sampleRows[1][14]}", 2026) => "${normalizeYmd_(sampleRows[1][14], 2026)}"`);

    const rows09 = extractShiftRows_(sh09, "2026/09/01", "2026/09/30", "kakutei");
    console.log(`④ extractShiftRows_([2026/09]) 抽出結果: ${rows09.length}件`);
  }

  // 3. buildBqShiftPayload_() を実際に組み立てて年月別の件数を確認（メール送信はしない）
  const payload = buildBqShiftPayload_();
  const ymCount = {};
  (payload.kakutei || []).forEach(r => {
    const ym = String(r.d || "").slice(0, 7);
    ymCount[ym] = (ymCount[ym] || 0) + 1;
  });
  console.log(`⑤ buildBqShiftPayload_ 結果: startDate=${payload.startDate} | maxDate=${payload.maxDate} | archiveSummary="${payload.archiveSummary}"`);
  console.log(`   kakutei 年月別内訳: ${JSON.stringify(ymCount)}`);

  console.log("=== ✅ mnys側 診断完了 ===");
}