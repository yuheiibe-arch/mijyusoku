/**
 * 週間および月間のサマリーテキストを組み立てる専用関数
 */
function buildWeeklyMonthlySummaryText(params) {
  const { startDate, endDate, isExtSsLoaded, hasCooData, weekly, monthly, areaMap, TARGET_SPLIT_CLINICS } = params;

  // --- 週間サマリーの計算 ---
  const wRate1st = weekly.req1st > 0 ? Math.floor((weekly.filled1st / weekly.req1st) * 100) : 100;
  const wReq1stH = Math.round(weekly.req1st / 60);
  const wFilled1stH = Math.round(weekly.filled1st / 60);
  
  const wRate2nd = weekly.req2nd > 0 ? Math.floor((weekly.filled2nd / weekly.req2nd) * 100) : 100;
  const wReq2ndH = Math.round(weekly.req2nd / 60);
  const wFilled2ndH = Math.round(weekly.filled2nd / 60);
  
  const wCooRate = weekly.cooReq > 0 ? Math.floor((weekly.cooFilled / weekly.cooReq) * 100) : 0;
  const wCooReqH = Math.round(weekly.cooReq / 60);
  const wCooFilledH = Math.round(weekly.cooFilled / 60);
  
  // ★ f判定時間と、純粋な応募時間を算出
  const wCooFH = Math.round(weekly.cooFMin / 60);
  const wCooOuboH = wCooFilledH - wCooFH;
  
  const wGapHours = Math.round(weekly.gapMin / 60);

  // エリアごとの集計ロジック
  const areaCount = { "東京": [], "神奈川": [], "埼玉": [], "千葉": [], "茨城": [], "大阪": [] };

  weekly.absenceClinics.forEach(record => {
    let displayName = record.clinic;
    if (TARGET_SPLIT_CLINICS.includes(record.clinic) && record.dept) {
      displayName = `${record.clinic}（${record.dept}）`;
    }
    const area = areaMap[record.clinic] || "その他";
    if (!areaCount[area]) areaCount[area] = [];
    areaCount[area].push(displayName);
  });

  let summaryText = `[info][title]週間医師充足数[/title]\n`;
  summaryText += `計測期間：${fastFormatDate(startDate)}~${fastFormatDate(endDate)}\n\n`;
  summaryText += `１診目充足率：${wRate1st}%（充足済み：${wFilled1stH}h/募集：${wReq1stH}h）\n`;
  
  if (isExtSsLoaded) {
    summaryText += `２診目充足率（全体）：${wRate2nd}%（充足済み：${wFilled2ndH}h/募集：${wReq2ndH}h）\n`;
    if (hasCooData) {
      summaryText += `└COO室依頼２診：${wCooRate}%（充足済み：${wCooFilledH}h/募集：${wCooReqH}h）\n`;
      summaryText += `（f判定含む依頼総数）\n`;
      summaryText += `f判定：${wCooFH}h\n`;
      summaryText += `応募時間数：${wCooOuboH}h\n`;
    }
  } else {
    summaryText += `２診目充足率：取得エラー\n`;
  }
  
  summaryText += `医師不在時間合計：${wGapHours}h\n`;
  summaryText += `医師不在拠点箇所（延べ数）：${weekly.absenceClinics.length}\n\n`;

  const kantouAreas = ["東京", "神奈川", "埼玉", "千葉", "茨城"];
  kantouAreas.forEach(area => {
    const clinics = areaCount[area] || [];
    if (clinics.length > 0) {
      const uniqueNames = [...new Set(clinics)];
      summaryText += `${area}：${uniqueNames.length}拠点（${uniqueNames.join('、')}）\n`;
    }
  });

  const osakaClinics = areaCount["大阪"] || [];
  if (osakaClinics.length > 0) {
    if (kantouAreas.some(a => (areaCount[a] || []).length > 0)) {
      summaryText += `[hr]\n`;
    }
    const uniqueNames = [...new Set(osakaClinics)];
    summaryText += `大阪：${uniqueNames.length}拠点（${uniqueNames.join('、')}）\n`;
  }

  // --- 月間サマリーの計算 ---
  const mRate1st = monthly.req1st > 0 ? Math.floor((monthly.filled1st / monthly.req1st) * 100) : 100;
  const mReq1stH = Math.round(monthly.req1st / 60);
  const mFilled1stH = Math.round(monthly.filled1st / 60);
  
  const mRate2nd = monthly.req2nd > 0 ? Math.floor((monthly.filled2nd / monthly.req2nd) * 100) : 100;
  const mReq2ndH = Math.round(monthly.req2nd / 60);
  const mFilled2ndH = Math.round(monthly.filled2nd / 60);
  
  const mCooRate = monthly.cooReq > 0 ? Math.floor((monthly.cooFilled / monthly.cooReq) * 100) : 0;
  const mCooReqH = Math.round(monthly.cooReq / 60);
  const mCooFilledH = Math.round(monthly.cooFilled / 60);
  
  // ★ f判定時間と、純粋な応募時間を算出
  const mCooFH = Math.round(monthly.cooFMin / 60);
  const mCooOuboH = mCooFilledH - mCooFH;
  
  const mGapHours = Math.round(monthly.gapMin / 60);

  const monthStartStr = fastFormatDate(new Date(startDate.getFullYear(), startDate.getMonth(), 1));
  const monthEndStr = fastFormatDate(new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0));

  summaryText += `[hr]\n`;
  summaryText += `月間集計数\n`;
  summaryText += `計測期間：${monthStartStr}~${monthEndStr}\n\n`;
  summaryText += `１診目充足率：${mRate1st}%（充足済み：${mFilled1stH}h/募集：${mReq1stH}h）\n`;
  summaryText += `２診目充足率（全体）：${mRate2nd}%（充足済み：${mFilled2ndH}h/募集：${mReq2ndH}h）\n`;
  
  if (isExtSsLoaded && hasCooData) {
    summaryText += `└COO室依頼２診：${mCooRate}%（充足済み：${mCooFilledH}h/募集：${mCooReqH}h）\n`;
    summaryText += `（f判定含む依頼総数）\n`;
    summaryText += `f判定：${mCooFH}h\n`;
    summaryText += `応募時間数：${mCooOuboH}h\n`;
  }
  
  summaryText += `医師不在時間合計：${mGapHours}h\n`;
  summaryText += `医師不在拠点箇所（延べ数）：${monthly.absenceCount}\n`;

  summaryText += `[/info]\n\n`;

  return summaryText;
}