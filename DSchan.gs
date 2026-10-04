/**
 * =========================================================
 * DSちゃん（専用UI）バックエンド処理
 * =========================================================
 */

function openDSChanUI() {
  const html = HtmlService.createTemplateFromFile('DS_UI')
    .evaluate()
    .setTitle('DSちゃん 統合管理パネル')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getDSData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('アプローチシート');
    if (!sheet) return [];
    
    const lastRow = Math.max(sheet.getLastRow(), 3);
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    if (lastRow <= 2) return [];
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const safeStr = (val) => val === null || val === undefined ? '' : String(val);

    return data.map((row, index) => {
      let rawDate = row[headers.indexOf('日付')] || '';
      let dObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
      let diffDays = 999;
      let dateStr = safeStr(rawDate);
      
      if (dObj && !isNaN(dObj.getTime())) {
        diffDays = Math.ceil((dObj.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
        dateStr = Utilities.formatDate(dObj, "Asia/Tokyo", "MM月dd日") + `（${weekdays[dObj.getDay()]}）`;
      }
      
      const getVal = (colName) => {
        const idx = headers.indexOf(colName);
        return idx >= 0 ? safeStr(row[idx]) : '';
      };

      return {
        rowIndex: index + 3,
        dateStr: dateStr,
        diffDays: diffDays,
        clinic: getVal('拠点名'),
        time: getVal('不在時間'),
        phase: getVal('Phaze') || getVal('Phase'),
        status: getVal('対応状況'),
        assignee: getVal('担当'),
        memo: getVal('作業メモ'),
        baseSalary: getVal('掲載時給'),
        specialSalary: getVal('特別時給')
      };
    });
  } catch (e) {
    Logger.log(e);
    throw new Error("データ読み込みに失敗しました: " + e.message);
  }
}

function saveDSData(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('アプローチシート');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const statusIdx = headers.indexOf('対応状況') + 1;
  const assigneeIdx = headers.indexOf('担当') + 1;
  const memoIdx = headers.indexOf('作業メモ') + 1;

  updates.forEach(update => {
    if (statusIdx > 0) sheet.getRange(update.rowIndex, statusIdx).setValue(update.status);
    if (assigneeIdx > 0) sheet.getRange(update.rowIndex, assigneeIdx).setValue(update.assignee);
    if (memoIdx > 0) sheet.getRange(update.rowIndex, memoIdx).setValue(update.memo);
  });
  return "シートに保存しました";
}

function getSlackThreadCandidates() {
  const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'); 
  const CHANNEL_ID = "C0AJTMLF0UD"; 

  const candidates = [];
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth() + 1;
  const d = today.getDate();
  
  const addCandidate = (year, month, isFirstHalf) => {
    const lastDay = new Date(year, month, 0).getDate();
    if (isFirstHalf) {
      candidates.push(`【${year}年${month}月 前半(1~15日)】超募集MTGログ`);
    } else {
      candidates.push(`【${year}年${month}月 後半(16~${lastDay}日)】超募集MTGログ`);
    }
  };

  if (d <= 15) {
    addCandidate(y, m, true);
    addCandidate(y, m, false);
  } else {
    addCandidate(y, m, false);
  }
  const ny = m === 12 ? y + 1 : y;
  const nm = m === 12 ? 1 : m + 1;
  addCandidate(ny, nm, true);
  addCandidate(ny, nm, false);
  
  let messages = [];
  let errorMsg = null;
  if (SLACK_TOKEN) {
    const url = `https://slack.com/api/conversations.history?channel=${CHANNEL_ID}&limit=50`;
    const options = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + SLACK_TOKEN },
      muteHttpExceptions: true
    };
    try {
      const response = UrlFetchApp.fetch(url, options);
      const result = JSON.parse(response.getContentText());
      if (result.ok && result.messages) {
        messages = result.messages;
      } else {
        errorMsg = result.error;
      }
    } catch (e) {
      errorMsg = e.message;
    }
  }

  const optionsList = [];
  const normalize = (str) => str.replace(/```/g, '').replace(/[  \n\r]/g, '').replace(/[～~ー-]/g, '~');

  candidates.forEach(title => {
    const normTitle = normalize(title);
    const match = messages.find(m => m.text && normalize(m.text).includes(normTitle));
    if (match) {
      optionsList.push({ label: title, ts: match.ts, isNew: false });
    } else {
      optionsList.push({ label: title, ts: null, isNew: true });
    }
  });

  messages.forEach(m => {
    if (!m.text) return;
    if (m.thread_ts && m.thread_ts !== m.ts) return; 
    if (optionsList.some(o => o.ts === m.ts)) return;

    const rawText = m.text.replace(/```/g, '').trim();
    const firstLine = rawText.split('\n')[0].substring(0, 50);
    
    if (firstLine.includes('募集') || firstLine.includes('MTG') || firstLine.includes('ログ') || firstLine.includes('【') || firstLine.includes('《')) {
      optionsList.push({ label: firstLine, ts: m.ts, isNew: false });
    }
  });

  return { options: optionsList, error: errorMsg };
}

// ★追加：既存スレッドの「過去の返信」をすべて取得する機能
function getSlackThreadReplies(threadTs) {
  const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'); 
  const CHANNEL_ID = "C0AJTMLF0UD"; 
  if (!SLACK_TOKEN || !threadTs) return [];

  const url = `https://slack.com/api/conversations.replies?channel=${CHANNEL_ID}&ts=${threadTs}`;
  const options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + SLACK_TOKEN },
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (result.ok && result.messages) {
      // 親メッセージ（タイトル）を除外して、スレッド内の返信だけを返す
      return result.messages.filter(m => m.ts !== threadTs).map(m => m.text);
    }
  } catch (e) {
    Logger.log(e);
  }
  return [];
}

// ★修正：引用モードの時はUI側で作成したテキストをそのまま（コードブロック化せずに）送信する
function postDSDataToSlack(threadOpt, groupedData) {
  const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'); 
  const CHANNEL_ID = "C0AJTMLF0UD"; 
  
  if (!SLACK_TOKEN) throw new Error("トークンが設定されていません。");

  let parentTs = threadOpt.ts;
  if (!parentTs) {
    const parentText = "```\n" + threadOpt.label + "\n```";
    parentTs = postToSlackAPI_(SLACK_TOKEN, CHANNEL_ID, parentText);
  }
  if (!parentTs) throw new Error("親メッセージの取得・作成に失敗しました。");

  groupedData.forEach(group => {
    // isQuote（引用モード）ならそのまま送信。通常モードなら黒枠のコードブロックで囲む
    let threadText = group.isQuote 
      ? group.text 
      : "`" + group.date + "`\n```\n" + group.text + "\n```";

    postToSlackAPI_(SLACK_TOKEN, CHANNEL_ID, threadText, parentTs);
    Utilities.sleep(1000); 
  });
  
  return "Slackへの投稿が完了しました！";
}

function getAssigneeOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('アプローチシート');
    if (!sheet) return [];
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIdx = headers.indexOf('担当') + 1;
    
    if (colIdx > 0) {
      const rule = sheet.getRange(3, colIdx).getDataValidation();
      if (rule) {
        const criteria = rule.getCriteriaType();
        if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          return rule.getCriteriaValues()[0];
        } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          return rule.getCriteriaValues()[0].getValues().flat().filter(String);
        }
      }
    }
    return [];
  } catch (e) {
    return [];
  }
}

function postToSlackAPI_(token, channelId, text, threadTs = null) {
  const url = 'https://slack.com/api/chat.postMessage';
  const payload = { channel: channelId, text: text };
  if (threadTs) payload.thread_ts = threadTs;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (result.ok) return result.ts;
  } catch (e) {}
  return null;
}