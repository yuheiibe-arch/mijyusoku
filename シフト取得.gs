const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyIddjdfR9CwbI6X5wyZGru5kITWoB8myC2XZtRsqFfTVG1L_2GkF_-aFjFrAIXPtU/exec';
const FIXED_SHIFT_SHEET_ID = "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co";

function callGetRecentShiftFromActual() {
  DriveApp.getFiles(); 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName('テスト'); // ※本番時は '貼付用'
  
  if (!targetSheet) {
    Logger.log("エラー: 'テスト' シートが見つかりません。");
    return;
  }

  try {
    ss.toast('1/4: 外部データの更新を要求しています...', '処理開始', 10);
    const payload = { sheetId: FIXED_SHIFT_SHEET_ID };
    const token = ScriptApp.getOAuthToken();
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    };

    const resp = UrlFetchApp.fetch(WEB_APP_URL, options);
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    
    // --- 【修正：無限ループ防止ロジック】 ---
    
    // 判定A：明らかにGoogleのログイン/許可画面(HTML)である時だけポップアップを出す
    const isLoginPage = body.includes('<html') || body.includes('google-signin') || body.includes('ログイン') || body.includes('Sign in');
    
    if (isLoginPage || code === 401 || code === 403) {
      Logger.log("認証画面または権限エラーを検知しました。ポップアップを表示します。");
      showAuthUrl();
      return; 
    }

    // 判定B：JSONパースを試みるが、失敗してもHTMLでなければ「認証は成功済み」とみなして強引に進む
    let result = { success: false };
    try { 
      result = JSON.parse(body); 
    } catch (e) { 
      Logger.log("JSONパースに失敗しましたが、認証画面ではないため続行します。応答内容: " + body.substring(0, 50));
      result = { success: true }; // 強制的に成功扱いで続行
    }

    // 通信自体がエラー（500系など）の場合はここで止める
    if (code >= 400 && code !== 401 && code !== 403) {
      throw new Error(`サーバーエラー (コード: ${code}) 内容: ${body.substring(0,50)}`);
    }
    // --- 【修正ここまで】 ---

    ss.toast('2/4: 更新要求完了。データを取得しています...', '進行中', 10);
    Utilities.sleep(5000); // 反映待ち

    const remoteSs = SpreadsheetApp.openById(FIXED_SHIFT_SHEET_ID);
    const remoteSheet = remoteSs.getSheetByName('確定シフト');
    if (!remoteSheet) throw new Error("'確定シフト' シートが見つかりません。");

    const remoteData = remoteSheet.getDataRange().getValues();
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth(); 
    let nextYear = currentYear;
    let nextMonth = currentMonth + 1;
    if (nextMonth > 11) { nextMonth = 0; nextYear++; }

    let filteredData = [];

    // 当月・翌月のデータを抽出し、A〜AM列(39列)を切り出す
    for (let i = 1; i < remoteData.length; i++) {
      const row = remoteData[i];
      const dateVal = row[5]; 
      if (!dateVal) continue;
      
      let d = new Date(dateVal);
      if (isNaN(d.getTime()) && typeof dateVal === 'string') {
        const cleanStr = dateVal.split('(')[0].split('（')[0].trim();
        d = new Date(cleanStr);
      }

      if (!isNaN(d.getTime())) {
        const rowYear = d.getFullYear();
        const rowMonth = d.getMonth();

        if ((rowYear === currentYear && rowMonth === currentMonth) || 
            (rowYear === nextYear && rowMonth === nextMonth)) {
          let rowData = row.slice(0, 39);
          while (rowData.length < 39) rowData.push(""); 
          filteredData.push(rowData);
        }
      }
    }

    if (filteredData.length === 0) throw new Error("当月・翌月のデータが0件でした。");

    ss.toast(`3/4: ${filteredData.length}件を転記し、関数をコピーしています...`, '進行中', 10);
    
    const lastRow = targetSheet.getLastRow();
    if (lastRow >= 3) targetSheet.getRange(3, 1, lastRow - 2, 39).clearContent();
    targetSheet.getRange(3, 1, filteredData.length, 39).setValues(filteredData);

    // 1899年バグ防止
    targetSheet.getRange(3, 6, filteredData.length, 1).setNumberFormat('yyyy/MM/dd');
    targetSheet.getRange(3, 7, filteredData.length, 10).setNumberFormat('HH:mm');

    // 保護列（AN, AO, AP）の関数を一番下までコピーダウンする
    if (filteredData.length > 0) {
      // 40列目=AN, 41列目=AO, 42列目=AP
      const formulaRange = targetSheet.getRange(3, 40, 1, 3); // 3行目の関数をコピー元とする
      const destRange = targetSheet.getRange(3, 40, filteredData.length, 3);
      formulaRange.copyTo(destRange);
    }

    SpreadsheetApp.flush();

    ss.toast('4/4: テスト集計を実行します...', '最終段階', 10);
    
    test_updateSheetRowAdjusted();

  } catch (e) {
    if (e.message.includes('権限') || e.message.includes('アクセス')) {
      showAuthUrl();
    } else {
      Logger.log("エラー: " + e.message);
      SpreadsheetApp.getUi().alert("エラーが発生しました: \n" + e.message);
    }
  }
}

function showAuthUrl() {
  const htmlTemplate = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h3 style="color: #d93025; margin-top: 0;">⚠️ 初回アクセス認証が必要です</h3>
      <p>外部システムへアクセスする権限がありません。<br>以下のボタンをクリックして、Googleの認証画面を進めてください。</p>
      <div style="text-align: center; margin: 25px 0;">
        <a href="${WEB_APP_URL}" target="_blank" style="background-color: #1a73e8; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">
          認証ページを開く
        </a>
      </div>
      <p style="font-size: 12px; color: #666; line-height: 1.5;">
        ※クリック後、アカウントを選択し「詳細」＞「〇〇に移動」＞「許可」を押してください。<br>
        ※画面に「{"success":false...」のような文字が出たら認証完了です。タブを閉じて、もう一度実行してください。
      </p>
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate).setWidth(450).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '【必須】初回認証');
}