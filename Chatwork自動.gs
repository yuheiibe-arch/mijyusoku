function postToChatwork1() {
  // 1. スクリプトプロパティからChatwork APIトークンを取得し、ルームIDを設定
  const properties = PropertiesService.getScriptProperties();
  const CHATWORK_API_TOKEN = properties.getProperty('CHATWORK_API_TOKEN'); // スクリプトプロパティで設定したキー名に合わせてください
  const CHATWORK_ROOM_ID = '165593914'; // 提供されたルームID

  // APIトークンが取得できなかった場合のチェック
  if (!CHATWORK_API_TOKEN) {
    Logger.log('Chatwork APIトークンがスクリプトプロパティに設定されていません。');
    SpreadsheetApp.getUi().alert('Chatwork APIトークンがスクリプトプロパティに設定されていません。プロジェクトの設定を確認してください。');
    return;
  }

  // 2. スプレッドシートからA6セルの値を取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet(); // アクティブなシートを取得
  // もし特定のシート名を指定したい場合は、以下のように変更してください
  // const sheet = spreadsheet.getSheetByName('文章自動作成');
  const message = sheet.getRange('A6').getValue();

  // 3. メッセージが空でないことを確認
  if (message === '' || message === null || typeof message === 'undefined') {
    Logger.log('A6セルにメッセージが入力されていません。');
    // 必要に応じてユーザーに通知
    // SpreadsheetApp.getUi().alert('A6セルにメッセージが入力されていません。');
    return;
  }

  // 4. Chatwork APIに送信するデータを作成
  const url = `https://api.chatwork.com/v2/rooms/${CHATWORK_ROOM_ID}/messages`;
  const options = {
    'method': 'post',
    'headers': {
      'X-ChatWorkToken': CHATWORK_API_TOKEN
    },
    'payload': {
      'body': String(message) // メッセージを文字列に変換
    },
    'muteHttpExceptions': true // エラーレスポンスも取得するため
  };

  // 5. Chatwork APIにリクエストを送信
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      Logger.log('メッセージをChatworkに投稿しました: ' + responseBody);
      // 必要に応じて成功をユーザーに通知
      // SpreadsheetApp.getUi().alert('メッセージをChatworkに投稿しました。');
    } else {
      Logger.log(`Chatworkへの投稿に失敗しました。ステータスコード: ${responseCode}, レスポンス: ${responseBody}`);
      SpreadsheetApp.getUi().alert(`Chatworkへの投稿に失敗しました。\nステータスコード: ${responseCode}\nエラー内容: ${responseBody}\nAPIトークンやルームID、メッセージ内容を確認してください。`);
    }
  } catch (e) {
    Logger.log('Chatworkへの投稿中に予期せぬエラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('Chatworkへの投稿中に予期せぬエラーが発生しました: ' + e.message);
  }
}