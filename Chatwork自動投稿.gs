function postToChatwork1() {
  const properties = PropertiesService.getScriptProperties();
  const CHATWORK_API_TOKEN = properties.getProperty('CHATWORK_API_TOKEN');
  const CHATWORK_ROOM_ID = '165593914';

  if (!CHATWORK_API_TOKEN) {
    Logger.log('Chatwork APIトークンがスクリプトプロパティに設定されていません。');
    SpreadsheetApp.getUi().alert('Chatwork APIトークンがスクリプトプロパティに設定されていません。プロジェクトの設定を確認してください。');
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const message = sheet.getRange('A6').getValue();

  if (message === '' || message === null || typeof message === 'undefined') {
    Logger.log('A6セルにメッセージが入力されていません。');
    return;
  }

  const url = `https://api.chatwork.com/v2/rooms/${CHATWORK_ROOM_ID}/messages`;
  const options = {
    'method': 'post',
    'headers': {
      'X-ChatWorkToken': CHATWORK_API_TOKEN
    },
    'payload': {
      'body': String(message)
    },
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      Logger.log('メッセージをChatworkに投稿しました: ' + responseBody);
    } else {
      Logger.log(`Chatworkへの投稿に失敗しました。ステータスコード: ${responseCode}, レスポンス: ${responseBody}`);
      SpreadsheetApp.getUi().alert(`Chatworkへの投稿に失敗しました。\nステータスコード: ${responseCode}\nエラー内容: ${responseBody}\nAPIトークンやルームID、メッセージ内容を確認してください。`);
    }
  } catch (e) {
    Logger.log('Chatworkへの投稿中に予期せぬエラーが発生しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('Chatworkへの投稿中に予期せぬエラーが発生しました: ' + e.message);
  }
}