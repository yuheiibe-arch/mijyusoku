/**
 * ============================================================================
 * Chatwork自動投稿.gs
 * スクリプトプロパティの CHATWORK_API_TOKEN を使用してDS部に投稿する
 * ＋ 投稿成功時に拠点連絡システム（⚙️_連携キュー）へ連動トリガーを書き込む
 * ============================================================================
 */

// 連携先の拠点連絡システム管理スプレッドシートID・シート名
const BRIDGE_QUEUE_SS_ID = '1jmRdwUEUNhC2iTZyqQC3IqdBIJGZiCdR9RR2An_yvw8';
const BRIDGE_QUEUE_SHEET_NAME = '⚙️_連携キュー';

function postToChatwork1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('文章自動作成');

  if (!sheet) {
    ui.alert('エラー', '「文章自動作成」シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  const message = String(sheet.getRange('A6').getValue() || '').trim();
  if (!message) {
    ui.alert('エラー', '「文章自動作成」シートのA6セルが空です。\n先に「本番シフト集計実行」を行ってください。', ui.ButtonSet.OK);
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const apiToken = (props.getProperty('CHATWORK_API_TOKEN') || '').trim();

  if (!apiToken) {
    ui.alert('エラー', 'スクリプトプロパティに「CHATWORK_API_TOKEN」が見つかりません。', ui.ButtonSet.OK);
    return;
  }

  // ルームIDをスクリプトプロパティから探す（複数のキー名に対応）
  let roomId = (
    props.getProperty('CHATWORK_ROOM_ID') ||
    props.getProperty('CHATWORK_ROOM_ID_1') ||
    props.getProperty('CW_ROOM_ID') ||
    props.getProperty('ROOM_ID') ||
    ''
  ).trim();

  // スクリプトプロパティにルームIDが無い場合、ダイアログで1回だけ聞いて自動保存する
  if (!roomId) {
    const promptRes = ui.prompt(
      'DS部 ルームIDの初回設定',
      'Chatworkの投稿先ルームID（URL末尾の rid に続く数字）を入力してください。\n※一度入力するとスクリプトプロパティ（CHATWORK_ROOM_ID）に保存され、次回から入力不要になります。',
      ui.ButtonSet.OK_CANCEL
    );
    if (promptRes.getSelectedButton() !== ui.Button.OK) {
      ss.toast('投稿をキャンセルしました。', 'キャンセル', 3);
      return;
    }
    // 「#!rid12345678」のまま貼られても数字だけ抜き出す
    roomId = promptRes.getResponseText().replace(/[^0-9]/g, '');
    if (!roomId) {
      ui.alert('エラー', '有効なルームID（数字）が入力されませんでした。', ui.ButtonSet.OK);
      return;
    }
    props.setProperty('CHATWORK_ROOM_ID', roomId);
  }

  // 送信確認ダイアログ
  const confirm = ui.alert(
    'Chatwork投稿確認',
    'DS部のChatworkへ未充足報告を投稿します。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    ss.toast('投稿をキャンセルしました。', 'キャンセル', 3);
    return;
  }

  try {
    const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
    const options = {
      method: 'post',
      headers: { 'X-ChatWorkToken': apiToken },
      payload: { body: message, self_unread: '0' },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const resBody = response.getContentText();

    if (code >= 200 && code < 300) {
      ss.toast('Chatwork（DS部）への投稿が完了しました！', '送信完了', 5);

      // ★復活：拠点連絡システムの「⚙️_連携キュー」へ MIJYUSOKU_POST を書き込む
      try {
        enqueueMijyusokuPostToBridge_();
      } catch (qErr) {
        Logger.log('連携キュー書き込みエラー: ' + qErr.message);
      }

      // AI連携の暗号化メール送信も連動させる
      try {
        if (typeof triggerAiBridgeExport === 'function') {
          triggerAiBridgeExport();
        }
      } catch (aiErr) {
        Logger.log('AI連携送信スキップ: ' + aiErr.message);
      }
    } else {
      throw new Error(`HTTP ${code}: ${resBody}`);
    }
  } catch (e) {
    Logger.log('Chatwork送信エラー: ' + e.message);
    ui.alert('送信エラー', 'Chatworkへの投稿に失敗しました。\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 拠点連絡システム（reamindkun）の「⚙️_連携キュー」シートへ未充足投稿イベントを追加するヘルパー
 * （見出し名から動的に列位置を特定して書き込み）
 */
function enqueueMijyusokuPostToBridge_() {
  const queueSs = SpreadsheetApp.openById(BRIDGE_QUEUE_SS_ID);
  const queueSheet = queueSs.getSheetByName(BRIDGE_QUEUE_SHEET_NAME);
  if (!queueSheet) {
    throw new Error(`シート「${BRIDGE_QUEUE_SHEET_NAME}」が見つかりません。`);
  }

  const lastCol = Math.max(5, queueSheet.getLastColumn());
  const headers = queueSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const idxReqAt  = headers.findIndex(h => h.includes('リクエスト日時'));
  const idxType   = headers.findIndex(h => h.includes('イベント種別'));
  const idxJson   = headers.findIndex(h => h.includes('ペイロード'));
  const idxStatus = headers.findIndex(h => h.includes('処理ステータス'));
  const idxDoneAt = headers.findIndex(h => h.includes('処理完了日時'));

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const newRow = new Array(lastCol).fill('');

  if (idxReqAt > -1)  newRow[idxReqAt]  = nowStr;
  if (idxType > -1)   newRow[idxType]   = 'MIJYUSOKU_POST';
  if (idxJson > -1)   newRow[idxJson]   = '';
  if (idxStatus > -1) newRow[idxStatus] = '未処理';
  if (idxDoneAt > -1) newRow[idxDoneAt] = '';

  // 万が一見出し名が取得できなかった場合のフォールバック
  if (idxType === -1 || idxStatus === -1) {
    queueSheet.appendRow([nowStr, 'MIJYUSOKU_POST', '', '未処理', '']);
  } else {
    queueSheet.appendRow(newRow);
  }
  Logger.log('✅ ⚙️_連携キューに MIJYUSOKU_POST (未処理) を追加しました。');
}

/**
 * 【リカバリー用】DS部へのChatwork再投稿はせず、拠点連絡システムへの連携キューだけを手動送信する関数
 * （今日16:30の未連携分を今すぐ反応させたい場合は、エディタからこの関数を1回実行してください）
 */
function resendQueueOnly() {
  enqueueMijyusokuPostToBridge_();
}