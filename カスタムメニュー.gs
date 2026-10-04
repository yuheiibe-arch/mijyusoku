/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('管理メニュー') // メニュー名
      
      // --- 本番・運用メニュー ---
      .addItem('本番シフト集計実行', 'updateSheetRowAdjusted_CallingCellSpecificFormatting') //
      .addItem('DSちゃんを起動', 'openDSChanUI')
      .addItem('データ一括削除', 'clearData')        
      .addItem('DS部に投稿する', 'postToChatwork1')
      .addToUi();
}