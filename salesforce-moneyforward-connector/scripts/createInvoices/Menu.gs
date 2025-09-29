//====================
// メニュー追加/その他
//====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Options')
    // .addItem('見積書', 'exportQuote')
    .addItem('請求情報をSalesforceから取り込み', 'exportInvoice')
    .addSeparator()
    .addItem('📄 Moneyforwardで請求書を発行', 'showInvoicePublishDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('マネーフォワード認証')
      .addItem('初期設定', 'setupMoneyForwardConfig')
      .addItem('認証開始', 'authorize')
      .addItem('認証状況確認', 'checkAuthStatus'))
      // .addItem('会社設定', 'testGetCompanies'))
    .addToUi();
}


 // 拡張ヘッダー
const headers = [
  'チェック', '実行日時', '商談URL', '商談ID', '取引先名称', '件名', 
  'アウトプットURL', '実行結果', '発行ステータス', '請求書番号', '発行日時', 'エラーメッセージ'
];

