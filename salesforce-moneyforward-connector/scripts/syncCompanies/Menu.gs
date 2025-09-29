
// =============================================================================
// メニュー作成
// =============================================================================

/**
 * スプレッドシート開時にメニューを自動作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('MF-SF統合')
    .addSubMenu(ui.createMenu('認証')
      .addItem('setupConfig','setupMoneyForwardConfig')
      .addItem('🔑 MoneyForward認証', 'menuAuth')
      .addItem('✅ 認証テスト', 'menuTestAuth'))
    .addSeparator()
    .addItem('🔄 完全同期実行', 'menuFullSync')
    .addItem('⚡ 増分同期実行', 'menuIncrementalSync')
    .addSeparator()
    .addItem('📊 定期実行トリガー更新', 'menuTrigger')
    .addSeparator()
    .addItem('🧹 自動クリーンアップ', 'menuCleanup')
    .addItem('⚙️ システム状態確認', 'menuHealthCheck')
    .addSeparator()
    .addItem('📋 初回セットアップ', 'menuSetup')
    .addToUi();
}

// =============================================================================
// メニュー関数
// =============================================================================

/**
 * 初回セットアップ
 */
function menuSetup() {
  try {
    runSetupWizard();
    SpreadsheetApp.getUi().alert('セットアップ完了', 'システムの初期設定が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('セットアップエラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * MoneyForward認証
 */
function menuAuth() {
  try {
    const authUrl = startMFAuth();
    SpreadsheetApp.getUi().alert('MoneyForward認証', 
      `以下のURLにアクセスして認証してください：\n\n${authUrl}`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('認証エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 認証テスト
 */
function menuTestAuth() {
  try {
    testAuth();
    SpreadsheetApp.getUi().alert('認証テスト', '✅ 認証が正常に動作しています', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('認証エラー', `❌ 認証に失敗しました：\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 完全同期実行
 */
function menuFullSync() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('完全同期実行', 
    '全データを同期します。時間がかかる場合があります。\n実行しますか？', 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    try {
      const result = runFullSync();
      ui.alert('完全同期完了', 
        `処理件数: ${result.processed}\n成功: ${result.success}\nエラー: ${result.errors}`, 
        ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('同期エラー', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * 増分同期実行
 */
function menuIncrementalSync() {
  try {
    const result = runIncrementalSync();
    SpreadsheetApp.getUi().alert('増分同期完了', 
      `処理件数: ${result.processed}\n成功: ${result.success}\nエラー: ${result.errors}`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('同期エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 統計レポート表示
 */
function menuStats() {
  try {
    const syncStats = getSyncStatistics();
    const deleteStats = getDeleteCandidateStatistics();
    
    const message = `=== 同期統計 ===\n` +
                   `総取引先: ${syncStats.total}\n` +
                   `連携済み: ${syncStats.linked}\n` +
                   `未連携: ${syncStats.noLink}\n\n` +
                   `=== 削除候補 ===\n` +
                   `承認待ち: ${deleteStats.pending}\n` +
                   `削除完了: ${deleteStats.completed}`;
    
    SpreadsheetApp.getUi().alert('統計レポート', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('統計エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 自動クリーンアップ実行
 */
function menuCleanup() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('自動クリーンアップ', 
    '承認済みの削除候補を削除します。\n実行しますか？', 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    try {
      const result = runAutoCleanup();
      ui.alert('クリーンアップ完了', 
        `処理件数: ${result.processed}\n削除完了: ${result.cleaned}`, 
        ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('クリーンアップエラー', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * 定期実行用トリガーセット
 */
function menuTrigger() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('定期実行用トリガー', 
    '設定を更新します。\n実行しますか？', 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    try {
      const result = setupScheduledTriggers();
      ui.alert(result, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('トリガー変更エラー', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * システム状態確認
 */
function menuHealthCheck() {
  try {
    const results = systemHealthCheck();
    const status = results.overall ? '✅ 正常' : '⚠️ 問題あり';
    const details = `シート: ${results.sheets ? 'OK' : 'NG'}\n` +
                   `設定: ${results.config ? 'OK' : 'NG'}\n` +
                   `認証: ${results.auth ? 'OK' : 'NG'}`;
    
    SpreadsheetApp.getUi().alert(`システム状態: ${status}`, details, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('状態確認エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
