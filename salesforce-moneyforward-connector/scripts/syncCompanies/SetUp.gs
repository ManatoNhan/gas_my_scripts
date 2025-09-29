/**
 * スプレッドシートセットアップ・初期化
 * 初回実行時にシートとヘッダーを作成
 */

// =============================================================================
// 初期セットアップ
// =============================================================================

/**
 * 全シートをセットアップ（初回実行用）
 */
function setupAllSheets() {
  console.log('=== データ同期システム セットアップ開始 ===');
  
  try {
    setupConfigSheet();
    setupSyncLogSheet();
    setupPartnerManagementSheet();
    setupErrorLogSheet();
    setupDeleteCandidatesSheet();
    
    console.log('=== セットアップ完了 ===');
    console.log('1. Configシートに認証情報を入力してください');
    console.log('2. testAuth()で認証テストを実行してください');
    console.log('3. runFullSync()で初回同期を実行してください');
    
  } catch (error) {
    console.error('セットアップエラー:', error);
    throw error;
  }
}

/*
 * 設定シートをセットアップ
 */
function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  
  if (!sheet) {
    sheet = ss.insertSheet('Config');
  }
  
  // 既存データをクリア
  sheet.clear();
  
  // ヘッダー設定
  const headers = [
    ['設定項目', '設定値', '説明'],
    ['MF_CLIENT_ID', '', 'MoneyForward API Client ID'],
    ['MF_CLIENT_SECRET', '', 'MoneyForward API Client Secret'],
    ['SF_CLIENT_ID', '', 'Salesforce Client ID'],
    ['SF_CLIENT_SECRET', '', 'Salesforce Client Secret'],
    ['SYNC_INTERVAL', '1440', '増分同期間隔（分）※1440分=24時間'],
    ['AUTO_SYNC_ENABLED', 'FALSE', '自動同期有効化（TRUE/FALSE）※手動運用推奨'],
    ['AUTO_CLEANUP', 'FALSE', '自動クリーンアップ有効化（TRUE/FALSE）'],
    ['BATCH_SIZE', '100', 'バッチ処理サイズ'],
    ['RATE_LIMIT_DELAY', '100', 'APIレート制限用待機時間（ミリ秒）'],
    ['LOG_RETENTION_DAYS', '30', 'ログ保持期間（日）']
  ];
  
  const range = sheet.getRange(1, 1, headers.length, headers[0].length);
  range.setValues(headers);
  
  // スタイル設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 300);
  
  console.log('✓ Configシート作成完了');
}

/**
 * 同期ログシートをセットアップ
 */
function setupSyncLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('SyncLog');
  
  if (!sheet) {
    sheet = ss.insertSheet('SyncLog');
  }
  
  // 既存データをクリア
  sheet.clear();
  
  // ヘッダー設定
  const headers = [
    '実行日時', '処理タイプ', '処理件数', '成功件数', 
    'エラー件数', '処理時間(秒)', 'ステータス', '備考'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#34a853').setFontColor('white').setFontWeight('bold');
  
  // 列幅設定
  sheet.setColumnWidth(1, 150); // 実行日時
  sheet.setColumnWidth(2, 120); // 処理タイプ
  sheet.setColumnWidth(3, 80);  // 処理件数
  sheet.setColumnWidth(4, 80);  // 成功件数
  sheet.setColumnWidth(5, 80);  // エラー件数
  sheet.setColumnWidth(6, 80);  // 処理時間
  sheet.setColumnWidth(7, 100); // ステータス
  sheet.setColumnWidth(8, 300); // 備考
  
  console.log('✓ SyncLogシート作成完了');
}

/**
 * 取引先管理シートをセットアップ
 */
function setupPartnerManagementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PartnerManagement');
  
  if (!sheet) {
    sheet = ss.insertSheet('PartnerManagement');
  }
  
  // 既存データをクリア
  sheet.clear();
  
  // ヘッダー設定
  const headers = [
    '更新日時', 'MF取引先ID', 'MF取引先名', '顧客コード', 
    'SF Account ID', 'SF Account名', '同期ステータス', '最終同期日時',
    'MF作成日', 'MF更新日'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#ff9800').setFontColor('white').setFontWeight('bold');
  
  // 列幅設定
  sheet.setColumnWidth(1, 130); // 更新日時
  sheet.setColumnWidth(2, 100); // MF取引先ID
  sheet.setColumnWidth(3, 200); // MF取引先名
  sheet.setColumnWidth(4, 150); // 顧客コード
  sheet.setColumnWidth(5, 150); // SF Account ID
  sheet.setColumnWidth(6, 200); // SF Account名
  sheet.setColumnWidth(7, 120); // 同期ステータス
  sheet.setColumnWidth(8, 130); // 最終同期日時
  sheet.setColumnWidth(9, 130); // MF作成日
  sheet.setColumnWidth(10, 130); // MF更新日
  
  // フィルタ設定
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  
  console.log('✓ PartnerManagementシート作成完了');
}

/**
 * エラーログシートをセットアップ
 */
function setupErrorLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ErrorLog');
  
  if (!sheet) {
    sheet = ss.insertSheet('ErrorLog');
  }
  
  // 既存データをクリア
  sheet.clear();
  
  // ヘッダー設定
  const headers = [
    '発生日時', 'エラータイプ', 'レコードID', 'エラー概要', 'エラー詳細'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#ea4335').setFontColor('white').setFontWeight('bold');
  
  // 列幅設定
  sheet.setColumnWidth(1, 150); // 発生日時
  sheet.setColumnWidth(2, 120); // エラータイプ
  sheet.setColumnWidth(3, 150); // レコードID
  sheet.setColumnWidth(4, 250); // エラー概要
  sheet.setColumnWidth(5, 400); // エラー詳細
  
  // フィルタ設定
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  
  console.log('✓ ErrorLogシート作成完了');
}

/**
 * 削除候補シートをセットアップ
 */
function setupDeleteCandidatesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DeleteCandidates');
  
  if (!sheet) {
    sheet = ss.insertSheet('DeleteCandidates');
  }
  
  // 既存データをクリア
  sheet.clear();
  
  // ヘッダー設定
  const headers = [
    '検知日時', 'MF取引先ID', '取引先名', '顧客コード', 
    '削除理由コード', '削除理由詳細', 'ステータス', '承認者',
    '承認日時', '備考'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#9c27b0').setFontColor('white').setFontWeight('bold');
  
  // 列幅設定
  sheet.setColumnWidth(1, 130); // 検知日時
  sheet.setColumnWidth(2, 100); // MF取引先ID
  sheet.setColumnWidth(3, 200); // 取引先名
  sheet.setColumnWidth(4, 150); // 顧客コード
  sheet.setColumnWidth(5, 120); // 削除理由コード
  sheet.setColumnWidth(6, 250); // 削除理由詳細
  sheet.setColumnWidth(7, 100); // ステータス
  sheet.setColumnWidth(8, 100); // 承認者
  sheet.setColumnWidth(9, 130); // 承認日時
  sheet.setColumnWidth(10, 200); // 備考
  
  // フィルタ設定
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  
  // データ検証設定（ステータス列）
  const statusRange = sheet.getRange(2, 7, 1000, 1); // ステータス列（G列）
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PENDING', 'APPROVED', 'REJECTED', 'COMPLETED', 'ERROR'])
    .build();
  statusRange.setDataValidation(statusValidation);
  
  console.log('✓ DeleteCandidatesシート作成完了');
}

// =============================================================================
// スプレッドシート構造確認・修復
// =============================================================================

/**
 * 全シートの構造を確認
 */
function checkSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    'Config', 'SyncLog', 'PartnerManagement', 'ErrorLog', 'DeleteCandidates'
  ];
  
  console.log('=== シート構造確認 ===');
  
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());
  const missingSheets = [];
  
  requiredSheets.forEach(sheetName => {
    if (existingSheets.includes(sheetName)) {
      console.log(`✓ ${sheetName}: 存在`);
    } else {
      console.log(`✗ ${sheetName}: 不足`);
      missingSheets.push(sheetName);
    }
  });
  
  if (missingSheets.length > 0) {
    console.log(`不足シート: ${missingSheets.join(', ')}`);
    console.log('setupAllSheets()を実行して修復してください');
    return false;
  } else {
    console.log('✓ 全シートが正常に存在します');
    return true;
  }
}

/**
 * 設定値の検証
 */
function validateConfig() {
  console.log('=== 設定値検証 ===');
  
  try {
    const config = getConfig();
    const requiredKeys = [
      'MF_CLIENT_ID', 'MF_CLIENT_SECRET',
      'SF_CLIENT_ID', 'SF_CLIENT_SECRET'
    ];
    
    let isValid = true;
    
    requiredKeys.forEach(key => {
      if (config[key] && config[key].trim() !== '') {
        console.log(`✓ ${key}: 設定済み`);
      } else {
        console.log(`✗ ${key}: 未設定または空`);
        isValid = false;
      }
    });
    
    // オプション設定の確認
    const optionalKeys = [
      'SYNC_INTERVAL', 'AUTO_CLEANUP', 'BATCH_SIZE', 
      'RATE_LIMIT_DELAY', 'LOG_RETENTION_DAYS'
    ];
    
    console.log('--- オプション設定 ---');
    optionalKeys.forEach(key => {
      const value = config[key] || 'デフォルト値使用';
      console.log(`${key}: ${value}`);
    });
    
    if (isValid) {
      console.log('✓ 設定値は有効です');
    } else {
      console.log('✗ 必須設定値が不足しています。Configシートを確認してください');
    }
    
    return isValid;
    
  } catch (error) {
    console.error('設定値検証エラー:', error);
    return false;
  }
}

// =============================================================================
// 初期データ・サンプルデータ設定
// =============================================================================

/**
 * サンプルログデータを作成（テスト用）
 */
function createSampleData() {
  console.log('=== サンプルデータ作成 ===');
  
  // SyncLogにサンプルデータ
  const syncLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SyncLog');
  const sampleSyncData = [
    [new Date(Date.now() - 86400000), 'FULL_SYNC', 150, 145, 5, 45, 'COMPLETED', '初回同期テスト'],
    [new Date(Date.now() - 43200000), 'INCREMENTAL', 25, 23, 2, 12, 'COMPLETED', '増分同期テスト'],
    [new Date(Date.now() - 3600000), 'CLEANUP', 5, 3, 2, 8, 'COMPLETED', 'クリーンアップテスト']
  ];
  
  sampleSyncData.forEach(row => {
    syncLogSheet.appendRow(row);
  });
  
  // ErrorLogにサンプルデータ
  const errorLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ErrorLog');
  const sampleErrorData = [
    [new Date(Date.now() - 7200000), 'MF_API', 'PARTNER_001', 'API レート制限エラー', 'Too Many Requests - Rate limit exceeded'],
    [new Date(Date.now() - 5400000), 'SF_UPDATE', '001XXXXXXXXXXXXXX', 'Account更新エラー', 'REQUIRED_FIELD_MISSING: Required fields are missing']
  ];
  
  sampleErrorData.forEach(row => {
    errorLogSheet.appendRow(row);
  });
  
  console.log('✓ サンプルデータ作成完了');
}

/**
 * サンプルデータをクリア
 */
function clearSampleData() {
  console.log('=== サンプルデータクリア ===');
  
  const sheets = ['SyncLog', 'PartnerManagement', 'ErrorLog', 'DeleteCandidates'];
  
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
        console.log(`✓ ${sheetName} データクリア完了`);
      }
    }
  });
}

// =============================================================================
// 保守・メンテナンス機能
// =============================================================================

/**
 * 古いログデータを削除
 */
function cleanupOldLogs() {
  const config = getConfig();
  const retentionDays = parseInt(config.LOG_RETENTION_DAYS) || 30;
  const cutoffDate = new Date(Date.now() - retentionDays * 24 * 60 * 60 * 1000);
  
  console.log(`=== ${retentionDays}日より古いログをクリーンアップ ===`);
  
  const logSheets = ['SyncLog', 'ErrorLog'];
  let totalDeleted = 0;
  
  logSheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    // ヘッダー行をスキップして、古いデータを特定
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]); // 最初の列が日時と仮定
      if (rowDate < cutoffDate) {
        rowsToDelete.push(i + 1); // 行番号は1ベース
      }
    }
    
    // 後ろから削除（行番号がずれないように）
    rowsToDelete.reverse().forEach(rowNumber => {
      sheet.deleteRow(rowNumber);
      totalDeleted++;
    });
    
    console.log(`${sheetName}: ${rowsToDelete.length}行削除`);
  });
  
  console.log(`✓ 古いログクリーンアップ完了: 合計${totalDeleted}行削除`);
  return totalDeleted;
}

/**
 * システム健全性チェック
 */
function systemHealthCheck() {
  console.log('=== システム健全性チェック ===');
  
  const results = {
    sheets: false,
    config: false,
    auth: false,
    overall: false
  };
  
  // 1. シート構造チェック
  results.sheets = checkSheetStructure();
  
  // 2. 設定値チェック
  results.config = validateConfig();
  
  // 3. 認証チェック
  try {
    testAuth();
    results.auth = true;
    console.log('✓ 認証: OK');
  } catch (error) {
    results.auth = false;
    console.log('✗ 認証: NG -', error.message);
  }
  
  // 総合判定
  results.overall = results.sheets && results.config && results.auth;
  
  console.log('=== システム健全性チェック結果 ===');
  console.log(`シート構造: ${results.sheets ? 'OK' : 'NG'}`);
  console.log(`設定値: ${results.config ? 'OK' : 'NG'}`);
  console.log(`認証: ${results.auth ? 'OK' : 'NG'}`);
  console.log(`総合: ${results.overall ? 'OK' : 'NG'}`);
  
  if (results.overall) {
    console.log('✓ システムは正常に動作可能です');
  } else {
    console.log('✗ システムに問題があります。上記の項目を確認してください');
  }
  
  return results;
}

// =============================================================================
// セットアップウィザード
// =============================================================================

/**
 * セットアップウィザード実行
 */
function runSetupWizard() {
  console.log('');
  console.log('='.repeat(60));
  console.log('    MoneyForward-Salesforce 統合システム');
  console.log('           セットアップウィザード');
  console.log('='.repeat(60));
  console.log('');
  
  // Step 1: シート作成
  console.log('Step 1: シート構造の作成...');
  try {
    setupAllSheets();
    console.log('✓ Step 1 完了');
  } catch (error) {
    console.log('✗ Step 1 失敗:', error.message);
    return false;
  }
  
  // Step 2: 設定確認
  console.log('');
  console.log('Step 2: 設定の確認...');
  const configValid = validateConfig();
  if (!configValid) {
    console.log('');
    console.log('⚠️  設定が不完全です。以下の手順を実行してください:');
    console.log('1. Configシートを開く');
    console.log('2. 以下の値を設定:');
    console.log('   - MF_CLIENT_ID: MoneyForward API Client ID');
    console.log('   - MF_CLIENT_SECRET: MoneyForward API Client Secret');
    console.log('   - SF_CLIENT_ID: Salesforce Client ID');
    console.log('   - SF_CLIENT_SECRET: Salesforce Client Secret');
    console.log('3. 設定完了後、再度 runSetupWizard() を実行');
    console.log('');
    return false;
  }
  console.log('✓ Step 2 完了');
  
  // Step 3: 認証テスト
  console.log('');
  console.log('Step 3: 認証テスト...');
  try {
    testAuth();
    console.log('✓ Step 3 完了');
  } catch (error) {
    console.log('✗ Step 3 失敗:', error.message);
    console.log('');
    console.log('⚠️  認証に失敗しました。以下を確認してください:');
    console.log('1. API認証情報が正しいかConfigシートで確認');
    console.log('2. MoneyForward認証: startMFAuth() を実行してOAuth認証完了');
    console.log('3. Salesforce認証: Client Credentialsが有効か確認');
    console.log('');
    return false;
  }
  
  // Step 4: 定期実行設定
  console.log('');
  console.log('Step 4: 定期実行トリガーの設定...');
  try {
    setupScheduledTriggers();
    console.log('✓ Step 4 完了');
  } catch (error) {
    console.log('⚠️ Step 4 部分的失敗:', error.message);
    console.log('手動で定期実行を設定することも可能です');
  }
  
  // 完了メッセージ
  console.log('');
  console.log('🎉 セットアップ完了！');
  console.log('');
  console.log('=== 次のステップ ===');
  console.log('1. runFullSync() - 初回完全同期実行');
  console.log('2. showStatisticsReport() - 同期結果確認');
  console.log('3. 以降は自動で増分同期が実行されます');
  console.log('');
  console.log('=== 管理機能 ===');
  console.log('- systemHealthCheck() - システム状態確認');
  console.log('- cleanupOldLogs() - 古いログ削除');
  console.log('- runAutoCleanup() - 削除候補の自動クリーンアップ');
  console.log('');
  
  return true;
}

/**
 * クイックスタート（設定済み環境用）
 */
function quickStart() {
  console.log('=== クイックスタート ===');
  
  // 健全性チェック
  const health = systemHealthCheck();
  if (!health.overall) {
    console.log('システムに問題があります。runSetupWizard()を実行してください');
    return false;
  }
  
  // 初回同期実行
  console.log('初回同期を実行します...');
  try {
    const result = runFullSync();
    console.log('✓ 初回同期完了');
    console.log(`処理件数: ${result.processed}, 成功: ${result.success}, エラー: ${result.errors}`);
  } catch (error) {
    console.log('✗ 初回同期失敗:', error.message);
    return false;
  }
  
  // 統計表示
  showStatisticsReport();
  
  console.log('');
  console.log('✓ クイックスタート完了！システムが稼働中です');
  return true;
}
