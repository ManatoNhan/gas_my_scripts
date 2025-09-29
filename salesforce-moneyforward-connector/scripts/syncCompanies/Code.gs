/**
 * MoneyForward-Salesforce データ同期システム
 * メイン処理ファイル
 */

// =============================================================================
// 設定・初期化
// =============================================================================

/**
 * 設定値を取得
 */
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getRange('A:B').getValues();
  const config = {};
  
  data.forEach(row => {
    if (row[0] && row[1]) {
      config[row[0]] = row[1];
    }
  });
  
  return config;
}

/**
 * ログを記録
 */
function writeLog(type, processedCount, successCount, errorCount, duration, status, note = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SyncLog');
  sheet.appendRow([
    new Date(),
    type,
    processedCount,
    successCount,
    errorCount,
    duration,
    status,
    note
  ]);
}

/**
 * エラーログを記録
 */
function writeErrorLog(errorType, recordId, description, details = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ErrorLog');
  sheet.appendRow([
    new Date(),
    errorType,
    recordId,
    description,
    details
  ]);
}

// =============================================================================
// API インスタンス
// =============================================================================

/**
 * Salesforce APIインスタンスを取得
 */
function getSalesforceAPI() {
  return new SalesforceAPI();
}

/**
 * MoneyForward APIインスタンスを取得
 */
function getMoneyForwardAPI() {
  const config = getConfig();
  const mfAPI = new MoneyForwardAPI();
  
  // 設定が未完了の場合は設定を実行
  try {
    mfAPI.getOAuthService();
  } catch (error) {
    if (error.message.includes('Client IDまたはClient Secret')) {
      mfAPI.setupConfig(config.MF_CLIENT_ID, config.MF_CLIENT_SECRET);
    }
  }
  
  return mfAPI;
}

// =============================================================================
// 認証・API呼び出しヘルパー
// =============================================================================

/**
 * MoneyForward API呼び出し（後方互換用）
 */
function callMFAPI(endpoint, method = 'GET', payload = null) {
  const mfAPI = getMoneyForwardAPI();
  return mfAPI.callMFAPI(endpoint, method, payload);
}

/**
 * Salesforce API呼び出し（後方互換用）
 */
function callSFAPI(endpoint, method = 'GET', payload = null) {
  const sfAPI = getSalesforceAPI();
  
  try {
    const result = sfAPI.call(endpoint, method, payload);
    
    // PATCH/PUT操作で空のレスポンスの場合は成功とみなす
    if (!result.success && result.error && result.error.includes('Unexpected end of JSON input')) {
      if (method === 'PATCH' || method === 'PUT') {
        console.log('✓ Salesforce更新成功（空レスポンス）');
        return { success: true, data: {} };
      }
    }
    
    return result;
  } catch (error) {
    // JSON解析エラーでPATCH/PUT操作の場合は成功とみなす
    if (error.message && error.message.includes('Unexpected end of JSON input')) {
      if (method === 'PATCH' || method === 'PUT') {
        console.log('✓ Salesforce更新成功（JSON解析エラー回避）');
        return { success: true, data: {} };
      }
    }
    
    console.error('SF API Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * MoneyForward認証開始
 */
function startMFAuth() {
  const mfAPI = getMoneyForwardAPI();
  const service = mfAPI.getOAuthService();
  const authorizationUrl = service.getAuthorizationUrl();
  console.log('以下のURLにアクセスして認証してください:');
  console.log(authorizationUrl);
  return authorizationUrl;
}

/**
 * OAuth認証コールバック
 */
function authCallback(request) {
  const mfAPI = getMoneyForwardAPI();
  const service = mfAPI.getOAuthService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('認証成功！このタブを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗。もう一度お試しください。');
  }
}

// =============================================================================
// 双方向ID連携ユーティリティ
// =============================================================================

/**
 * Account IDから顧客コードを生成
 */
function generateCustomerCode(accountId) {
  return `SF:${accountId}`;
}

/**
 * CustomAccount IDから顧客コードを生成
 */
function generateCustomCustomerCode(accountId) {
  return `取引先ID:${accountId}`;
}

/**
 * 顧客コードからAccount IDを抽出
 */
function extractSalesforceId(customerCode) {
  if (!customerCode) return null;
  const match = customerCode.match(/^SF:(.+)$/);
  return match ? match[1] : null;
}

/**
 * Salesforce ID連携の妥当性チェック
 */
function isValidSalesforceLink(customerCode) {
  if (!customerCode) return false;
  // カスタムAccount IDの形式をチェック（より柔軟に）
  return /^SF:.+$/.test(customerCode) && customerCode.length > 3;
}

/**
 * カスタムAccount IDでSalesforce Accountを検索
 */
function findSFAccountByCustomId(customAccountId) {
  try {
    const sfAPI = getSalesforceAPI();
    
    // まずaccountId__cで検索
    let query = `SELECT Id, Name, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
                 MoneyForward_Partner_ID__c, MF_Last_Sync_Date__c, MF_Address_Hash__c, accountId__c 
                 FROM Account 
                 WHERE accountId__c = '${customAccountId}' 
                 LIMIT 1`;
    
    let result = sfAPI.query(query);
    
    if (result.success && result.data.records.length > 0) {
      return result.data.records[0];
    }
    
    // accountId__cで見つからない場合は標準IDで検索（フォールバック）
    if (customAccountId.match(/^[a-zA-Z0-9]{15,18}$/)) {
      query = `SELECT Id, Name, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
               MoneyForward_Partner_ID__c, MF_Last_Sync_Date__c, MF_Address_Hash__c, accountId__c 
               FROM Account 
               WHERE Id = '${customAccountId}' 
               LIMIT 1`;
      
      result = sfAPI.query(query);
      
      if (result.success && result.data.records.length > 0) {
        return result.data.records[0];
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('SF Account検索エラー:', error);
    return null;
  }
}

/**
 * 住所ハッシュを生成（変更検知用）
 */
function generateAddressHash(address) {
  if (!address) return '';
  // 住所文字列をハッシュ化（簡易版）
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, address));
}

// =============================================================================
// メイン同期処理
// =============================================================================

/**
 * 完全同期処理（手動実行用）
 */
function runFullSync() {
  const startTime = new Date();
  let processedCount = 0;
  let successCount = 0;
  let errorCount = 0;
  
  try {
    console.log('=== 完全同期開始 ===');
    
    // 1. 全MF取引先取得
    console.log('MoneyForward取引先を取得中...');
    const mfPartners = getAllMFPartners();
    console.log(`MoneyForward取引先数: ${mfPartners.length}`);
    
    // 2. 取引先管理シート更新
    updatePartnerManagementSheet(mfPartners);
    
    // 3. データ整合性チェック
    console.log('データ整合性チェック実行中...');
    const checkResult = performDataIntegrityCheck(mfPartners);
    processedCount += checkResult.processed;
    successCount += checkResult.success;
    errorCount += checkResult.errors;
    
    // 4. SF側変更検知・更新
    console.log('Salesforce変更検知中...');
    const syncResult = syncSalesforceChanges();
    processedCount += syncResult.processed;
    successCount += syncResult.success;
    errorCount += syncResult.errors;
    
    const duration = Math.round((new Date() - startTime) / 1000);
    writeLog('FULL_SYNC', processedCount, successCount, errorCount, duration, 'COMPLETED', 
             `MF取引先: ${mfPartners.length}件`);
    
    console.log('=== 完全同期完了 ===');
    console.log(`処理件数: ${processedCount}, 成功: ${successCount}, エラー: ${errorCount}`);
    
    return {
      success: true,
      processed: processedCount,
      success: successCount,
      errors: errorCount,
      duration: duration
    };
    
  } catch (error) {
    const duration = Math.round((new Date() - startTime) / 1000);
    writeLog('FULL_SYNC', processedCount, successCount, errorCount, duration, 'ERROR', error.message);
    writeErrorLog('SYSTEM', 'FULL_SYNC', 'システムエラー', error.message);
    
    console.error('完全同期エラー:', error);
    throw error;
  }
}

/**
 * 増分同期処理（定期実行用）
 */
function runIncrementalSync() {
  const startTime = new Date();
  let processedCount = 0;
  let successCount = 0;
  let errorCount = 0;
  
  try {
    console.log('=== 増分同期開始 ===');
    
    // 最近変更されたSF Accountのみチェック
    const syncResult = syncRecentSalesforceChanges();
    processedCount = syncResult.processed;
    successCount = syncResult.success;
    errorCount = syncResult.errors;
    
    const duration = Math.round((new Date() - startTime) / 1000);
    writeLog('INCREMENTAL', processedCount, successCount, errorCount, duration, 'COMPLETED');
    
    console.log('=== 増分同期完了 ===');
    
    return {
      success: true,
      processed: processedCount,
      success: successCount,
      errors: errorCount,
      duration: duration
    };
    
  } catch (error) {
    const duration = Math.round((new Date() - startTime) / 1000);
    writeLog('INCREMENTAL', processedCount, successCount, errorCount, duration, 'ERROR', error.message);
    writeErrorLog('SYSTEM', 'INCREMENTAL', 'システムエラー', error.message);
    
    console.error('増分同期エラー:', error);
    throw error;
  }
}

// =============================================================================
// テスト・デバッグ用関数
// =============================================================================

/**
 * 認証テスト
 */
function testAuth() {
  try {
    // MoneyForward認証テスト
    console.log('MoneyForward認証テスト...');
    const mfAPI = getMoneyForwardAPI();
    const mfResult = mfAPI.get('/partners?page=1&per_page=1');
    if (mfResult.success) {
      console.log('✓ MoneyForward認証成功');
    } else {
      console.log('✗ MoneyForward認証失敗:', mfResult.error);
    }
    
    // Salesforce認証テスト
    console.log('Salesforce認証テスト...');
    const sfAPI = getSalesforceAPI();
    const sfResult = sfAPI.query('SELECT Id FROM Account LIMIT 1');
    if (sfResult.success) {
      console.log('✓ Salesforce認証成功');
    } else {
      console.log('✗ Salesforce認証失敗:', sfResult.error);
    }
    
  } catch (error) {
    console.error('認証テストエラー:', error);
  }
}

/**
 * 設定値確認
 */
function checkConfig() {
  const config = getConfig();
  console.log('=== 設定値確認 ===');
  
  const requiredKeys = [
    'MF_CLIENT_ID', 'MF_CLIENT_SECRET',
    'SF_CLIENT_ID', 'SF_CLIENT_SECRET'
  ];
  
  requiredKeys.forEach(key => {
    if (config[key]) {
      console.log(`✓ ${key}: 設定済み`);
    } else {
      console.log(`✗ ${key}: 未設定`);
    }
  });
}
