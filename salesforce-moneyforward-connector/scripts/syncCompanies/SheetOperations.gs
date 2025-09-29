/**
 * スプレッドシート操作・削除候補管理
 */

// =============================================================================
// 削除候補管理
// =============================================================================

/**
 * 削除候補シートをクリア
 */
function clearDeleteCandidatesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeleteCandidates');
  
  // ヘッダー行以外をクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

/**
 * 削除候補を追加
 */
function addDeleteCandidate(partner, reason, description) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeleteCandidates');
  
  sheet.appendRow([
    new Date(),           // 検知日時
    partner.id,           // MF取引先ID
    partner.name,         // 取引先名
    partner.code || '',   // 顧客コード
    reason,               // 削除理由コード
    description,          // 削除理由詳細
    'PENDING',            // ステータス（PENDING/APPROVED/REJECTED）
    '',                   // 承認者
    '',                   // 承認日時
    ''                    // 備考
  ]);
  
  console.log(`削除候補追加: ${partner.name} (${partner.id}) - ${description}`);
}

/**
 * 削除候補を取得
 */
function getDeleteCandidates(status = 'PENDING') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeleteCandidates');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return [];
  }
  
  const candidates = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[6] === status) { // ステータス列をチェック
      candidates.push({
        row: i + 1,
        detectedDate: row[0],
        mfPartnerId: row[1],
        partnerName: row[2],
        customerCode: row[3],
        reasonCode: row[4],
        description: row[5],
        status: row[6],
        approver: row[7],
        approvedDate: row[8],
        note: row[9]
      });
    }
  }
  
  return candidates;
}

/**
 * 削除候補のステータスを更新
 */
function updateDeleteCandidateStatus(mfPartnerId, status, approver = '', note = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeleteCandidates');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === mfPartnerId) { // MF取引先IDでマッチ
      const row = i + 1;
      sheet.getRange(row, 7).setValue(status);        // ステータス
      sheet.getRange(row, 8).setValue(approver);      // 承認者
      sheet.getRange(row, 9).setValue(new Date());    // 承認日時
      sheet.getRange(row, 10).setValue(note);         // 備考
      
      console.log(`削除候補ステータス更新: ${mfPartnerId} -> ${status}`);
      return true;
    }
  }
  
  return false;
}

// =============================================================================
// 取引先管理シート操作
// =============================================================================

/**
 * 取引先管理シートを更新
 */
function updatePartnerManagementSheet(mfPartners) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartnerManagement');
  
  // 既存データをクリア（ヘッダー行は残す）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  console.log('取引先管理シート更新中...');
  
  const rows = [];
  let processedCount = 0;
  
  for (const partner of mfPartners) {
    try {
      const customerCode = partner.customer_code || '';
      const sfAccountId = extractSalesforceId(customerCode);
      let sfAccountName = '';
      let syncStatus = '';
      let lastSyncDate = '';
      
      // Salesforce Account情報を取得
      if (sfAccountId) {
        const sfAccount = getSFAccount(sfAccountId);
        if (sfAccount) {
          sfAccountName = sfAccount.Name;
          syncStatus = 'LINKED';
          lastSyncDate = sfAccount.MF_Last_Sync_Date__c || '';
        } else {
          syncStatus = 'SF_NOT_FOUND';
        }
      } else if (customerCode && !isValidSalesforceLink(customerCode)) {
        syncStatus = 'INVALID_CODE';
      } else {
        syncStatus = 'NO_LINK';
      }
      
      rows.push([
        new Date(),                    // 更新日時
        partner.id,                   // MF取引先ID
        partner.name,                 // MF取引先名
        customerCode,                 // 顧客コード
        sfAccountId || '',            // SF Account ID
        sfAccountName,                // SF Account名
        syncStatus,                   // 同期ステータス
        lastSyncDate,                 // 最終同期日時
        partner.created_at,           // MF作成日
        partner.updated_at            // MF更新日
      ]);
      
      processedCount++;
      
    } catch (error) {
      console.error(`取引先管理シート更新エラー (Partner ID: ${partner.id}):`, error);
    }
    
    // 進捗表示
    if (processedCount % 100 === 0) {
      console.log(`取引先管理シート進捗: ${processedCount}/${mfPartners.length}`);
    }
  }
  
  // 一括でデータを書き込み
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  console.log(`取引先管理シート更新完了: ${rows.length}件`);
}

/**
 * 取引先管理シートから特定の取引先情報を取得
 */
function getPartnerFromManagementSheet(mfPartnerId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartnerManagement');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === mfPartnerId) { // MF取引先ID列
      return {
        updateDate: row[0],
        mfPartnerId: row[1],
        mfPartnerName: row[2],
        customerCode: row[3],
        sfAccountId: row[4],
        sfAccountName: row[5],
        syncStatus: row[6],
        lastSyncDate: row[7],
        mfCreatedAt: row[8],
        mfUpdatedAt: row[9]
      };
    }
  }
  
  return null;
}

// =============================================================================
// 自動クリーンアップ機能
// =============================================================================

/**
 * 自動クリーンアップ実行
 */
function runAutoCleanup() {
  const config = getConfig();
  
  if (config.AUTO_CLEANUP !== 'TRUE') {
    console.log('自動クリーンアップは無効化されています');
    return { success: true, processed: 0, cleaned: 0 };
  }
  
  const startTime = new Date();
  let processed = 0;
  let cleaned = 0;
  let errors = 0;
  
  try {
    console.log('=== 自動クリーンアップ開始 ===');
    
    // 承認済みの削除候補を取得
    const approvedCandidates = getDeleteCandidates('APPROVED');
    console.log(`承認済み削除候補: ${approvedCandidates.length}件`);
    
    const mfAPI = getMoneyForwardAPI();
    
    for (const candidate of approvedCandidates) {
      processed++;
      
      try {
        // MoneyForward取引先を削除
        const result = mfAPI.callMFAPI(`/partners/${candidate.mfPartnerId}`, 'DELETE');
        
        if (result.success) {
          cleaned++;
          updateDeleteCandidateStatus(
            candidate.mfPartnerId, 
            'COMPLETED', 
            'SYSTEM', 
            '自動削除完了'
          );
          console.log(`✓ MF取引先削除完了: ${candidate.partnerName} (${candidate.mfPartnerId})`);
        } else {
          errors++;
          updateDeleteCandidateStatus(
            candidate.mfPartnerId, 
            'ERROR', 
            'SYSTEM', 
            `削除エラー: ${result.error}`
          );
          writeErrorLog('CLEANUP', candidate.mfPartnerId, '自動削除エラー', result.error);
        }
        
      } catch (error) {
        errors++;
        updateDeleteCandidateStatus(
          candidate.mfPartnerId, 
          'ERROR', 
          'SYSTEM', 
          `削除エラー: ${error.message}`
        );
        writeErrorLog('CLEANUP', candidate.mfPartnerId, '自動削除システムエラー', error.message);
        console.error(`自動削除エラー (Partner ID: ${candidate.mfPartnerId}):`, error);
      }
    }
    
    const duration = Math.round((new Date() - startTime) / 1000);
    writeLog('CLEANUP', processed, cleaned, errors, duration, 'COMPLETED', 
             `削除完了: ${cleaned}件`);
    
    console.log('=== 自動クリーンアップ完了 ===');
    console.log(`処理: ${processed}, 削除: ${cleaned}, エラー: ${errors}`);
    
    return {
      success: true,
      processed: processed,
      cleaned: cleaned,
      errors: errors,
      duration: duration
    };
    
  } catch (error) {
    const duration = Math.round((new Date() - startTime) / 1000);
    writeLog('CLEANUP', processed, cleaned, errors, duration, 'ERROR', error.message);
    writeErrorLog('SYSTEM', 'CLEANUP', 'システムエラー', error.message);
    
    console.error('自動クリーンアップエラー:', error);
    throw error;
  }
}

/**
 * 手動削除承認
 */
function approveDeleteCandidate(mfPartnerId, approver, note = '') {
  return updateDeleteCandidateStatus(mfPartnerId, 'APPROVED', approver, note);
}

/**
 * 手動削除拒否
 */
function rejectDeleteCandidate(mfPartnerId, approver, note = '') {
  return updateDeleteCandidateStatus(mfPartnerId, 'REJECTED', approver, note);
}

// =============================================================================
// 統計・レポート機能
// =============================================================================

/**
 * 同期統計を取得
 */
function getSyncStatistics() {
  try {
    // 取引先管理シートから統計を算出
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartnerManagement');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        total: 0,
        linked: 0,
        noLink: 0,
        invalidCode: 0,
        sfNotFound: 0
      };
    }
    
    const stats = {
      total: 0,
      linked: 0,
      noLink: 0,
      invalidCode: 0,
      sfNotFound: 0
    };
    
    for (let i = 1; i < data.length; i++) {
      const syncStatus = data[i][6]; // 同期ステータス列
      stats.total++;
      
      switch (syncStatus) {
        case 'LINKED':
          stats.linked++;
          break;
        case 'NO_LINK':
          stats.noLink++;
          break;
        case 'INVALID_CODE':
          stats.invalidCode++;
          break;
        case 'SF_NOT_FOUND':
          stats.sfNotFound++;
          break;
      }
    }
    
    return stats;
    
  } catch (error) {
    console.error('統計取得エラー:', error);
    return null;
  }
}

/**
 * 削除候補統計を取得
 */
function getDeleteCandidateStatistics() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeleteCandidates');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        total: 0,
        pending: 0,
        approved: 0,
        rejected: 0,
        completed: 0,
        error: 0
      };
    }
    
    const stats = {
      total: 0,
      pending: 0,
      approved: 0,
      rejected: 0,
      completed: 0,
      error: 0
    };
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][6]; // ステータス列
      stats.total++;
      
      switch (status) {
        case 'PENDING':
          stats.pending++;
          break;
        case 'APPROVED':
          stats.approved++;
          break;
        case 'REJECTED':
          stats.rejected++;
          break;
        case 'COMPLETED':
          stats.completed++;
          break;
        case 'ERROR':
          stats.error++;
          break;
      }
    }
    
    return stats;
    
  } catch (error) {
    console.error('削除候補統計取得エラー:', error);
    return null;
  }
}

/**
 * 統計レポートを表示
 */
function showStatisticsReport() {
  console.log('=== 同期統計レポート ===');
  
  const syncStats = getSyncStatistics();
  if (syncStats) {
    console.log(`総取引先数: ${syncStats.total}`);
    console.log(`連携済み: ${syncStats.linked} (${Math.round(syncStats.linked/syncStats.total*100)}%)`);
    console.log(`未連携: ${syncStats.noLink} (${Math.round(syncStats.noLink/syncStats.total*100)}%)`);
    console.log(`無効コード: ${syncStats.invalidCode}`);
    console.log(`SF未発見: ${syncStats.sfNotFound}`);
  }
  
  console.log('=== 削除候補統計 ===');
  
  const deleteStats = getDeleteCandidateStatistics();
  if (deleteStats) {
    console.log(`削除候補総数: ${deleteStats.total}`);
    console.log(`承認待ち: ${deleteStats.pending}`);
    console.log(`承認済み: ${deleteStats.approved}`);
    console.log(`拒否: ${deleteStats.rejected}`);
    console.log(`削除完了: ${deleteStats.completed}`);
    console.log(`エラー: ${deleteStats.error}`);
  }
}

// =============================================================================
// 定期実行設定
// =============================================================================

/**
 * 定期実行用トリガーを設定
 */
function setupScheduledTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runIncrementalSync' || 
        trigger.getHandlerFunction() === 'runAutoCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  const config = getConfig();
  let msg = null;

  // 自動同期が有効化されている場合のみトリガー設定
  console.log('AUTO_SYNC_ENABLED='+ config.AUTO_SYNC_ENABLED)
  if (!config.AUTO_SYNC_ENABLED || (config.AUTO_SYNC_ENABLED !== true && config.AUTO_SYNC_ENABLED !== 'TRUE')) {
    msg = 'ℹ️ 自動同期が無効化されています。手動実行での運用となります。\n💡 必要に応じてメニューから「完全同期実行」または「増分同期実行」を使用してください。';
    console.log(msg);
    showManualOperationGuide();
    return msg;
  }
  
  const syncInterval = parseInt(config.SYNC_INTERVAL) || 1440; // デフォルト1440分（24時間）
  
  // 増分同期用トリガー
  if (syncInterval >= 1440) {
    // 24時間以上の場合は日次実行
    ScriptApp.newTrigger('runIncrementalSync')
      .timeBased()
      .everyDays(Math.floor(syncInterval / 1440))
      .atHour(9) // 午前9時実行
      .create();
    msg = `定期実行トリガー設定完了: 増分同期 ${Math.floor(syncInterval / 1440)}日間隔（午前9時）`
    console.log(msg);
  } else {
    // 24時間未満の場合は時間間隔実行
    const hours = Math.floor(syncInterval / 60);
    if (hours >= 1) {
      ScriptApp.newTrigger('runIncrementalSync')
        .timeBased()
        .everyHours(hours)
        .create();
      msg = `定期実行トリガー設定完了: 増分同期 ${hours}時間間隔`;
      console.log(msg);
    } else {
      msg = '⚠️ 同期間隔が短すぎます。手動実行を推奨します。'
      console.log(msg);
    }
  }
  
  // 自動クリーンアップ用トリガー（週1回実行）AUTO_CLEANUP
  if (config.AUTO_CLEANUP || (config.AUTO_CLEANUP == true && config.AUTO_CLEANUP == 'TRUE')) {  
    ScriptApp.newTrigger('runAutoCleanup')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(2) // 日曜午前2時実行
      .create();
    msg = msg + '\n定期実行トリガー設定完了: 自動クリーンアップ 週1回（日曜午前2時）'
    console.log(msg);
  }
}

/**
 * 手動運用推奨の確認メッセージ
 */
function showManualOperationGuide() {
  console.log('');
  console.log('=== 手動運用ガイド ===');
  console.log('');
  console.log('📅 推奨実行スケジュール:');
  console.log('• 完全同期: 月1回（月初など）');
  console.log('• 増分同期: 週1-2回または必要時');
  console.log('• 統計確認: 同期実行後');
  console.log('• クリーンアップ: 月1回');
  console.log('');
  console.log('🔧 メニューからの実行:');
  console.log('1. MF-SF統合 > 完全同期実行');
  console.log('2. MF-SF統合 > 統計レポート');
  console.log('3. 必要に応じてクリーンアップ実行');
  console.log('');
  console.log('💡 自動実行を有効化したい場合:');
  console.log('Configシートで AUTO_SYNC_ENABLED を TRUE に設定');
  console.log('');
}

/**
 * 定期実行用トリガーを削除
 */
function removeScheduledTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runIncrementalSync' || 
        trigger.getHandlerFunction() === 'runAutoCleanup') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  
  console.log(`定期実行トリガー削除完了: ${removed}件`);
}
