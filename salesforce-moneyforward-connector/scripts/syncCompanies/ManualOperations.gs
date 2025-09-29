/**
 * 手動操作・段階的同期処理
 * 初動処理と日常運用の手動フローを管理
 */

// =============================================================================
// 初動処理（直近1年の商談企業）
// =============================================================================

/**
 * 初動対象企業の件数確認
 * 直近1年間で商談が作成された企業を対象
 */
function checkInitialSyncCandidates() {
  console.log('=== 初動同期対象企業の件数確認 ===');
  
  try {
    const sfAPI = getSalesforceAPI();
    
    // 1年前の日付を計算
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    const cutoffDate = oneYearAgo.toISOString();
    
    console.log(`対象期間: ${oneYearAgo.toLocaleDateString()} 以降`);
    
    // まず商談のあるAccountIdを取得
    const opportunityQuery = `
      SELECT AccountId, COUNT(Id) OpportunityCount
      FROM Opportunity 
      WHERE CreatedDate >= ${cutoffDate}
      AND AccountId != null
      GROUP BY AccountId
      ORDER BY COUNT(Id) DESC
    `;
    
    console.log('商談集計クエリ実行中...');
    const oppResult = sfAPI.query(opportunityQuery);
    
    if (!oppResult.success) {
      throw new Error(`商談クエリエラー: ${oppResult.error}`);
    }
    
    const opportunityData = oppResult.data.records;
    console.log(`商談のある企業数: ${opportunityData.length}件`);
    
    if (opportunityData.length === 0) {
      console.log('対象となる商談が見つかりませんでした');
      return {
        totalCandidates: 0,
        linkedCount: 0,
        unlinkedCount: 0,
        customIdCount: 0,
        noCustomIdCount: 0,
        unlinkedAccounts: []
      };
    }
    
    // 大量データのためバッチ処理
    const batchSize = 50; // URL長制限を考慮
    const allAccounts = [];
    const totalBatches = Math.ceil(opportunityData.length / batchSize);
    
    console.log(`バッチ処理開始: ${totalBatches}バッチで処理`);
    
    for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
      const startIndex = batchIndex * batchSize;
      const endIndex = Math.min(startIndex + batchSize, opportunityData.length);
      const batchData = opportunityData.slice(startIndex, endIndex);
      
      console.log(`バッチ ${batchIndex + 1}/${totalBatches}: ${batchData.length}件処理中...`);
      
      // AccountIdのリストを作成（IN句用）
      const accountIds = batchData.map(record => `'${record.AccountId}'`).join(',');
      
      // 該当するAccountの詳細情報を取得
      const accountQuery = `
        SELECT Id, Name, accountId__c, MoneyForward_Partner_ID__c
        FROM Account 
        WHERE Id IN (${accountIds})
      `;
      
      const accountResult = sfAPI.query(accountQuery);
      
      if (!accountResult.success) {
        console.error(`バッチ ${batchIndex + 1} エラー: ${accountResult.error}`);
        continue; // エラーがあっても他のバッチは継続
      }
      
      allAccounts.push(...accountResult.data.records);
      
      // レート制限対策
      if (batchIndex < totalBatches - 1) {
        Utilities.sleep(500); // 0.5秒待機
      }
    }
    
    console.log(`Account詳細取得完了: ${allAccounts.length}件`);
    
    console.log('=== 件数確認結果 ===');
    console.log(`総対象企业数: ${allAccounts.length}件`);
    
    // MF連携状況の分析
    let linkedCount = 0;
    let unlinkedCount = 0;
    let customIdCount = 0;
    let noCustomIdCount = 0;
    
    const unlinkedAccounts = [];
    
    // 商談数とAccount情報をマッピング（重複削除処理追加）
    const accountMap = {};
    const processedAccountIds = new Set(); // 重複チェック用

    allAccounts.forEach(account => {
      accountMap[account.Id] = account;
    });

    opportunityData.forEach(oppRecord => {
      const account = accountMap[oppRecord.AccountId];
      if (!account || processedAccountIds.has(oppRecord.AccountId)) {
        return; // Account情報がないか、既に処理済みの場合はスキップ
      }
      
      processedAccountIds.add(oppRecord.AccountId); // 処理済みとしてマーク
      
      
      if (account.MoneyForward_Partner_ID__c) {
        linkedCount++;
      } else {
        unlinkedCount++;
        unlinkedAccounts.push({
          accountId: account.Id,
          accountName: account.Name,
          customId: account.accountId__c,
          opportunityCount: oppRecord.OpportunityCount
        });
      }
      
      if (account.accountId__c) {
        customIdCount++;
      } else {
        noCustomIdCount++;
      }
    });
    
    console.log(`既にMF連携済み: ${linkedCount}件`);
    console.log(`MF未連携（作成対象）: ${unlinkedCount}件`);
    console.log(`カスタムID設定済み: ${customIdCount}件`);
    console.log(`カスタムID未設定: ${noCustomIdCount}件`);
    
    // 未連携企業の詳細（上位10件）
    console.log('=== 未連携企業詳細（商談数上位10件） ===');
    unlinkedAccounts
      .sort((a, b) => b.opportunityCount - a.opportunityCount)
      .slice(0, 10)
      .forEach((account, index) => {
        console.log(`${index + 1}. ${account.accountName}`);
        console.log(`   Account ID: ${account.accountId}`);
        console.log(`   カスタムID: ${account.customId || '未設定'}`);
        console.log(`   商談数: ${account.opportunityCount}件`);
      });
    
    return {
      totalCandidates: allAccounts.length,
      linkedCount: linkedCount,
      unlinkedCount: unlinkedCount,
      customIdCount: customIdCount,
      noCustomIdCount: noCustomIdCount,
      unlinkedAccounts: unlinkedAccounts
    };
    
  } catch (error) {
    console.error('初動同期対象確認エラー:', error);
    throw error;
  }
}

/**
 * 初動同期対象企業の詳細リスト取得
 */
function getInitialSyncAccountDetails() {
  console.log('=== 初動同期対象企業の詳細取得 ===');
  
  try {
    const sfAPI = getSalesforceAPI();
    
    // 1年前の日付を計算
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    const cutoffDate = oneYearAgo.toISOString();
    const subCompanies = subGetCompanies(cutoffDate)
    
    // 詳細情報を取得（未連携のみ）
    const query = `
      SELECT Id, Name, Type,
            accountId__c, MoneyForward_Partner_ID__c,
            (SELECT Id FROM Opportunities WHERE CreatedDate >= ${cutoffDate}) Opportunities
      FROM Account 
      WHERE Id IN ('${subCompanies.join("','")}')
      AND MoneyForward_Partner_ID__c = null
      ORDER BY Name
          `;

    // const query = `
    //   SELECT Id, Name, Type, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
    //         accountId__c, MoneyForward_Partner_ID__c,
    //         (SELECT Id FROM Opportunities WHERE CreatedDate >= ${cutoffDate}) Opportunities
    //   FROM Account 
    //   WHERE Id IN ('${subCompanies}')
    //   AND MoneyForward_Partner_ID__c = null
    //   ORDER BY Name
    //       `;

    const result = sfAPI.query(query);
    
    if (result.success) {
      const accounts = result.data.records;
      console.log(`取得した詳細アカウント数: ${accounts.length}件`);
      
      // 住所情報の完成度チェック
      // let completeAddressCount = 0;
      // let incompleteAddressCount = 0;
      
      // accounts.forEach(account => {
      //   const hasCompleteAddress = account.BillingPostalCode && 
      //                             account.BillingState && 
      //                             account.BillingCity;
      //   if (hasCompleteAddress) {
      //     completeAddressCount++;
      //   } else {
      //     incompleteAddressCount++;
      //   }
      // });
      
      // console.log('=== 住所データ完成度 ===');
      // console.log(`住所完成: ${completeAddressCount}件`);
      // console.log(`住所不完全: ${incompleteAddressCount}件`);
      
      return accounts;
      
    } else {
      throw new Error(`詳細取得エラー: ${result.error}`);
    }
    
  } catch (error) {
    console.error('詳細取得エラー:', error);
    throw error;
  }
}

function subGetCompanies(cutoffDate, limit = 50){
  console.log('=== サブ初動企業一覧取得 ===')
  try {
    const sfAPI = getSalesforceAPI();
    
    const query = `
      SELECT Opportunity.Account.Id
      FROM Opportunity
      WHERE CreatedDate >= ${cutoffDate}
      AND accountId__c != null
      AND Account.MoneyForward_Partner_ID__c = null
      AND Opportunity.AccountId != null
      ORDER BY CreatedDate DESC
    `;

    const result = sfAPI.query(query);
    if (result.success) {
      const accounts = result.data.records;
      // 重複削除してから件数制限
      const orgids = [...new Set(accounts.map(item => item.Account.Id))]
      console.log(`取得した残りアカウント数: ${orgids.length}件`);
      const ids = orgids.slice(0, limit);
      console.log(`取得した詳細アカウント数: ${ids.length}件`);
      console.log(ids);
      return ids;

    } else {
      throw new Error(`詳細取得エラー: ${result.error}`);
    } 
  } catch (error) {
    console.error('詳細取得エラー:', error);
    throw error;
  }
}

/**
 * 初動同期実行（バッチ処理）
 * 注意: 大量データのため段階的実行を推奨
 */
function runInitialSync(maxCount = 50) {
  console.log(`=== 初動同期実行（最大${maxCount}件） ===`);
  
  try {
    const accounts = getInitialSyncAccountDetails();
    const processTargets = accounts.slice(0, maxCount);
    
    console.log(`処理対象: ${processTargets.length}件`);
    
    let successCount = 0;
    let errorCount = 0;
    const errors = [];
    
    processTargets.forEach((account, index) => {
      try {
        console.log(`処理中 ${index + 1}/${processTargets.length}: ${account.Name}`);
        
        // MF取引先作成
        const mfPartner = createMFPartner(account);
        
        // 双方向リンク設定
        linkSFAccountToMFPartner(account.Id, mfPartner.id);
        
        successCount++;
        console.log(`✓ 完了: ${account.Name} (MF ID: ${mfPartner.id})`);
        
        // レート制限対策
        if (index % 10 === 9) {
          console.log('レート制限対策: 2秒待機');
          Utilities.sleep(2000);
        }
        
      } catch (error) {
        errorCount++;
        errors.push({
          accountName: account.Name,
          accountId: account.Id,
          error: error.message
        });
        console.log(`✗ エラー: ${account.Name} - ${error.message}`);
      }
    });
    
    console.log('=== 初動同期結果 ===');
    console.log(`成功: ${successCount}件`);
    console.log(`エラー: ${errorCount}件`);
    
    if (errors.length > 0) {
      console.log('=== エラー詳細 ===');
      errors.forEach(error => {
        console.log(`${error.accountName}: ${error.error}`);
      });
    }
    
    return {
      processed: processTargets.length,
      success: successCount,
      errors: errorCount,
      errorDetails: errors
    };
    
  } catch (error) {
    console.error('初動同期実行エラー:', error);
    throw error;
  }
}

// =============================================================================
// 日常運用（直近30日）
// =============================================================================

/**
 * 直近30日で動きがあった企業の件数確認
 */
function checkRecentActivityCandidates() {
  console.log('=== 直近30日活動企業の件数確認 ===');
  
  try {
    const sfAPI = getSalesforceAPI();
    
    // 30日前の日付を計算
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    const cutoffDate = thirtyDaysAgo.toISOString();
    
    console.log(`対象期間: ${thirtyDaysAgo.toLocaleDateString()} 以降`);
    
    // 複数の活動を統合して確認
    const queries = {
      // 商談活動
      opportunities: `
        SELECT COUNT(DISTINCT AccountId) AccountCount
        FROM Opportunity 
        WHERE (CreatedDate >= ${cutoffDate} OR LastModifiedDate >= ${cutoffDate})
        AND AccountId != null
      `,
      
      // Account更新
      accounts: `
        SELECT COUNT() AccountCount
        FROM Account 
        WHERE LastModifiedDate >= ${cutoffDate}
        AND MoneyForward_Partner_ID__c = null
      `,
      
      // Task/Event活動
      activities: `
        SELECT COUNT(DISTINCT WhatId) AccountCount
        FROM Task 
        WHERE CreatedDate >= ${cutoffDate}
        AND What.Type = 'Account'
      `
    };
    
    const results = {};
    
    for (const [key, query] of Object.entries(queries)) {
      try {
        const result = sfAPI.query(query);
        if (result.success && result.data.records.length > 0) {
          results[key] = result.data.records[0].AccountCount || 0;
        } else {
          results[key] = 0;
        }
      } catch (error) {
        console.log(`${key}クエリでエラー: ${error.message}`);
        results[key] = 0;
      }
    }
    
    console.log('=== 活動別件数 ===');
    console.log(`商談関連活動: ${results.opportunities}件`);
    console.log(`Account更新: ${results.accounts}件`);
    console.log(`タスク関連活動: ${results.activities}件`);
    
    return results;
    
  } catch (error) {
    console.error('直近活動確認エラー:', error);
    throw error;
  }
}

/**
 * 直近30日活動企業の詳細取得（MF未連携のみ）
 */
function getRecentActivityAccountDetails() {
  console.log('=== 直近30日活動企業の詳細取得 ===');
  
  try {
    const sfAPI = getSalesforceAPI();
    
    // 30日前の日付を計算
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    const cutoffDate = thirtyDaysAgo.toISOString();
    
    // 活動があった企業を取得（重複除去）
    const query = `
      SELECT Id, Name, Type, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
             accountId__c, MoneyForward_Partner_ID__c, LastModifiedDate
      FROM Account 
      WHERE (
        Id IN (SELECT Account.Id FROM Opportunity WHERE LastModifiedDate >= ${cutoffDate})
        OR LastModifiedDate >= ${cutoffDate}
      )
      AND MoneyForward_Partner_ID__c = null
      ORDER BY LastModifiedDate DESC
    `;
    
    const result = sfAPI.query(query);
    
    if (result.success) {
      const accounts = result.data.records;
      console.log(`対象企業数: ${accounts.length}件`);
      
      // 上位10件の詳細表示
      console.log('=== 最近の活動企業（上位10件） ===');
      accounts.slice(0, 10).forEach((account, index) => {
        console.log(`${index + 1}. ${account.Name}`);
        console.log(`   最終更新: ${new Date(account.LastModifiedDate).toLocaleDateString()}`);
        console.log(`   カスタムID: ${account.accountId__c || '未設定'}`);
      });
      
      return accounts;
      
    } else {
      throw new Error(`詳細取得エラー: ${result.error}`);
    }
    
  } catch (error) {
    console.error('直近活動詳細取得エラー:', error);
    throw error;
  }
}

// =============================================================================
// ユーティリティ関数
// =============================================================================

/**
 * 段階的実行のガイド表示
 */
function showManualSyncGuide() {
  console.log('');
  console.log('='.repeat(60));
  console.log('      手動同期実行ガイド');
  console.log('='.repeat(60));
  console.log('');
  
  console.log('📊 Step 1: 件数確認');
  console.log('初動対象: checkInitialSyncCandidates()');
  console.log('日常対象: checkRecentActivityCandidates()');
  console.log('');
  
  console.log('🔍 Step 2: 詳細確認');
  console.log('初動詳細: getInitialSyncAccountDetails()');
  console.log('日常詳細: getRecentActivityAccountDetails()');
  console.log('');
  
  console.log('🚀 Step 3: 実行');
  console.log('初動実行: runInitialSync(50)  // 最大50件ずつ');
  console.log('日常実行: メニューから実行予定');
  console.log('');
  
  console.log('⚠️ 重要事項:');
  console.log('• 初動処理は段階的実行を推奨（50件ずつ）');
  console.log('• レート制限に注意');
  console.log('• エラー発生時は個別確認');
  console.log('');
}
