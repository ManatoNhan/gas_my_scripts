/**
 * 取引先データ操作関数
 * MoneyForward取引先とSalesforce Accountの連携処理
 */

// =============================================================================
// MoneyForward取引先操作
// =============================================================================

/**
 * 全MoneyForward取引先を取得（ページネーション対応）
 */
function getAllMFPartners() {
  const mfAPI = getMoneyForwardAPI();
  const allPartners = [];
  let page = 1;
  const perPage = 100; // API制限内の最大値
  
  console.log('MoneyForward取引先取得開始...');
  
  while (true) {
    console.log(`ページ ${page} を取得中...`);
    
    const result = mfAPI.get(`/partners?page=${page}&per_page=${perPage}`);
    
    if (!result.success) {
      writeErrorLog('MF_API', `PAGE_${page}`, 'MF取引先取得エラー', result.error);
      throw new Error(`MoneyForward取引先取得エラー (ページ${page}): ${result.error}`);
    }
    
    const partners = result.data.data || [];
    allPartners.push(...partners);
    
    console.log(`ページ ${page}: ${partners.length}件取得`);
    
    // ページネーション終了判定
    const pagination = result.data.pagination;
    if (!pagination || page >= pagination.total_pages) {
      break;
    }
    
    page++;
    Utilities.sleep(100); // レート制限対策
  }
  
  console.log(`MoneyForward取引先取得完了: 合計 ${allPartners.length}件`);
  return allPartners;
}

/**
 * MoneyForward取引先を顧客コードで検索
 */
function findMFPartnerByCustomerCode(customerCode) {
  if (!customerCode) return null;
  
  try {
    // 顧客コードでの検索は全件取得して絞り込み
    // （MF APIに顧客コード検索機能がないため）
    const allPartners = getAllMFPartners();
    
    return allPartners.find(partner => 
      partner.code === customerCode
    ) || null;
    
  } catch (error) {
    console.error('MF取引先検索エラー:', error);
    return null;
  }
}

/**
 * MoneyForward取引先を作成
 */
function createMFPartner(accountData) {
  const mfAPI = getMoneyForwardAPI();
  
  const partnerData = {
    name: accountData.Name,
    code: generateCustomerCode(accountData.Id),
    memo: generateCustomCustomerCode(accountData.accountId__c),
    name_suffix: "御中"
    // zip: accountData.BillingPostalCode || '',
    // prefecture: extractPrefecture(accountData.BillingState || ''),
    // address1: `${accountData.BillingCity || ''} ${accountData.BillingStreet || ''}`.trim(),
    // phone: accountData.Phone || '',
    // email: getAccountEmail(accountData.Id) || ''
  };
  
  console.log(`MF取引先作成: ${partnerData.name} (${partnerData.code})`);
  
  const result = mfAPI.post('/partners', partnerData);
  
  if (result.success) {
    console.log(`✓ MF取引先作成成功: ID ${result.data.id}`);
    return result.data;
  } else {
    writeErrorLog('MF_CREATE', accountData.Id, 'MF取引先作成エラー', result.error);
    throw new Error(`MoneyForward取引先作成エラー: ${result.error}`);
  }
}

/**
 * MoneyForward取引先を更新
 */
function updateMFPartner(partnerId, accountData) {
  const mfAPI = getMoneyForwardAPI();
  
  const partnerData = {
    name: accountData.Name,
    // zip: accountData.BillingPostalCode || '',
    // prefecture: extractPrefecture(accountData.BillingState || ''),
    // address1: `${accountData.BillingCity || ''} ${accountData.BillingStreet || ''}`.trim(),
    // phone: accountData.Phone || '',
    // email: getAccountEmail(accountData.Id) || ''
  };
  
  console.log(`MF取引先更新: ID ${partnerId}`);
  
  const result = mfAPI.put(`/partners/${partnerId}`, partnerData);
  
  if (result.success) {
    console.log(`✓ MF取引先更新成功: ID ${partnerId}`);
    return result.data;
  } else {
    writeErrorLog('MF_UPDATE', partnerId, 'MF取引先更新エラー', result.error);
    throw new Error(`MoneyForward取引先更新エラー: ${result.error}`);
  }
}

// =============================================================================
// Salesforce Account操作
// =============================================================================

/**
 * Salesforce Accountを取得
 */
function getSFAccount(accountId) {
  const sfAPI = getSalesforceAPI();
  
  const query = `SELECT Id, Name, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
                 MoneyForward_Partner_ID__c, MF_Last_Sync_Date__c, MF_Address_Hash__c, accountId__c 
                 FROM Account WHERE Id = '${accountId}'`;
  
  const result = sfAPI.query(query);
  
  if (result.success && result.data.records.length > 0) {
    return result.data.records[0];
  }
  
  return null;
}

/**
 * 全Salesforce Accountを取得（MF連携対象のみ）
 */
function getAllSFAccounts() {
  const sfAPI = getSalesforceAPI();
  
  const query = `SELECT Id, Name, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
                 MoneyForward_Partner_ID__c, MF_Last_Sync_Date__c, MF_Address_Hash__c, accountId__c 
                 FROM Account 
                 WHERE MoneyForward_Partner_ID__c != null`;
  
  const result = sfAPI.query(query);
  
  if (result.success) {
    return result.data.records || [];
  } else {
    throw new Error(`Salesforce Account取得エラー: ${result.error}`);
  }
}

/**
 * 最近変更されたSalesforce Accountを取得
 */
function getRecentSFAccounts(hoursAgo = 24) {
  const sfAPI = getSalesforceAPI();
  
  const cutoffDate = new Date(Date.now() - hoursAgo * 60 * 60 * 1000);
  const isoDate = cutoffDate.toISOString();
  
  const query = `SELECT Id, Name, Phone, BillingPostalCode, BillingState, BillingCity, BillingStreet, 
                 MoneyForward_Partner_ID__c, MF_Last_Sync_Date__c, MF_Address_Hash__c, accountId__c 
                 FROM Account 
                 WHERE LastModifiedDate > ${isoDate}
                   AND MF_Last_Sync_Date__c <= ${isoDate}`;
  
  const result = sfAPI.query(query);
  
  if (result.success) {
    return result.data.records || [];
  } else {
    throw new Error(`最近変更されたSalesforce Account取得エラー: ${result.error}`);
  }
}

/**
 * Salesforce Accountを更新
 */
function updateSFAccount(accountId, updateData) {
  console.log(`SF Account更新開始: ${accountId}`);
  console.log(`更新データ:`, JSON.stringify(updateData, null, 2));
  
  const result = callSFAPI(`/sobjects/Account/${accountId}`, 'PATCH', updateData);
  
  if (result.success) {
    console.log(`✓ SF Account更新成功: ${accountId}`);
    return true;
  } else {
    console.error(`SF Account更新エラー: ${accountId}`, result.error);
    writeErrorLog('SF_UPDATE', accountId, 'SF Account更新エラー', result.error);
    throw new Error(`Salesforce Account更新エラー: ${result.error}`);
  }
}

/**
 * Salesforce AccountにMF取引先IDを設定
 */
function linkSFAccountToMFPartner(accountId, mfPartnerId) {
  const updateData = {
    MoneyForward_Partner_ID__c: mfPartnerId.toString(),
    MF_Last_Sync_Date__c: new Date().toISOString()
  };
  
  return updateSFAccount(accountId, updateData);
}

// =============================================================================
// データ整合性チェック
// =============================================================================

/**
 * データ整合性チェック実行
 */
function performDataIntegrityCheck(mfPartners) {
  let processed = 0;
  let success = 0;
  let errors = 0;
  
  console.log('=== データ整合性チェック開始 ===');
  
  // 削除候補シートをクリア
  clearDeleteCandidatesSheet();
  
  for (const partner of mfPartners) {
    processed++;
    
    try {
      const customerCode = partner.code;
      
      // 1. 顧客コードの妥当性チェック
      if (!customerCode) {
        // 顧客コードが空の場合は削除候補
        addDeleteCandidate(partner, 'NO_CUSTOMER_CODE', '顧客コードが設定されていません');
        continue;
      }
      
      if (!isValidSalesforceLink(customerCode)) {
        // SF:形式でない場合は削除候補
        addDeleteCandidate(partner, 'INVALID_FORMAT', `無効な顧客コード形式: ${customerCode}`);
        continue;
      }
      
      // 2. Salesforce側の存在確認
      const sfAccountId = extractSalesforceId(customerCode);
      const sfAccount = getSFAccount(sfAccountId);
      
      if (!sfAccount) {
        // SF Accountが見つからない場合は削除候補
        addDeleteCandidate(partner, 'SF_NOT_FOUND', `対応するSF Account(${sfAccountId})が見つかりません`);
        continue;
      }
      
      // 3. 双方向リンクの整合性チェック
      if (sfAccount.MoneyForward_Partner_ID__c != partner.id) {
        // SF側のMF IDが不一致の場合は修正
        console.log(`双方向リンク修正: SF ${sfAccountId} -> MF ${partner.id}`);
        linkSFAccountToMFPartner(sfAccountId, partner.id);
      }
      
      success++;
      
    } catch (error) {
      errors++;
      writeErrorLog('INTEGRITY_CHECK', partner.id, '整合性チェックエラー', error.message);
      console.error(`整合性チェックエラー (Partner ID: ${partner.id}):`, error);
    }
    
    // 進捗表示
    if (processed % 50 === 0) {
      console.log(`整合性チェック進捗: ${processed}/${mfPartners.length}`);
    }
  }
  
  console.log('=== データ整合性チェック完了 ===');
  console.log(`処理: ${processed}, 成功: ${success}, エラー: ${errors}`);
  
  return { processed, success, errors };
}

/**
 * Salesforce変更の同期
 */
function syncSalesforceChanges() {
  let processed = 0;
  let success = 0;
  let errors = 0;
  
  console.log('=== Salesforce変更同期開始 ===');
  
  try {
    const sfAccounts = getAllSFAccounts();
    
    for (const account of sfAccounts) {
      processed++;
      
      try {
        // 住所変更検知
        const currentAddressHash = generateAddressHash(
          `${account.BillingPostalCode || ''}${account.BillingState || ''}${account.BillingCity || ''}${account.BillingStreet || ''}`
        );
        
        if (account.MF_Address_Hash__c !== currentAddressHash) {
          console.log(`住所変更検知: SF Account ${account.Id}`);
          
          // MF取引先を更新
          if (account.MoneyForward_Partner_ID__c) {
            updateMFPartner(account.MoneyForward_Partner_ID__c, account);
            
            // SF側のハッシュも更新
            updateSFAccount(account.Id, {
              MF_Address_Hash__c: currentAddressHash,
              MF_Last_Sync_Date__c: new Date().toISOString()
            });
          }
        }
        
        success++;
        
      } catch (error) {
        errors++;
        writeErrorLog('SF_SYNC', account.Id, 'SF変更同期エラー', error.message);
        console.error(`SF変更同期エラー (Account ID: ${account.Id}):`, error);
      }
    }
    
  } catch (error) {
    errors++;
    writeErrorLog('SF_SYNC', 'SYSTEM', 'SF変更同期システムエラー', error.message);
    throw error;
  }
  
  console.log('=== Salesforce変更同期完了 ===');
  console.log(`処理: ${processed}, 成功: ${success}, エラー: ${errors}`);
  
  return { processed, success, errors };
}

/**
 * 最近のSalesforce変更の同期
 */
function syncRecentSalesforceChanges(hoursAgo = 24) {
  let processed = 0;
  let success = 0;
  let errors = 0;
  
  console.log(`=== 最近${hoursAgo}時間のSF変更同期開始 ===`);
  
  try {
    const recentAccounts = getRecentSFAccounts(hoursAgo);
    console.log(`最近変更されたAccount: ${recentAccounts.length}件`);
    
    for (const account of recentAccounts) {
      processed++;
      
      try {
        // MF取引先IDが設定されている場合のみ同期
        if (account.MoneyForward_Partner_ID__c) {
          updateMFPartner(account.MoneyForward_Partner_ID__c, account);
          
          // 住所ハッシュ更新
          const currentAddressHash = generateAddressHash(
            `${account.BillingPostalCode || ''}${account.BillingState || ''}${account.BillingCity || ''}${account.BillingStreet || ''}`
          );
          
          updateSFAccount(account.Id, {
            MF_Address_Hash__c: currentAddressHash,
            MF_Last_Sync_Date__c: new Date().toISOString()
          });
        }
        
        success++;
        
      } catch (error) {
        errors++;
        writeErrorLog('RECENT_SYNC', account.Id, '最近変更同期エラー', error.message);
        console.error(`最近変更同期エラー (Account ID: ${account.Id}):`, error);
      }
    }
    
  } catch (error) {
    errors++;
    writeErrorLog('RECENT_SYNC', 'SYSTEM', '最近変更同期システムエラー', error.message);
    throw error;
  }
  
  console.log('=== 最近変更同期完了 ===');
  console.log(`処理: ${processed}, 成功: ${success}, エラー: ${errors}`);
  
  return { processed, success, errors };
}

// =============================================================================
// ユーティリティ関数
// =============================================================================

/**
 * 都道府県を抽出（住所から）
 */
function extractPrefecture(address) {
  if (!address) return '';
  
  const prefectures = [
    '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
    '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
    '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県',
    '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県',
    '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県',
    '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県',
    '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
  ];
  
  for (const pref of prefectures) {
    if (address.includes(pref)) {
      return pref;
    }
  }
  
  return '';
}

/**
 * AccountのEmailを取得（ContactまたはAccountから）
 */
function getAccountEmail(accountId) {
  const sfAPI = getSalesforceAPI();
  
  try {
    // 関連ContactのEmailをチェック（Accountの標準Emailフィールドは使用しない）
    const contactQuery = `SELECT Email FROM Contact WHERE AccountId = '${accountId}' AND Email != null LIMIT 1`;
    const contactResult = sfAPI.query(contactQuery);
    
    if (contactResult.success && contactResult.data.records.length > 0) {
      return contactResult.data.records[0].Email;
    }
    
  } catch (error) {
    console.error('Email取得エラー:', error);
  }
  
  return '';
}
