// ===================================
// Phase 1: 基本発行機能実装
// ===================================

// ===================================
// 1. 取引先ID紐付け機能
// ===================================

/**
 * SalesforceのAccount IDからMoneyForward Partner IDを取得
 * @param {string} accountId - Salesforce Account ID
 * @param {string} accountName - 取引先名（エラー表示用）
 * @return {string} MoneyForward Partner ID
 */
function getPartnerIdFromAccount(accountId, accountName) {
  try {
    const accessToken = getSalesforceAccessToken();
    const instanceUrl = PropertiesService.getScriptProperties().getProperty('SF_INSTANCE_URL');
    
    const query = `SELECT Id, Name, MoneyForward_Partner_ID__c FROM Account WHERE Id = '${accountId}'`;
    const encodedQuery = encodeURIComponent(query);
    const url = `${instanceUrl}/services/data/v58.0/query?q=${encodedQuery}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.totalSize === 0) {
      throw new Error(`取引先が見つかりません: ${accountName} (ID: ${accountId})`);
    }
    
    const account = result.records[0];
    
    if (!account.MoneyForward_Partner_ID__c) {
      throw new Error(`取引先が同期されていません: ${accountName} (ID: ${accountId})`);
    }
    
    return account.MoneyForward_Partner_ID__c;
    
  } catch (error) {
    console.error('Partner ID取得エラー:', error);
    throw error;
  }
}

/**
 * 取引先紐付けの検証
 * @param {string} accountId - Salesforce Account ID
 * @param {string} accountName - 取引先名
 * @return {Object} {valid: boolean, partnerId: string, error: string}
 */
function validatePartnerLink(accountId, accountName) {
  try {
    const partnerId = getPartnerIdFromAccount(accountId, accountName);
    return {
      valid: true,
      partnerId: partnerId,
      error: null
    };
  } catch (error) {
    return {
      valid: false,
      partnerId: null,
      error: error.message
    };
  }
}

// ===================================
// 2. アウトプットデータ変換機能
// ===================================

/**
 * スプレッドシートURLからデータを取得
 * @param {string} url - スプレッドシートURL
 * @return {Array[]} スプレッドシートデータ
 */
function getSpreadsheetDataFromUrl(url) {
  try {
    const spreadsheetId = url.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheets()[0]; // 最初のシートを使用
    return sheet.getDataRange().getValues();
  } catch (error) {
    throw new Error(`スプレッドシートの取得に失敗しました: ${error.message}`);
  }
}

/**
 * 税率を税区分（excise）に変換
 * @param {string|number} taxRateValue - 税率値
 * @return {string} 税区分文字列
 */
function convertTaxRateToExcise(taxRateValue) {
  // 文字列から数値を抽出
  const taxRate = typeof taxRateValue === 'string' 
    ? parseFloat(taxRateValue.replace(/[^\d.]/g, '')) 
    : taxRateValue;
  
  // 軽減税率の判定
  const isReduced = String(taxRateValue).includes('軽減');
  
  switch (taxRate) {
    case 10:
      return 'ten_percent';
    case 8:
      return isReduced ? 'eight_percent_as_reduced_tax_rate' : 'eight_percent';
    case 5:
      return 'five_percent';
    case 0:
      return 'non_taxable';
    default:
      return 'ten_percent'; // デフォルトは10%
  }
}

/**
 * パートナーIDから部門IDを取得
 * @param {string} partnerId - MoneyForward Partner ID
 * @return {string} 部門ID
 */
function getDepartmentIdFromPartnerId(partnerId) {
  try {
    const mfApi = new MoneyForwardAPI();
    const service = mfApi.getOAuthService();
    
    const response = UrlFetchApp.fetch(`${mfApi.apiBaseUrl}/partners/${partnerId}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${service.getAccessToken()}`,
        'Accept': 'application/json'
      }
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.departments && data.departments.length > 0) {
      // 通常は最初の部門を使用
      return data.departments[0].id;
    }
    
    throw new Error('部門情報が見つかりません');
    
  } catch (error) {
    console.error('部門ID取得エラー:', error);
    throw error;
  }
}

/**
 * 日付形式を変換（YYYY-MM-DD形式に統一）
 * @param {any} dateValue - 日付値
 * @return {string} YYYY-MM-DD形式の日付文字列
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  let date;
  
  if (dateValue instanceof Date) {
    date = dateValue;
  } else if (typeof dateValue === 'string') {
    date = new Date(dateValue);
  } else {
    return '';
  }
  
  if (isNaN(date.getTime())) return '';
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * アウトプットデータをMoneyForward形式に変換
 * @param {Object} quoteData - 請求書ログデータ
 * @return {Object} MoneyForward API形式のデータ
 */
function convertToMFQuoteFormat(quoteData) {
  try {
    // スプレッドシートデータ取得
    const data = getSpreadsheetDataFromUrl(quoteData.outputUrl);
    
    if (data.length < 3) {
      throw new Error('アウトプットデータの形式が不正です');
    }
    
    // 会社情報（2行目）
    const companyInfo = data[1];
    
    // 取引先ID取得（商談情報から）
    const partnerId = quoteData.partnerId || getPartnerIdFromAccount(quoteData.accountId, quoteData.accountName);
    
    // 部門ID取得
    const departmentId = getDepartmentIdFromPartnerId(partnerId);
    
    // 明細データ（3行目以降）
    const details = data.slice(2).filter(row => row[29]); // 品名がある行のみ
    
    if (details.length === 0) {
      throw new Error('明細データがありません');
    }
    
    // MoneyForward形式に変換
    const mfData = {
      department_id: departmentId,
      title: companyInfo[3] || quoteData.subject || '見積書',
      quote_date: formatDate(companyInfo[4]) || formatDate(new Date()),
      expired_date: formatDate(companyInfo[5]) || '',
      // sales_date: formatDate(companyInfo[7]) || formatDate(companyInfo[4]) || formatDate(new Date()),
      memo: companyInfo[6] || '',
      items: details.map((detail, index) => ({
        name: detail[29] || `明細${index + 1}`,  // 商品名をそのまま使用
        detail: detail[35] || detail[34] || '', // 詳細または備考
        price: parseFloat(detail[31]) || 0,  // unit_price → price
        quantity: parseFloat(detail[32]) || 1,
        unit: detail[33] || '',
        excise: convertTaxRateToExcise(detail[37]),  // 税率を税区分に変換
        is_deduct_withholding_tax: false
      }))
    };
    
    // 必須項目の検証
    validateMFQuoteData(mfData);
    
    return mfData;
    
  } catch (error) {
    console.error('データ変換エラー:', error);
    throw error;
  }
}

/**
 * MoneyForwardデータの検証
 * @param {Object} mfData - MoneyForward形式のデータ
 */
function validateMFQuoteData(mfData) {
  const errors = [];
  
  if (!mfData.department_id) {
    errors.push('部門IDが設定されていません');
  }
  
  if (!mfData.quote_date) {
    errors.push('見積日が設定されていません');
  }
  
  if (!mfData.expired_date) {
    errors.push('有効期日が設定されていません');
  }
  
  if (!mfData.items || mfData.items.length === 0) {
    errors.push('明細データがありません');
  }
  
  mfData.items.forEach((item, index) => {
    if (!item.name) {
      errors.push(`明細${index + 1}の品名が設定されていません`);
    }
    if (item.price < 0) {
      errors.push(`明細${index + 1}の単価が不正です`);
    }
    if (!item.excise) {
      errors.push(`明細${index + 1}の税区分が設定されていません`);
    }
  });
  
  if (errors.length > 0) {
    throw new Error('データ検証エラー:\n' + errors.join('\n'));
  }
}

// ===================================
// 3. 見積書発行処理
// ===================================

/**
 * MoneyForward請求書を作成
 * @param {Object} mfData - MoneyForward形式のデータ
 * @return {Object} {success: boolean, invoiceNumber: string, invoiceId: string, error: string}
 */
function createQuoteWithPartner(mfData) {
  try {
    const mfApi = new MoneyForwardAPI();
    const service = mfApi.getOAuthService();
    
    if (!service.hasAccess()) {
      throw new Error('MoneyForward認証が必要です');
    }
    
    const response = UrlFetchApp.fetch(`${mfApi.apiBaseUrl}/quotes`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${service.getAccessToken()}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(mfData)
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 201 || response.getResponseCode() === 200) {
      return {
        success: true,
        quoteNumber: result.quote_number,
        quoteId: result.id,
        error: null
      };
    } else {
      throw new Error(`API エラー: ${result.error || 'Unknown error'}`);
    }
    
  } catch (error) {
    console.error('見積書作成エラー:', error);
    return {
      success: false,
      quoteNumber: null,
      quoteId: null,
      error: error.message
    };
  }
}


/**
 * ログからデータを取得
 * @param {number} rowIndex - 行番号（0ベース）
 * @return {Object} 請求書データ
 */
function getInvoiceDataFromLog(sheet, rowIndex) {
  const row = sheet.getRange(rowIndex + 1, 1, 1, 12).getValues()[0];
  
  return {
    checked: row[0],
    executionDate: row[1],
    opportunityUrl: row[2],
    opportunityId: row[3],
    accountName: row[4],
    subject: row[5],
    outputUrl: row[6],
    result: row[7],
    status: row[8],
    invoiceNumber: row[9],
    publishDate: row[10],
    errorMessage: row[11]
  };
}

/**
 * Salesforceから商談情報を取得
 * @param {string} opportunityId - 商談ID
 * @return {Object} 商談情報
 */
function getOpportunityInfo(opportunityId) {
  try {
    const accessToken = getSalesforceAccessToken();
    const instanceUrl = PropertiesService.getScriptProperties().getProperty('SF_INSTANCE_URL');
    
    const query = `SELECT Id, Name, AccountId, Account.Name, Account.MoneyForward_Partner_ID__c FROM Opportunity WHERE Id = '${opportunityId}'`;
    const encodedQuery = encodeURIComponent(query);
    const url = `${instanceUrl}/services/data/v58.0/query?q=${encodedQuery}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.totalSize === 0) {
      throw new Error(`商談が見つかりません: ${opportunityId}`);
    }
    
    return result.records[0];
    
  } catch (error) {
    console.error('商談情報取得エラー:', error);
    throw error;
  }
}

/**
 * ログのステータスを更新
 * @param {Sheet} sheet - 対象シート
 * @param {number} rowIndex - 行番号（0ベース）
 * @param {string} status - ステータス
 * @param {string} invoiceNumber - 請求書番号
 * @param {string} errorMessage - エラーメッセージ
 */
function updateLogStatus(sheet, rowIndex, status, invoiceNumber = '', errorMessage = '') {
  const row = rowIndex + 1;
  
  // ステータス更新
  sheet.getRange(row, 9).setValue(status);
  
  // 請求書番号
  if (invoiceNumber) {
    sheet.getRange(row, 10).setValue(invoiceNumber);
  }
  
  // 発行日時
  if (status === '完了') {
    sheet.getRange(row, 11).setValue(new Date());
    sheet.getRange(row, 8).setValue('成功');
  }
  
  // エラーメッセージ
  if (errorMessage) {
    sheet.getRange(row, 12).setValue(errorMessage);
    sheet.getRange(row, 8).setValue('失敗');
  }
  
  // 色分け
  const statusCell = sheet.getRange(row, 9);
  switch (status) {
    case '処理中':
      statusCell.setBackground('#fff3cd'); // 黄色
      break;
    case '完了':
      statusCell.setBackground('#d4edda'); // 緑色
      break;
    case 'エラー':
      statusCell.setBackground('#f8d7da'); // 赤色
      break;
  }
  
  SpreadsheetApp.flush();
}

/**
 * 選択された行を取得
 * @return {Array} 選択された行のインデックス配列
 */
function getSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('見積書ログ');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const selectedRows = [];
  
  // ヘッダー行をスキップして、チェックされた行を探す
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === true || values[i][0] === 'TRUE') {
      selectedRows.push(i);
    }
  }
  
  return selectedRows;
}

