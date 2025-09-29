
// 見積書ログのステータスを更新
function updateQuoteLogStatus(logSheet, rowIndex, status, quoteNumber, errorMessage) {
  try {
    const now = new Date();
    
    // I列: 発行ステータス
    logSheet.getRange(rowIndex, 9).setValue(status);
    
    // J列: 見積書番号
    if (quoteNumber) {
      logSheet.getRange(rowIndex, 10).setValue(quoteNumber);
    }
    
    // K列: 発行日時
    if (status === '完了') {
      logSheet.getRange(rowIndex, 11).setValue(Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss'));
    }
    
    // L列: エラーメッセージ
    if (errorMessage) {
      logSheet.getRange(rowIndex, 12).setValue(errorMessage);
    } else if (status === '完了') {
      logSheet.getRange(rowIndex, 12).setValue('');
    }
    
    // 色分け
    const range = logSheet.getRange(rowIndex, 1, 1, 12);
    switch (status) {
      case '処理中':
        range.setBackground('#fff2cc'); // 黄色
        break;
      case '完了':
        range.setBackground('#d9ead3'); // 緑色
        break;
      case 'エラー':
      case '一部エラー':
        range.setBackground('#f4cccc'); // 赤色
        break;
      default:
        range.setBackground('white');
    }
    
  } catch (error) {
    console.error('Quote log update error:', error);
  }
}

// Salesforceの見積書番号を更新
function updateSalesforceQuoteNumber(opportunityId, quoteNumber) {
  try {
    const accessToken = getSalesforceAccessToken();
    if (!accessToken) {
      return {
        success: false,
        error: 'Salesforce認証に失敗しました'
      };
    }
    
    const updateData = {
      'Quotenumber_a__c': quoteNumber // 実際のフィールド名に要変更
    };
    
    const url = `https://wed.my.salesforce.com/services/data/v58.0/sobjects/Opportunity/${opportunityId}`;
    
    const options = {
      'method': 'PATCH',
      'headers': {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(updateData),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 204) {
      console.log(`Salesforce quote updated: ${opportunityId} → ${quoteNumber}`);
      return { success: true };
    } else {
      const errorData = JSON.parse(response.getContentText());
      const errorMessage = errorData[0]?.message || 'Unknown error';
      return {
        success: false,
        error: errorMessage
      };
    }
    
  } catch (error) {
    console.error('Salesforce quote update error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}


// ===========================================
// 請求書発行処理
// ===========================================

// 選択行の請求書発行
function publishSelectedInvoices() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  
  if (!activeRange) {
    SpreadsheetApp.getUi().alert('エラー', '発行する行を選択してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  
  processInvoicePublish('selected', startRow, numRows);
}

// チェック済み行の請求書発行
function publishCheckedInvoices() {
  processInvoicePublish('checked');
}

// エラー行の再試行
function retryFailedInvoices() {
  processInvoicePublish('retry');
}

// メイン処理：請求書発行
function processInvoicePublish(mode, startRow = null, numRows = null) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('請求書ログ');
    
    if (!logSheet) {
      throw new Error('請求書ログシートが見つかりません');
    }
    
    // 認証確認
    const api = new MoneyForwardAPI();
    try {
      api.getValidAccessToken();
    } catch (error) {
      SpreadsheetApp.getUi().alert('エラー', '認証が必要です。メニューから認証を開始してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 会社ID確認
    const companyId = api.props.getProperty('MF_COMPANY_ID');
    if (!companyId) {
      SpreadsheetApp.getUi().alert('エラー', '会社が設定されていません。メニューから会社設定を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 対象行を特定
    const targetRows = getTargetRows(logSheet, mode, startRow, numRows);
    
    if (targetRows.length === 0) {
      SpreadsheetApp.getUi().alert('情報', '発行対象の行がありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 処理開始の確認
    const confirmResult = SpreadsheetApp.getUi().alert(
      '確認',
      `${targetRows.length}件の請求書を発行します。よろしいですか？`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (confirmResult !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    // 各行を処理
    for (let i = 0; i < targetRows.length; i++) {
      const rowIndex = targetRows[i];
      console.log(`Processing ${i + 1}/${targetRows.length} (row: ${rowIndex})`);
      
      try {
        // ステータスを「処理中」に更新
        updateInvoiceLogStatus(logSheet, rowIndex, '処理中', '', '');
        
        // アウトプットデータを取得
        const invoiceData = getInvoiceDataFromLog(logSheet, rowIndex);
        
        // マネーフォワード形式に変換
        const mfInvoiceData = convertToMFFormat(invoiceData);
        
        // 請求書作成
        const result = api.createInvoice(companyId, mfInvoiceData);
        
        if (result.success) {
          // Salesforceに請求書番号を更新
          const sfUpdateResult = updateSalesforceInvoiceNumber(invoiceData.opportunityId, result.invoiceNumber);
          
          if (sfUpdateResult.success) {
            updateInvoiceLogStatus(logSheet, rowIndex, '完了', result.invoiceNumber, '');
            successCount++;
            console.log(`Success: Invoice ${result.invoiceNumber}`);
          } else {
            updateInvoiceLogStatus(logSheet, rowIndex, '一部エラー', result.invoiceNumber, `SF更新失敗: ${sfUpdateResult.error}`);
            errorCount++;
            console.error(`SF update failed: ${sfUpdateResult.error}`);
          }
          
        } else {
          updateInvoiceLogStatus(logSheet, rowIndex, 'エラー', '', result.error);
          errorCount++;
          console.error(`Failed: ${result.error}`);
        }
        
      } catch (error) {
        updateInvoiceLogStatus(logSheet, rowIndex, 'エラー', '', error.message);
        errorCount++;
        console.error(`Row ${rowIndex} processing error:`, error);
      }
      
      // API制限対策（1秒待機）
      Utilities.sleep(1000);
    }
    
    // 完了通知
    const message = `請求書発行完了\n✅ 成功: ${successCount}件\n❌ エラー: ${errorCount}件`;
    SpreadsheetApp.getUi().alert('完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Invoice publish error:', error);
    SpreadsheetApp.getUi().alert('エラー', `処理中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 対象行を特定
function getTargetRows(logSheet, mode, startRow, numRows) {
  const data = logSheet.getDataRange().getValues();
  const targetRows = [];
  
  for (let i = 1; i < data.length; i++) { // ヘッダー行をスキップ
    const rowData = data[i];
    const checkboxValue = rowData[0]; // A列: チェックボックス
    const publishStatus = rowData[8]; // I列: 発行ステータス
    
    switch (mode) {
      case 'selected':
        if (i >= startRow - 1 && i < startRow - 1 + numRows) {
          if (publishStatus !== '完了') {
            targetRows.push(i + 1);
          }
        }
        break;
        
      case 'checked':
        if (checkboxValue === true && publishStatus !== '完了') {
          targetRows.push(i + 1);
        }
        break;
        
      case 'retry':
        if (publishStatus === 'エラー' || publishStatus === '一部エラー') {
          targetRows.push(i + 1);
        }
        break;
    }
  }
  
  return targetRows;
}

// ログからアウトプートデータを取得
function getInvoiceDataFromLog(logSheet, rowIndex) {
  const rowData = logSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
  
  return {
    executionDate: rowData[1], // B列: 実行日時
    opportunityUrl: rowData[2], // C列: 商談URL
    opportunityId: rowData[3], // D列: 商談ID
    accountName: rowData[4], // E列: 取引先名称  
    opportunityName: rowData[5], // F列: 件名
    outputUrl: rowData[6], // G列: アウトプットURL
    executionResult: rowData[7] // H列: 実行結果
  };
}

// アウトプットデータをマネーフォワード形式に変換
function convertToMFFormat(invoiceData) {
  // アウトプットURLからスプレッドシートデータを取得
  const outputSpreadsheet = getSpreadsheetFromUrl(invoiceData.outputUrl);
  const outputSheet = outputSpreadsheet.getSheets()[0];
  const data = outputSheet.getDataRange().getValues();
  
  // 会社情報行（2行目）
  const companyInfo = data[1];
  
  // 詳細行（3行目以降）
  const details = data.slice(2);
  
  // マネーフォワード形式のデータ作成
  const mfData = {
    billing: {
      department: companyInfo[2] || '', // 取引先名称
      title: companyInfo[3] || '', // 件名
      billing_date: companyInfo[4] || '', // 請求日
      due_date: companyInfo[5] || '', // お支払期限
      sales_date: companyInfo[7] || '', // 売上計上日
      memo: companyInfo[8] || '', // メモ
      document_name: '請求書',
      tags: companyInfo[9] || '', // タグ
      
      // 取引先情報
      partner_name: companyInfo[2] || '',
      partner_detail: companyInfo[13] || '', // 取引先敬称
      partner_zip_code: companyInfo[14] || '', // 郵便番号
      partner_prefecture: companyInfo[15] || '', // 都道府県
      partner_address1: companyInfo[16] || '', // 住所1
      partner_address2: companyInfo[17] || '', // 住所2
      
      // 明細
      items: []
    }
  };
  
  // 明細行を追加
  details.forEach(detail => {
    if (detail[29]) { // 品名がある場合
      mfData.billing.items.push({
        name: detail[29] || '', // 品名
        code: detail[30] || '', // 品目コード
        detail: detail[35] || '', // 詳細
        unit_price: detail[31] || 0, // 単価
        quantity: detail[32] || 1, // 数量
        unit: detail[33] || '個', // 単位
        price: detail[36] || 0, // 金額
        tax_rate: convertTaxRate(detail[37] || '10%') // 税率
      });
    }
  });
  
  return mfData;
}

// 税率文字列を数値に変換
function convertTaxRate(taxRateStr) {
  if (taxRateStr === '非課税') {
    return 0;
  }
  
  const match = taxRateStr.match(/(\d+)%/);
  return match ? parseInt(match[1]) : 10;
}

// URLからスプレッドシートを取得
function getSpreadsheetFromUrl(url) {
  try {
    const spreadsheetId = url.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(`スプレッドシートの取得に失敗しました: ${url}`);
  }
}

// Salesforceの請求書番号を更新
function updateSalesforceInvoiceNumber(opportunityId, invoiceNumber) {
  try {
    const accessToken = getSalesforceAccessToken();
    if (!accessToken) {
      return {
        success: false,
        error: 'Salesforce認証に失敗しました'
      };
    }
    
    const updateData = {
      'Invoicenumber_a__c': invoiceNumber // 実際のフィールド名に要変更
    };
    
    const url = `https://wed.my.salesforce.com/services/data/v58.0/sobjects/Opportunity/${opportunityId}`;
    
    const options = {
      'method': 'PATCH',
      'headers': {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(updateData),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 204) {
      console.log(`Salesforce updated: ${opportunityId} → ${invoiceNumber}`);
      return { success: true };
    } else {
      const errorData = JSON.parse(response.getContentText());
      const errorMessage = errorData[0]?.message || 'Unknown error';
      return {
        success: false,
        error: errorMessage
      };
    }
    
  } catch (error) {
    console.error('Salesforce update error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ログのステータスを更新
function updateInvoiceLogStatus(logSheet, rowIndex, status, invoiceNumber, errorMessage) {
  try {
    const now = new Date();
    
    // I列: 発行ステータス
    logSheet.getRange(rowIndex, 9).setValue(status);
    
    // J列: 請求書番号
    if (invoiceNumber) {
      logSheet.getRange(rowIndex, 10).setValue(invoiceNumber);
    }
    
    // K列: 発行日時
    if (status === '完了') {
      logSheet.getRange(rowIndex, 11).setValue(Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss'));
    }
    
    // L列: エラーメッセージ
    if (errorMessage) {
      logSheet.getRange(rowIndex, 12).setValue(errorMessage);
    } else if (status === '完了') {
      logSheet.getRange(rowIndex, 12).setValue('');
    }
    
    // 色分け
    const range = logSheet.getRange(rowIndex, 1, 1, 12);
    switch (status) {
      case '処理中':
        range.setBackground('#fff2cc'); // 黄色
        break;
      case '完了':
        range.setBackground('#d9ead3'); // 緑色
        break;
      case 'エラー':
      case '一部エラー':
        range.setBackground('#f4cccc'); // 赤色
        break;
      default:
        range.setBackground('white');
    }
    
  } catch (error) {
    console.error('Log update error:', error);
  }
}

// ===========================================
// 初期設定・セットアップ関数
// ===========================================

function setupInvoiceLogHeaders() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = null;
  
  try {
    logSheet = spreadsheet.getSheetByName('請求書ログ');
  } catch (e) {
    logSheet = spreadsheet.insertSheet('請求書ログ');
  }
  
  
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  logSheet.autoResizeColumns(1, headers.length);
  
  // A列をチェックボックス列として設定
  const dataRange = logSheet.getRange(2, 1, 100, 1); // 100行分のチェックボックス
  dataRange.insertCheckboxes();
  
  console.log('請求書ログヘッダーを設定しました');
}

// システム全体の初期化
function initializeSystem() {
  console.log('=== システム初期化開始 ===');
  
  try {
    // 1. 請求書ログ設定
    setupInvoiceLogHeaders();
    
    // 2. Salesforce認証情報確認
    const salesforceToken = PropertiesService.getScriptProperties().getProperty('SF_ACCESS_TOKEN');
    if (!salesforceToken) {
      console.log('⚠️ Salesforce認証が必要です');
    } else {
      console.log('✅ Salesforce認証済み');
    }
    
    console.log('=== システム初期化完了 ===');
    console.log('次のステップ:');
    console.log('1. メニューから「マネーフォワード」→「初期設定」を実行');
    console.log('2. スクリプトをWeb Appsとして公開');
    console.log('3. マネーフォワードAPIにリダイレクトURIを登録');
    console.log('4. 認証開始');
    
    SpreadsheetApp.getUi().alert(
      '初期化完了',
      'システムの初期化が完了しました。\n\n次にメニューから「マネーフォワード」→「初期設定」を実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('System initialization error:', error);
    SpreadsheetApp.getUi().alert('エラー', `初期化に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
