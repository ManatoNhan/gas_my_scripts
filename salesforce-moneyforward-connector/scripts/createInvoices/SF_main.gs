// ===========================================
// メイン処理：スプレッドシート統合システム
// ===========================================



// 見積書エクスポート
function exportQuote() {
  processExport('quote');
}

// 請求書エクスポート
function exportInvoice() {
  processExport('invoice');
}

// メイン処理関数
function processExport(documentType) {
  const docName = documentType === 'quote' ? '見積書' : '請求書';
  console.log(`=== ${docName}エクスポート開始 ===`);
  
  try {
    // 1. B3セルから商談URLを取得
    const sheet = SpreadsheetApp.getActiveSheet();
    const opportunityUrl = sheet.getRange('B3').getValue();
    
    if (!opportunityUrl) {
      SpreadsheetApp.getUi().alert('エラー', 'B3セルに商談URLを入力してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      logResult(documentType, '', '', '', '', '', 'エラー: 商談URLが入力されていません');
      return;
    }
    
    console.log('商談URL:', opportunityUrl);
    
    // 2. URLから商談IDを抽出
    const opportunityId = extractOpportunityIdFromUrl(opportunityUrl);
    if (!opportunityId) {
      SpreadsheetApp.getUi().alert('エラー', '有効な商談URLを入力してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      logResult(documentType, opportunityUrl, '', '', '', '', 'エラー: 無効な商談URL');
      return;
    }
    
    console.log('商談ID:', opportunityId);
    
    // 3. 商談IDから売上詳細を取得
    const itemDetails = getItemDetailsByOpportunityId(opportunityId);
    if (!itemDetails || itemDetails.length === 0) {
      SpreadsheetApp.getUi().alert('エラー', '指定された商談の売上詳細が見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      logResult(documentType, opportunityUrl, opportunityId, '', '', '', 'エラー: 売上詳細が見つかりません');
      return;
    }
    
    // 4. アウトプット用スプレッドシートを作成
    const outputSpreadsheetUrl = createOutputSpreadsheet(itemDetails, documentType);
    if (!outputSpreadsheetUrl) {
      SpreadsheetApp.getUi().alert('エラー', 'スプレッドシートの作成に失敗しました。', SpreadsheetApp.getUi().ButtonSet.OK);
      logResult(documentType, opportunityUrl, opportunityId, '', '', '', 'エラー: スプレッドシート作成失敗');
      return;
    }
    
    // 5. 指定フォルダに移動
    const finalUrl = moveToTargetFolder(outputSpreadsheetUrl, itemDetails[0], documentType);
    
    // 6. 結果をログに記録
    const opportunity = itemDetails[0].Opportunity__r || {};
    const account = opportunity.Account || {};
    logResult(
      documentType, 
      opportunityUrl, 
      opportunityId, 
      account.Name || '', 
      opportunity.Name || '', 
      finalUrl || outputSpreadsheetUrl, 
      '成功'
    );
    
    // 7. ユーザーに完了通知
    SpreadsheetApp.getUi().alert(
      '完了', 
      `${docName}の作成が完了しました。\n\nファイルURL:\n${finalUrl || outputSpreadsheetUrl}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log(`=== ${docName}エクスポート完了 ===`);
    
  } catch (error) {
    console.error(`${docName}エクスポートエラー:`, error);
    SpreadsheetApp.getUi().alert('エラー', `処理中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    logResult(documentType, '', '', '', '', '', `エラー: ${error.message}`);
  }
}

// URLから商談IDを抽出
function extractOpportunityIdFromUrl(url) {
  try {
    // Salesforce URL パターンをチェック
    const patterns = [
      /\/([a-zA-Z0-9]{15,18})\//, // 一般的なSalesforce ID
      /\/lightning\/r\/Opportunity\/([a-zA-Z0-9]{15,18})\//, // Lightning URL
      /opportunity_id=([a-zA-Z0-9]{15,18})/ // パラメータ形式
    ];
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match && match[1]) {
        return match[1];
      }
    }
    
    return null;
  } catch (error) {
    console.error('URL解析エラー:', error);
    return null;
  }
}

// 商談IDから売上詳細を取得
function getItemDetailsByOpportunityId(opportunityId) {
  const accessToken = getSalesforceAccessToken();
  if (!accessToken) {
    throw new Error('Salesforce認証に失敗しました');
  }
  
  const query = `
    SELECT Id, Name, PlanningQuantity__c, ResultQuantity__c, 
           UnitPrice__c, BillingAmount__c, AccountingAmount__c,
           Opportunity__r.Id, Opportunity__r.Name, Opportunity__r.Amount, 
           Opportunity__r.CloseDate, Opportunity__r.StageName,
           Opportunity__r.deliveryDate__c,
           Opportunity__r.Account.Name, Opportunity__r.Account.BillingStreet,
           Opportunity__r.Account.BillingCity, Opportunity__r.Account.BillingPostalCode,
           Opportunity__r.Account.BillingState,
           OpportunityLineItem__r.Name, OpportunityLineItem__r.Quantity,
           OpportunityLineItem__r.UnitPrice, OpportunityLineItem__r.TotalPrice,
           ProductAccountDetail__r.Name,
           Opportunity__r.Account.MoneyForward_Partner_ID__c
    FROM ItemAccountDetail__c 
    WHERE Opportunity__r.Id = '${opportunityId}'
    ORDER BY Name
  `;
  
  const apiUrl = `https://wed.my.salesforce.com/services/data/v58.0/query/?q=${encodeURIComponent(query)}`;
  
  const options = {
    'method': 'GET',
    'headers': {
      'Authorization': 'Bearer ' + accessToken
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Salesforce API エラー: ${result[0]?.message || 'Unknown error'}`);
    }
    
    console.log(`取得した売上詳細件数: ${result.totalSize}`);
    return result.records || [];
    
  } catch (error) {
    console.error('売上詳細取得エラー:', error);
    throw error;
  }
}

// アウトプット用スプレッドシートを作成
function createOutputSpreadsheet(itemDetails, documentType) {
  if (!itemDetails || itemDetails.length === 0) {
    return null;
  }
  
  const docName = documentType === 'quote' ? '見積書' : '請求書';
  const opportunity = itemDetails[0].Opportunity__r || {};
  const account = opportunity.Account || {};
  
  // ファイル名作成
  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'JST', 'yyyyMMdd');
  const fileName = `${docName}_${account.Name || '取引先不明'}_${opportunity.Name || '件名不明'}_${dateStr}`;
  
  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName(`${docName}_${dateStr}`);
  
  // ヘッダー設定
  const headers = [
    'csv_type(変更不可)', '行形式', '取引先名称', '件名', '請求日', 'お支払期限', '請求書番号',
    '売上計上日', 'メモ', 'タグ', '小計', '消費税', '合計金額', '取引先敬称', '取引先郵便番号',
    '取引先都道府県', '取引先住所1', '取引先住所2', '取引先部署', '取引先担当者役職', '取引先担当者氏名',
    '自社担当者氏名', '備考', '振込先', '入金ステータス', 'メール送信ステータス', '郵送ステータス',
    'ダウンロードステータス', '納品日', '品名', '品目コード', '単価', '数量', '単位', '納品書番号',
    '詳細', '金額', '品目消費税率'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // データ作成
  const invoiceDate = new Date();
  const paymentDueDate = getNextMonthLastDay(invoiceDate);
  
  // 合計計算
  let groupSubTotal = 0;
  let groupTax = 0;
  
  itemDetails.forEach(item => {
    const billingAmount = item.BillingAmount__c || 0;
    const itemName = item.Name || '';
    const isCBAmount = itemName.includes('CB金額');
    const tax = isCBAmount ? 0 : Math.round(billingAmount * 0.1);
    
    groupSubTotal += billingAmount;
    groupTax += tax;
  });
  
  const groupTotal = groupSubTotal + groupTax;
  
  // 1行目: 会社情報
  const companyInfoRow = Array(headers.length).fill('');
  companyInfoRow[1] = docName; // B列
  companyInfoRow[2] = account.Name || ''; // C列
  companyInfoRow[3] = opportunity.Name || ''; // D列
  companyInfoRow[4] = Utilities.formatDate(invoiceDate, 'JST', 'yyyy/MM/dd'); // E列
  companyInfoRow[5] = Utilities.formatDate(paymentDueDate, 'JST', 'yyyy/MM/dd'); // F列
  companyInfoRow[7] = opportunity.CloseDate || ''; // H列
  companyInfoRow[10] = groupSubTotal; // K列
  companyInfoRow[11] = groupTax; // L列
  companyInfoRow[12] = groupTotal; // M列
  companyInfoRow[13] = '御中'; // N列
  companyInfoRow[14] = account.BillingPostalCode || ''; // O列
  companyInfoRow[15] = account.BillingState || ''; // P列
  companyInfoRow[16] = account.BillingStreet || ''; // Q列
  companyInfoRow[24] = '未入金'; // Y列
  companyInfoRow[25] = '未送信'; // Z列
  companyInfoRow[26] = '未送信'; // AA列
  companyInfoRow[27] = '未ダウンロード'; // AB列
  
  sheet.getRange(2, 1, 1, headers.length).setValues([companyInfoRow]);
  
  // 2行目以降: 詳細行
  const detailRows = [];
  
  itemDetails.forEach(item => {
    const billingAmount = item.BillingAmount__c || 0;
    const itemName = item.Name || '';
    const isCBAmount = itemName.includes('CB金額');
    const taxRate = isCBAmount ? '非課税' : '10%';
    
    // 数量決定
    let quantity;
    if (documentType === 'quote') {
      quantity = item.PlanningQuantity__c || 1;
    } else {
      // 請求書：実績数量またはデフォルト0
      quantity = item.ResultQuantity__c ?? 0;
    }
    
    const detailRow = Array(headers.length).fill('');
    detailRow[1] = '品目'; // B列
    detailRow[28] = opportunity.deliveryDate__c || ''; // AC列（納品日）
    detailRow[29] = itemName; // AD列（品名）
    detailRow[31] = item.UnitPrice__c || 0; // AF列（単価）
    detailRow[32] = quantity; // AG列（数量）
    detailRow[33] = '個'; // AH列（単位）
    detailRow[36] = billingAmount; // AK列（金額）
    detailRow[37] = taxRate; // AL列（品目消費税率）
    
    detailRows.push(detailRow);
  });
  
  if (detailRows.length > 0) {
    sheet.getRange(3, 1, detailRows.length, headers.length).setValues(detailRows);
  }
  
  // フォーマット調整
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  console.log(`アウトプットスプレッドシート作成完了: ${fileName}`);
  return spreadsheet.getUrl();
}

// 指定フォルダにファイル移動
function moveToTargetFolder(spreadsheetUrl, firstItem, documentType) {
  try {
    const spreadsheetId = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
    const file = DriveApp.getFileById(spreadsheetId);
    
    // ターゲットフォルダ
    const targetFolderId = '1Wmg6XLDkXgEREi026pE17EI3BgEVy2-L';
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    
    // ファイルを移動
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      parent.removeFile(file);
    }
    targetFolder.addFile(file);
    
    console.log('ファイルを指定フォルダに移動しました');
    return spreadsheetUrl;
    
  } catch (error) {
    console.error('ファイル移動エラー:', error);
    return null;
  }
}

function getNextMonthLastDay(date) {
  const nextMonth = new Date(date.getFullYear(), date.getMonth() + 1, 1);
  const lastDayOfNextMonth = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0);
  return lastDayOfNextMonth;
}

// ログ記録
// ログ記録
function logResult(documentType, opportunityUrl, opportunityId, accountName, opportunityName, outputUrl, result) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const docName = documentType === 'quote' ? '見積書' : '請求書';
    const sheetName = `${docName}ログ`;
    
    let logSheet = null;
    try {
      logSheet = spreadsheet.getSheetByName(sheetName);
    } catch (e) {
      // シートが存在しない場合は作成
      logSheet = spreadsheet.insertSheet(sheetName);
      
      // ヘッダー行を設定
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // B列の最終行を取得
    const lastRowInB = logSheet.getRange('B:B').getLastRow();
    const nextRow = lastRowInB + 1;
    
    // ログデータを追加
    const now = new Date();
    const checkbx = false;
    const logData = [
      checkbx,
      Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss'),
      opportunityUrl,
      opportunityId,
      accountName,
      opportunityName,
      outputUrl,
      result
    ];
    
    // B列を基準とした最終行の次の行にデータを挿入
    logSheet.getRange(nextRow, 1, 1, logData.length).setValues([logData]);
    
    console.log(`${docName}ログに記録しました`);
    
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}

// GAS駆動用スプレッドシートを作成する関数
function createGASDriverSpreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // メインシート設定
  const mainSheet = spreadsheet.getActiveSheet();
  mainSheet.setName('メイン');
  
  // 説明とフォーマット
  mainSheet.getRange('A1').setValue('Salesforce 請求書・見積書 エクスポートシステム');
  mainSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  
  mainSheet.getRange('A3').setValue('商談URL:');
  mainSheet.getRange('A3').setFontWeight('bold');
  mainSheet.getRange('B3').setValue('ここに商談URLを貼り付けてください');
  
  mainSheet.getRange('A5').setValue('使用方法:');
  mainSheet.getRange('A5').setFontWeight('bold');
  mainSheet.getRange('A6').setValue('1. B3セルに商談URLを貼り付け');
  mainSheet.getRange('A7').setValue('2. メニューバーの「Export」から「見積書」または「請求書」を選択');
  mainSheet.getRange('A8').setValue('3. 処理完了後、結果URLが表示されます');
  
  // 列幅調整
  mainSheet.setColumnWidth(1, 100);
  mainSheet.setColumnWidth(2, 400);
  
  console.log('GAS駆動用スプレッドシートを作成しました');
  console.log('URL:', spreadsheet.getUrl());
  
  return spreadsheet.getUrl();
}
