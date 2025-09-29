// ===================================
// Phase 2: 運用性向上
// ===================================

// ===================================
// 4. UI・操作性改善
// ===================================

/**
 * スプレッドシートを開いた時の処理
 */


/**
 * 請求書発行ダイアログを表示
 */
function showInvoicePublishDialog() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>📄 請求書発行</h2>
      <p>選択された請求書を発行します。</p>
      
      <div id="status" style="margin: 20px 0;">
        <p id="statusText">準備完了</p>
        <div style="width: 100%; background-color: #f0f0f0; border-radius: 5px;">
          <div id="progressBar" style="width: 0%; height: 20px; background-color: #4CAF50; border-radius: 5px; transition: width 0.3s;"></div>
        </div>
        <p id="progressText" style="font-size: 12px; color: #666; margin-top: 5px;">0 / 0</p>
      </div>
      
      <div style="margin-top: 20px;">
        <button onclick="startPublish()" style="padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer;">
          発行開始
        </button>
        <button onclick="google.script.host.close()" style="padding: 10px 20px; margin-left: 10px; cursor: pointer;">
          キャンセル
        </button>
      </div>
      
      <div id="result" style="margin-top: 20px; display: none;">
        <h3>処理結果</h3>
        <p id="resultText"></p>
      </div>
    </div>
    
    <script>
      function startPublish() {
        document.getElementById('statusText').textContent = '処理中...';
        google.script.run
          .withSuccessHandler(onSuccess)
          .withFailureHandler(onFailure)
          .publishSelectedInvoicesWithProgress();
      }
      
      function updateProgress(current, total) {
        const percentage = Math.round((current / total) * 100);
        document.getElementById('progressBar').style.width = percentage + '%';
        document.getElementById('progressText').textContent = current + ' / ' + total;
      }
      
      function onSuccess(result) {
        document.getElementById('statusText').textContent = '完了';
        document.getElementById('result').style.display = 'block';
        document.getElementById('resultText').innerHTML = 
          '成功: ' + result.successCount + '件<br>' +
          'エラー: ' + result.errorCount + '件';
      }
      
      function onFailure(error) {
        document.getElementById('statusText').textContent = 'エラーが発生しました';
        alert('エラー: ' + error.message);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(html, '請求書発行');
}

/**
 * 進捗表示付き請求書発行処理
 */
function publishSelectedInvoicesWithProgress() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請求書ログ');
  
  try {
    const targetRows = getSelectedRows();
    
    if (targetRows.length === 0) {
      throw new Error('発行対象が選択されていません');
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    for (let i = 0; i < targetRows.length; i++) {
      const rowIndex = targetRows[i];
      
      try {
        updateLogStatus(sheet, rowIndex, '処理中');
        
        const invoiceData = getInvoiceDataFromLog(sheet, rowIndex);
        
        if (!invoiceData.outputUrl) {
          throw new Error('アウトプットURLが設定されていません');
        }
        
        // 商談情報を取得
        const opportunity = getOpportunityInfo(invoiceData.opportunityId);
        
        // invoiceDataに必要な情報を追加
        invoiceData.accountId = opportunity.AccountId;
        invoiceData.accountName = opportunity.Account.Name;
        invoiceData.partnerId = opportunity.Account.MoneyForward_Partner_ID__c;
        
        if (!invoiceData.partnerId) {
          throw new Error(`取引先が同期されていません: ${invoiceData.accountName}`);
        }
        
        // データ変換
        const mfData = convertToMFInvoiceFormat(invoiceData);
        
        const result = createInvoiceWithPartner(mfData);
        
        if (result.success) {
          updateSalesforceInvoiceId(invoiceData.opportunityId, result.invoiceNumber);
          updateLogStatus(sheet, rowIndex, '完了', result.invoiceNumber);
          successCount++;
        } else {
          updateLogStatus(sheet, rowIndex, 'エラー', '', result.error);
          errorCount++;
        }
        
      } catch (error) {
        updateLogStatus(sheet, rowIndex, 'エラー', '', error.message);
        errorCount++;
        logError('請求書発行', error, rowIndex);
      }
      
      Utilities.sleep(1000);
    }
    
    return {
      successCount: successCount,
      errorCount: errorCount
    };
    
  } catch (error) {
    logError('バッチ処理', error);
    throw error;
  }
}

/**
 * 全て選択
 */
function selectAllInvoices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請求書ログ');
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, 1);
    const values = [];
    for (let i = 0; i < lastRow - 1; i++) {
      values.push([true]);
    }
    range.setValues(values);
  }
}

/**
 * 全て選択解除
 */
function deselectAllInvoices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請求書ログ');
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, 1);
    const values = [];
    for (let i = 0; i < lastRow - 1; i++) {
      values.push([false]);
    }
    range.setValues(values);
  }
}

/**
 * 未処理のみ選択
 */
function selectUnprocessedInvoices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請求書ログ');
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const statusRange = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
    const checkRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const values = [];
    
    for (let i = 0; i < statusRange.length; i++) {
      const status = statusRange[i][0];
      values.push([status === '未処理' || status === '' || status === 'エラー']);
    }
    
    checkRange.setValues(values);
  }
}

// ===================================
// 5. エラーハンドリング強化
// ===================================

/**
 * エラーログシートにエラーを記録
 * @param {string} process - 処理名
 * @param {Error} error - エラーオブジェクト
 * @param {number} rowIndex - 関連する行番号
 */
function logError(process, error, rowIndex = null) {
  try {
    const errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('エラーログ');
    
    if (!errorSheet) {
      console.error('エラーログシートが見つかりません');
      return;
    }
    
    const errorData = [
      new Date(),                    // タイムスタンプ
      process,                       // 処理名
      rowIndex ? `行 ${rowIndex + 1}` : '-',  // 関連行
      error.name || 'Error',         // エラー種別
      error.message,                 // エラーメッセージ
      error.stack || ''              // スタックトレース
    ];
    
    errorSheet.appendRow(errorData);
    
  } catch (logError) {
    console.error('エラーログ記録失敗:', logError);
  }
}

/**
 * エラーログを表示
 */
function showErrorLog() {
  const errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('エラーログ');
  
  if (!errorSheet) {
    SpreadsheetApp.getUi().alert('エラーログシートが見つかりません');
    return;
  }
  
  const lastRow = errorSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('エラーログはありません');
    return;
  }
  
  // 最新10件のエラーを取得
  const startRow = Math.max(2, lastRow - 9);
  const numRows = Math.min(10, lastRow - 1);
  const errors = errorSheet.getRange(startRow, 1, numRows, 5).getValues();
  
  let html = '<div style="font-family: Arial, sans-serif; padding: 20px;">';
  html += '<h2>🚨 最新のエラーログ</h2>';
  html += '<table style="width: 100%; border-collapse: collapse;">';
  html += '<tr style="background-color: #f0f0f0;">';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">日時</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">処理</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">行</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">エラー</th>';
  html += '</tr>';
  
  errors.reverse().forEach(error => {
    html += '<tr>';
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${Utilities.formatDate(error[0], 'JST', 'yyyy-MM-dd HH:mm:ss')}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${error[1]}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${error[2]}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${error[4]}</td>`;
    html += '</tr>';
  });
  
  html += '</table>';
  html += '</div>';
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'エラーログ');
}

/**
 * 発行統計を表示
 */
function showStatistics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請求書ログ');
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('データがありません');
    return;
  }
  
  const statusData = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  
  let stats = {
    total: statusData.length,
    unprocessed: 0,
    processing: 0,
    completed: 0,
    error: 0
  };
  
  statusData.forEach(row => {
    const status = row[0];
    switch (status) {
      case '未処理':
      case '':
        stats.unprocessed++;
        break;
      case '処理中':
        stats.processing++;
        break;
      case '完了':
        stats.completed++;
        break;
      case 'エラー':
        stats.error++;
        break;
    }
  });
  
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>📊 発行統計</h2>
      
      <table style="width: 100%; margin-top: 20px;">
        <tr>
          <td style="padding: 10px;">
            <div style="background-color: #f0f0f0; padding: 20px; border-radius: 10px; text-align: center;">
              <h3 style="margin: 0;">全体</h3>
              <p style="font-size: 36px; margin: 10px 0; font-weight: bold;">${stats.total}</p>
            </div>
          </td>
        </tr>
      </table>
      
      <table style="width: 100%; margin-top: 20px;">
        <tr>
          <td style="padding: 10px; width: 25%;">
            <div style="background-color: #e3f2fd; padding: 15px; border-radius: 10px; text-align: center;">
              <h4 style="margin: 0;">未処理</h4>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.unprocessed}</p>
            </div>
          </td>
          <td style="padding: 10px; width: 25%;">
            <div style="background-color: #fff3cd; padding: 15px; border-radius: 10px; text-align: center;">
              <h4 style="margin: 0;">処理中</h4>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.processing}</p>
            </div>
          </td>
          <td style="padding: 10px; width: 25%;">
            <div style="background-color: #d4edda; padding: 15px; border-radius: 10px; text-align: center;">
              <h4 style="margin: 0;">完了</h4>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.completed}</p>
            </div>
          </td>
          <td style="padding: 10px; width: 25%;">
            <div style="background-color: #f8d7da; padding: 15px; border-radius: 10px; text-align: center;">
              <h4 style="margin: 0;">エラー</h4>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.error}</p>
            </div>
          </td>
        </tr>
      </table>
      
      <div style="margin-top: 30px;">
        <p style="color: #666;">完了率: ${Math.round((stats.completed / stats.total) * 100)}%</p>
        <div style="width: 100%; background-color: #f0f0f0; border-radius: 5px;">
          <div style="width: ${(stats.completed / stats.total) * 100}%; height: 20px; background-color: #4CAF50; border-radius: 5px;"></div>
        </div>
      </div>
    </div>
  `)
  .setWidth(500)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, '発行統計');
}

/**
 * 認証設定ダイアログを表示
 */
function showAuthDialog() {
  const mfApi = new MoneyForwardAPI();
  const service = mfApi.getOAuthService();
  const isAuthorized = service.hasAccess();
  
  let html = '<div style="font-family: Arial, sans-serif; padding: 20px;">';
  html += '<h2>🔧 認証設定</h2>';
  
  html += '<h3>MoneyForward認証</h3>';
  if (isAuthorized) {
    html += '<p style="color: green;">✅ 認証済み</p>';
    html += '<button onclick="google.script.run.resetMFAuth()">認証をリセット</button>';
  } else {
    html += '<p style="color: red;">❌ 未認証</p>';
    html += '<p>以下のURLにアクセスして認証を行ってください:</p>';
    html += `<a href="${service.getAuthorizationUrl()}" target="_blank">認証ページを開く</a>`;
  }
  
  html += '<h3 style="margin-top: 30px;">Salesforce認証</h3>';
  try {
    getSalesforceAccessToken();
    html += '<p style="color: green;">✅ 認証済み</p>';
  } catch (error) {
    html += '<p style="color: red;">❌ 認証エラー</p>';
    html += '<p>スクリプトプロパティを確認してください。</p>';
  }
  
  html += '</div>';
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '認証設定');
}

/**
 * MoneyForward認証をリセット
 */
function resetMFAuth() {
  const mfApi = new MoneyForwardAPI();
  const service = mfApi.getOAuthService();
  service.reset();
  SpreadsheetApp.getUi().alert('MoneyForward認証をリセットしました。');
}
