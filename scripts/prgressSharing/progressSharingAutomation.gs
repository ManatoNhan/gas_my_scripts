// ========================================
// クライアント報告管理システム - Google Apps Script
// ========================================

// ========================================
// 設定・定数
// ========================================
const SHEET_CONFIG = {
  CLIENT_LIST: 'Client_List',
  URL_FORMAT: 'URLフォーマット', 
  EXECUTION_LOG: 'ExecutionLog'
};

// メール送信設定
const EMAIL_CONFIG = {
  FROM_ADDRESS: 'sales@wed.company', // エイリアスメールアドレスを指定
  FROM_NAME: 'WED株式会社 - 自動報告システム',
  COMPANY_NAME: 'WED株式会社'
};

const COLUMN_INDEX = {
  CLIENT_NAME: 1,      // A列
  CONTACT_NAME: 2,     // B列
  TO_EMAIL: 3,         // C列
  CC_EMAIL: 4,         // D列
  SALES_EMAIL: 5,      // E列
  SALES_NAME: 6,       // F列
  PRODUCT_NAME: 7,     // G列
  URL_FORMAT: 8,       // H列
  MISSION_TYPE: 9,     // I列
  CAMPAIGN_ID: 10,     // J列
  PARAM_URL: 11,       // K列
  LAST_REPORT: 12,     // L列
  INTERVAL: 13,        // M列
  NEXT_REPORT: 14,     // N列
  PROCESS_FLAG: 15,    // O列
  ADDITIONAL_TEXT: 16, // P列
  API_STATUS: 17,      // Q列
  LAST_PROCESS: 18,    // R列
  RESULT: 19,          // S列
  MESSAGE: 20,         // T列
  URL_PARAMS: 21       // U列
};

// ========================================
// メニュー追加（スプレッドシート起動時）
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('クライアント報告管理')
    .addItem('🔍 報告対象をチェック（参考）', 'checkReportTargets')
    .addItem('📧 メール送信実行', 'executeReporting')
    .addSeparator()
    .addItem('⚙️ 初期設定', 'initializeSheets')
    .addSeparator()
    .addSubMenu(ui.createMenu('📤 エイリアス設定')
      .addItem('エイリアステスト送信', 'testEmailAlias')
      .addItem('エイリアス設定確認', 'getAvailableAliases'))
    .addToUi();
}

// ========================================
// 初期設定：シート作成とヘッダー設定
// ========================================
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Client_Listシートの作成・設定
    createClientListSheet(ss);
    
    // URLフォーマットシートの作成・設定
    createUrlFormatSheet(ss);
    
    // ExecutionLogシートの作成・設定
    createExecutionLogSheet(ss);
    
    SpreadsheetApp.getUi().alert('初期設定が完了しました！');
    
  } catch (error) {
    console.error('初期設定エラー:', error);
    SpreadsheetApp.getUi().alert('初期設定中にエラーが発生しました: ' + error.message);
  }
}

function createClientListSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_CONFIG.CLIENT_LIST);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG.CLIENT_LIST);
  }
  
  // ヘッダー設定
  const headers = [
    'クライアント名', '担当者名', '担当者メールアドレス（To）', 'CCメールアドレス',
    '営業担当メールアドレス', '営業担当者名', '商材名', 'URLフォーマット指定',
    'ミッションタイプ', 'Campaign_id', 'パラメータを反映させたURL', '最終報告日',
    '報告インターバル（日数）', '次回報告予定日', '処理フラグ', 'メール追加文言',
    'API処理ステータス', '最終処理日時', '結果', 'メッセージ', 'URL実行パラメータ'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // 実データ行のみにチェックボックスと数式を設定
  setupDataValidation(sheet);
  
  sheet.autoResizeColumns(1, headers.length);
}

function createUrlFormatSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_CONFIG.URL_FORMAT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG.URL_FORMAT);
  }
  
  const headers = ['フォーマット番号', 'URL', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // サンプルデータ
  const sampleData = [
    [0, 'https://example.com/dashboard?campaign_id={CAMPAIGN_ID}', 'デフォルトダッシュボード'],
    [1, 'https://example.com/analytics?id={CAMPAIGN_ID}&type={MISSION_TYPE}', 'アナリティクス用']
  ];
  sheet.getRange(2, 1, sampleData.length, 3).setValues(sampleData);
  
  sheet.autoResizeColumns(1, 3);
}

function createExecutionLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_CONFIG.EXECUTION_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG.EXECUTION_LOG);
  }
  
  const headers = [
    '実行日時', '処理タイプ', 'クライアント名', '担当者名', 'メールアドレス',
    '結果', 'エラーメッセージ', '送信ID'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  sheet.autoResizeColumns(1, headers.length);
}

// ========================================
// 実データ行の最終行を取得（C列とJ列両方が入力されている行）
// ========================================
function getActualLastRow(sheet) {
  const maxRows = sheet.getMaxRows();
  const cColumnValues = sheet.getRange(2, COLUMN_INDEX.TO_EMAIL, maxRows - 1, 1).getValues();
  const jColumnValues = sheet.getRange(2, COLUMN_INDEX.CAMPAIGN_ID, maxRows - 1, 1).getValues();
  
  let lastRow = 1; // ヘッダー行
  
  for (let i = 0; i < cColumnValues.length; i++) {
    const cValue = cColumnValues[i][0];
    const jValue = jColumnValues[i][0];
    
    // C列とJ列の両方が入力されている場合のみカウント
    if (cValue && cValue.toString().trim() !== '' && 
        jValue && jValue.toString().trim() !== '') {
      lastRow = i + 2; // 配列のインデックスは0始まりなので+2
    }
  }
  
  return lastRow;
}

// ========================================
// データ検証とフォーミュラの設定
// ========================================
function setupDataValidation(sheet) {
  // O列（処理フラグ）にチェックボックスを設定
  const maxDataRows = 100; // 想定される最大データ行数
  const checkboxRange = sheet.getRange(2, COLUMN_INDEX.PROCESS_FLAG, maxDataRows, 1);
  checkboxRange.insertCheckboxes();
  
  // N列（次回報告予定日）に数式を設定
  for (let row = 2; row <= maxDataRows + 1; row++) {
    const formula = `=IF(AND(L${row}<>"", M${row}<>""), L${row}+M${row}, "")`;
    sheet.getRange(row, COLUMN_INDEX.NEXT_REPORT).setFormula(formula);
  }
  
  // K列（パラメータを反映させたURL）に数式を設定
  setupUrlGenerationFormulas(sheet, maxDataRows);
}

// ========================================
// K列のURL生成数式を設定
// ========================================
function setupUrlGenerationFormulas(sheet, maxRows) {
  for (let row = 2; row <= maxRows + 1; row++) {
    // VLOOKUP + 条件分岐でURL生成
    const formula = `=IF(AND(H${row}<>"",J${row}<>""),
      CONCATENATE(
        IFERROR(INDEX(URLフォーマット!B:B,MATCH(H${row},URLフォーマット!A:A,0)),""),
        IF(INDEX(URLフォーマット!B:B,MATCH(H${row},URLフォーマット!A:A,0))<>"",
          CONCATENATE(
            IF(COUNT(FIND("?",INDEX(URLフォーマット!B:B,MATCH(H${row},URLフォーマット!A:A,0))))>0,"&","?"),
            IF(I${row}=1,"receipt_campaign_group_id=","receipt_campaign_id="),
            J${row}
          ),
          ""
        )
      ),
      ""
    )`;
    
    sheet.getRange(row, COLUMN_INDEX.PARAM_URL).setFormula(formula);
  }
}

// ========================================
// 機能1: 報告対象をチェック（補助機能）
// ========================================
function checkReportTargets() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.CLIENT_LIST);
    if (!sheet) {
      throw new Error('Client_Listシートが見つかりません');
    }
    
    // 実データの最終行を取得
    const lastRow = getActualLastRow(sheet);
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('データが登録されていません');
      return;
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0); // 時間を00:00:00に設定
    
    let checkedCount = 0;
    let uncheckedCount = 0;
    
    // 実データ範囲を取得（2行目から実際の最終行まで）
    const dataRange = sheet.getRange(2, 1, lastRow - 1, COLUMN_INDEX.URL_PARAMS);
    const values = dataRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const nextReportDate = row[COLUMN_INDEX.NEXT_REPORT - 1];
      
      // 次回報告予定日が設定されているかチェック
      if (nextReportDate && nextReportDate instanceof Date) {
        const nextReport = new Date(nextReportDate);
        nextReport.setHours(0, 0, 0, 0);
        
        if (nextReport <= today) {
          // 処理フラグをTrueに設定（参考情報として）
          sheet.getRange(i + 2, COLUMN_INDEX.PROCESS_FLAG).setValue(true);
          checkedCount++;
        } else {
          // 処理フラグをFalseに設定（参考情報として）
          sheet.getRange(i + 2, COLUMN_INDEX.PROCESS_FLAG).setValue(false);
          uncheckedCount++;
        }
      } else {
        // 次回報告予定日が未設定の場合はFalse（参考情報として）
        sheet.getRange(i + 2, COLUMN_INDEX.PROCESS_FLAG).setValue(false);
        uncheckedCount++;
      }
    }
    
    const message = `報告対象チェック完了（参考情報）\n推奨対象: ${checkedCount}件\n推奨非対象: ${uncheckedCount}件\n\n※実際の送信は処理フラグ（O列）で制御されます\n※手動でチェックボックスを調整してください`;
    SpreadsheetApp.getUi().alert(message);
    
    // ログ記録
    addExecutionLog('報告対象チェック', '-', '-', '-', '成功', message, '-');
    
  } catch (error) {
    console.error('報告対象チェックエラー:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    addExecutionLog('報告対象チェック', '-', '-', '-', 'エラー', error.message, '-');
  }
}

// ========================================
// 機能2: メール送信実行（処理フラグ優先）
// ========================================
function executeReporting() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.CLIENT_LIST);
    if (!sheet) {
      throw new Error('Client_Listシートが見つかりません');
    }
    
    // 実データの最終行を取得
    const lastRow = getActualLastRow(sheet);
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('データが登録されていません');
      return;
    }
    
    // 処理フラグ（O列）がTrueの行のみを取得
    const dataRange = sheet.getRange(2, 1, lastRow - 1, COLUMN_INDEX.URL_PARAMS);
    const values = dataRange.getValues();
    
    const targets = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const processFlag = row[COLUMN_INDEX.PROCESS_FLAG - 1];
      
      // 処理フラグがTrueの場合のみ送信対象に追加
      if (processFlag === true) {
        targets.push({
          rowIndex: i + 2,
          data: row
        });
      }
    }
    
    if (targets.length === 0) {
      SpreadsheetApp.getUi().alert('送信対象がありません\n\n処理フラグ（O列）にチェックが入っている行がありません。\n送信したい行のO列にチェックを入れてください。');
      return;
    }
    
    // 送信対象の詳細を表示
    let targetList = '送信対象一覧:\n';
    targets.forEach((target, index) => {
      const clientName = target.data[COLUMN_INDEX.CLIENT_NAME - 1] || '（未設定）';
      const contactName = target.data[COLUMN_INDEX.CONTACT_NAME - 1] || '（未設定）';
      const toEmail = target.data[COLUMN_INDEX.TO_EMAIL - 1] || '（未設定）';
      targetList += `${index + 1}. ${clientName} - ${contactName} (${toEmail})\n`;
    });
    
    // 確認ダイアログ
    const response = SpreadsheetApp.getUi().alert(
      `${targets.length}件のメール送信を実行しますか？\n\n${targetList}`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
    
    let successCount = 0;
    let errorCount = 0;
    const errorDetails = [];
    
    // 各対象にメール送信
    for (const target of targets) {
      try {
        sendReportEmail(target.data, target.rowIndex, sheet);
        successCount++;
        
        // 処理後の更新
        sheet.getRange(target.rowIndex, COLUMN_INDEX.ADDITIONAL_TEXT).setValue(''); // P列をクリア
        sheet.getRange(target.rowIndex, COLUMN_INDEX.PROCESS_FLAG).setValue(false); // O列をFalse
        sheet.getRange(target.rowIndex, COLUMN_INDEX.LAST_PROCESS).setValue(new Date()); // R列に実行日時
        sheet.getRange(target.rowIndex, COLUMN_INDEX.RESULT).setValue('送信成功'); // S列
        sheet.getRange(target.rowIndex, COLUMN_INDEX.LAST_REPORT).setValue(new Date()); // L列を更新
        sheet.getRange(target.rowIndex, COLUMN_INDEX.MESSAGE).setValue('正常送信完了'); // T列
        
      } catch (emailError) {
        console.error('メール送信エラー:', emailError);
        errorCount++;
        
        const clientName = target.data[COLUMN_INDEX.CLIENT_NAME - 1] || '（未設定）';
        errorDetails.push(`${clientName}: ${emailError.message}`);
        
        // エラー情報を記録
        sheet.getRange(target.rowIndex, COLUMN_INDEX.PROCESS_FLAG).setValue(false); // エラー時もフラグはリセット
        sheet.getRange(target.rowIndex, COLUMN_INDEX.LAST_PROCESS).setValue(new Date());
        sheet.getRange(target.rowIndex, COLUMN_INDEX.RESULT).setValue('送信エラー');
        sheet.getRange(target.rowIndex, COLUMN_INDEX.MESSAGE).setValue(emailError.message);
        
        // エラーログ
        addExecutionLog(
          'メール送信',
          clientName,
          target.data[COLUMN_INDEX.CONTACT_NAME - 1] || '',
          target.data[COLUMN_INDEX.TO_EMAIL - 1] || '',
          'エラー',
          emailError.message,
          '-'
        );
      }
    }
    
    // 結果メッセージ作成
    let resultMessage = `メール送信完了\n\n成功: ${successCount}件\nエラー: ${errorCount}件`;
    
    if (errorCount > 0) {
      resultMessage += '\n\nエラー詳細:\n' + errorDetails.join('\n');
    }
    
    SpreadsheetApp.getUi().alert(resultMessage);
    
    // 全体ログ
    addExecutionLog('メール送信実行', '-', '-', '-', '完了', `成功:${successCount}件 エラー:${errorCount}件`, '-');
    
  } catch (error) {
    console.error('メール送信実行エラー:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    addExecutionLog('メール送信実行', '-', '-', '-', 'エラー', error.message, '-');
  }
}

// ========================================
// メール送信処理
// ========================================
function sendReportEmail(rowData, rowIndex, sheet) {
  // データ取得
  const clientName = rowData[COLUMN_INDEX.CLIENT_NAME - 1] || '';
  const contactName = rowData[COLUMN_INDEX.CONTACT_NAME - 1] || '';
  const toEmail = rowData[COLUMN_INDEX.TO_EMAIL - 1] || '';
  const ccEmail = rowData[COLUMN_INDEX.CC_EMAIL - 1] || '';
  const salesEmail = rowData[COLUMN_INDEX.SALES_EMAIL - 1] || '';
  const salesName = rowData[COLUMN_INDEX.SALES_NAME - 1] || '';
  const productName = rowData[COLUMN_INDEX.PRODUCT_NAME - 1] || '';
  const additionalText = rowData[COLUMN_INDEX.ADDITIONAL_TEXT - 1] || '';
  const metabaseurl = rowData[COLUMN_INDEX.PARAM_URL - 1] || '';
  
  // バリデーション
  if (!toEmail) {
    throw new Error('送信先メールアドレスが設定されていません');
  }
  
  // メールアドレスの分割・整形
  const toAddresses = parseEmailAddresses(toEmail);
  const ccAddresses = ccEmail ? parseEmailAddresses(ccEmail) : [];
  
  // メール本文作成
  const subject = `【ONE: ${productName}】進捗報告`;
  let body = `
    ${clientName}\n${contactName} 様\n\nお世話になっております。\n
    ${EMAIL_CONFIG.COMPANY_NAME}です。\n\n
    ※本メールは自動配信になります\n
    ※お問い合わせは担当の${salesName}(${salesEmail})までご連絡ください\n
    \n
    現在、実施中のキャンペーン「${productName}」の進捗を下記URLからご確認ください。\n
    【進捗報告】${metabaseurl}
    `;
  
  // 追加文言があれば挿入
  if (additionalText) {
    body += `\n${additionalText}\n`;
  }
  
  body += '\nよろしくお願いいたします。';
  
  // メール送信オプション
  const options = {
    cc: ccAddresses.join(','),
    from: EMAIL_CONFIG.FROM_ADDRESS, // エイリアスメールアドレス
    name: EMAIL_CONFIG.FROM_NAME
  };
  
  // 空のCCを除去
  if (!options.cc) {
    delete options.cc;
  }
  
  try {
    // メール送信
    GmailApp.sendEmail(toAddresses.join(','), subject, body, options);
    
    // 送信ログ
    addExecutionLog(
      'メール送信',
      clientName,
      contactName,
      toEmail,
      '成功',
      `From: ${EMAIL_CONFIG.FROM_ADDRESS} で送信完了`,
      Utilities.getUuid()
    );
    
  } catch (gmailError) {
    // Gmailエラーの詳細をログに記録
    const errorMsg = `Gmail送信エラー: ${gmailError.message}`;
    console.error(errorMsg);
    
    addExecutionLog(
      'メール送信',
      clientName,
      contactName,
      toEmail,
      'エラー',
      errorMsg,
      '-'
    );
    
    throw new Error(errorMsg);
  }
}

// ========================================
// エイリアス設定確認・テスト機能
// ========================================
function testEmailAlias() {
  try {
    // テストメール送信でエイリアス確認（権限不要な方法を使用）
    const testEmail = PropertiesService.getScriptProperties().getProperty('TEST_EMAIL') || 'test@example.com';
    const testSubject = 'エイリアステスト - WED報告システム';
    const testBody = `エイリアスメール送信テストです。\n\nFrom: ${EMAIL_CONFIG.FROM_ADDRESS}\nName: ${EMAIL_CONFIG.FROM_NAME}\n\n送信日時: ${new Date()}\n\n※このメールでエイリアス設定が正しく動作しているか確認してください。`;
    
    const options = {
      from: EMAIL_CONFIG.FROM_ADDRESS,
      name: EMAIL_CONFIG.FROM_NAME
    };
    
    // テスト送信先メールアドレスの入力を求める
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'エイリアステスト',
      'テストメールの送信先アドレスを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const inputEmail = response.getResponseText().trim();
    if (!inputEmail || !inputEmail.includes('@')) {
      ui.alert('❌ 有効なメールアドレスを入力してください');
      return;
    }
    
    // 確認ダイアログ
    const confirmResponse = ui.alert(
      `エイリアステストメールを送信しますか？\n\n` +
      `送信先: ${inputEmail}\n` +
      `From: ${EMAIL_CONFIG.FROM_ADDRESS}\n` +
      `Name: ${EMAIL_CONFIG.FROM_NAME}`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    GmailApp.sendEmail(inputEmail, testSubject, testBody, options);
    
    // テストメールアドレスを保存（次回使用のため）
    PropertiesService.getScriptProperties().setProperty('TEST_EMAIL', inputEmail);
    
    ui.alert(
      `✅ テストメール送信完了\n\n` +
      `送信先: ${inputEmail}\n` +
      `From: ${EMAIL_CONFIG.FROM_ADDRESS}\n` +
      `Name: ${EMAIL_CONFIG.FROM_NAME}\n\n` +
      `📥 受信ボックスを確認して、送信者情報をチェックしてください。\n` +
      `エイリアスが正しく表示されていれば設定完了です。`
    );
    
    // ログ記録
    addExecutionLog('エイリアステスト', '-', '-', inputEmail, '成功', 'テストメール送信', '-');
    
  } catch (error) {
    const errorMsg = `エイリアステストエラー: ${error.message}`;
    console.error(errorMsg);
    
    let userFriendlyMsg = '❌ エイリアステスト送信エラー\n\n';
    
    if (error.message.includes('Invalid from')) {
      userFriendlyMsg += `指定されたエイリアス「${EMAIL_CONFIG.FROM_ADDRESS}」が無効です。\n\n`;
      userFriendlyMsg += '対処方法:\n';
      userFriendlyMsg += '1. Gmail設定でエイリアスが正しく設定されているか確認\n';
      userFriendlyMsg += '2. EMAIL_CONFIGの設定を確認\n';
      userFriendlyMsg += '3. エイリアス認証が完了しているか確認';
    } else {
      userFriendlyMsg += `エラー詳細: ${error.message}`;
    }
    
    SpreadsheetApp.getUi().alert(userFriendlyMsg);
    addExecutionLog('エイリアステスト', '-', '-', '-', 'エラー', errorMsg, '-');
  }
}

// ========================================
// エイリアス設定確認・ガイド機能
// ========================================
function getAvailableAliases() {
  try {
    // Session.getActiveUser()の代わりに、権限不要な方法を使用
    const guideMessage = 
      `📧 エイリアス設定ガイド\n\n` +
      `🔧 Gmail側でのエイリアス設定手順:\n` +
      `1. Gmail設定 → アカウントとインポート\n` +
      `2. 「名前」セクション → 「メールアドレスを追加」\n` +
      `3. エイリアスメールアドレスを入力\n` +
      `4. 確認手続きを完了\n\n` +
      `⚙️ 現在のGAS設定:\n` +
      `FROM_ADDRESS: ${EMAIL_CONFIG.FROM_ADDRESS}\n` +
      `FROM_NAME: ${EMAIL_CONFIG.FROM_NAME}\n` +
      `COMPANY_NAME: ${EMAIL_CONFIG.COMPANY_NAME}\n\n` +
      `💡 エイリアス設定のポイント:\n` +
      `• G Workspace の場合は管理者がエイリアス作成\n` +
      `• 個人Gmail でも「送信メールアドレス」として追加可能\n` +
      `• 確認コードでの認証が必要\n\n` +
      `✅ 設定完了後は「エイリアステスト送信」で動作確認してください。`;
    
    SpreadsheetApp.getUi().alert(guideMessage);
    
    // ログ記録（ユーザー情報なしで記録）
    addExecutionLog('エイリアス設定確認', '-', '-', '-', '成功', '設定ガイド表示', '-');
    
  } catch (error) {
    const errorMsg = `エイリアス設定確認エラー: ${error.message}`;
    console.error(errorMsg);
    SpreadsheetApp.getUi().alert('❌ ' + errorMsg);
    addExecutionLog('エイリアス設定確認', '-', '-', '-', 'エラー', errorMsg, '-');
  }
}

// ========================================
// ユーティリティ関数
// ========================================
function parseEmailAddresses(emailString) {
  if (!emailString) return [];
  
  // 区切り文字で分割（,、;、全角カンマ）
  const emails = emailString.split(/[,;、；]/)
    .map(email => email.trim())
    .filter(email => email.length > 0);
  
  return emails;
}

function addExecutionLog(processType, clientName, contactName, emailAddress, result, message, sendId) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.EXECUTION_LOG);
    if (!logSheet) return;
    
    const logData = [
      new Date(), // 実行日時
      processType, // 処理タイプ
      clientName, // クライアント名
      contactName, // 担当者名
      emailAddress, // メールアドレス
      result, // 結果
      message, // エラーメッセージ
      sendId // 送信ID
    ];
    
    logSheet.appendRow(logData);
    
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}
