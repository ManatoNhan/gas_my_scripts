// ==============================================
// 初期設定・認証関数
// ==============================================

/**
 * 初期設定
 */
function setupMoneyForwardConfig() {
  const ui = SpreadsheetApp.getUi();
  
  // Client ID入力
  const clientIdResult = ui.prompt(
    'マネーフォワード初期設定',
    'Client IDを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (clientIdResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const clientId = clientIdResult.getResponseText().trim();
  if (!clientId) {
    ui.alert('エラー', 'Client IDが入力されていません', ui.ButtonSet.OK);
    return;
  }
  
  // Client Secret入力
  const clientSecretResult = ui.prompt(
    'マネーフォワード初期設定',
    'Client Secretを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (clientSecretResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const clientSecret = clientSecretResult.getResponseText().trim();
  if (!clientSecret) {
    ui.alert('エラー', 'Client Secretが入力されていません', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const api = new MoneyForwardAPI();
    api.setupConfig(clientId, clientSecret);
    
    const redirectUri = `https://script.google.com/macros/d/${ScriptApp.getScriptId()}/usercallback`;
    
    ui.alert(
      '設定完了',
      `初期設定が完了しました。\n\nリダイレクトURI:\n${redirectUri}\n\n※このURIをマネーフォワードAPIの設定に登録してください。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Setup error:', error);
    ui.alert('エラー', `設定に失敗しました: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 認証開始
 */
function authorize() {
  try {
    const api = new MoneyForwardAPI();
    const service = api.getOAuthService();
    
    if (!service.hasAccess()) {
      const authorizationUrl = service.getAuthorizationUrl();
      console.log('認証URL:', authorizationUrl);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
          <h2 style="color: #333;">マネーフォワード認証</h2>
          <p>以下のリンクをクリックして認証を開始してください：</p>
          <p style="margin: 20px 0;">
            <a href="${authorizationUrl}" target="_blank" style="
              display: inline-block;
              padding: 12px 24px;
              background-color: #4caf50;
              color: white;
              text-decoration: none;
              border-radius: 4px;
              font-weight: bold;
            ">認証を開始</a>
          </p>
        </div>
      `).setWidth(500).setHeight(250);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'マネーフォワード認証');
      
    } else {
      console.log('既に認証されています');
      SpreadsheetApp.getUi().alert('認証済み', '既に認証されています', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('Auth start error:', error);
    SpreadsheetApp.getUi().alert('エラー', `認証開始に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * OAuth2コールバック処理
 */
function authCallback_(request) {
  try {
    console.log('=== OAuth Callback ===');
    console.log('Request:', JSON.stringify(request.parameter || {}));
    
    const api = new MoneyForwardAPI();
    const service = api.getOAuthService();
    const isAuthorized = service.handleCallback(request);
    
    if (isAuthorized) {
      console.log('認証成功');
      return HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
          <h2 style="color: #4caf50;">✅ 認証成功！</h2>
          <p>マネーフォワードとの連携が完了しました</p>
          <script>setTimeout(() => window.close(), 3000);</script>
        </div>
      `);
    } else {
      console.log('認証失敗');
      return HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
          <h2 style="color: #f44336;">❌ 認証失敗</h2>
          <p>認証に失敗しました。もう一度お試しください。</p>
        </div>
      `);
    }
    
  } catch (error) {
    console.error('OAuth callback error:', error);
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
        <h2 style="color: #f44336;">エラー</h2>
        <p>認証処理中にエラーが発生しました: ${error.message}</p>
      </div>
    `);
  }
}

/**
 * 認証状況確認
 */
function checkAuthStatus() {
  try {
    const api = new MoneyForwardAPI();
    const service = api.getOAuthService();
    
    if (service.hasAccess()) {
      SpreadsheetApp.getUi().alert('認証状況', '✅ 認証済みです', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('認証状況', '認証されていません。authorize()を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('Auth status check error:', error);
    SpreadsheetApp.getUi().alert('エラー', `認証状況の確認に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 認証リセット
 */
function resetAuth() {
  try {
    const api = new MoneyForwardAPI();
    const service = api.getOAuthService();
    service.reset();
    
    console.log('認証をリセットしました');
    SpreadsheetApp.getUi().alert('完了', '認証をリセットしました', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Auth reset error:', error);
    SpreadsheetApp.getUi().alert('エラー', `認証リセットに失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 会社一覧取得テスト
 */
function testGetCompanies() {
  try {
    const api = new MoneyForwardAPI();
    const companies = api.getCompanies();
    
    console.log('取得した会社数:', companies.length);
    companies.forEach((company, index) => {
      console.log(`会社${index + 1}:`, {
        id: company.id,
        name: company.name
      });
    });
    
    if (companies.length > 0) {
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2 style="color: #4caf50;">✅ API動作確認成功</h2>
          <p><strong>取得した会社数:</strong> ${companies.length}</p>
          <h3>会社一覧:</h3>
          <ul>
            ${companies.map(company => `<li><strong>${company.name}</strong> (ID: ${company.id})</li>`).join('')}
          </ul>
        </div>
      `).setWidth(600).setHeight(400);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'API動作確認結果');
    }
    
  } catch (error) {
    console.error('会社一覧取得エラー:', error);
    SpreadsheetApp.getUi().alert('API エラー', `会社一覧の取得に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
