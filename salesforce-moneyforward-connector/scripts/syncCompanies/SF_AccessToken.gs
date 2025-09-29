// GASスクリプトでアクセストークンを取得する
function getSalesforceAccessToken() {
  const clientId = PropertiesService.getScriptProperties().getProperty('SF_CLIENT_ID'); // SalesforceのConnected AppのClient ID
  const clientSecret = PropertiesService.getScriptProperties().getProperty('SF_CLIENT_SECRET'); // SalesforceのConnected AppのClient Secret
  const tokenUrl = 'https://wed.my.salesforce.com/services/oauth2/token'; // エンドポイント
  
  // リクエスト用のペイロード
  const payload = {
    'grant_type': 'client_credentials',
    'client_id': clientId,
    'client_secret': clientSecret
  };
  
  // POSTリクエストのオプション
  const options = {
    'method': 'post',
    'payload': payload
  };
  
  // Salesforceにアクセストークンをリクエスト
  const response = UrlFetchApp.fetch(tokenUrl, options);
  
  // レスポンスをパースしてアクセストークンを取得
  const result = JSON.parse(response.getContentText());
  
  // エラーチェック
  if (result.error) {
    Logger.log('Error: ' + result.error_description);
    return null;
  }
  
  Logger.log('Access Token: ' + result.access_token);
  return result.access_token;
}

// Salesforce APIを呼び出す
function callSalesforceApi() {
  const accessToken = getSalesforceAccessToken();
  if (!accessToken) {
    Logger.log('アクセストークンの取得に失敗しました。');
    return;
  }
  
  // APIエンドポイント (例: アカウント情報取得)
  const apiUrl = 'https://wed.my.salesforce.com/services/data/v58.0/sobjects/Account';
  
  // APIリクエストのオプション
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + accessToken
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    Logger.log(result);
  } catch (error) {
    Logger.log('API呼び出し中にエラーが発生しました: ' + error);
  }
}
