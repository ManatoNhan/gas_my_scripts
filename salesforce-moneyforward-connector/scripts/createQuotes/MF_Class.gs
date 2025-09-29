

// ===========================================
// セキュアOAuth2管理クラス
// ===========================================

/**
 * CLIENT_SECRET_POST対応 MoneyForwardAPI
 * OAuth2ライブラリID: 1B7FSrxSEb9zLgAE4AhPzpTJSDiSCpEF1MrZ8m7iIx4u-B_JzQH70cYr-
 */

class MoneyForwardAPI {
  constructor() {
    this.props = PropertiesService.getScriptProperties();
    this.apiBaseUrl = 'https://invoice.moneyforward.com/api/v3';
  }
  
  /**
   * OAuth2設定
   */
  setupConfig(clientId, clientSecret) {
    this.props.setProperties({
      'CLIENT_ID': clientId,
      'CLIENT_SECRET': clientSecret
    });
    console.log('OAuth2設定完了');
  }
  
  /**
   * CLIENT_SECRET_POST対応のOAuth2サービス取得
   */
  getOAuthService() {
    const serviceName = 'mfInvoice';
    const authorizationBaseUrl = 'https://api.biz.moneyforward.com/authorize';
    const tokenUrl = 'https://api.biz.moneyforward.com/token';
    
    const clientId = this.props.getProperty('MF_CLIENT_ID');
    const clientSecret = this.props.getProperty('MF_CLIENT_SECRET');
    
    if (!clientId || !clientSecret) {
      throw new Error('Client IDまたはClient Secretが設定されていません。setupConfig()を実行してください。');
    }
    
    return OAuth2.createService(serviceName)
        .setAuthorizationBaseUrl(authorizationBaseUrl)
        .setTokenUrl(tokenUrl)
        .setClientId(clientId)
        .setClientSecret(clientSecret)
        .setCallbackFunction('authCallback_')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope('mfc/invoice/data.read mfc/invoice/data.write')
        // CLIENT_SECRET_POST強制設定
        .setTokenPayloadHandler(function(payload) {
          payload.client_id = clientId;
          payload.client_secret = clientSecret;
          return payload;
        });
  }
  
  /**
   * 認証ヘッダー取得
   */
  getAuthHeaders() {
    const service = this.getOAuthService();
    if (!service.hasAccess()) {
      throw new Error('認証が必要です。authorize()を実行してください。');
    }
    
    const accessToken = service.getAccessToken();
    return {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };
  }
  
  /**
   * 会社一覧を取得
   */
  getCompanies() {
    const options = {
      'method': 'GET',
      'headers': this.getAuthHeaders(),
      'muteHttpExceptions': true
    };
    
    try {
      const response = UrlFetchApp.fetch(`${this.apiBaseUrl}/partners`, options);
      const responseData = JSON.parse(response.getContentText());
      
      if (response.getResponseCode() !== 200) {
        throw new Error(`API Error: ${responseData.message || 'Unknown error'}`);
      }
      
      return responseData.companies || [];
      
    } catch (error) {
      console.error('Get companies error:', error);
      throw error;
    }
  }
  
  /**
   * 請求書作成
   */
  createInvoice(companyId, invoiceData) {
    const url = `${this.apiBaseUrl}/companies/${companyId}/billings`;
    
    const options = {
      'method': 'POST',
      'headers': this.getAuthHeaders(),
      'payload': JSON.stringify(invoiceData),
      'muteHttpExceptions': true
    };
    
    try {
      console.log('Creating invoice via MoneyForward API...');
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());
      
      if (response.getResponseCode() !== 201 && response.getResponseCode() !== 200) {
        return {
          success: false,
          error: `API Error (${response.getResponseCode()}): ${responseData.message || responseData.error || 'Unknown error'}`
        };
      }
      
      return {
        success: true,
        invoiceNumber: responseData.billing_number || responseData.number || responseData.id,
        invoiceId: responseData.id,
        data: responseData
      };
      
    } catch (error) {
      console.error('MoneyForward API error:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * 見積書作成
   */
  createQuote(companyId, quoteData) {
    const url = `${this.apiBaseUrl}/companies/${companyId}/quotes`;
    
    const options = {
      'method': 'POST',
      'headers': this.getAuthHeaders(),
      'payload': JSON.stringify(quoteData),
      'muteHttpExceptions': true
    };
    
    try {
      console.log('Creating quote via MoneyForward API...');
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());
      
      if (response.getResponseCode() !== 201 && response.getResponseCode() !== 200) {
        return {
          success: false,
          error: `API Error (${response.getResponseCode()}): ${responseData.message || responseData.error || 'Unknown error'}`
        };
      }
      
      return {
        success: true,
        quoteNumber: responseData.quote_number || responseData.number || responseData.id,
        quoteId: responseData.id,
        data: responseData
      };
      
    } catch (error) {
      console.error('MoneyForward Quote API error:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
}



