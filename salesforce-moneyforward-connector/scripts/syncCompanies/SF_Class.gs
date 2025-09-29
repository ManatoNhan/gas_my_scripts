/**
 * Salesforce API呼び出しクラス
 */
class SalesforceAPI {
  constructor() {
    this.baseURL = 'https://wed.my.salesforce.com/services/data/v58.0';
    this.token = null;
  }

  /**
   * アクセストークンを取得・キャッシュ
   */
  getToken() {
    if (!this.token) {
      this.token = getSalesforceAccessToken();
    }
    return this.token;
  }

  /**
   * トークンをリフレッシュ
   */
  refreshToken() {
    this.token = null;
    return this.getToken();
  }

  /**
   * 基本的なAPI呼び出し
   * @param {string} endpoint - APIエンドポイント
   * @param {string} method - HTTPメソッド (GET, POST, PATCH, DELETE)
   * @param {Object} payload - リクエストボディ
   * @returns {Object} レスポンス結果
   */
  call(endpoint, method = 'GET', payload = null) {
    const options = {
      method: method,
      headers: {
        'Authorization': 'Bearer ' + this.getToken(),
        'Content-Type': 'application/json'
      }
    };

    if (payload && (method === 'POST' || method === 'PATCH')) {
      options.payload = JSON.stringify(payload);
    }

    try {
      const response = UrlFetchApp.fetch(`${this.baseURL}${endpoint}`, options);
      const responseData = JSON.parse(response.getContentText());

      if (response.getResponseCode() >= 400) {
        throw new Error(`SF API Error: ${responseData[0]?.message || 'Unknown error'}`);
      }

      return {
        success: true,
        data: responseData
      };
    } catch (error) {
      console.error('SF API Error:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * GETリクエスト
   * @param {string} endpoint - APIエンドポイント
   * @returns {Object} レスポンス結果
   */
  get(endpoint) {
    return this.call(endpoint, 'GET');
  }

  /**
   * POSTリクエスト
   * @param {string} endpoint - APIエンドポイント
   * @param {Object} payload - リクエストボディ
   * @returns {Object} レスポンス結果
   */
  post(endpoint, payload) {
    return this.call(endpoint, 'POST', payload);
  }

  /**
   * PATCHリクエスト（更新）
   * @param {string} endpoint - APIエンドポイント
   * @param {Object} payload - リクエストボディ
   * @returns {Object} レスポンス結果
   */
  patch(endpoint, payload) {
    return this.call(endpoint, 'PATCH', payload);
  }

  /**
   * DELETEリクエスト
   * @param {string} endpoint - APIエンドポイント
   * @returns {Object} レスポンス結果
   */
  delete(endpoint) {
    return this.call(endpoint, 'DELETE');
  }

  /**
   * SOQL クエリ実行
   * @param {string} query - SOQLクエリ文字列
   * @returns {Object} クエリ結果
   */
  query(query) {
    const encodedQuery = encodeURIComponent(query);
    return this.get(`/query/?q=${encodedQuery}`);
  }

  /**
   * レコード作成
   * @param {string} objectType - Salesforceオブジェクト名 (Account, Contact等)
   * @param {Object} recordData - レコードデータ
   * @returns {Object} 作成結果
   */
  createRecord(objectType, recordData) {
    return this.post(`/sobjects/${objectType}/`, recordData);
  }

  /**
   * レコード更新
   * @param {string} objectType - Salesforceオブジェクト名
   * @param {string} recordId - レコードID
   * @param {Object} recordData - 更新データ
   * @returns {Object} 更新結果
   */
  updateRecord(objectType, recordId, recordData) {
    return this.patch(`/sobjects/${objectType}/${recordId}`, recordData);
  }

  /**
   * レコード取得
   * @param {string} objectType - Salesforceオブジェクト名
   * @param {string} recordId - レコードID
   * @param {string} fields - 取得フィールド（カンマ区切り）
   * @returns {Object} レコードデータ
   */
  getRecord(objectType, recordId, fields = null) {
    let endpoint = `/sobjects/${objectType}/${recordId}`;
    if (fields) {
      endpoint += `?fields=${encodeURIComponent(fields)}`;
    }
    return this.get(endpoint);
  }

  /**
   * レコード削除
   * @param {string} objectType - Salesforceオブジェクト名
   * @param {string} recordId - レコードID
   * @returns {Object} 削除結果
   */
  deleteRecord(objectType, recordId) {
    return this.delete(`/sobjects/${objectType}/${recordId}`);
  }
}

// 使用例
/*
const sf = new SalesforceAPI();

// レコード取得
const account = sf.getRecord('Account', '001XXXXXXXXXXXXXX', 'Name,Type,Industry');

// SOQL クエリ実行
const contacts = sf.query("SELECT Id, Name, Email FROM Contact WHERE AccountId = '001XXXXXXXXXXXXXX'");

// レコード作成
const newContact = sf.createRecord('Contact', {
  FirstName: 'John',
  LastName: 'Doe',
  Email: 'john.doe@example.com'
});

// レコード更新
const updatedContact = sf.updateRecord('Contact', '003XXXXXXXXXXXXXX', {
  Email: 'newemail@example.com'
});
*/
