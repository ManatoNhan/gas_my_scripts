function registerToMetabase(email, firstName, lastName) {
  const metabaseUrl = PropertiesService.getScriptProperties().getProperty('metabase_url');
  const apiToken = PropertiesService.getScriptProperties().getProperty('metabase_api_token');
  
  const payload = {
    email: email,
    first_name: firstName,
    last_name: lastName,
    // login_attributesやuser_group_membershipsは必要に応じて追加
  };
  
  try {
    const response = UrlFetchApp.fetch(`${metabaseUrl}/api/user`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-API-KEY': apiToken  // Token認証
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      console.log('Metabaseユーザー登録成功:', responseText);
      return { success: true, message: 'ユーザー登録が完了しました', data: JSON.parse(responseText) };
    } else if (responseCode === 400) {
      console.log('Metabaseユーザー登録エラー:', responseText);
      return { success: false, message: 'メールアドレスが既に使用されています', error: responseText };
    } else {
      console.log('Metabase予期しないエラー:', responseCode, responseText);
      return { success: false, message: 'ユーザー登録に失敗しました', error: responseText };
    }
  } catch (error) {
    console.error('Metabase API呼び出しエラー:', error);
    return { success: false, message: 'API呼び出しに失敗しました', error: error.toString() };
  }
}

