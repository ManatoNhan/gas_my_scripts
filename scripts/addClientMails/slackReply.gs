function editTrigger(e) {
  console.log("editTrigger (onChange) function has been called!");
  console.log("Event object:", e); // eオブジェクト全体をログに出力して確認

  // 変更の種類が「行の挿入」であることを確認
  if (e.changeType === 'INSERT_ROW') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row = sheet.getLastRow(); // 追加された行が最終行であると仮定

    // スプレッドシートからデータを取得
    // Slackフォームが具体的にどの列にどのデータを書き込むかを確認してください。
    // 例: A列に会社名、D列にメールアドレスが書き込まれると仮定します。
    const company = sheet.getRange(row, 1).getValue();    // A列: 会社名
    const lname = sheet.getRange(row, 2).getValue();      // B列: 担当者名(LastName)
    const fname = sheet.getRange(row, 3).getValue();      // C列: 担当者名(FirstName)
    const email = sheet.getRange(row, 4).getValue();      // D列: email
    const slackAcc = sheet.getRange(row, 5).getValue();   // E列: 送信者
    const msgURL = sheet.getRange(row, 8).getValue();     // H列: messageURL
    const channelId = sheet.getRange(row, 10).getValue(); // J列: channel_id

    // メールアドレスが入力されていることを最低条件とする
    if (email) { // ここで処理実行の条件を設定（例: メールアドレスが空でない）
      const name = lname+' '+fname;

      // msgURLが有効なURLであることを確認してからthreadTsを変換
      let threadTs = null;
      if (msgURL && typeof msgURL === 'string' && msgURL.startsWith('http')) {
          threadTs = convertUrlToThreadTs(msgURL);
      }
      if (threadTs) {
          sheet.getRange(row, 9).setValue(threadTs);
      } else {
          console.warn("Invalid or missing msgURL, skipping threadTs conversion for row:", row);
      }

      // Metabaseユーザー登録実行
      const result = registerToMetabase(email, fname, lname);

      // 結果に応じてSlackに返信
      const message = result.success ?
        `✅ <@${slackAcc}>\n\n${company}の${name}さん（${email}）のMetabaseアカウント登録が完了しました!` :
        `❌ <@${slackAcc}>\n\n${company}の${name}さん（${email}）のMetabaseアカウント登録に失敗しました: ${result.message}`;

      // channelIdとthreadTsが有効な場合にのみSlackに返信
      if (channelId && threadTs) {
          replyToSlackThread(threadTs, channelId, message);
      } else {
          console.warn("Missing channelId or threadTs, skipping Slack reply for row:", row);
      }

      // ステータス更新 (G列に書き込む)
      sheet.getRange(row, 7).setValue(result.success ? "完了" : "エラー");
    } else {
      console.log("Skipping row", row, "due to missing email or other conditions.");
    }
  }
}



function convertUrlToThreadTs(url) {
  // pXXXXXXXXXXXXXXXX の形式から抽出
  const match = url.match(/p(\d+)/);
  if (match) {
    const timestamp = match[1];
    // 最後の6桁の前にピリオドを挿入
    const threadTs = timestamp.slice(0, -6) + '.' + timestamp.slice(-6);
    return threadTs;
  }
  return null;
}

function replyToSlackThread(threadTs, channelId, message) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('webhook_url')
  
  const payload = {
    channel: channelId,
    text: message,
    thread_ts: threadTs
  };
  
  try {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload)
    });
    
    console.log('Slack返信成功:', response.getContentText());
  } catch (error) {
    console.error('Slack返信エラー:', error);
  }
}
