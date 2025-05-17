function onFormSubmit(e) {
  var values = e.values;
  var namedValues = e.namedValues;
  var timestamp = values[0];
  var formattedDate = Utilities.formatDate(new Date(timestamp), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  // ====== 委員会ごとのメールアドレス ======
  var committeeEmails = {
    "総務委員会": "soumu@example.com",
    "福祉活動委員会": "fukushi@example.com",
    "保健体育委員会": "hoken@example.com",
    "環境衛生委員会": "kankyo@example.com",
    "防災防犯委員会": "bousai@example.com",
    "ふれあい活動委員会": "fureai@example.com",
    "文化教養委員会": "bunka@example.com"
  };

  // ====== イベント → 委員会 のマッピング ======
  var eventCommittee = {
    "新年福引会": "総務委員会",
    "成人祝い": "総務委員会",
    "敬老のつどい": "福祉活動委員会",
    "茶話会": "福祉活動委員会",
    "地区連合運動会": "保健体育委員会",
    "公民館地区球技大会": "保健体育委員会",
    "春のごみゼロキャンペーン": "環境衛生委員会",
    "秋のごみゼロキャンペーン": "環境衛生委員会",
    "避難場所運営訓練": "防災防犯委員会",
    "夜警": "防災防犯委員会", 
    "自主防災会訓練": "防災防犯委員会",
    "校門警備": "防災防犯委員会",
    "わんわんパトロール": "防災防犯委員会",
    "おたのしみ会": "ふれあい活動委員会",
    "ファミリー納涼のつどい": "ふれあい活動委員会",
    "日帰り研修旅行": "文化教養委員会",
    "料理教室": "文化教養委員会"
  };

  // ====== フォームの質問キー：正確なタイトルを記述 ======
  var eventKey = "お申し込みのイベント"; // ← ここが重要！
  var committeeKey = "参加するイベントの委員会"; // ← 任意で指定されている場合用

  var selectedEvent = namedValues[eventKey]?.[0] || "";
  var selectedCommittee = namedValues[committeeKey]?.[0] || "";

  // イベントから委員会名を逆引き
  if (!selectedCommittee && selectedEvent && eventCommittee[selectedEvent]) {
    selectedCommittee = eventCommittee[selectedEvent];
  }

  // ====== イベント名でシート（タブ）を分ける ======
  var sheetName = selectedEvent || "未分類イベント";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName(sheetName);

  if (!targetSheet) {
    targetSheet = ss.insertSheet(sheetName);
    var headers = ["タイムスタンプ"];
    for (var key in namedValues) {
      headers.push(key);
    }
    targetSheet.appendRow(headers);
  }

  // 回答行を作成
  var row = [formattedDate];
  for (var key in namedValues) {
    row.push(namedValues[key][0]);
  }

  targetSheet.appendRow(row);

  // ====== メール送信 ======
  var emailTo = "default@example.com"; // fallback
  if (committeeEmails[selectedCommittee]) {
    emailTo = committeeEmails[selectedCommittee];
  }

  var subject = "【" + (selectedEvent || selectedCommittee || "イベント") + "】新しい申し込みがありました";
  var body = `以下の内容で申し込みがありました。\n\n`;
  body += `◆申し込み日時: ${formattedDate}\n`;
  if (selectedEvent) body += `◆イベント名: ${selectedEvent}\n`;
  if (selectedCommittee) body += `◆担当委員会: ${selectedCommittee}\n`;
  body += `\n==== 回答内容 ====\n`;
  for (var key in namedValues) {
    body += `${key}: ${namedValues[key][0]}\n`;
  }
  body += `\n============================\n`;
  body += `スプレッドシートを確認: ${ss.getUrl()}`;

  try {
    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      body: body
    });
    Logger.log("メール送信成功: " + emailTo);
  } catch (error) {
    Logger.log("メール送信エラー: " + error.toString());
  }
}
