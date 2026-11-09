/**
 * レシート読み取り用 Google Apps Script
 *
 * 役割
 *  1. doPost    : アプリから届いた写真を Claude API で読み取り、結果を返す
 *  2. processInbox : olive.an.osaka@gmail.com に届いたレシートメールを読み取り、シートに追記する
 *
 * 設定（プロジェクトの設定 > スクリプト プロパティ）
 *  ANTHROPIC_API_KEY : Claude API のキー
 *  APP_TOKEN         : アプリと共有する合言葉（自由な文字列）
 *  SHEET_ID          : 記録先スプレッドシートのID
 */

var MODEL = "claude-sonnet-5-5";

// 勘定科目の候補。TKCで使っている科目名に合わせて書き換えてください。
var ACCOUNTS = [
  "消耗品費", "食材料費", "水道光熱費", "旅費交通費", "通信費",
  "会議費", "福利厚生費", "修繕費", "車両費", "事務用品費", "雑費"
];

var HOMES = [
  "西九条ホーム", "九条ホーム", "酉島ホーム",
  "新郡山ホーム", "春日出ホーム", "出来島ホーム", "本社"
];

function props_() { return PropertiesService.getScriptProperties(); }

function buildPrompt_() {
  return [
    "あなたは日本の福祉事業所の経理補助です。画像はレシートまたは領収書です。",
    "読み取れた内容だけを、次のJSONで返してください。説明文やコードブロックは付けないでください。",
    "読み取れない項目は null にしてください。推測で数字を作らないでください。",
    "{",
    '  "date": "YYYY-MM-DD",',
    '  "store": "店名",',
    '  "total": 税込合計(整数),',
    '  "tax8": 8%対象の税込金額(整数),',
    '  "tax10": 10%対象の税込金額(整数),',
    '  "invoice_no": "T+13桁の登録番号",',
    '  "account": "次の候補から1つ: ' + ACCOUNTS.join("、") + '",',
    '  "is_business": 事業用の経費として経理処理が必要そうなら true、個人的な買い物や判断できないなら false,',
    '  "reason": "判定の理由を30字以内",',
    '  "readable": 画像が読み取れたか true/false',
    "}"
  ].join("\n");
}

function readReceipt_(base64, mediaType) {
  var body = {
    model: MODEL,
    max_tokens: 800,
    messages: [{
      role: "user",
      content: [
        { type: "image", source: { type: "base64", media_type: mediaType || "image/jpeg", data: base64 } },
        { type: "text", text: buildPrompt_() }
      ]
    }]
  };
  var res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": props_().getProperty("ANTHROPIC_API_KEY"),
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error("Claude API " + res.getResponseCode() + ": " + res.getContentText().slice(0, 200));
  }
  var text = JSON.parse(res.getContentText()).content[0].text;
  var m = text.match(/\{[\s\S]*\}/);
  if (!m) throw new Error("読み取り結果がJSONではありません");
  return JSON.parse(m[0]);
}

/** アプリからの受け口 */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.token !== props_().getProperty("APP_TOKEN")) throw new Error("合言葉が違います");
    out = { ok: true, data: readReceipt_(req.image, req.mediaType) };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * メール受信箱の処理。時間主導型トリガー（例: 1時間ごと）に登録して使う。
 * 件名に拠点名（例: 酉島ホーム）が入っていれば、それを拠点にする。
 * 処理したスレッドには「処理済」ラベルを付ける。
 */
function processInbox() {
  var sheet = SpreadsheetApp.openById(props_().getProperty("SHEET_ID")).getSheetByName("レシート")
    || SpreadsheetApp.openById(props_().getProperty("SHEET_ID")).insertSheet("レシート");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["受信日時", "送信者", "拠点", "日付", "店名", "税込合計", "8%対象", "10%対象",
                     "T番号", "勘定科目", "経理対象", "理由", "確認済"]);
  }
  var label = GmailApp.getUserLabelByName("処理済") || GmailApp.createLabel("処理済");
  var threads = GmailApp.search("has:attachment -label:処理済 newer_than:14d", 0, 30);

  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      var home = "";
      HOMES.forEach(function (h) { if (msg.getSubject().indexOf(h) >= 0) home = h; });
      msg.getAttachments().forEach(function (att) {
        var type = att.getContentType();
        if (type.indexOf("image/") !== 0) return;
        try {
          var r = readReceipt_(Utilities.base64Encode(att.getBytes()), type);
          sheet.appendRow([msg.getDate(), msg.getFrom(), home, r.date, r.store, r.total, r.tax8, r.tax10,
                           r.invoice_no, r.account, r.is_business ? "対象" : "対象外", r.reason, ""]);
        } catch (err) {
          sheet.appendRow([msg.getDate(), msg.getFrom(), home, "", "", "", "", "", "", "", "要確認", String(err.message), ""]);
        }
      });
    });
    th.addLabel(label);
  });
}
