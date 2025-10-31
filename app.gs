// ========================================
// 設定セクション
// ========================================

// 書籍リスト: ここに管理する書籍のタイトルを配列で記載する
// 新しい書籍を追加する場合はこのリストに追加するだけでOK
const BOOK_TITLES = [
  "BookA",
  "BookB",
  "BookC",
  "BookD",
];

// スプレッドシートID
const SPREADSHEET_ID = "XXX"; // ココにスプレッドシートのIDを記載する

// イベント名（シート名）
const EVENT_NAME = "xxx"; // ココに該当のスプレッドシートの該当シート名を記載する

// Slack Webhook URLs
const SLACK_WEBHOOK_URLS = [
  "https://hooks.slack.com/services/xxx",
  "https://hooks.slack.com/services/yyy",
];

// グラフ投稿用のSlack設定（最初のWebhook URLを使用する場合に有効）
const SLACK_BOT_TOKEN = "xoxb-xxx"; // ココにBot User OAuth Tokenを記載する
const SLACK_CHANNEL_ID = "xxx"; // ココに通知先のチャンネルIDを記載する

// グラフ配置の設定
const CHART_POSITION_ROW = 2;
const CHART_POSITION_COLUMN_OFFSET = 2; // 書籍列の右側に配置するための列オフセット

// ========================================
// ヘルパー関数
// ========================================

/**
 * 書籍リストから列番号のマップを生成する
 * @return {Object} 書籍名をキー、列番号（2から開始）を値とするオブジェクト
 */
function createColumnMap() {
  const columnMap = {};
  BOOK_TITLES.forEach((title, index) => {
    columnMap[title] = index + 2; // 列番号は2から開始（1列目は日付用）
  });
  return columnMap;
}

/**
 * 書籍の総数を取得する
 * @return {number} 書籍の総数
 */
function getBookCount() {
  return BOOK_TITLES.length;
}

/**
 * グラフ描画に使用する総列数を取得する（日付列 + 書籍列）
 * @return {number} 総列数
 */
function getTotalColumns() {
  return 1 + getBookCount(); // 1列目（日付）+ 書籍の列数
}

/**
 * グラフ配置の列位置を取得する
 * @return {number} グラフを配置する列番号
 */
function getChartPositionColumn() {
  return getTotalColumns() + CHART_POSITION_COLUMN_OFFSET;
}

// ========================================
// メイン処理
// ========================================

function checkBuyMail() {
  // 該当のメールを検索
  // 受信トレイにある未開封のメールのうち、「ファン」というキーワードが含まれているメールを拾ってくる
  const threads = GmailApp.search("in:Inbox is:Unread ファン", 0, 100);

  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      if(!message.isUnread()) { return }
      const text = create_message(message);
      SLACK_WEBHOOK_URLS.forEach((webhookUrl) => {
        sendTextToSlack(text, webhookUrl);
        if (webhookUrl === SLACK_WEBHOOK_URLS[0]) { // 最初のWebhook URLの場合、グラフを投稿する
          calcBuyData(message);
        }
      });
      message.markRead();
    })
    thread.moveToTrash();
  })
}

function create_message(message) {
  var body = extMessage(getMessageBodyAsText(message));
  return `:book:${message.getSubject()}:shopping_trolley:${body}\n`;
}

function getMessageBodyAsText(message) {
  // HTMLメールかプレーンテキストかを確認し、HTMLメールの場合はテキスト化する
  var htmlBody = message.getBody();
  if (htmlBody) {
    return convertHtmlToText(htmlBody).split(/\r\n|\n/);
  } else {
    return message.getPlainBody().split(/\r\n|\n/);
  }
}

function convertHtmlToText(html) {
  // HTMLタグを削除し、テキストのみを抽出
  return html.replace(/<[^>]+>/g, '');
}

function extMessage(texts) {
  const keywords = ["電子版", "電子&#43;紙", "会場（電子&#43;紙）", "会場（電子版）"];
  for (var i = 0; i < texts.length; i++) {
    for (var j = 0; j < keywords.length; j++) {
      if (texts[i].includes(keywords[j])) {
        return keywords[j];
      }
    }
  }
  return "販売形態が見つかりません";
}

function decodeHtmlEntities(text) {
  // HTMLエンコードされた文字をデコード
  return text.replace(/&#43;/g, '+');
}

function sendTextToSlack(text, webhookUrl) {
  const data = { "text": text }
  var payload = JSON.stringify(data);
  sendToSlack(payload, webhookUrl);
}

function sendToSlack(payload, webhookUrl) {
  const headers = { "Content-type": "application/json" }
  const options = {
    "method": "post",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": false
  }
  UrlFetchApp.fetch(webhookUrl, options);
}

function extBookTitle(message, titles) {
  const subject = message.getSubject()
  // メールから書籍のタイトルを取り出す
  const res = titles.filter((title) => {
    const result = subject.match(title);
    if (result !== null) return title;
  })
  return res[0];
}

function getTargetSheet(sheetName){
  // 集計結果を書き込むスプレッドシートを取得
  var spread = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = spread.getSheetByName(sheetName);
  return sheet;
}

function getStartDate(sheetName) {
  // 集計の開始日を取得する
  // 基本的には技術書典の開催日
  var sheet = getTargetSheet(sheetName);

  var startDate = sheet.getRange("A2").getValue();
  if (startDate === "") {
    // 未記入の場合は今日の日付を開催日とする
    startDate = new Date();
  }

  return startDate;
}

function calcDiffDates(startDate) {
  // 集計開始日からの経過日数を算出する
  var today = new Date();

  var timeDiff = Math.abs(today.getTime() - startDate.getTime());
  var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24));

  return diffDays;
}

function writeDatesFromStartDate(sheetName) {
  // 集計開始日から今日までの日付を埋める
  const sheet = getTargetSheet(sheetName);
  const startDate = getStartDate(sheetName);
  const diffDays = calcDiffDates(startDate);

  for (var i = 0; i < diffDays; i++) {
    var row = i + 2;
    var cell = sheet.getRange(row, 1);
    cell.setValue(new Date(startDate.getTime() + i * (1000 * 3600 * 24)));
  }
  if (diffDays === 0) {
    var cell = sheet.getRange(2, 1);
    cell.setValue(today);
  }
}

function incrementCellValue(sheetName, row, columName, columnMap) {
  // 該当する日付, 書籍のセルの値をインクリメントする
  const sheet = getTargetSheet(sheetName);

  const columnNumber = columnMap[columName];

  if (!columnNumber) {
    throw new Error("指定した書籍名は存在しません");
  }

  const cell = sheet.getRange(row, columnNumber);

  const value = cell.getValue();

  var incrementValue = value + 1;

  cell.setValue(incrementValue);
}

function createLineChartWithMultipleSeries(sheetName, maxRow) {
  // 集計結果をグラフにする
  const sheet = getTargetSheet(sheetName);

  // 既存のシートを削除する
  var charts = sheet.getCharts();
  charts.forEach((c) => {
    sheet.removeChart(c);
  })

  // グラフを描画する範囲（日付列 + 書籍の列数分）
  const totalColumns = getTotalColumns();
  const dataRange = sheet.getRange(1, 1, maxRow, totalColumns);

  // グラフの配置位置を動的に計算
  const chartColumn = getChartPositionColumn();

  const chart = sheet.newChart()
  .asBarChart()
  .addRange(dataRange)
  .setChartType(Charts.ChartType.BAR)
  .setPosition(CHART_POSITION_ROW, chartColumn, 0, 0)
  .setOption("useFirstColumnAsDomain", "true")
  .setNumHeaders(1)
  .setOption("legend", {position: "bottom"})
  .build();

  sheet.insertChart(chart)
  sendChartToSlack(chart);
}

function sendChartToSlack(chart) {
  var fileName = "book_chart.png";
  var chartImage = chart.getBlob().getAs("image/png").setName(fileName);
  var fileSize = chartImage.getBytes().length;

  // Step 1: Get Upload URL and File ID
  var uploadInfo = getSlackUploadURL(SLACK_BOT_TOKEN, fileSize, fileName);
  var uploadUrl = uploadInfo.upload_url;
  var fileId = uploadInfo.file_id;

  // Step 2: Upload file to Slack
  uploadFileToSlack(uploadUrl, chartImage);

  // Step 3: Complete the file upload and post it to the channel
  completeSlackFileUpload(SLACK_BOT_TOKEN, fileId, SLACK_CHANNEL_ID);
}

function getSlackUploadURL(token, fileSize, fileName) {
  var url = "https://slack.com/api/files.getUploadURLExternal";
  var payload = {
    'token': token,
    'length': fileSize.toString(),
    'filename': fileName
  };

  var options = {
    'method': 'post',
    'payload': payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  if (json.ok) {
    return {
      'upload_url': json.upload_url,
      'file_id': json.file_id
    };
  } else {
    throw new Error("Failed to get upload URL: " + json.error);
  }
}

function uploadFileToSlack(uploadUrl, fileBlob) {
  var options = {
    'method': 'post',
    'payload': fileBlob
  };

  var response = UrlFetchApp.fetch(uploadUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error("Failed to upload file: " + response.getContentText());
  }
}

function completeSlackFileUpload(token, fileId, channel) {
  var url = "https://slack.com/api/files.completeUploadExternal";
  var payload = {
    'files': [{'id': fileId, 'title': '売り上げチャート'}],
    'channel_id': channel
  };

  var options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  if (!json.ok) {
    throw new Error("Failed to complete file upload: " + json.error);
  }
}

function calcBuyData(message) {
  // 書籍リストから列番号のマップを動的に生成
  const columnMap = createColumnMap();
  const titles = BOOK_TITLES;

  const bookTitle = extBookTitle(message, titles);
  var startDate = getStartDate(EVENT_NAME);
  var diffDays = calcDiffDates(startDate);
  writeDatesFromStartDate(EVENT_NAME);
  incrementCellValue(EVENT_NAME, diffDays+1, bookTitle, columnMap);
  createLineChartWithMultipleSeries(EVENT_NAME, diffDays+1);
}
