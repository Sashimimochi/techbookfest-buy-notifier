// ========================================
// 設定セクション
// ========================================

/**
 * スクリプトプロパティから書籍リストを取得する
 * @return {Array<string>} 書籍タイトルの配列
 */
function getBookTitles() {
  const properties = PropertiesService.getScriptProperties();
  const bookTitlesStr = properties.getProperty('BOOK_TITLES');
  
  if (!bookTitlesStr) {
    throw new Error('BOOK_TITLESが設定されていません。スクリプトプロパティに設定してください。');
  }
  
  return JSON.parse(bookTitlesStr);
}

/**
 * スクリプトプロパティからスプレッドシートIDを取得する
 * @return {string} スプレッドシートID
 */
function getSpreadsheetId() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_IDが設定されていません。スクリプトプロパティに設定してください。');
  }
  
  return spreadsheetId;
}

/**
 * スクリプトプロパティからイベント名を取得する
 * @return {string} イベント名（シート名）
 */
function getEventName() {
  const properties = PropertiesService.getScriptProperties();
  const eventName = properties.getProperty('EVENT_SHEET_NAME');
  
  if (!eventName) {
    throw new Error('EVENT_SHEET_NAMEが設定されていません。スクリプトプロパティに設定してください。');
  }
  
  return eventName;
}

/**
 * スクリプトプロパティからSlack Webhook URLsを取得する
 * @return {Array<{url: string, postChart: boolean}>} Webhook URLの配列
 */
function getSlackWebhookUrls() {
  const properties = PropertiesService.getScriptProperties();
  const webhookUrlsStr = properties.getProperty('SLACK_WEBHOOK_URLS');
  
  if (!webhookUrlsStr) {
    throw new Error('SLACK_WEBHOOK_URLsが設定されていません。スクリプトプロパティに設定してください。');
  }
  
  return JSON.parse(webhookUrlsStr);
}

/**
 * スクリプトプロパティからSlack Bot Tokenを取得する
 * @return {string} Slack Bot Token
 */
function getSlackBotToken() {
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('SLACK_BOT_TOKEN');
  
  if (!token) {
    throw new Error('SLACK_BOT_TOKENが設定されていません。スクリプトプロパティに設定してください。');
  }
  
  return token;
}

/**
 * スクリプトプロパティからSlack Channel IDを取得する
 * @return {string} Slack Channel ID
 */
function getSlackChannelId() {
  const properties = PropertiesService.getScriptProperties();
  const channelId = properties.getProperty('SLACK_CHANNEL_ID');
  
  if (!channelId) {
    throw new Error('SLACK_CHANNEL_IDが設定されていません。スクリプトプロパティに設定してください。');
  }
  
  return channelId;
}

// グラフ配置の設定
const CHART_POSITION_ROW = 2;
const CHART_POSITION_COLUMN_OFFSET = 2; // 書籍列の右側に配置するための列オフセット

// ========================================
// ヘルパー関数
// ========================================

/**
 * 書籍リストから列番号のマップを生成する
 * @return {Object} 書籍名をキー、列番号（1列目は日付、2列目から書籍データ）を値とするオブジェクト
 */
function createColumnMap() {
  const bookTitles = getBookTitles();
  const columnMap = {};
  bookTitles.forEach((title, index) => {
    columnMap[title] = index + 2; // 列番号は2から開始（1列目は日付用）
  });
  return columnMap;
}

/**
 * 書籍の総数を取得する
 * @return {number} 書籍の総数
 */
function getBookCount() {
  return getBookTitles().length;
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
      const webhookUrls = getSlackWebhookUrls();
      webhookUrls.forEach((webhook) => {
        sendTextToSlack(text, webhook.url);
        if (webhook.postChart) { // グラフを投稿するフラグが立っている場合
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
  const spreadsheetId = getSpreadsheetId();
  var spread = SpreadsheetApp.openById(spreadsheetId);
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

  const token = getSlackBotToken();
  const channel = getSlackChannelId();

  // Step 1: Get Upload URL and File ID
  var uploadInfo = getSlackUploadURL(token, fileSize, fileName);
  var uploadUrl = uploadInfo.upload_url;
  var fileId = uploadInfo.file_id;

  // Step 2: Upload file to Slack
  uploadFileToSlack(uploadUrl, chartImage);

  // Step 3: Complete the file upload and post it to the channel
  completeSlackFileUpload(token, fileId, channel);
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
  const titles = getBookTitles();
  const eventName = getEventName();

  const bookTitle = extBookTitle(message, titles);
  var startDate = getStartDate(eventName);
  var diffDays = calcDiffDates(startDate);
  writeDatesFromStartDate(eventName);
  incrementCellValue(eventName, diffDays+1, bookTitle, columnMap);
  createLineChartWithMultipleSeries(eventName, diffDays+1);
}
