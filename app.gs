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
 * スクリプトプロパティから価格マスタシート名を取得する
 * @return {string} 価格マスタシート名
 */
function getPriceMasterSheetName() {
  const properties = PropertiesService.getScriptProperties();
  const sheetName = properties.getProperty('PRICE_MASTER_SHEET_NAME');
  
  // 設定されていない場合はデフォルト値を使用
  return sheetName || 'price_master';
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
        return decodeHtmlEntities(keywords[j]);
      }
    }
  }
  return "販売形態が見つかりません";
}

/**
 * メール本文から価格を抽出する
 * @param {Array<string>} texts メール本文の行配列
 * @return {number|null} 価格（見つからない場合はnull）
 */
function extractPriceFromEmail(texts) {
  // メール本文から「頒布価格」を含む行を探す
  // 全角コロン（：）と半角コロン（:）の両方に対応
  const pricePattern = /頒布価格[：:]\s*(\d+)\s*円/;
  for (var i = 0; i < texts.length; i++) {
    const match = texts[i].match(pricePattern);
    if (match) {
      return parseInt(match[1]);
    }
  }
  return null;
}

/**
 * 価格マスタシートから頒布形式に対応する価格を取得する
 * @param {string} distributionFormat 頒布形式
 * @param {string} bookTitle 書籍タイトル
 * @return {number|null} 価格（見つからない場合はnull）
 */
function getPriceFromMaster(distributionFormat, bookTitle) {
  try {
    const priceMasterSheetName = getPriceMasterSheetName();
    const sheet = getTargetSheet(priceMasterSheetName);
    
    if (!sheet) {
      Logger.log('価格マスタシートが見つかりません: ' + priceMasterSheetName);
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップして検索
    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const masterBookTitle = String(row[0]).trim();
      const masterFormat = String(row[1]).trim();
      const price = row[2];
      
      if (masterBookTitle === bookTitle.trim() && masterFormat === distributionFormat.trim()) {
        return parseInt(price);
      }
    }
    
    Logger.log('価格マスタに該当データが見つかりません: ' + bookTitle + ', ' + distributionFormat);
    return null;
  } catch (e) {
    Logger.log('価格マスタ取得エラー: ' + e.message);
    return null;
  }
}

/**
 * 価格を取得する（メール本文優先、見つからない場合は価格マスタを参照）
 * @param {GmailMessage} message メールメッセージ
 * @param {string} distributionFormat 頒布形式
 * @param {string} bookTitle 書籍タイトル
 * @return {number} 価格（見つからない場合は0）
 */
function getPrice(message, distributionFormat, bookTitle) {
  const texts = getMessageBodyAsText(message);
  
  // まずメール本文から価格を抽出
  const priceFromEmail = extractPriceFromEmail(texts);
  if (priceFromEmail !== null) {
    return priceFromEmail;
  }
  
  // メール本文に価格がない場合は価格マスタから取得
  const priceFromMaster = getPriceFromMaster(distributionFormat, bookTitle);
  if (priceFromMaster !== null) {
    return priceFromMaster;
  }
  
  // どちらも見つからない場合は0を返す
  Logger.log('価格が取得できませんでした: ' + bookTitle + ', ' + distributionFormat);
  return 0;
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
  
  // シートが存在しない場合やA2セルが空の場合は、イベントシートから取得
  if (!sheet) {
    const eventName = getEventName();
    sheet = getTargetSheet(eventName);
  }
  
  if (!sheet) {
    // イベントシートも存在しない場合は今日の日付を返す
    return new Date();
  }

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
  
  // シートが存在しない場合はスキップ
  if (!sheet) {
    return;
  }
  
  const startDate = getStartDate(sheetName);
  const diffDays = calcDiffDates(startDate);

  for (var i = 0; i < diffDays; i++) {
    var row = i + 2;
    var cell = sheet.getRange(row, 1);
    cell.setValue(new Date(startDate.getTime() + i * (1000 * 3600 * 24)));
  }
  if (diffDays === 0) {
    var cell = sheet.getRange(2, 1);
    cell.setValue(new Date());
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

/**
 * セルの値に指定した金額を加算する（売上集計用）
 * @param {string} sheetName シート名
 * @param {number} row 行番号
 * @param {string} columName 列名（書籍タイトル）
 * @param {Object} columnMap 列マップ
 * @param {number} amount 加算する金額
 */
function addCellValue(sheetName, row, columName, columnMap, amount) {
  const sheet = getTargetSheet(sheetName);
  const columnNumber = columnMap[columName];

  if (!columnNumber) {
    throw new Error("指定した書籍名は存在しません");
  }

  const cell = sheet.getRange(row, columnNumber);
  const value = cell.getValue() || 0;
  cell.setValue(value + amount);
}

/**
 * メール受信時刻から時間帯を取得する（0-23）
 * @param {Date} date 日付時刻
 * @return {number} 時間（0-23）
 */
function getHourFromDate(date) {
  return date.getHours();
}

/**
 * 時間帯別シート名を取得する
 * @param {string} eventName イベント名
 * @return {string} 時間帯別シート名
 */
function getHourlySheetName(eventName) {
  return eventName + '_hourly';
}

/**
 * 売上集計シート名を取得する
 * @param {string} eventName イベント名
 * @return {string} 売上集計シート名
 */
function getSalesSheetName(eventName) {
  return eventName + '_sales';
}

/**
 * 時間帯別シートに集計データを書き込む
 * @param {string} sheetName シート名
 * @param {number} hour 時間帯（0-23）
 * @param {string} bookTitle 書籍タイトル
 * @param {Object} columnMap 列マップ
 * @param {number} price 価格
 */
function writeHourlyData(sheetName, hour, bookTitle, columnMap, price) {
  const sheet = getTargetSheet(sheetName);
  
  // 時間帯の行を探す（A列に時間帯が記録されている）
  const lastRow = Math.max(sheet.getLastRow(), 1);
  var targetRow = -1;
  
  for (var i = 2; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (cellValue === hour) {
      targetRow = i;
      break;
    }
  }
  
  // 該当時間帯の行が見つからない場合は新規追加
  if (targetRow === -1) {
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1).setValue(hour);
  }
  
  // 頒布数をインクリメント（書籍列）
  incrementCellValue(sheetName, targetRow, bookTitle, columnMap);
  
  // 売上を加算（書籍列 + オフセット）
  const bookCount = getBookCount();
  const salesColumnMap = {};
  Object.keys(columnMap).forEach(key => {
    salesColumnMap[key] = columnMap[key] + bookCount;
  });
  addCellValue(sheetName, targetRow, bookTitle, salesColumnMap, price);
}

/**
 * 時間帯別シートを初期化する
 * @param {string} sheetName シート名
 */
function initializeHourlySheet(sheetName) {
  const spreadsheetId = getSpreadsheetId();
  const spread = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spread.getSheetByName(sheetName);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spread.insertSheet(sheetName);
    
    // ヘッダー行の設定
    const bookTitles = getBookTitles();
    const bookCount = bookTitles.length;
    sheet.getRange(1, 1).setValue('時間帯');
    
    // 頒布数列のヘッダー
    bookTitles.forEach((title, index) => {
      sheet.getRange(1, index + 2).setValue(title + ' 頒布数');
    });
    
    // 売上列のヘッダー
    bookTitles.forEach((title, index) => {
      sheet.getRange(1, index + 2 + bookCount).setValue(title + ' 売上');
    });
    
    // 0-23時の行を事前に作成
    for (var hour = 0; hour < 24; hour++) {
      sheet.getRange(hour + 2, 1).setValue(hour);
    }
  }
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
  if (!bookTitle) {
    Logger.log('書籍タイトルが見つかりません');
    return;
  }
  
  // 頒布形式を取得
  const texts = getMessageBodyAsText(message);
  const distributionFormat = extMessage(texts);
  
  // 価格を取得
  const price = getPrice(message, distributionFormat, bookTitle);
  
  // メール受信時刻を取得
  const messageDate = message.getDate();
  
  // 開始日と経過日数を取得（全シートで共通）
  var startDate = getStartDate(eventName);
  var diffDays = calcDiffDates(startDate);
  
  // イベントシート（日別頒布数）を初期化
  initializeEventSheet(eventName);
  writeDatesFromStartDate(eventName);
  
  // 日別頒布数を記録（行番号 = diffDays + 2、1行目はヘッダー、2行目から日付データ）
  const targetRow = diffDays + 2;
  incrementCellValue(eventName, targetRow, bookTitle, columnMap);
  
  // 日別売上を記録
  const salesSheetName = getSalesSheetName(eventName);
  initializeSalesSheet(salesSheetName);
  writeDatesFromStartDate(salesSheetName);
  addCellValue(salesSheetName, targetRow, bookTitle, columnMap, price);
  
  // 時間帯別集計（オフライン開催日当日のみ）
  if (isEventDay(messageDate, startDate)) {
    const hourlySheetName = getHourlySheetName(eventName);
    initializeHourlySheet(hourlySheetName);
    const hour = getHourFromDate(messageDate);
    writeHourlyData(hourlySheetName, hour, bookTitle, columnMap, price);
  }
  
  createLineChartWithMultipleSeries(eventName, targetRow);
  createLineChartWithMultipleSeries(salesSheetName, targetRow);
}

/**
 * イベントシート（日別頒布数シート）を初期化する
 * @param {string} sheetName シート名
 */
function initializeEventSheet(sheetName) {
  const spreadsheetId = getSpreadsheetId();
  const spread = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spread.getSheetByName(sheetName);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spread.insertSheet(sheetName);
    
    // ヘッダー行の設定
    const bookTitles = getBookTitles();
    sheet.getRange(1, 1).setValue('日付');
    
    bookTitles.forEach((title, index) => {
      sheet.getRange(1, index + 2).setValue(title);
    });
  }
}

/**
 * 売上集計シートを初期化する
 * @param {string} sheetName シート名
 */
function initializeSalesSheet(sheetName) {
  const spreadsheetId = getSpreadsheetId();
  const spread = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spread.getSheetByName(sheetName);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spread.insertSheet(sheetName);
    
    // ヘッダー行の設定
    const bookTitles = getBookTitles();
    sheet.getRange(1, 1).setValue('日付');
    
    bookTitles.forEach((title, index) => {
      sheet.getRange(1, index + 2).setValue(title);
    });
  }
}

/**
 * 指定日時がイベント開催日かどうかを判定する
 * @param {Date} date 判定する日時
 * @param {Date} eventStartDate イベント開始日
 * @return {boolean} イベント開催日の場合true
 */
function isEventDay(date, eventStartDate) {
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const eventDateStr = Utilities.formatDate(eventStartDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return dateStr === eventDateStr;
}
