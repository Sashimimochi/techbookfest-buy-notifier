# techbookfest-buy-notifier

技術書典の購入履歴をSpreadSheetで管理し、通知にSlackするためのスクリプト。

詳しくは以下参照。
https://zenn.dev/sashimimochi/articles/98d1e41dd5123b

## 設定方法

このスクリプトは、Google Apps Scriptのスクリプトプロパティを使用して機密情報を管理します。
以下の手順に従って、スクリプトプロパティを設定してください。

### スクリプトプロパティの設定手順

1. Google Apps Scriptエディタで、このスクリプト（`app.gs`）を開きます
2. 左側のメニューから「プロジェクトの設定」（歯車アイコン）をクリックします
3. 「スクリプト プロパティ」セクションまでスクロールします
4. 「スクリプト プロパティを追加」をクリックして、以下のプロパティを追加します

### 必要なスクリプトプロパティ

| プロパティ名 | 説明 | 例 |
|------------|------|-----|
| `BOOK_TITLES` | 管理する書籍のタイトルリスト（JSON配列形式） | `["BookA","BookB","BookC","BookD"]` |
| `SLACK_WEBHOOK_URLS` | SlackのWebhook URLのリスト（JSON配列形式、postChart: trueでグラフも投稿） | `[{"url":"https://hooks.slack.com/services/YOUR_URL_1","postChart":true},{"url":"https://hooks.slack.com/services/YOUR_URL_2","postChart":false}]` |
| `SLACK_BOT_TOKEN` | SlackのBot User OAuth Token | `xoxb-YOUR-BOT-TOKEN-HERE` |
| `SLACK_CHANNEL_ID` | 画像を投稿するSlackチャンネルID | `C0123456789` |
| `SPREADSHEET_ID` | Google SpreadsheetのID | `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms` |
| `EVENT_SHEET_NAME` | イベント名（スプレッドシートのシート名） | `techbookfest16` |
| `PRICE_MASTER_SHEET_NAME` | 価格マスタシート名（オプション、デフォルト: `price_master`） | `price_master` |

### 設定例

```
BOOK_TITLES: ["BookA","BookB","BookC","BookD"]
SLACK_WEBHOOK_URLS: [{"url":"https://hooks.slack.com/services/YOUR_URL_1","postChart":true},{"url":"https://hooks.slack.com/services/YOUR_URL_2","postChart":false}]
SLACK_BOT_TOKEN: xoxb-YOUR-BOT-TOKEN-HERE
SLACK_CHANNEL_ID: C0123456789
SPREADSHEET_ID: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
EVENT_SHEET_NAME: techbookfest16
PRICE_MASTER_SHEET_NAME: price_master
```

### 価格マスタシートの設定

価格マスタシートは、頒布形式ごとの価格を管理するためのシートです。以下の形式で作成してください。

#### 価格マスタシートの構造

| 書籍タイトル | 頒布形式 | 価格 |
|------------|---------|------|
| BookA | 電子版 | 500 |
| BookA | 電子+紙 | 1000 |
| BookB | 電子版 | 800 |
| BookB | 電子+紙 | 1500 |

- 1行目: ヘッダー行（書籍タイトル、頒布形式、価格）
- 2行目以降: 各書籍の頒布形式ごとの価格データ

**注意**: 
- メール本文に「頒布価格: XXX円」の記載がある場合、メール本文の価格が優先されます
- メール本文に価格の記載がない場合、価格マスタシートから頒布形式に基づいて価格を取得します
- 価格が取得できない場合は、売上は0円として記録されます

### 生成されるシート

このスクリプトは以下のシートを自動生成します：

1. **{イベント名}** (例: `techbookfest16`)
   - 日別の頒布数を記録
   - 各書籍の日別頒布数とグラフを表示

2. **{イベント名}_sales** (例: `techbookfest16_sales`)
   - 日別の売上を記録
   - 各書籍の日別売上金額とグラフを表示

3. **{イベント名}_hourly** (例: `techbookfest16_hourly`)
   - オフライン開催日当日の時間帯別集計（0-23時）
   - 各書籍の時間帯別頒布数と売上を記録
   - イベント開催日当日のメールのみ集計対象

### 注意事項

- スクリプトプロパティはブラウザ上で設定でき、コードに機密情報を含める必要がありません
- `BOOK_TITLES`はJSON配列形式で指定してください。新しい書籍を追加する場合はこのリストに追加するだけでOKです
- `SLACK_WEBHOOK_URLS`はJSON配列形式で指定してください。各URLオブジェクトには`url`（Webhook URL）と`postChart`（グラフを投稿するか）を指定します
- 列番号は自動的に計算されるため、書籍の順序はBOOK_TITLESの配列順序で決まります（1列目は日付、2列目から書籍データ）
