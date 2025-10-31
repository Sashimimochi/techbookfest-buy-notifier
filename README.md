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

### 設定例

```
BOOK_TITLES: ["BookA","BookB","BookC","BookD"]
SLACK_WEBHOOK_URLS: [{"url":"https://hooks.slack.com/services/YOUR_URL_1","postChart":true},{"url":"https://hooks.slack.com/services/YOUR_URL_2","postChart":false}]
SLACK_BOT_TOKEN: xoxb-YOUR-BOT-TOKEN-HERE
SLACK_CHANNEL_ID: C0123456789
SPREADSHEET_ID: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
EVENT_SHEET_NAME: techbookfest16
```

### 注意事項

- スクリプトプロパティはブラウザ上で設定でき、コードに機密情報を含める必要がありません
- `BOOK_TITLES`はJSON配列形式で指定してください。新しい書籍を追加する場合はこのリストに追加するだけでOKです
- `SLACK_WEBHOOK_URLS`はJSON配列形式で指定してください。各URLオブジェクトには`url`（Webhook URL）と`postChart`（グラフを投稿するか）を指定します
- 列番号は自動的に計算されるため、書籍の順序はBOOK_TITLESの配列順序で決まります（1列目は日付、2列目から書籍データ）
