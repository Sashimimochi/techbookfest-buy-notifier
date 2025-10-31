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
| `SLACK_WEBHOOK_URLS` | SlackのWebhook URLのリスト（カンマ区切り） | `https://hooks.slack.com/services/YOUR_WEBHOOK_URL_1,https://hooks.slack.com/services/YOUR_WEBHOOK_URL_2` |
| `SLACK_PRIMARY_WEBHOOK_URL` | 画像を投稿するプライマリーのWebhook URL | `https://hooks.slack.com/services/YOUR_PRIMARY_WEBHOOK_URL` |
| `SLACK_BOT_TOKEN` | SlackのBot User OAuth Token | `xoxb-YOUR-BOT-TOKEN-HERE` |
| `SLACK_CHANNEL_ID` | 画像を投稿するSlackチャンネルID | `C0123456789` |
| `SPREADSHEET_ID` | Google SpreadsheetのID | `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms` |
| `EVENT_SHEET_NAME` | イベント名（スプレッドシートのシート名） | `techbookfest16` |
| `BOOK_COLUMN_MAP` | 書籍タイトルと列番号のマッピング（JSON形式） | `{"BookA":2,"BookB":3,"BookC":4,"BookD":5}` |

### 設定例

```
SLACK_WEBHOOK_URLS: https://hooks.slack.com/services/YOUR_WEBHOOK_URL_1
SLACK_PRIMARY_WEBHOOK_URL: https://hooks.slack.com/services/YOUR_PRIMARY_WEBHOOK_URL
SLACK_BOT_TOKEN: xoxb-YOUR-BOT-TOKEN-HERE
SLACK_CHANNEL_ID: C0123456789
SPREADSHEET_ID: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
EVENT_SHEET_NAME: techbookfest16
BOOK_COLUMN_MAP: {"BookA":2,"BookB":3,"BookC":4,"BookD":5}
```

### 注意事項

- スクリプトプロパティはブラウザ上で設定でき、コードに機密情報を含める必要がありません
- 複数のWebhook URLを設定する場合は、カンマ区切りで指定してください
- `BOOK_COLUMN_MAP`はJSON形式で指定してください。書籍タイトルをキーとし、スプレッドシートの列番号を値とします
