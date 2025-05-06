# カレンダー同期スクリプト

このGoogle Apps Scriptは、Googleカレンダーのイベントをスプレッドシートに自動的に同期するためのスクリプトです。

## 機能

- カレンダーのイベントをスプレッドシートに自動バックアップ
- 以下の情報を同期:
  - イベントタイトル
  - 開始日時
  - 終了日時
  - 場所
  - 説明
  - 参加者
  - イベントID
- 毎日午前6時に自動同期
- 手動同期機能（メニューから実行可能）

## セットアップ手順

1. Google Apps Scriptプロジェクトを作成
2. このスクリプトをコピー＆ペースト
3. カレンダーIDの設定:
   ```javascript
   function setupCalendarId(calendarId) {
     const scriptProperties = PropertiesService.getScriptProperties();
     scriptProperties.setProperty('CALENDAR_ID', calendarId);
   }
   ```
   - `calendarId`は同期したいカレンダーのIDを指定
   - カレンダーIDはカレンダーの設定から取得可能

4. 自動同期の設定:
   ```javascript
   function setupDailyTrigger() {
     deleteTriggers();
     ScriptApp.newTrigger('backupCalendarEvents')
       .timeBased()
       .everyDays(1)
       .atHour(6)
       .create();
   }
   ```

## 使用方法

1. スプレッドシートを開く
2. メニューバーに「カレンダー同期」が表示される
3. 「今すぐ同期」をクリックして手動で同期を実行

## 注意事項

- スクリプトを初めて実行する際は、必要な権限の承認が必要です
- カレンダーIDが正しく設定されていない場合、エラーが発生します
- 同期は現在の日付から2099年12月31日までのイベントを対象とします

## トラブルシューティング

1. カレンダーIDが見つからないエラー:
   - スクリプトプロパティに`CALENDAR_ID`が正しく設定されているか確認
   - カレンダーIDが正しいか確認

2. カレンダーが見つからないエラー:
   - 指定したカレンダーIDが存在するか確認
   - カレンダーへのアクセス権限があるか確認

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。 