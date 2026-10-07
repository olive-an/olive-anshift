# レシート読み取り（試作）

スマホで撮ったレシートを Claude が読み取り、経理処理が必要なものだけ一覧・CSVにするアプリです。

## 2つの入口
1. アプリ（index.html）: 撮影 → 読み取り → 確認 → CSV
2. メール: 従業員が olive.an.osaka@gmail.com に写真を送る（件名に拠点名）→ processInbox がスプレッドシートに追記

## 設定手順
1. Googleスプレッドシートを作り、IDを控える
2. そのスプレッドシートの「拡張機能 > Apps Script」に Code.gs を貼る
3. スクリプト プロパティに ANTHROPIC_API_KEY、APP_TOKEN、SHEET_ID を登録
4. 「デプロイ > ウェブアプリ」（実行: 自分、アクセス: 全員）でURLを得る
5. メール入口を使うなら、processInbox を時間主導型トリガー（1時間ごと）に登録
6. アプリを開き「設定」にURLと合言葉を入れる

## 未対応
- TKC取込用の形式（見本待ち）
- 勘定科目は Code.gs の ACCOUNTS を TKC の科目名に合わせて変更
