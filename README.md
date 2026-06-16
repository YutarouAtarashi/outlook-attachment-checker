# Outlook Smart Alerts 添付忘れチェックアドイン（検証用）

## 重要
このパッケージは、Outlook Smart Alerts / OnMessageSend の個人検証用スターターです。
Outlook on the web へ「ファイルから追加」する場合、基本的に `manifest.xml` を指定します。

ただし、manifest.xml 内の URL は既定で `https://localhost:3000/` になっています。
Outlook on the web で実際に動かすには、`commands.html`、`launchevent.html`、`launchevent.js`、`commands.js`、アイコンファイル等を HTTPS でアクセス可能な場所に配置し、manifest.xml 内の URL をその配置先に変更してください。

## ファイル構成

- `manifest.xml` : Outlookに追加するマニフェストファイル
- `launchevent.js` : 送信時チェック本体
- `launchevent.html` : イベントランタイム用HTML
- `commands.html` : タスクペイン表示用HTML
- `commands.js` : タスクペイン用JS

## 初期検知キーワード

- 添付
- 添付します
- 添付いたします
- 添付ファイル
- 別添
- 別紙
- 資料を送付
- ファイルを送付
- データを送付
- PDF
- Excel
- エクセル
- CSV
- attached
- attachment
- attach
- enclosed
- please find attached
- see attached

## 除外キーワード

- 添付不要
- 添付なし
- 添付は不要
- 資料は後送
- 別途送付
- 後ほど送付
- ファイルなし

## 動作概要

1. メール送信時に OnMessageSend が発火
2. 件名・本文を取得
3. 添付ファイル数を取得
4. 添付関連キーワードあり、除外キーワードなし、添付0件の場合に送信前警告

## 注意

- 実運用前にI情/RPAサポート/SECの確認を受けてください。
- 本文や件名の扱いはセキュリティレビュー対象になり得ます。
- OneDrive/SharePointリンクを添付扱いにするかは追加実装が必要です。
- 初期版ではエラー時は送信を許可する設計にしています。
