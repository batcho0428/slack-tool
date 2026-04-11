# Slackがちょっと便利になりそうなツール

学生団体でSlackを使用しているとき、チャンネルのメンバー追加や複数人へのDM送信が面倒だな～と感じて作ったツールです。

## 現在の構成

- バックエンド: Google Apps Script
- フロントエンド: Next.js (`frontend`)
- フロントからGASへの接続: Next.js Route Handler (`frontend/app/api/gas/route.ts`) 経由
- コンテナ: Docker / Docker Compose

## 追加した主なファイル

- GAS APIエンドポイント: `Api.gs`
- Next.jsアプリ: `frontend/` 一式
- Docker: `docker-compose.yml`, `frontend/Dockerfile`

## Google Apps Script 側の準備

以下は「初期セットアップを最初に1回行う」ための手順です。

### 1. Apps Scriptプロジェクトを開く

1. Google Apps Script エディタで対象プロジェクトを開く
2. `Code.gs` があることを確認する
3. `Api.gs` を追加して保存する

### 2. スクリプトプロパティを設定する

1. 左メニューの歯車アイコン（プロジェクトの設定）を開く
2. 「スクリプト プロパティ」セクションで「スクリプト プロパティを追加」を押す
3. 以下のキー/値を追加して保存する

- キー: `FRONTEND_API_SHARED_SECRET`
- 値: 推測されにくいランダムな長い文字列

例:

```text
FRONTEND_API_SHARED_SECRET=3x0h...十分に長いランダム値...
```

この値は、後でフロントエンド側（Cloudflare Pages の環境変数 / ローカルの `.env`）にも同じ値を設定します。

このプロジェクトで使う Script Properties は、まとめて次のとおりです。

#### フロントエンド連携用

- `FRONTEND_API_SHARED_SECRET`: フロントエンドとGASのAPI認証に使う共有シークレット

#### 基本設定

- `SPREADSHEET_ID`: Slack名簿や各種データを保存しているスプレッドシートID
- `SLACK_BOT_TOKEN`: Slack API 呼び出し用の Bot Token
- `SLACK_CLIENT_ID`: Slack OAuth の Client ID
- `SLACK_CLIENT_SECRET`: Slack OAuth の Client Secret
- `SHAREABLE_URL`: 共有リンクや送信文面で使うURL

#### 機能に応じて必要な設定

- `EXPORT_SHARED_FOLDER_ID`: 名簿エクスポートで使う共有フォルダID
- `NEXT_YEAR_CONTINUE_ENABLED`: 次年度継続の機能を有効化する場合に `true`

#### 初期設定の順番

1. `FRONTEND_API_SHARED_SECRET` を決める
2. `SPREADSHEET_ID` と Slack 認証系の値を入れる
3. 共有リンクや出力先フォルダを使う機能があれば `SHAREABLE_URL` と `EXPORT_SHARED_FOLDER_ID` を入れる
4. 次年度継続を使うなら `NEXT_YEAR_CONTINUE_ENABLED=true` を設定する
5. すべての値は GAS の Script Properties に保存し、Git には含めない

例:

```text
FRONTEND_API_SHARED_SECRET=replace-with-long-random-secret
SPREADSHEET_ID=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
SLACK_BOT_TOKEN=xoxb-xxxxxxxxxxxxxxxxxxxxxxxxxxxx
SLACK_CLIENT_ID=1234567890.1234567890
SLACK_CLIENT_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
SHAREABLE_URL=https://example.com/share
EXPORT_SHARED_FOLDER_ID=xxxxxxxxxxxxxxxxxxxxxxxxxxxx
NEXT_YEAR_CONTINUE_ENABLED=true
```

### 3. Webアプリとしてデプロイする

1. 右上の「デプロイ」→「新しいデプロイ」を選択
2. 種類で「ウェブアプリ」を選択
3. 実行ユーザーとアクセス権を要件に合わせて設定
4. 「デプロイ」を実行
5. 発行されたWebアプリURLを控える

### 4. 動作確認する

1. フロントエンドからAPIを呼び出す
2. GAS側で `Unauthorized request` が出ないことを確認する
3. もし `Unauthorized request` が出る場合は、
	- GASの `FRONTEND_API_SHARED_SECRET`
	- フロントエンドの `FRONTEND_API_SHARED_SECRET`
	の一致を確認する

### 5. 追加で有効化が必要なサービス

このプロジェクトでは、通常の `SpreadsheetApp` や `UrlFetchApp` は追加設定なしで使えます。
ただし、`Drive.Files` を使う機能があるため、Google Apps Script の **高度な Google サービス** で **Drive API** を有効化してください。

有効化の流れは次のとおりです。

1. Apps Script エディタの左メニューで「サービス」を開く
2. 一覧から「Drive API」を追加する
3. 必要なら Google Cloud 側でも同じプロジェクトで Drive API を有効化する
4. 保存して再デプロイする

補足:

- これは名簿エクスポートやフォルダ参照など、`Drive.Files` を使う処理で必要です
- `SpreadsheetApp` を使うだけの処理には追加サービスは不要です
- `PropertiesService` も追加サービスではなく、そのまま利用できます

## Slack API アプリの準備

Slack 連携に必要なアプリ設定は、最初に 1 回だけ行います。

### 1. Slackアプリを作成する

1. Slack API の管理画面で「Create New App」を開く
2. 「From scratch」を選ぶ
3. App Name を入力する
4. 対象の Workspace を選んで作成する

### 2. OAuth設定をする

1. 左メニューで「OAuth & Permissions」を開く
2. 「Redirect URLs」に Google Apps Script の WebアプリURLを追加する
	- ここには `ScriptApp.getService().getUrl()` で返るURLを入れる
	- 例: `https://script.google.com/macros/s/XXXX/exec`
3. 「User Token Scopes」に次の権限を追加する
	- `chat:write`
	- `users:read`
	- `users:read.email`
	- `channels:read`
	- `groups:read`
	- `channels:write`
	- `groups:write`
4. 必要に応じて「Bot Token Scopes」に次の権限も追加する
	- `chat:write`
	- `users:read`
	- `users:read.email`
	- `channels:read`
	- `groups:read`
	- `channels:write`
	- `groups:write`
5. 「Install to Workspace」または再インストールを実行する
6. 表示された `Client ID` と `Client Secret` を控える

### 3. Script Properties と対応させる

Slackアプリの設定値は、GAS の Script Properties に以下のように対応させます。

- `SLACK_CLIENT_ID`: Slackアプリの Client ID
- `SLACK_CLIENT_SECRET`: Slackアプリの Client Secret
- `SLACK_BOT_TOKEN`: Slackアプリをワークスペースにインストールした後に取得できる Bot Token

### 4. 動作確認する

1. GAS の Webアプリを開く
2. Slackログイン画面から認証を実行する
3. `users.lookupByEmail` や `chat.postMessage` が成功することを確認する
4. 権限不足エラーが出る場合は、Slackアプリの scopes を見直す

補足:

- このプロジェクトでは Events API や Slash Commands は使っていません
- App Home や Socket Mode も必須ではありません
- Slack 側の設定は、OAuth と必要な scopes を中心に合わせれば動きます

## ローカル（Dockerのみ）準備

1. `.env.example` を `.env` にコピー
2. 以下を設定
	- `GAS_WEB_APP_URL`: Apps Script の WebアプリURL
	- `FRONTEND_API_SHARED_SECRET`: GAS側と同じ値
	- `NGROK_AUTHTOKEN`: ngrokのトークン

例:

```env
GAS_WEB_APP_URL=https://script.google.com/macros/s/XXXX/exec
FRONTEND_API_SHARED_SECRET=replace-with-long-random-secret
NGROK_AUTHTOKEN=replace-with-ngrok-authtoken
```

`.env` はGit管理対象外です（`.gitignore` で除外）。

## Cloudflare Pages デプロイ時の設定

Cloudflare Pages では、ローカルの `.env` をアップロードせず、Pagesの環境変数機能で値を設定します。

### 1. デプロイ設定

Cloudflare Pages で新規プロジェクトを作成し、GitHubリポジトリを接続して以下を設定してください。

- Framework preset: Next.js
- Root directory: `frontend`
- Build command: `npm run build`
- Build output directory: `.next`

※ Next.js向け設定が自動検出される場合は、その値を優先して問題ありません。

### 2. 環境変数の設定

Cloudflare Pages の Project Settings → Environment Variables で、以下を設定してください（Production/Preview 両方推奨）。

- `GAS_WEB_APP_URL`
- `FRONTEND_API_SHARED_SECRET`

これにより、フロントエンドからGASへ送る全APIリクエストに共有シークレットが付与され、
GAS側で未認証リクエストを拒否します。

### 3. envファイルの扱い

- ローカル開発用の `.env` はGit管理対象外です。
- 本番・プレビュー用の値は、Cloudflare Pages の Environment Variables にのみ設定してください。
- リポジトリにはサンプルとして `.env.example` のみを置きます。

### 4. デプロイ後の確認

デプロイ後、Pagesの公開URLで以下を確認してください。

- 画面表示ができること
- ログインなどAPI通信が成功すること
- GAS側で `Unauthorized request` が出ていないこと

## Docker起動

```bash
docker compose build
docker compose up
```

これで `next` と `ngrok` が起動します。

- アプリ: `http://localhost:3333`
- ngrok管理画面: `http://localhost:4040`

## ngrokでDockerフロントエンドを外部公開

Slack認証などで公開URLが必要な場合に使います。

1. `.env` に `NGROK_AUTHTOKEN` を設定
2. `docker compose build && docker compose up` で起動
3. 公開URL確認

- ターミナル出力、または `http://localhost:4040` のngrok管理画面から確認

補足: `NGROK_AUTHTOKEN` が未設定でもcompose起動は継続します（ngrokトンネルは無効化）。

## コンポーネント分割方針

Next.js側は機能ごとに分割済みです。

- 認証: `frontend/components/auth`
- レイアウト: `frontend/components/layout`
- 共通UI: `frontend/components/common`
- タブ機能: `frontend/components/tabs`
- APIクライアント: `frontend/lib/api.ts`

現在、主要機能は次の状態です。

- 実装済み: ログイン、DM送信、チャンネル招待、マイページ基本更新
- 段階移行: 名簿、アンケート、集金（既存GAS APIを使って順次移行）
