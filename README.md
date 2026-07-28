# Claude × Google Drive / Threads 連携ツール

Google Drive のファイルを Claude AI で要約・質問応答するCLIツールです。
また、Threads への投稿下書き生成・投稿にも対応しています。

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 2. Anthropic API キーの設定

`.env.example` をコピーして `.env` を作成し、APIキーを設定します。

```bash
cp .env.example .env
# .env を編集して ANTHROPIC_API_KEY を設定
```

### 3. Google Drive API の認証情報を取得

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. プロジェクトを作成（または既存のものを選択）
3. **APIとサービス** → **有効なAPI** → **Google Drive API** を有効化
4. **APIとサービス** → **認証情報** → **認証情報を作成** → **OAuthクライアントID**
5. アプリケーションの種類: **デスクトップアプリ**
6. 作成した認証情報をダウンロードし、`credentials.json` として保存

### 4. 初回認証

初回実行時にブラウザが開き、Googleアカウントでの認証が求められます。
認証後、`token.json` が自動生成され、次回以降は自動ログインします。

### 5. Threads API の認証情報を取得

1. [Meta for Developers](https://developers.facebook.com/apps) にアクセスし、Meta（Facebook）アカウントでログイン
2. **アプリを作成** → ユースケースで **「Threads API のアクセス」** を選択
3. 作成したアプリの **Threads API** → **設定** 画面で:
   - **Threads Tester** に自分のThreadsアカウント（例: `kiri_gohan_`）を追加
   - 追加後、Threadsアプリ側（設定 → アカウント → アプリとWebサイトの権限 → テスター招待）で招待を承認する
   - **有効なOAuthリダイレクトURI** に任意のURL（例: `https://localhost/callback`）を登録
     - 実際にサーバーを立てる必要はありません。認可後にブラウザがこのURLへリダイレクトされ、読み込みエラーになっても構いません。アドレスバーに表示される `code=` の値を使います。
4. **アプリID** と **アプリシークレット** を控える
5. `.env` に以下を設定
   ```bash
   THREADS_APP_ID=取得したアプリID
   THREADS_APP_SECRET=取得したアプリシークレット
   THREADS_REDIRECT_URI=登録したリダイレクトURI
   ```
6. 初回投稿時（`python main.py threads post`）に認可URLが表示されるので、ブラウザで開いて認可し、リダイレクト後のURLに含まれる `code` の値をターミナルに貼り付けてください。
   以降は `threads_token.json` に保存されたトークンを自動更新しながら使用します（60日ごとに自動リフレッシュ）。

## 使い方

```bash
# Google Drive のファイル一覧を表示
python main.py list

# キーワードでファイルを検索
python main.py list 議事録

# ファイルを要約（file_id は list コマンドで確認）
python main.py summarize <file_id>

# ファイルについて質問
python main.py ask <file_id> "このドキュメントの主な結論は何ですか？"

# テーマを指定してThreads投稿の下書きを生成
python main.py threads draft "ファスティングと体の力"

# 下書きの内容を確認し、承認してThreadsに投稿
python main.py threads post
```

`threads draft` は下書きを `threads_draft.json` に保存するだけで投稿はしません。
`threads post` は保存された下書きを表示し、`y` で承認した場合のみ実際に投稿します。

## 対応ファイル形式

- Google ドキュメント（テキストとして読み込み）
- Google スプレッドシート（CSVとして読み込み）
- Google スライド（テキストとして読み込み）
- プレーンテキスト（.txt）
- その他のテキストベースのファイル

## ファイル構成

```
├── main.py            # メインCLI
├── google_drive.py    # Google Drive API ラッパー
├── threads_auth.py    # Threads API OAuth認証
├── threads_client.py  # Threads API 投稿ラッパー
├── requirements.txt   # 依存パッケージ
├── .env                # APIキー（gitignore済み）
├── credentials.json    # Google OAuth認証情報（gitignore済み）
├── token.json           # Google 認証トークン（自動生成・gitignore済み）
├── threads_token.json   # Threads 認証トークン（自動生成・gitignore済み）
└── threads_draft.json   # Threads 下書き（自動生成・gitignore済み）
```
