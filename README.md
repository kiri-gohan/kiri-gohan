# Claude × Google Drive 連携ツール

Google Drive のファイルを Claude AI で要約・質問応答するCLIツールです。

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
```

## 対応ファイル形式

- Google ドキュメント（テキストとして読み込み）
- Google スプレッドシート（CSVとして読み込み）
- Google スライド（テキストとして読み込み）
- プレーンテキスト（.txt）
- その他のテキストベースのファイル

## ファイル構成

```
├── main.py          # メインCLI
├── google_drive.py  # Google Drive API ラッパー
├── requirements.txt # 依存パッケージ
├── .env             # APIキー（gitignore済み）
├── credentials.json # Google OAuth認証情報（gitignore済み）
└── token.json       # 認証トークン（自動生成・gitignore済み）
```
