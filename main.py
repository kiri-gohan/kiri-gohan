#!/usr/bin/env python3
"""
Claude × Google Drive 連携ツール

Google Drive のファイルを Claude で要約・質問応答するCLIツール。

使い方:
  python main.py list                          # ファイル一覧を表示
  python main.py summarize <file_id>           # ファイルを要約
  python main.py ask <file_id> "<質問>"        # ファイルについて質問
  python main.py search "<キーワード>" ask "<質問>"  # 検索してから質問
"""

import sys
import os
import anthropic
from dotenv import load_dotenv
from google_drive import get_drive_service, list_files, read_file_content

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    print("エラー: ANTHROPIC_API_KEY が設定されていません。.env ファイルを確認してください。")
    sys.exit(1)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def ask_claude(prompt: str) -> str:
    with client.messages.stream(
        model="claude-opus-4-8",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        return stream.get_final_message().content[-1].text


def cmd_list(service, args: list[str]):
    query = None
    if args:
        keyword = " ".join(args)
        query = f"name contains '{keyword}'"
        print(f"「{keyword}」を含むファイルを検索中...\n")
    else:
        print("Google Drive のファイル一覧:\n")

    files = list_files(service, query=query)
    if not files:
        print("ファイルが見つかりませんでした。")
        return

    for i, f in enumerate(files, 1):
        size = f.get("size", "-")
        print(f"[{i:2d}] {f['name']}")
        print(f"      ID: {f['id']}")
        print(f"      種類: {f['mimeType']}")
        print(f"      更新日: {f.get('modifiedTime', '不明')[:10]}")
        print()


def cmd_summarize(service, file_id: str):
    print("ファイルを取得中...")
    files = list_files(service, query=f"'{file_id}' in parents or id = '{file_id}'", max_results=1)
    # ファイルIDで直接取得
    result = service.files().get(fileId=file_id, fields="id,name,mimeType").execute()
    name = result["name"]
    mime = result["mimeType"]

    print(f"「{name}」を読み込み中...")
    content = read_file_content(service, file_id, mime)

    if len(content) > 50000:
        content = content[:50000] + "\n\n[... 長すぎるため省略 ...]"

    print("Claudeで要約中...\n")
    prompt = f"""以下はGoogle Driveのファイル「{name}」の内容です。
日本語で簡潔に要約してください。重要なポイントを箇条書きでまとめてください。

---
{content}
---
"""
    summary = ask_claude(prompt)
    print(f"=== 「{name}」の要約 ===\n")
    print(summary)


def cmd_ask(service, file_id: str, question: str):
    print("ファイルを取得中...")
    result = service.files().get(fileId=file_id, fields="id,name,mimeType").execute()
    name = result["name"]
    mime = result["mimeType"]

    print(f"「{name}」を読み込み中...")
    content = read_file_content(service, file_id, mime)

    if len(content) > 50000:
        content = content[:50000] + "\n\n[... 長すぎるため省略 ...]"

    print("Claudeに質問中...\n")
    prompt = f"""以下はGoogle Driveのファイル「{name}」の内容です。
この内容をもとに、以下の質問に日本語で答えてください。

質問: {question}

---
{content}
---
"""
    answer = ask_claude(prompt)
    print(f"=== 回答 ===\n")
    print(answer)


def print_usage():
    print(__doc__)


def main():
    args = sys.argv[1:]
    if not args:
        print_usage()
        sys.exit(0)

    cmd = args[0]

    try:
        service = get_drive_service()
    except FileNotFoundError as e:
        print(f"エラー: {e}")
        sys.exit(1)

    if cmd == "list":
        cmd_list(service, args[1:])
    elif cmd == "summarize":
        if len(args) < 2:
            print("使い方: python main.py summarize <file_id>")
            sys.exit(1)
        cmd_summarize(service, args[1])
    elif cmd == "ask":
        if len(args) < 3:
            print('使い方: python main.py ask <file_id> "<質問>"')
            sys.exit(1)
        cmd_ask(service, args[1], " ".join(args[2:]))
    else:
        print(f"不明なコマンド: {cmd}")
        print_usage()
        sys.exit(1)


if __name__ == "__main__":
    main()
