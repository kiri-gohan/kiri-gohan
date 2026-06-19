#!/usr/bin/env python3
"""
Claude × Google Drive / Notion 連携ツール

使い方:

  --- Google Drive ---
  python main.py list                               # ファイル一覧を表示
  python main.py summarize <file_id>                # ファイルを要約
  python main.py ask <file_id> "<質問>"             # ファイルについて質問

  --- Notion ---
  python main.py notion list                        # ページ一覧を表示
  python main.py notion summarize <page_id>         # ページを要約
  python main.py notion ask <page_id> "<質問>"      # ページについて質問
  python main.py notion write <page_id> "<指示>"    # ページに追記
  python main.py notion create <page_id> "<タイトル>" "<指示>"  # 新規ページを作成
"""

import sys
import os
import anthropic
from dotenv import load_dotenv
from google_drive import get_drive_service, list_files, read_file_content
from notion_helper import (
    get_notion_client,
    list_pages,
    read_page_content,
    create_page,
    append_to_page,
)

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


# ─── Google Drive コマンド ───────────────────────────────────────────────────

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
        print(f"[{i:2d}] {f['name']}")
        print(f"      ID: {f['id']}")
        print(f"      種類: {f['mimeType']}")
        print(f"      更新日: {f.get('modifiedTime', '不明')[:10]}")
        print()


def cmd_summarize(service, file_id: str):
    print("ファイルを取得中...")
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
    print(f"=== 「{name}」の要約 ===\n")
    print(ask_claude(prompt))


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
    print("=== 回答 ===\n")
    print(ask_claude(prompt))


# ─── Notion コマンド ─────────────────────────────────────────────────────────

def cmd_notion_list(notion):
    print("Notion のページ一覧:\n")
    pages = list_pages(notion)
    if not pages:
        print("ページが見つかりませんでした。")
        return
    for i, p in enumerate(pages, 1):
        print(f"[{i:2d}] {p['title']}")
        print(f"      ID: {p['id']}")
        print(f"      更新日: {p['last_edited']}")
        print()


def cmd_notion_summarize(notion, page_id: str):
    print("ページを取得中...")
    title, content = read_page_content(notion, page_id)
    if len(content) > 50000:
        content = content[:50000] + "\n\n[... 長すぎるため省略 ...]"

    print(f"「{title}」を要約中...\n")
    prompt = f"""以下はNotionのページ「{title}」の内容です。
日本語で簡潔に要約してください。重要なポイントを箇条書きでまとめてください。

---
{content}
---
"""
    print(f"=== 「{title}」の要約 ===\n")
    print(ask_claude(prompt))


def cmd_notion_ask(notion, page_id: str, question: str):
    print("ページを取得中...")
    title, content = read_page_content(notion, page_id)
    if len(content) > 50000:
        content = content[:50000] + "\n\n[... 長すぎるため省略 ...]"

    print("Claudeに質問中...\n")
    prompt = f"""以下はNotionのページ「{title}」の内容です。
この内容をもとに、以下の質問に日本語で答えてください。

質問: {question}

---
{content}
---
"""
    print("=== 回答 ===\n")
    print(ask_claude(prompt))


def cmd_notion_write(notion, page_id: str, instruction: str):
    print("ページを取得中...")
    title, content = read_page_content(notion, page_id)

    print("Claudeで文章を生成中...\n")
    prompt = f"""以下はNotionのページ「{title}」の現在の内容です。

---
{content}
---

次の指示に従って、このページに追記する文章を日本語で書いてください。
マークダウン形式（# 見出し、- 箇条書きなど）で書いてください。

指示: {instruction}
"""
    new_content = ask_claude(prompt)
    print("生成した内容:\n")
    print(new_content)
    print("\nNotionに追記中...")
    append_to_page(notion, page_id, new_content)
    print("追記しました！")


def cmd_notion_create(notion, parent_page_id: str, title: str, instruction: str):
    print("Claudeで文章を生成中...\n")
    prompt = f"""次の指示に従って、Notionページの内容を日本語で書いてください。
マークダウン形式（# 見出し、- 箇条書きなど）で書いてください。

タイトル: {title}
指示: {instruction}
"""
    content = ask_claude(prompt)
    print("生成した内容:\n")
    print(content)
    print("\nNotionにページを作成中...")
    url = create_page(notion, parent_page_id, title, content)
    print(f"作成しました！\nURL: {url}")


def cmd_notion(args: list[str]):
    if not args:
        print("使い方: python main.py notion <サブコマンド>")
        print("サブコマンド: list / summarize / ask / write / create")
        sys.exit(1)

    try:
        notion = get_notion_client()
    except ValueError as e:
        print(f"エラー: {e}")
        sys.exit(1)

    sub = args[0]
    rest = args[1:]

    if sub == "list":
        cmd_notion_list(notion)
    elif sub == "summarize":
        if not rest:
            print("使い方: python main.py notion summarize <page_id>")
            sys.exit(1)
        cmd_notion_summarize(notion, rest[0])
    elif sub == "ask":
        if len(rest) < 2:
            print('使い方: python main.py notion ask <page_id> "<質問>"')
            sys.exit(1)
        cmd_notion_ask(notion, rest[0], " ".join(rest[1:]))
    elif sub == "write":
        if len(rest) < 2:
            print('使い方: python main.py notion write <page_id> "<指示>"')
            sys.exit(1)
        cmd_notion_write(notion, rest[0], " ".join(rest[1:]))
    elif sub == "create":
        if len(rest) < 3:
            print('使い方: python main.py notion create <page_id> "<タイトル>" "<指示>"')
            sys.exit(1)
        cmd_notion_create(notion, rest[0], rest[1], " ".join(rest[2:]))
    else:
        print(f"不明なサブコマンド: {sub}")
        sys.exit(1)


# ─── メイン ──────────────────────────────────────────────────────────────────

def print_usage():
    print(__doc__)


def main():
    args = sys.argv[1:]
    if not args:
        print_usage()
        sys.exit(0)

    cmd = args[0]

    if cmd == "notion":
        cmd_notion(args[1:])
        return

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
