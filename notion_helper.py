import os
from notion_client import Client


def get_notion_client() -> Client:
    token = os.getenv("NOTION_TOKEN")
    if not token:
        raise ValueError(
            "NOTION_TOKEN が設定されていません。.env ファイルを確認してください。"
        )
    return Client(auth=token)


def list_pages(client: Client, max_results: int = 20) -> list[dict]:
    response = client.search(
        filter={"property": "object", "value": "page"},
        page_size=max_results,
    )
    pages = []
    for item in response.get("results", []):
        title = _get_title(item)
        pages.append({
            "id": item["id"],
            "title": title,
            "url": item.get("url", ""),
            "last_edited": item.get("last_edited_time", "")[:10],
        })
    return pages


def read_page_content(client: Client, page_id: str) -> tuple[str, str]:
    page = client.pages.retrieve(page_id=page_id)
    title = _get_title(page)

    blocks = client.blocks.children.list(block_id=page_id)
    lines = []
    for block in blocks.get("results", []):
        text = _extract_block_text(block)
        if text:
            lines.append(text)

    return title, "\n".join(lines)


def create_page(client: Client, parent_page_id: str, title: str, content: str) -> str:
    children = _text_to_blocks(content)
    response = client.pages.create(
        parent={"type": "page_id", "page_id": parent_page_id},
        properties={
            "title": {"title": [{"text": {"content": title}}]}
        },
        children=children,
    )
    return response["url"]


def append_to_page(client: Client, page_id: str, content: str) -> None:
    blocks = _text_to_blocks(content)
    client.blocks.children.append(block_id=page_id, children=blocks)


def _get_title(item: dict) -> str:
    props = item.get("properties", {})
    for prop in props.values():
        if prop.get("type") == "title":
            texts = prop.get("title", [])
            return "".join(t.get("plain_text", "") for t in texts)
    return "（タイトルなし）"


def _extract_block_text(block: dict) -> str:
    block_type = block.get("type", "")
    data = block.get(block_type, {})
    rich_texts = data.get("rich_text", [])
    text = "".join(t.get("plain_text", "") for t in rich_texts)

    prefix_map = {
        "heading_1": "# ",
        "heading_2": "## ",
        "heading_3": "### ",
        "bulleted_list_item": "・",
        "numbered_list_item": "1. ",
        "to_do": "☐ ",
        "quote": "> ",
        "code": "```\n",
    }
    suffix_map = {"code": "\n```"}

    prefix = prefix_map.get(block_type, "")
    suffix = suffix_map.get(block_type, "")
    return f"{prefix}{text}{suffix}" if text else ""


def _text_to_blocks(content: str) -> list[dict]:
    blocks = []
    for line in content.split("\n"):
        line = line.rstrip()
        if line.startswith("# "):
            blocks.append(_heading_block(1, line[2:]))
        elif line.startswith("## "):
            blocks.append(_heading_block(2, line[3:]))
        elif line.startswith("### "):
            blocks.append(_heading_block(3, line[4:]))
        elif line.startswith("・") or line.startswith("- "):
            text = line[1:].lstrip() if line.startswith("・") else line[2:]
            blocks.append(_bullet_block(text))
        else:
            blocks.append(_paragraph_block(line))
    return blocks


def _rich_text(text: str) -> list[dict]:
    return [{"type": "text", "text": {"content": text}}]


def _paragraph_block(text: str) -> dict:
    return {"object": "block", "type": "paragraph",
            "paragraph": {"rich_text": _rich_text(text)}}


def _heading_block(level: int, text: str) -> dict:
    t = f"heading_{level}"
    return {"object": "block", "type": t, t: {"rich_text": _rich_text(text)}}


def _bullet_block(text: str) -> dict:
    return {"object": "block", "type": "bulleted_list_item",
            "bulleted_list_item": {"rich_text": _rich_text(text)}}
