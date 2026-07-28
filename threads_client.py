import requests

GRAPH_BASE = "https://graph.threads.net/v1.0"


def create_text_post(access_token: str, user_id: str, text: str) -> str:
    """テキストのみのThreads投稿を作成して公開し、投稿IDを返す。"""
    container = requests.post(
        f"{GRAPH_BASE}/{user_id}/threads",
        data={
            "media_type": "TEXT",
            "text": text,
            "access_token": access_token,
        },
    )
    container.raise_for_status()
    creation_id = container.json()["id"]

    publish = requests.post(
        f"{GRAPH_BASE}/{user_id}/threads_publish",
        data={
            "creation_id": creation_id,
            "access_token": access_token,
        },
    )
    publish.raise_for_status()
    return publish.json()["id"]


def get_permalink(access_token: str, media_id: str) -> str:
    resp = requests.get(
        f"{GRAPH_BASE}/{media_id}",
        params={"fields": "permalink", "access_token": access_token},
    )
    resp.raise_for_status()
    return resp.json().get("permalink", "")
