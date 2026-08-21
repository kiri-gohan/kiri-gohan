import json
import os
import time
import urllib.parse

import requests

AUTH_BASE = "https://threads.net/oauth/authorize"
GRAPH_BASE = "https://graph.threads.net"
TOKEN_FILE = "threads_token.json"

SCOPES = "threads_basic,threads_content_publish"


def _load_token() -> dict | None:
    if not os.path.exists(TOKEN_FILE):
        return None
    with open(TOKEN_FILE) as f:
        return json.load(f)


def _save_token(data: dict):
    with open(TOKEN_FILE, "w") as f:
        json.dump(data, f)


def build_authorize_url(app_id: str, redirect_uri: str) -> str:
    params = {
        "client_id": app_id,
        "redirect_uri": redirect_uri,
        "scope": SCOPES,
        "response_type": "code",
    }
    return f"{AUTH_BASE}?{urllib.parse.urlencode(params)}"


def exchange_code_for_token(app_id: str, app_secret: str, redirect_uri: str, code: str) -> dict:
    resp = requests.post(
        f"{GRAPH_BASE}/oauth/access_token",
        data={
            "client_id": app_id,
            "client_secret": app_secret,
            "grant_type": "authorization_code",
            "redirect_uri": redirect_uri,
            "code": code,
        },
    )
    resp.raise_for_status()
    return resp.json()


def exchange_for_long_lived_token(app_secret: str, short_lived_token: str) -> dict:
    resp = requests.get(
        f"{GRAPH_BASE}/access_token",
        params={
            "grant_type": "th_exchange_token",
            "client_secret": app_secret,
            "access_token": short_lived_token,
        },
    )
    resp.raise_for_status()
    return resp.json()


def refresh_long_lived_token(access_token: str) -> dict:
    resp = requests.get(
        f"{GRAPH_BASE}/refresh_access_token",
        params={
            "grant_type": "th_refresh_token",
            "access_token": access_token,
        },
    )
    resp.raise_for_status()
    return resp.json()


def run_authorization_flow(app_id: str, app_secret: str, redirect_uri: str) -> dict:
    url = build_authorize_url(app_id, redirect_uri)
    print("以下のURLをブラウザで開き、Threadsアカウントで認可してください:\n")
    print(url)
    print(
        f"\n認可後、ブラウザは {redirect_uri} へリダイレクトされます"
        "（ページ自体は表示されなくても問題ありません）。"
        "\nアドレスバーのURLに含まれる `code=` の値をコピーして貼り付けてください。"
    )
    code = input("\ncode: ").strip()
    if code.startswith("http"):
        parsed = urllib.parse.urlparse(code)
        code = urllib.parse.parse_qs(parsed.query).get("code", [""])[0]

    short_lived = exchange_code_for_token(app_id, app_secret, redirect_uri, code)
    long_lived = exchange_for_long_lived_token(app_secret, short_lived["access_token"])

    token_data = {
        "access_token": long_lived["access_token"],
        "user_id": short_lived["user_id"],
        "obtained_at": time.time(),
        "expires_in": long_lived.get("expires_in", 60 * 24 * 60 * 60),
    }
    _save_token(token_data)
    return token_data


def get_access_token() -> dict:
    """有効なアクセストークンを返す。必要なら自動更新し、なければ認可フローを実行する。"""
    app_id = os.getenv("THREADS_APP_ID")
    app_secret = os.getenv("THREADS_APP_SECRET")
    redirect_uri = os.getenv("THREADS_REDIRECT_URI")
    if not all([app_id, app_secret, redirect_uri]):
        raise EnvironmentError(
            "THREADS_APP_ID / THREADS_APP_SECRET / THREADS_REDIRECT_URI が .env に設定されていません。"
        )

    token_data = _load_token()

    if token_data is None:
        return run_authorization_flow(app_id, app_secret, redirect_uri)

    age = time.time() - token_data["obtained_at"]
    expires_in = token_data["expires_in"]

    # 期限まで7日を切ったら更新（更新にはトークンが24時間以上経過している必要がある）
    if age > 86400 and age > expires_in - 7 * 86400:
        refreshed = refresh_long_lived_token(token_data["access_token"])
        token_data["access_token"] = refreshed["access_token"]
        token_data["expires_in"] = refreshed.get("expires_in", expires_in)
        token_data["obtained_at"] = time.time()
        _save_token(token_data)

    return token_data
