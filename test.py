import json
import requests
import msal

CLIENT_ID = "4765445b-32c6-49b0-83e6-1d93765276ca"
TENANT_ID = "organizations"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# SCOPES = ["https://graph.microsoft.com/.default"]



# CLIENT_ID = "你的 Application (client) ID"
# TENANT_ID = "你的 Directory (tenant) ID"
# AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["User.Read"]

token_cache = msal.SerializableTokenCache()


def build_app():
    return msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=token_cache,
    )


def get_access_token():
    app = build_app()

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    result = app.acquire_token_interactive(
        scopes=SCOPES,
        port=53100,
    )

    if "access_token" not in result:
        raise RuntimeError("取得 token 失敗：\n" + json.dumps(result, indent=2, ensure_ascii=False))

    return result["access_token"]


def call_me(token: str):
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def main():
    token = get_access_token()
    print("Access token 取得成功：")
    print(token[:80] + "...")

    me = call_me(token)
    print("\n/me 結果：")
    print(json.dumps(me, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
