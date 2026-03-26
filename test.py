import requests
import msal

CLIENT_ID = "4765445b-32c6-49b0-83e6-1d93765276ca"
TENANT_ID = "organizations"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["https://graph.microsoft.com/.default"]


def get_token():
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
    )

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"建立 device flow 失敗: {flow}")

    print(flow["message"])  # 去瀏覽器登入
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception(f"拿 token 失敗: {result}")

    return result["access_token"]


def create_conversation(token):
    url = "https://graph.microsoft.com/beta/copilot/conversations"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    resp = requests.post(url, headers=headers, json={})
    resp.raise_for_status()
    return resp.json()["id"]


def ask_copilot(token, conversation_id, prompt):
    url = f"https://graph.microsoft.com/beta/copilot/conversations/{conversation_id}/chat"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    payload = {"message": {"text": prompt}, "locationHint": {"timeZone": "Asia/Taipei"}}

    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()

    data = resp.json()
    messages = data.get("messages", [])
    texts = [m.get("text", "").strip() for m in messages if m.get("text")]
    return texts[-1] if texts else data


if __name__ == "__main__":
    token = get_token()
    conversation_id = create_conversation(token)
    answer = ask_copilot(token, conversation_id, "請用繁體中文簡單介紹 Python 是什麼")
    print("Copilot 回答：")
    print(answer)
