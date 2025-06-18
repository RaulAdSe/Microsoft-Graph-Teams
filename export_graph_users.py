import requests
import json

# === CONFIGURATION ===
ACCESS_TOKEN = ""
OUTPUT_FILE = "graph_users.json"

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}"
}

# === Who is running the script (sender) ===
ME_URL = "https://graph.microsoft.com/v1.0/me"
me_response = requests.get(ME_URL, headers=headers)
if me_response.status_code != 200:
    raise Exception(f"Failed to get current user info: {me_response.text}")

me = me_response.json()
me_id = me["id"]

# === Try to find existing 1:1 chat with self ===
self_chat_id = None
chats_response = requests.get("https://graph.microsoft.com/v1.0/me/chats", headers=headers)
if chats_response.status_code == 200:
    for chat in chats_response.json().get("value", []):
        if chat.get("chatType") == "oneOnOne":
            members_url = f"https://graph.microsoft.com/v1.0/chats/{chat['id']}/members"
            members_response = requests.get(members_url, headers=headers)
            if members_response.status_code != 200:
                continue
            members = members_response.json().get("value", [])
            if len(members) == 1 and members[0].get("userId") == me_id:
                self_chat_id = chat["id"]
                break

# === Get all users and create/find chats ===
url = "https://graph.microsoft.com/v1.0/users"
users = []

while url:
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Error {response.status_code}: {response.text}")
    
    data = response.json()
    for user in data.get("value", []):
        user_id = user.get("id")
        display_name = user.get("displayName")
        user_principal = user.get("userPrincipalName")

        if not user_id or not user_principal:
            continue

        if user_id == me_id:
            # Current user — assign known self_chat_id
            users.append({
                "displayName": display_name,
                "userPrincipalName": user_principal,
                "id": user_id,
                "chatId": self_chat_id
            })
            continue

        # Try to create/find 1:1 chat with this user
        chat_payload = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{me_id}')"
                },
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')"
                }
            ]
        }

        chat_response = requests.post(
            "https://graph.microsoft.com/v1.0/chats",
            headers={**headers, "Content-Type": "application/json"},
            json=chat_payload
        )

        if chat_response.status_code == 201:
            chat_id = chat_response.json().get("id")
        else:
            print(f"⚠️ Could not create/find chat with {display_name}: {chat_response.status_code}")
            chat_id = None

        users.append({
            "displayName": display_name,
            "userPrincipalName": user_principal,
            "id": user_id,
            "chatId": chat_id
        })

    url = data.get("@odata.nextLink")  # pagination

# === Write to JSON file ===
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(users, f, indent=2, ensure_ascii=False)

print(f"✅ Exported {len(users)} users to {OUTPUT_FILE}")
