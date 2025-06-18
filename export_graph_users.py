import requests
import json

# === CONFIGURATION ===
ACCESS_TOKEN = ""
OUTPUT_FILE = "graph_users.json"

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}"
}

url = "https://graph.microsoft.com/v1.0/users"
users = []

while url:
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Error {response.status_code}: {response.text}")
    
    data = response.json()
    for user in data.get("value", []):
        users.append({
            "displayName": user.get("displayName"),
            "userPrincipalName": user.get("userPrincipalName"),
            "id": user.get("id")  # Azure AD Object ID
        })

    url = data.get("@odata.nextLink")  # handle pagination

# === WRITE TO FILE ===
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(users, f, indent=2, ensure_ascii=False)

print(f"✅ Exported {len(users)} users to {OUTPUT_FILE}")
