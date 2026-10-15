import requests

# Set up your access token and headers
ACCESS_TOKEN = "YOUR_ACCESS_TOKEN"
headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

base_url = "https://graph.microsoft.com/v1.0/me"

# Step 1: Get the Inbox folder ID
inbox_res = requests.get(f"{base_url}/mailFolders/inbox", headers=headers)
inbox_id = inbox_res.json().get("id")

# Step 2: Get all messages from the Junk Email folder
# (Note: Microsoft Graph usually names the folder 'junkemail')
junk_res = requests.get(f"{base_url}/mailFolders/junkemail/messages", headers=headers)
messages = junk_res.json().get("value", [])

print(f"Found {len(messages)} messages in Junk Email. Moving them to Inbox...")

# Step 3: Move each message to the Inbox
for msg in messages:
    msg_id = msg["id"]
    move_url = f"{base_url}/messages/{msg_id}/move"
    payload = {
        "destinationId": inbox_id
    }
    
    response = requests.post(move_url, headers=headers, json=payload)
    if response.status_code == 201:
        print(f"Successfully moved message: {msg.get('subject')}")
    else:
        print(f"Failed to move message {msg_id}: {response.text}")

print("All done, meow! 😺")