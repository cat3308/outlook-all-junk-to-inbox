import os
import threading
import time
import msal
import requests
from flask import Flask, redirect, request, session, url_for, render_template_string

app = Flask(__name__)
app.secret_key = os.urandom(24)

CONFIG_FILE = "/app/data/config.json"
CACHE_FILE = "/app/data/token_cache.bin"
REDIRECT_URI = "http://localhost:9050/gettoken"

def get_client_id():
    # Check environment variable first, then fallback to saved config file
    client_id = os.getenv("CLIENT_ID")
    if client_id:
        return client_id
    if os.path.exists(CONFIG_FILE):
        import json
        with open(CONFIG_FILE, "r") as f:
            return json.load(f).get("client_id")
    return None

def save_client_id(client_id):
    import json
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump({"client_id": client_id}, f)

def get_msal_app(cache):
    client_id = get_client_id()
    if not client_id:
        return None
    return msal.PublicClientApplication(
        client_id,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache
    )

def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        new_client_id = request.form.get("client_id")
        if new_client_id:
            save_client_id(new_client_id.strip())
        return redirect(url_for("index"))

    client_id = get_client_id()
    if not client_id:
        # Show setup page if Client ID is not configured yet
        return """
        <html>
            <head><title>Setup - Outlook Junk Cleaner</title></head>
            <body style="font-family: Arial; text-align: center; margin-top: 80px; background-color: #f4f6f9;">
                <div style="background: white; padding: 40px; border-radius: 10px; display: inline-block; box-shadow: 0 4px 10px rgba(0,0,0,0.1); width: 400px;">
                    <h2>Welcome! 🐾</h2>
                    <p style="color: #555;">Please enter your Microsoft Azure Client ID to get started:</p>
                    <form method="POST">
                        <input type="text" name="client_id" placeholder="Enter Client ID here..." style="width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;" required>
                        <button type="submit" style="width: 100%; padding: 12px; background-color: #0078d4; color: white; border: none; border-radius: 5px; font-size: 16px; cursor: pointer;">Save & Continue</button>
                    </form>
                </div>
            </body>
        </html>
        """

    cache = load_cache()
    msal_app = get_msal_app(cache)
    accounts = msal_app.get_accounts() if msal_app else []
    logged_in = len(accounts) > 0
    email = accounts[0]["username"] if logged_in else None
    
    return f"""
    <html>
        <head><title>Outlook Junk Cleaner</title></head>
        <body style="font-family: Arial; text-align: center; margin-top: 80px; background-color: #f4f6f9;">
            <div style="background: white; padding: 40px; border-radius: 10px; display: inline-block; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                <h2>Outlook Junk to Inbox Cleaner 🐾</h2>
                {f"<p style='color: green;'>Logged in as: <b>{email}</b></p>" if logged_in else "<p style='color: #d9534f;'>Not logged in yet.</p>"}
                <br>
                {
                    '<a href="/login"><button style="padding: 12px 24px; font-size: 16px; background-color: #0078d4; color: white; border: none; border-radius: 5px; cursor: pointer;">Login with Microsoft</button></a>' 
                    if not logged_in else 
                    '<p style="color: #0078d4; font-weight: bold;">Service is active and running in the background! 😺</p>'
                }
                <br><br>
                <a href="/reset" style="color: #888; font-size: 12px; text-decoration: none;">Reset Settings</a>
            </div>
        </body>
    </html>
    """

@app.route("/login")
def login():
    cache = load_cache()
    msal_app = get_msal_app(cache)
    if not msal_app:
        return redirect(url_for("index"))
    
    auth_url = msal_app.initiate_auth_code_flow(
        ["Mail.ReadWrite", "offline_access"],
        redirect_uri=REDIRECT_URI
    )
    session["flow"] = auth_url
    return redirect(auth_url["auth_uri"])

@app.route("/gettoken")
def authorized():
    cache = load_cache()
    msal_app = get_msal_app(cache)
    flow = session.get("flow", {})
    result = msal_app.acquire_token_by_auth_code_flow(flow, request.args.to_dict())
    
    if "error" in result:
        return f"Login failed: {result.get('error_description')}"
    
    save_cache(cache)
    return redirect(url_for("index"))

@app.route("/reset")
def reset():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
    if os.path.exists(CACHE_FILE):
        os.remove(CACHE_FILE)
    return redirect(url_for("index"))

def background_sync():
    while True:
        time.sleep(60)
        try:
            client_id = get_client_id()
            if not client_id:
                continue
            cache = load_cache()
            msal_app = get_msal_app(cache)
            if not msal_app:
                continue
            accounts = msal_app.get_accounts()
            if not accounts:
                continue
            
            result = msal_app.acquire_token_silent(["Mail.ReadWrite", "offline_access"], account=accounts[0])
            if not result or "access_token" not in result:
                continue
            
            token = result["access_token"]
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            base_url = "https://graph.microsoft.com/v1.0/me"
            
            inbox_res = requests.get(f"{base_url}/mailFolders/inbox", headers=headers)
            inbox_id = inbox_res.json().get("id")
            
            junk_res = requests.get(f"{base_url}/mailFolders/junkemail/messages", headers=headers)
            messages = junk_res.json().get("value", [])
            
            if messages:
                print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Moving {len(messages)} junk emails...")
                for msg in messages:
                    msg_id = msg["id"]
                    requests.post(f"{base_url}/messages/{msg_id}/move", headers=headers, json={"destinationId": inbox_id})
        except Exception as e:
            print(f"Sync error: {e}")

if __name__ == "__main__":
    t = threading.Thread(target=background_sync, daemon=True)
    t.start()
    app.run(host="0.0.0.0", port=5000)