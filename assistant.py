import os
from flask import Flask, redirect, url_for, session, request, render_template, jsonify
import msal
import requests
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.urandom(24)

# -------------------------
# Config
# -------------------------
CLIENT_ID = os.environ.get("CLIENT_ID")
TENANT_ID = os.environ.get("TENANT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")  # optional for public client
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_API = "https://graph.microsoft.com/v1.0"
SCOPES = ["User.Read.All", "Chat.ReadWrite"]  # Delegated permissions

# -------------------------
# MSAL Public Client for interactive login
# -------------------------
msal_app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)


# -------------------------
# Login route
# -------------------------
@app.route("/login")
def login():
    # Get auth URL
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=url_for("authorized", _external=True))
    return redirect(auth_url)


# -------------------------
# Redirect URI after login
# -------------------------
@app.route("/authorized")
def authorized():
    code = request.args.get("code")
    if not code:
        return "No auth code received", 400

    result = msal_app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=url_for("authorized", _external=True)
    )

    if "access_token" not in result:
        return f"Login failed: {result}", 400

    session["access_token"] = result["access_token"]

    # Get current user info
    me_resp = requests.get(f"{GRAPH_API}/me", headers={"Authorization": f"Bearer {session['access_token']}"})
    session["user"] = me_resp.json()

    return redirect(url_for("chat_page"))


# -------------------------
# Logout
# -------------------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))



if __name__ == "__main__":
    app.run(debug=True)
