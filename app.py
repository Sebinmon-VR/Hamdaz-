from flask import Flask, redirect, url_for, session, request, render_template
from msal import ConfidentialClientApplication
import requests
import os
from dotenv import load_dotenv
load_dotenv(override=True)
app = Flask(__name__)
app.secret_key = os.urandom(24)

# ---------------- Azure AD Config ----------------
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["User.Read"]

SUPERUSERS = ["sebin@hamdaz.com"]  # full access
LIMITED_USERS = ["hello@hamdaz.com" , "marwa@hamdaz.com"]  

# Initialize MSAL
msal_app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

# ---------------- Routes ----------------
@app.route("/")
def index():
    if "user" in session:
        email = session["user"]["email"]
        if email.lower() in SUPERUSERS:
            role = "super"
            return render_template("admin_dashboard.html" , role=role , email=email)
        else:
            role = "limited"
            return render_template("user_dashboard.html" , role=role , email=email)
    return render_template("login.html")

@app.route("/login")
def login():
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route("/getAToken")
def authorized():
    code = request.args.get("code")
    if code:
        result = msal_app.acquire_token_by_authorization_code(
            code,
            scopes=SCOPE,
            redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            # Get user info from Graph
            graph_data = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {result['access_token']}"}
            ).json()
            session["user"] = {"email": graph_data.get("mail") or graph_data.get("userPrincipalName")}
            return redirect("/")
    return "Login failed"

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")




if __name__ == "__main__":
    app.run(debug=True)

