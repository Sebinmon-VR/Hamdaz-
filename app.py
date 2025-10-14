from flask import Flask, redirect, url_for, session, request, render_template
from msal import ConfidentialClientApplication
import requests
import os
from dotenv import load_dotenv
import base64
from datetime import datetime
from sharepoint_data import *
from sharepoint_items import *
load_dotenv(override=True)

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "supersecretkey123")  # fixed key

# ---------------- Azure AD Config ----------------
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")  # Must match Azure
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["User.Read"]

SUPERUSERS = ["sebin@hamdaz.com"]
LIMITED_USERS = ["hello@hamdaz.com"]

# Initialize MSAL
msal_app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
SITE_DOMAIN = "hamdaz1.sharepoint.com"
SITE_PATH = "/sites/ProposalTeam"
LIST_NAME = "Proposals"
EXCLUDED_USERS = ["Sebin", "Shamshad", "Jaymon", "Hisham Arackal", "Althaf", "Nidal", "Nayif Muhammed S", "Afthab"]


# ---------------- Helper Function ----------------
def fetch_user_photo(access_token):
    """
    Fetches the user's profile photo from Microsoft Graph.
    Returns a base64-encoded string or None if not available.
    """
    try:
        photo_resp = requests.get(
            "https://graph.microsoft.com/v1.0/me/photo/$value",
            headers={"Authorization": f"Bearer {access_token}"},
            stream=True
        )
        if photo_resp.status_code == 200:
            return "data:image/jpeg;base64," + base64.b64encode(photo_resp.content).decode()
        else:
            return None
    except Exception as e:
        print("Error fetching profile photo:", e)
        return None

def greetings():
    now = datetime.now()
    hour = now.hour
    if 5 <= hour < 12:
        return "Good Morning"
    elif 12 <= hour < 17:
        return "Good Afternoon"
    elif 17 <= hour < 21:
        return "Good Evening"
    else:
        return "Hello"
    


# ---------------- Routes ----------------
@app.route("/")
def index():
    if "user" in session:
        user = session["user"]
        email = user.get("mail") or user.get("userPrincipalName")
        access_token = session.get("access_token")
        if access_token and "photo" not in user:
            photo = fetch_user_photo(access_token)
            session["user"] = user  # update session with photo
            
        role = "admin" if email.lower() in SUPERUSERS else "user"
        
        # items = get_sharepoint_list_items('hamdaz1.sharepoint.com', '/sites/ProposalTeam', 'Proposals')
        
        greeting=greetings()
        tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
        df=items_to_dataframe(tasks)
        analytics= compute_overall_analytics(df)
        per_user=compute_user_analytics_with_last_date(df ,EXCLUDED_USERS)
        
        return render_template(f"{role}_dashboard.html", role=role, user=user , greeting=greeting , tasks=tasks ,photo=photo , analytics=analytics ,per_user=per_user)

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
        # Use the same REDIRECT_URI as registered in Azure
        result = msal_app.acquire_token_by_authorization_code(
            code,
            scopes=SCOPE,
            redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            access_token = result["access_token"]

            # Get user info
            graph_data = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {access_token}"}
            ).json()

            # Save user info and token in session
            session["user"] = graph_data
            session["access_token"] = access_token

            return redirect("/")

    return "Login failed"


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


if __name__ == "__main__":
    app.run(debug=True)
