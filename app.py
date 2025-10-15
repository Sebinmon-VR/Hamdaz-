from flask import Flask, redirect, url_for, session, request, render_template, jsonify
from flask_caching import Cache
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

# Configure Flask-Caching
cache = Cache(app, config={
    'CACHE_TYPE': 'simple',
    'CACHE_DEFAULT_TIMEOUT': 300  # Cache timeout in seconds (5 minutes)
})

# ---------------- Azure AD Config ----------------
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")  # Must match Azure
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = [
    "User.Read",
    "Sites.Read.All",
    "Sites.ReadWrite.All",
    "User.ReadBasic.All",
    "profile",
    "email",
    "offline_access"
]

SUPERUSERS = ["sebin@hamdaz.com","jishad@hamdaz.com", "hisham@hamdaz.com", "mustaq@hamdaz.com"]
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

def is_admin(user):
    """Check if the user is an admin"""
    if not user:
        return False
    email = user.get("mail") or user.get("userPrincipalName")
    return email and email.lower() in SUPERUSERS

# ---------------- Routes ----------------
@app.route("/")
@cache.cached(timeout=300, key_prefix='index_view')  # Cache for 5 minutes
def index():
    if "user" in session:
        user = session["user"]
        email = user.get("mail") or user.get("userPrincipalName")
        access_token = session.get("access_token")
        
        # Cache the photo in the session
        if access_token and "photo" not in user:
            photo = fetch_user_photo(access_token)
            if photo:
                user["photo"] = photo
                session["user"] = user
        
        # Set admin status
        user["is_admin"] = is_admin(user)
        session["user"] = user
        
        role = "admin" if user["is_admin"] else "user"
        
        # Cache the SharePoint data and computations
        @cache.memoize(300)  # Cache for 5 minutes
        def get_dashboard_data():
            tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
            df = items_to_dataframe(tasks)
            analytics = compute_overall_analytics(df)
            per_user = compute_user_analytics_with_last_date(df, EXCLUDED_USERS)
            return tasks, analytics, per_user
        
        tasks, analytics, per_user = get_dashboard_data()
        greeting = greetings()
        photo = user.get("photo")
        
        return render_template(
            f"{role}_dashboard.html",
            role=role,
            user=user,
            greeting=greeting,
            tasks=tasks,
            photo=photo,
            analytics=analytics,
            per_user=per_user
        )

    return render_template("login.html")

@app.route("/dashboard")
def dashboard():
    return redirect("/")

@app.route("/teams")
@cache.cached(timeout=300, key_prefix='teams_view')  # Cache for 5 minutes
def teams():
    if "user" not in session or "access_token" not in session:
        return redirect(url_for('login'))
        
    user = session["user"]
    access_token = session["access_token"]
    
    # Ensure photo is loaded
    if access_token and "photo" not in user:
        photo = fetch_user_photo(access_token)
        if photo:
            user["photo"] = photo
            session["user"] = user
    
    @cache.memoize(300)  # Cache for 5 minutes
    def get_teams_data():
        tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
        df = items_to_dataframe(tasks)
        return generate_user_analytics(df, exclude_users=EXCLUDED_USERS)
    
    try:
        user_analytics = get_teams_data()
        email = user.get("mail") or user.get("userPrincipalName")
        photo = user.get("photo")
        return render_template("teams.html", user=user, email=email, user_analytics=user_analytics, photo=photo)
    except Exception as e:
        print(f"Error in teams route: {str(e)}")
        session.clear()
        cache.clear()
        return redirect(url_for('login'))


@app.route("/user/<username>")
def user_profile(username):
    if "user" not in session or "access_token" not in session:
        return redirect(url_for('login'))

    user = session["user"]
    access_token = session["access_token"]
    
    # Ensure photo is loaded
    if access_token and "photo" not in user:
        photo = fetch_user_photo(access_token)
        if photo:
            user["photo"] = photo
            session["user"] = user

    # Fetch SharePoint tasks
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    df = items_to_dataframe(tasks)

    # Get analytics for specific user
    user_analytics = get_user_analytics_specific(df, username)  # returns dict

    # Filter ongoing tasks: BCD in future & not submitted
    now_utc = pd.Timestamp.utcnow()
    ongoing_filtered = [
        t for t in user_analytics['OngoingTasks']
        if pd.to_datetime(t['BCD']) > now_utc and t.get('SubmissionStatus','') != 'Submitted'
    ]
    user_analytics['OngoingTasks'] = ongoing_filtered
    user_analytics['OngoingTasksCount'] = len(ongoing_filtered)
    email = user.get("mail") or user.get("userPrincipalName")
    photo = user.get("photo")
    
    if username=="dashboard":
        return redirect("/")
    elif username== "customer":
        return redirect("/customer")
    elif username=="businesscard":
        return redirect("/businesscard")
    elif username == "orders":
        return redirect("/orders")
    elif username == "payments":
        return redirect("/payments")
    elif username == "reports":
        return redirect("/reports")
    
    return render_template("profile.html", user=user, email=email, user_analytics=user_analytics, photo=photo)


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
    if not code:
        print("Error: No code received")
        return redirect(url_for('login'))

    try:
        # Use the same REDIRECT_URI as registered in Azure
        result = msal_app.acquire_token_by_authorization_code(
            code,
            scopes=SCOPE,
            redirect_uri=REDIRECT_URI
        )
        
        if "error" in result:
            print(f"Token acquisition error: {result.get('error_description', 'No error description')}")
            return redirect(url_for('login'))

        if "access_token" in result:
            access_token = result["access_token"]

            # Get user info
            graph_response = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {access_token}"}
            )
            
            if graph_response.status_code != 200:
                print(f"Graph API error: {graph_response.status_code}, {graph_response.text}")
                return redirect(url_for('login'))

            graph_data = graph_response.json()

            # Save user info and tokens in session
            session["user"] = graph_data
            session["access_token"] = access_token
            if "refresh_token" in result:
                session["refresh_token"] = result["refresh_token"]

            return redirect(url_for('index'))

        print("No access token in result")
        return redirect(url_for('login'))

    except Exception as e:
        print(f"Authorization error: {str(e)}")
        return redirect(url_for('login'))


@app.route("/logout")
def logout():
    session.clear()
    # Clear all caches
    cache.clear()
    return redirect("/")




# ==============================================================================
# == BUSINESS CARDS ROUTES ==
# ==============================================================================

@app.route("/businesscard")
@cache.cached(timeout=300, key_prefix='business_cards_view')  # Cache for 5 minutes
def business_cards():
    if "user" not in session or "access_token" not in session:
        return redirect(url_for('login'))
    
    user = session["user"]
    access_token = session["access_token"]
    
    # Ensure photo is loaded
    if access_token and "photo" not in user:
        photo = fetch_user_photo(access_token)
        if photo:
            user["photo"] = photo
            session["user"] = user
    
    photo = user.get("photo")

    @cache.memoize(300)  # Cache for 5 minutes
    def get_contacts_data():
        return get_all_contacts_from_onedrive()

    contacts = get_contacts_data()
    return render_template("pages/business_cards.html", contacts=contacts, user=user, photo=photo)


@app.route("/api/update-contact", methods=['POST'])
def api_update_contact():
    if "user" not in session:
        return jsonify({"success": False, "error": "Unauthorized"}), 401
    
    data = request.get_json()
    row_id = data.get('row_id')
    contact_data = data.get('contact_data')

    if not row_id or not contact_data:
        return jsonify({"success": False, "error": "Missing data"}), 400

    success = update_contact_in_onedrive_excel(row_id, contact_data)

    if success:
        return jsonify({"success": True})
    else:
        return jsonify({"success": False, "error": "Failed to update Excel file"}), 500


# ==============================================================================

# ==============================================================================
# == ADMIN ROUTES ==
# ==============================================================================

def admin_required(f):
    """Decorator to check if user is admin"""
    from functools import wraps
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for('login'))
        
        user = session["user"]
        if not is_admin(user):
            return redirect(url_for('index'))
            
        return f(*args, **kwargs)
    return decorated_function



@app.route("/vendorpartnership")
@admin_required
def admin_users():
    user = session["user"]
    photo = user.get("photo")
    # Implement user management logic here
    return render_template("pages/vendor_partnership.html", user=user, photo=photo)

# ==============================================================================
if __name__ == "__main__":
    app.run(debug=True)
