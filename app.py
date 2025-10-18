import email
from flask import Flask, redirect, url_for, session, request, render_template, jsonify
from msal import ConfidentialClientApplication
import requests
import os
from dotenv import load_dotenv
import base64
from datetime import datetime
from sharepoint_data import *
from sharepoint_items import *
from zoho import *

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

SUPERUSERS = [""]
LIMITED_USERS = ["hello@hamdaz.com"]

# Initialize MSAL
msal_app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# ==============================================================================

# Template helper functions
def is_admin(email):
    return email.lower() in SUPERUSERS if email else False

# ==============================================================================


app.jinja_env.globals.update(is_admin=is_admin)

SITE_DOMAIN = "hamdaz1.sharepoint.com"
SITE_PATH = "/sites/ProposalTeam"
LIST_NAME = "Proposals"
EXCLUDED_USERS = ["Sebin", "Shamshad", "Jaymon", "Hisham Arackal", "Althaf", "Nidal", "Nayif Muhammed S", "Afthab"]




# ---------------- Helper Function ----------------
# def fetch_user_photo(access_token):
#     """
#     Fetches the user's profile photo from Microsoft Graph.
#     Returns a base64-encoded string or None if not available.
#     """
#     try:
#         photo_resp = requests.get(
#             "https://graph.microsoft.com/v1.0/me/photo/$value",
#             headers={"Authorization": f"Bearer {access_token}"},
#             stream=True
#         )
#         if photo_resp.status_code == 200:
#             return "data:image/jpeg;base64," + base64.b64encode(photo_resp.content).decode()
#         else:
#             return None
#     except Exception as e:
#         print("Error fetching profile photo:", e)
#         return None

# ==============================================================================

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
def get_analytics_data(df, period_type='month', year=None, month=None):
    """Helper function to compute analytics data"""
    if year is None:
        year = datetime.now().year
    if month is None:
        month = datetime.now().month

    period = {
        'type': period_type,
        'year': year,
        'month': month
    } if period_type != 'all' else None

    analytics = compute_overall_analytics(df, period)
    per_user = compute_user_analytics_with_last_date(df, EXCLUDED_USERS, period)
    
    return analytics, per_user
# ==============================================================================

@app.route("/update_analytics")
def update_analytics():
    """AJAX endpoint for updating analytics data"""
    if "user" not in session:
        return jsonify({"error": "Unauthorized"}), 401
    
    period_type = request.args.get('period', 'month')
    year = int(request.args.get('year', datetime.now().year))
    month = int(request.args.get('month', datetime.now().month))
    
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    df = items_to_dataframe(tasks)
    
    analytics, per_user = get_analytics_data(df, period_type, year, month)
    
    return jsonify({
        "analytics": analytics,
        "per_user": per_user
    })

# ==============================================================================

@app.route("/")
def index():
    if "user" in session:
        user = session["user"]
        email = user.get("mail") or user.get("userPrincipalName")
        user_id = user.get("id")
        access_token = session.get("access_token")

        # Fetch user photo if not in session
        # if access_token and "photo" not in user:
        #     photo = fetch_user_photo(access_token)
        #     user["photo"] = photo
        #     session["user"] = user

        # Load Excel data
        user_flag_data = get_user_details_from_excell()

        # Find current user in Excel
        current_user = next((u for u in user_flag_data if u.get("email", "").lower() == email.lower()), None)

        # --- SAFELY CHECK FLAG ---
        flag_value = current_user.get("flag") if current_user else 0
        try:
            flag = int(flag_value) if str(flag_value).strip() else 0
        except (ValueError, TypeError):
            flag = 0

        # Redirect to form if flag != 1 or user not found
        if not current_user or flag != 1:
            return redirect(url_for("user_form"))
# --- Determine dashboard based on SUPERUSERS first ---
        if email.lower() in SUPERUSERS:
            dashboard_role = "admin_dashboard"
            excel_role = "admin"
        else:
            excel_role = current_user.get("role", "").strip().lower() if current_user else ""
            if excel_role == "pre-sales":
                dashboard_role = "pre_sales_dashboard"
            elif excel_role == "business development":
                dashboard_role = "business_dev_dashboard"
            elif excel_role == "customer success":
                dashboard_role = "customer_success_dashboard"
            else:
                dashboard_role = "user_dashboard"  # default

        # --- Dashboard logic ---
        period_type = request.args.get('period', 'month')
        year = int(request.args.get('year') or datetime.now().year)
        month = int(request.args.get('month') or datetime.now().month)

        greeting = greetings()
        tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
        df = items_to_dataframe(tasks)

        analytics, per_user = get_analytics_data(df, period_type, year, month)
        username = user.get("displayName", "").replace(" ", "")
        user_analytics = get_user_analytics_specific(df, username)
        now_utc = pd.Timestamp.utcnow()

        ongoing_filtered = [
            t for t in user_analytics['OngoingTasks']
            if pd.to_datetime(t['BCD']) > now_utc and t.get('SubmissionStatus','') != 'Submitted'
        ]
        user_analytics['OngoingTasks'] = ongoing_filtered
        ongoing_tasks_count = len(ongoing_filtered)

        if 'Created' in df.columns:
            df['Created'] = pd.to_datetime(df['Created'])
            available_years = sorted(df['Created'].dt.year.unique().tolist())
        else:
            available_years = [datetime.now().year]

        return render_template(
            f"{dashboard_role}.html",  # dynamic template based on Excel role
            role=excel_role,
            user=user,
            greeting=greeting,
            tasks=tasks,
            # photo=user.get("photo"),
            analytics=analytics,
            per_user=per_user,
            current_period=period_type,
            current_year=year,
            current_month=month,
            available_years=available_years,
            user_analytics=user_analytics,
            email=email,
            ongoing_tasks_count=ongoing_tasks_count,
            due_today_tasks_count=len([
                t for t in user_analytics['OngoingTasks']
                if pd.to_datetime(t['BCD']).date() == now_utc.date()
            ]),
            user_flag_data=user_flag_data
        )

    return render_template("login.html")

# ==============================================================================

@app.route("/user_form", methods=["GET", "POST"])
def user_form():
    if "user" not in session:
        return redirect(url_for("login"))

    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    user_id = user.get("id")

    if request.method == "POST":
        role = request.form.get("role")
        # photo_file = request.files.get("photo")

        # Save or update user in Excel
        success = add_or_update_user_in_excel(email, user_id, user.get("displayName"), role)
        if success:
            return redirect(url_for("index"))
        else:
            return "Error saving user data. Please try again."

    return render_template("user_form.html", user=user)

# ==============================================================================

@app.route("/dashboard")
def dashboard():
    return redirect("/")

# ==============================================================================

@app.route("/teams")
def teams():
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    df=items_to_dataframe(tasks)
    access_token=session.get("access_token")
    user_analytics=generate_user_analytics(df,exclude_users= EXCLUDED_USERS)
    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    return render_template("teams.html" ,user=user , email=email ,user_analytics=user_analytics )

# ==============================================================================

@app.route("/user/<username>")
def user_profile(username):
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
    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    
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
    
    return render_template("profile.html",user=user,email=email,user_analytics=user_analytics)

# ==============================================================================

@app.route("/login")
def login():
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

# ==============================================================================

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

# ==============================================================================


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ==============================================================================

@app.route("/task_details/<title>")
def task_details(title):
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    # photo = session.get("user", {}).get("photo")
    df = items_to_dataframe(fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME))
    task = get_task_details(df, title)
    return render_template("pages/task_details.html", task=task, user=user)

# ==============================================================================
# == BUSINESS CARDS ROUTES ==
# ==============================================================================

@app.route("/businesscard")
def business_cards():
    if "user" not in session:
        return redirect(url_for('login'))
    
    # THE FIX IS HERE: We now get the user and photo from the session
    user = session.get("user")
    # Safely get photo, providing a default empty dict if user is not found
    # photo = session.get("user", {}).get("photo")

    contacts = get_all_contacts_from_onedrive()

    # THE FIX IS HERE: We now pass the user and photo to the template
    return render_template("pages/business_cards.html", contacts=contacts, user=user)

# ==============================================================================

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

@app.route("/customers")
def customer():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    # photo = session.get("user", {}).get("photo")
    raw_customers = fetch_customers()
    structured_customers = structure_customers_data(raw_customers)
    return render_template("pages/customers.html", customers=structured_customers, user=user)




if __name__ == "__main__":
    app.run(debug=True)
