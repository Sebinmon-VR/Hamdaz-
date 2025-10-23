import email
from flask import Flask, redirect, url_for, session, request, render_template, jsonify
from msal import ConfidentialClientApplication
import requests
import os
from dotenv import load_dotenv
import base64
from datetime import datetime
import threading
import time
import pandas as pd  # Required for timestamp conversion
from sharepoint_data import *
from sharepoint_items import *
from zoho import *
from auto_assign import *
import jwt
# ================== LOAD ENVIRONMENT ==================
load_dotenv(override=True)

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "supersecretkey123")

# ---------------- Azure AD Config ----------------
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["User.Read"]

SUPERUSERS = ["sebin@hamdaz.com"]
LIMITED_USERS = [""]

# Initialize MSAL
msal_app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# ---------------- SharePoint Config ----------------
SITE_DOMAIN = "hamdaz1.sharepoint.com"
SITE_PATH = "/sites/ProposalTeam"
LIST_NAME = "Proposals"
EXCLUDED_USERS = ["Sebin", "Shamshad", "Jaymon", "Hisham Arackal", "Althaf", "Nidal", "Nayif Muhammed S", "Afthab"]

# ==============================================================
# Initialize global data (first load)
print("[INIT] Fetching initial SharePoint data...")
tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
df = items_to_dataframe(tasks)

user_analytics = generate_user_analytics(df, exclude_users=EXCLUDED_USERS)
print("[INIT] Data loaded successfully.")

# ==============================================================
# HELPER FUNCTIONS
# ==============================================================

def is_admin(email):
    return email.lower() in SUPERUSERS if email else False

app.jinja_env.globals.update(is_admin=is_admin , current_date=datetime.now() )

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

def get_analytics_data(df, period_type='month', year=None, month=None):
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

# ==============================================================
# BACKGROUND DATA UPDATER
# ==============================================================

def background_updater():
    """Runs in background to refresh SharePoint data periodically."""
    global tasks, df, user_analytics
    while True:
        try:
            print("[BG] Updating SharePoint data...")
            tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
            df = items_to_dataframe(tasks)
            user_analytics = generate_user_analytics(df, exclude_users=EXCLUDED_USERS)

            print(f"[BG] Data updated successfully at {datetime.now()}")


        except Exception as e:
            print("[BG] Error during update:", e)

        time.sleep(500)
# ==============================================================
# ROUTES
# ==============================================================

@app.route("/update_analytics")
def update_analytics():
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

@app.route("/")
def index():
    if "user" in session:
        user = session["user"]
        email = user.get("mail") or user.get("userPrincipalName")
        user_id = user.get("id")

        user_flag_data = get_user_details_from_excell()
        current_user = next((u for u in user_flag_data if u.get("email", "").lower() == email.lower()), None)
        flag_value = current_user.get("flag") if current_user else 0
        try:
            flag = int(flag_value) if str(flag_value).strip() else 0
        except (ValueError, TypeError):
            flag = 0

        if not current_user or flag != 1:
            return redirect(url_for("user_form"))

        if email.lower() in SUPERUSERS:
            dashboard_role = "admin_dashboard"
            excel_role = "admin"
        else:
            excel_role = current_user.get("role", "").strip().lower() if current_user else ""
            if excel_role == "pre-sales":
                app.jinja_env.globals.update(excel_role="pre-sales")
                dashboard_role = "pre_sales_dashboard"
            elif excel_role == "business development":
                app.jinja_env.globals.update(excel_role="bd")
                dashboard_role = "business_dev_dashboard"
            elif excel_role == "customer success":
                app.jinja_env.globals.update(excel_role="cs")
                dashboard_role = "customer_success_dashboard"
            else:
                
                dashboard_role = "user_dashboard"

        period_type = request.args.get('period', 'month')
        year = int(request.args.get('year') or datetime.now().year)
        month = int(request.args.get('month') or datetime.now().month)

        greeting = greetings()
        analytics, per_user = get_analytics_data(df, period_type, year, month)
        username = user.get("displayName", "").replace(" ", "")
        user_analytics_specific = get_user_analytics_specific(df, username)
        now_utc = pd.Timestamp.utcnow()
        ongoing_filtered = [
            t for t in user_analytics_specific['OngoingTasks']
            if pd.to_datetime(t['BCD']) > now_utc and t.get('SubmissionStatus','') != 'Submitted'
        ]
        user_analytics_specific['OngoingTasks'] = ongoing_filtered
        ongoing_tasks_count = len(ongoing_filtered)
        if 'Created' in df.columns:
            df['Created'] = pd.to_datetime(df['Created'])
            available_years = sorted(df['Created'].dt.year.unique().tolist())
        else:
            available_years = [datetime.now().year]

        return render_template(
            f"{dashboard_role}.html",
            role=excel_role,
            user=user,
            greeting=greeting,
            tasks=tasks,
            analytics=analytics,
            per_user=per_user,
            current_period=period_type,
            current_year=year,
            current_month=month,
            available_years=available_years,
            user_analytics=user_analytics_specific,
            email=email,
            ongoing_tasks_count=ongoing_tasks_count,
            due_today_tasks_count=len([
                t for t in user_analytics_specific['OngoingTasks']
                if pd.to_datetime(t['BCD']).date() == now_utc.date()
            ]),
            user_flag_data=user_flag_data
        )
    return render_template("login.html")

@app.route("/user_form", methods=["GET", "POST"])
def user_form():
    if "user" not in session:
        return redirect(url_for("login"))
    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    user_id = user.get("id")
    if request.method == "POST":
        role = request.form.get("role")
        success = add_or_update_user_in_excel(email, user_id, user.get("displayName"), role)
        if success:
            return redirect(url_for("index"))
        else:
            return "Error saving user data. Please try again."
    return render_template("user_form.html", user=user)

@app.route("/dashboard")
def dashboard():
    return redirect("/")

@app.route("/teams")
def teams():
    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    return render_template("teams.html", user=user, email=email, user_analytics=user_analytics)

@app.route("/bd")
def bd():
    return render_template("business_dev_team.html")

@app.route("/cs")
def cs():
    return render_template("customer_success_team.html")

@app.route("/user/<username>")
def user_profile(username):
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    df = items_to_dataframe(tasks)
    user_analytics_specific = get_user_analytics_specific(df, username)
    
    now_utc = pd.Timestamp.utcnow()
    ongoing_filtered = [
        t for t in user_analytics_specific['OngoingTasks']
        if pd.to_datetime(t['BCD']) > now_utc and t.get('SubmissionStatus','') != 'Submitted'
    ]
    user_analytics_specific['OngoingTasks'] = ongoing_filtered
    user_analytics_specific['OngoingTasksCount'] = len(ongoing_filtered)
    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    if username in ["dashboard", "customer", "businesscard", "orders", "payments", "reports"]:
        return redirect(f"/{username}")
    return render_template("profile.html", user=user, email=email, user_analytics=user_analytics_specific)

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
            access_token = result["access_token"]
            graph_data = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {access_token}"}
            ).json()
            session["user"] = graph_data
            session["access_token"] = access_token
            return redirect("/")
    return "Login failed"

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/task_details/<title>")
def task_details(title):
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    df = items_to_dataframe(fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME))
    task = get_task_details(df, title)
    return render_template("pages/task_details.html", task=task, user=user)

@app.route("/businesscard")
def business_cards():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    contacts = get_all_contacts_from_onedrive()
    return render_template("pages/business_cards.html", contacts=contacts, user=user)

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

@app.route("/customers")
def customer():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    raw_customers = fetch_customers()
    structured_customers = structure_customers_data(raw_customers)
    return render_template("pages/customers.html", customers=structured_customers, user=user)

@app.route("/quote")
def quote():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    raw_customers = fetch_customers()
    structured_customers = structure_customers_data(raw_customers)
    
    return render_template("pages/quote.html", user=user, customers=structured_customers)
# ==============================================================
# ==============================================================
def get_first(quote_data, key, index=0, default=""):
    """
    Safely get the first or indexed value from a list in quote_data.
    """
    values = quote_data.get(key, [default])
    if isinstance(values, list):
        if len(values) > index:
            return values[index]
        else:
            return default
    return values or default

@app.route("/send_quote_for_approval", methods=["POST"])
def send_for_approval():
    if "user" not in session:
        return redirect(url_for("login"))

    user = session.get("user")
    submitter_email = user.get("mail") or user.get("userPrincipalName")

    quote_data = request.form.to_dict(flat=False)

    # ✅ Combine all items into a single list
    num_items = len(quote_data.get("item_details[]", []))
    combined_items = []

    for i in range(num_items):
        try:
            tax_val = get_first(quote_data, "tax[]", i, 0)
            try:
                tax_val = float(tax_val)
            except ValueError:
                tax_val = 0

            combined_items.append({
                "ItemDetails": get_first(quote_data, "item_details[]", i),
                "Brand": get_first(quote_data, "brand[]", i),
                "Quantity": int(get_first(quote_data, "quantity[]", i, 0)),
                "Rate": float(get_first(quote_data, "rate[]", i, 0)),
                "Margin": float(get_first(quote_data, "margin[]", i, 0)),
                "Tax": tax_val,
                "Amount": float(get_first(quote_data, "amount[]", i, 0))
            })
        except Exception as e:
            print(f"Error parsing item {i}: {e}")

    # ✅ Store all items together in one SharePoint row
    item_fields = {
        "Title": get_first(quote_data, "reference", 0, "No Title"),
        "CustomerID": get_first(quote_data, "customer_id", 0),
        "Currency": get_first(quote_data, "currency", 0),
        "PaymentTerms": get_first(quote_data, "payment_terms", 0),
        "Email": get_first(quote_data, "email", 0),
        "TaxTreatment": get_first(quote_data, "tax_treatment", 0),
        "Reference": get_first(quote_data, "reference", 0),
        "QuoteDate": get_first(quote_data, "quote_date", 0),
        "ExpiryDate": get_first(quote_data, "expiry_date", 0),
        "Portal": get_first(quote_data, "portal", 0),
        "QuoteCreator": get_first(quote_data, "quote_creator", 0),
        "BCD": get_first(quote_data, "bcd", 0),
        "ApprovalStatus": "Pending",
        "Amount": sum(item["Amount"] for item in combined_items),
        "Margin": sum(item["Margin"] for item in combined_items),
        "Rate": sum(item["Rate"] for item in combined_items) / len(combined_items) if combined_items else 0,
        # ✅ Convert list of items to a readable string (or JSON)
        "AllItems": json.dumps(combined_items, indent=2)
        
    }

    try:
        add_sharepoint_list_item(item_fields)
        return render_template("pages/quote_success.html", user=user, added_items=1)
    except Exception as e:
        return f"❌ Error adding quote to SharePoint: {str(e)}", 500





# ==============================================================
# START FLASK + BACKGROUND UPDATER
# ==============================================================

if __name__ == "__main__":
    threading.Thread(target=background_updater, daemon=True).start()
    app.run(debug=True)
