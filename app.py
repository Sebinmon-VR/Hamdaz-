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
import jwt
import re
import json
import html

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

SUPERUSERS = ["jishad@hamdaz.com", "sebin@hamdaz.com"]
approvers = ["shibit@hamdaz.com", "althaf@hamdaz.com" ,"sebin@hamdaz.com"]
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

def is_approver(email):
    return email.lower() in approvers if email else False

app.jinja_env.globals.update(is_admin=is_admin, is_approver=is_approver, current_date=datetime.now())

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
            app.jinja_env.globals.update(excel_role="admin")
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
    user = session["user"]
    return render_template("business_dev_team.html", user=user)

@app.route("/cs")
def cs():
    user = session["user"]
    return render_template("customer_success_team.html", user =user)

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
    supplier_file = request.files.get("supplier_quote")  # Uploaded file

    def safe_float(val, default=0):
        try:
            return float(val)
        except (TypeError, ValueError):
            return default

    # Process quote line items as before
    num_items = len(quote_data.get("item_details[]", []))
    combined_items = []

    total_amount = 0
    total_selling = 0
    total_discount = 0
    total_tax_amount = 0
    total_margin = 0
    total_rate = 0
    count = 0

    for i in range(num_items):
        qty = int(quote_data.get("quantity[]", [])[i] or 0)
        rate = safe_float(quote_data.get("rate[]", [])[i])
        margin = safe_float(quote_data.get("margin[]", [])[i])
        discount = safe_float(quote_data.get("discount[]", [])[i])
        tax = safe_float(quote_data.get("tax[]", [])[i])

        base_amount = qty * rate * (1 + margin / 100)
        amount = base_amount - discount
        tax_amount = amount * tax
        selling = amount + tax_amount

        combined_items.append({
            "ItemDetails": quote_data.get("item_details[]", [])[i],
            "Brand": quote_data.get("brand[]", [])[i],
            "Quantity": qty,
            "Rate": rate,
            "Margin": margin,
            "Tax": tax,
            "Discount": discount,
            "Amount": amount,
            "SellingPrice": selling
        })

        total_amount += amount
        total_selling += selling
        total_discount += discount
        total_tax_amount += tax_amount
        total_margin += margin
        total_rate += rate * qty
        count += 1

    avg_tax_percentage = (total_tax_amount / total_amount * 100) if total_amount else 0

    item_fields = {
        "Title": quote_data.get("reference", ["No Title"])[0],
        "CustomerID": quote_data.get("customer_id", [""])[0],
        "Currency": quote_data.get("currency", [""])[0],
        "PaymentTerms": quote_data.get("payment_terms", [""])[0],
        "Email": quote_data.get("email", [""])[0],
        "TaxTreatment": quote_data.get("tax_treatment", [""])[0],
        "Reference": quote_data.get("reference", [""])[0],
        "QuoteDate": quote_data.get("quote_date", [""])[0],
        "ExpiryDate": quote_data.get("expiry_date", [""])[0],
        "Portal": quote_data.get("portal", [""])[0],
        "QuoteCreator": quote_data.get("quote_creator", [""])[0],
        "BCD": quote_data.get("bcd", [""])[0],
        "ApprovalStatus": "Pending",
        "Tax": avg_tax_percentage,
        "Amount": total_amount,
        "TotalDiscount": total_discount,
        "TotalSellingPrice": total_selling,
        "Margin": (total_margin / count) if count else 0,
        "Rate": (total_rate / sum([int(q) for q in quote_data.get("quantity[]", [])])) if count else 0,
        "AllItems": json.dumps(combined_items, indent=2)
    }

    print("DEBUG SharePoint payload:", json.dumps(item_fields, indent=2))

    try:
        # Add SharePoint list item
        resp_json = add_sharepoint_list_item(item_fields)
        if not resp_json:
            raise Exception("Failed to create SharePoint item")
        item_id = resp_json.get("id")

        if supplier_file and supplier_file.filename:
            print("File detected in request")
            file_name = supplier_file.filename  # Use secure_filename if needed
            file_bytes = supplier_file.read()
            print(f"File size: {len(file_bytes)} bytes")

            # Upload file to SharePoint document library folder (replace 'QuoteAttachments' with your folder path)
            share_link = upload_file_to_sharepoint_folder( folder_path="QuoteAttachments", file_name=file_name, file_bytes=file_bytes)
            print("File uploaded. Share link:", share_link)

            # Update SharePoint list item with link
            update_sharepoint_item_with_link(item_id, share_link)
        else:
            print("No file in request or filename missing")

        return render_template("pages/quote_success.html", user=user, added_items=1)

    except Exception as e:
        print("SharePoint Error:", e)
        return f"❌ Error adding quote to SharePoint: {str(e)}", 500


@app.route("/quote_decision")
def quote_decision():
    if "user" not in session:
        return redirect(url_for("login"))

    user = session.get("user")
    user_name = user.get("displayName")  # user's name in session

    try:
        site_domain = "hamdaz1.sharepoint.com"
        site_path = "/sites/Test"
        list_name = "Quotes"

        quote_items = fetch_sharepoint_list(site_domain, site_path, list_name)
        # Filter only quotes created by this user
        quote_items = [q for q in quote_items if q.get("QuoteCreator") == user_name]
    except Exception as e:
        print(f"Error fetching SharePoint list items: {e}")
        quote_items = []

    return render_template("pages/quote_decision.html", user=user, quote_items=quote_items)


@app.route("/quote_details/<quote_id>")
def quote_details(quote_id):
    if "user" not in session:
        return redirect(url_for("login"))

    user = session.get("user")

    # Fetch quote by ID from SharePoint
    site_domain = "hamdaz1.sharepoint.com"
    site_path = "/sites/Test"
    list_name = "Quotes"
    quote_items = fetch_sharepoint_list(site_domain, site_path, list_name)
    quote = next((q for q in quote_items if str(q.get("id")) == str(quote_id)), None)
    if not quote:
        return "Quote not found", 404

    # Get customer name from Zoho
    customer_id = quote.get("CustomerID", "")
    customer_name = get_customer_name_from_zoho(customer_id) or ""
    quote["CustomerName"] = customer_name

    quote["attachmentlink"] = quote.get("attachmentlink", "")
    print(quote["attachmentlink"])

    # Check Approval Status
    approval_status = quote.get("ApprovalStatus", "")
    is_approved = approval_status.lower() == "approved"
    quote["IsApproved"] = is_approved

    if is_approved:
        print("approved")

    # -----------------------------
    # Parse AllItems JSON
    # -----------------------------
    all_items_raw = quote.get("AllItems", "")
    try:
        match = re.search(r'\[.*\]', html.unescape(all_items_raw), re.DOTALL)
        if match:
            items = json.loads(match.group(0))
            for item in items:
                item.setdefault('Discount', 0)
                item.setdefault('Quantity', 1)  # default to 1 if missing

            quote['AllItems_parsed'] = items

            def to_float(val):
                try:
                    if val is None:
                        return 0.0
                    if isinstance(val, str):
                        val = val.replace("%", "").strip()
                    return float(val)
                except:
                    return 0.0

            # -----------------------------
            # Calculate Totals correctly
            # -----------------------------
            total_rate = total_amount = total_selling_price = total_tax = total_discount = 0.0
            total_margin_value = 0.0
            margin_count = 0
            total_quantity = 0.0

            for item in items:
                quantity = to_float(item.get('Quantity', 1))
                rate = to_float(item.get('Rate', 0))
                margin = to_float(item.get('Margin', 0))
                tax = to_float(item.get('Tax', 0))
                discount = to_float(item.get('Discount', 0))

                amount = rate * quantity
                item['Amount'] = amount

                tax_amount = amount * tax / 100.0
                selling_price = amount * (1 + margin / 100.0) + tax_amount - (discount * quantity)
                qty = item['Quantity']
                
                base_amount = (qty * rate * (1 + margin / 100))
                amount = (base_amount - discount)
                tax_amount = (amount * tax)
                selling = (amount + tax_amount)
                item['SellingPrice'] = selling

                # Accumulate totals per item
                total_selling_price += selling
                total_tax += tax_amount
                total_rate += amount
                total_amount += amount
                total_discount += discount * quantity

                if margin > 0:
                    total_margin_value += margin
                    margin_count += 1

                total_quantity += quantity

            avg_margin = total_margin_value / margin_count if margin_count > 0 else 0
            avg_rate = total_rate / total_quantity if total_quantity > 0 else 0

            quote['Totals'] = {
                "Rate": total_rate,
                "AvgRate": avg_rate,
                "AvgMargin": avg_margin,
                "Tax": total_tax,
                "Amount": total_amount,
                "SellingPrice": total_selling_price,
                "Discount": total_discount
            }
        else:
            quote['AllItems_parsed'] = []
            quote['Totals'] = {}
    except Exception as e:
        print("Error parsing AllItems:", e)
        quote['AllItems_parsed'] = []
        quote['Totals'] = {}

    return render_template("pages/quote_details.html", quote=quote, user=user)

@app.route("/vendors")
def vendors():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    vendors = get_excel_data_from_onedrive("Partnership_Status.xlsx", "Vendor-Partnership")
    df=pd.DataFrame(vendors)
    # Check the columns first
    print(df.columns)

    # Example: select only useful columns
    columns_of_interest = ['text', 'values']  # or actual column names in your df
    df_selected = df[columns_of_interest]

    # If 'values' contains the actual row data
    # Convert 'values' column (list) into separate columns
    df_expanded = pd.DataFrame(df['values'].tolist(), columns=[
        'ID', 'Vendor', 'Col3', 'Status', 'Col5', 'Comments', 'URL', 'Admin User', 'Password', 'Col10', 'Contact', 'Col12'
    ])

    # View the cleaned DataFrame
    print(df_expanded.head())

    return render_template("pages/vendors.html", data=df_expanded.to_dict(orient="records"), user=user)


@app.route("/vendor/<vendor_id>")
def vendor_detail(vendor_id):
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")

    vendors = get_excel_data_from_onedrive("Partnership_Status.xlsx", "Vendor-Partnership")
    df = pd.DataFrame(vendors)
    df_expanded = pd.DataFrame(df['values'].tolist(), columns=[
        'ID', 'Vendor', 'Col3', 'Status', 'Col5', 'Comments', 'URL', 'Admin User', 'Password', 'Col10', 'Contact', 'Col12'
    ])

    # Convert ID to string for comparison
    vendor = df_expanded[df_expanded['ID'].astype(str) == str(vendor_id)].to_dict(orient="records")
    if not vendor:
        return "Vendor not found", 404

    return render_template("pages/vendor_detail.html", vendor=vendor[0], user=user)

# ==============================================================

@app.route("/approvals")
def approvals():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    site_domain = "hamdaz1.sharepoint.com"
    site_path = "/sites/Test"
    list_name = "Quotes"
    quote_items = fetch_sharepoint_list(site_domain, site_path, list_name)
    # Show all the quotes with ApprovalStatus as both 'Pending' and 'Approved'
    quote_items = [q for q in quote_items if q.get("ApprovalStatus") in ["Pending", "Approved"]]
    
    return render_template("pages/quote_decision.html", user=user, quote_items=quote_items)

    





# ==============================================================

@app.route("/chatbot")
def chatbot():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    return render_template("pages/chatbot.html", user=user)


# ==============================================================

@app.route("/user_report")
def user_report():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    user_name = user.get("displayName").replace(" ", "")
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    df = items_to_dataframe(tasks)
    user_analytics_specific = get_user_analytics_specific(df, user_name)

    return render_template("pages/user_report.html", user=user, user_analytics=user_analytics_specific)

@app.route("/admin_report")
def admin_report():
    if "user" not in session:
        return redirect(url_for('login'))
    user = session.get("user")
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    df = items_to_dataframe(tasks)
    overall_analytics, per_user_analytics = get_analytics_data(df, period_type='all')
    return render_template("pages/admin_report.html", user=user, overall_analytics=overall_analytics, per_user_analytics=per_user_analytics)



# ==============================================================




# ==============================================================
# START FLASK + BACKGROUND UPDATER
# ==============================================================

if __name__ == "__main__":
    threading.Thread(target=background_updater, daemon=True).start()
    app.run(debug=True)
