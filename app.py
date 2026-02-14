import email
from flask import Flask, redirect, url_for, session, request, render_template, jsonify ,abort ,send_file
from msal import ConfidentialClientApplication
import requests
import os
from dotenv import load_dotenv
import msal
from datetime import datetime
import threading
import time
import pandas as pd  # Required for timestamp conversion
from sharepoint_data import *
from sharepoint_items import *
from zoho import *
import openai  
import re
import json
import html
from pinecone import Pinecone, ServerlessSpec
from uuid import uuid4
# from openai import OpenAI

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

SUPERUSERS = ["jishad@hamdaz.com", "hisham@hamdaz.com" , "sebin@hamdaz.com" , "sujeel@hamdaz.com","shibit@hamdaz.com", "althaf@hamdaz.com"]
approvers = ["shibit@hamdaz.com", "althaf@hamdaz.com" ,"sebin@hamdaz.com" , "sujeel@hamdaz.com"]
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
EXCLUDED_USERS = excludeusers_from_sl() 


# ✅ Initialize the OpenAI Client properly
openai.api_key = os.getenv("OPENAI_API_KEY")

PINECONE_API_KEY = os.getenv("PINECONE_API_KEY")  
pc = Pinecone(api_key=PINECONE_API_KEY)

# Initialize the index
index = pc.Index("hamdaz")

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

            # Fetch latest SharePoint list
            tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
            # list_columns= get_list_columns(SITE_DOMAIN , SITE_PATH , LIST_NAME)
        
            
            df = items_to_dataframe(tasks)
            user_analytics = generate_user_analytics(df, exclude_users=EXCLUDED_USERS)

            # Calculate priority score and rank
            user_analytics = calculate_priority_score(user_analytics)
            user_analytics = assign_priority_rank(user_analytics)
            
            

            # Fetch existing SharePoint items
            existing_items = get_existing_useranalytics_items()
            
            # jobcount= user_with_jobs_ls() ////

            for _, row in user_analytics.iterrows():
                username = row["User"]
                item_fields = {
                    "Username": username,
                    "ActiveTasks": int(row["OngoingTasksCount"]),
                    "RecentDate": row["LastAssignedDate"].isoformat() if isinstance(row["LastAssignedDate"], datetime) else row["LastAssignedDate"],
                    "Priority": int(row["PriorityRank"]),
                    # "jobcount": int(jobcount.get(username, 0)) ////
                }

                # ✅ Use safe helper to find existing user
                existing_item = find_existing_user_item(existing_items, username)

                if existing_item:
                    update_user_analytics_in_sharepoint(existing_item["id"], item_fields)
                    print(f"🔄 Updated {username} in SharePoint")
                else:
                    add_item_to_sharepoint(item_fields)  
                    print(f"➕ Added new user {username} to SharePoint")
            
            # Perform smart rotation
            swp()
            
            # print("Indexing tasks to PineCone")
            # index_tasks()    # Update indexing for Pinecone  
            print(f"[BG] Data updated successfully at {datetime.now()}", flush=True)

        except Exception as e:
            print("[BG] Error during update:", e)

        time.sleep(100)


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
            elif excel_role == "ai":
                dashboard_role = "admin_dashboard"
                app.jinja_env.globals.update(excel_role="ai")
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

from urllib.parse import unquote
import time

def get_partnership_data_processed(search="", status="", page=1, per_page=50):
    raw_data = get_partnership_data()# Replace with actual loading!
    filtered = [
        row for row in raw_data
        if (search.lower() in row["Product"].lower() or search.lower() in row["Competitor Company"].lower())
        and (not status or row.get("Status", "Not Started").strip() == status)
    ]
    total_count = len(filtered)
    start = (page - 1) * per_page
    end = start + per_page
    paginated = filtered[start:end]
    return paginated, total_count

from math import ceil

@app.route("/bd")
def view_data():
    user = session.get("user", "BD_USER")  # placeholder session user
    search = request.args.get("search", "")
    status = request.args.get("status", "")
    page = int(request.args.get("page", 1))
    per_page = 50
    data, total_count = get_partnership_data_processed(search, status, page, per_page)

    # Group by Product Group Number
    grouped_data = {}
    for row in data:
        key = row.get("Product Group Number", "N/A")
        grouped_data.setdefault(key, []).append(row)

    total_pages = ceil(total_count / per_page)

    return render_template(
        "business_dev_team.html",
        grouped_data=grouped_data,
        user=user,
        search=search,
        status=status,
        page=page,
        total_pages=total_pages,
        per_page=per_page,
        total_count=total_count
    )

@app.route('/competitor/<path:competitor_name>/<path:product_name>/<path:manufacturer>', methods=['GET', 'POST'])
def competitor_profile(competitor_name, product_name, manufacturer):
    user = session["user"]
    data = get_partnership_data()  # fetch fresh data every time

    # Decode URL and restore slashes
    competitor_name = unquote(competitor_name).strip()
    product_name = unquote(product_name).replace("_slash_", "/").strip()
    manufacturer = unquote(manufacturer).replace("_slash_", "/").strip()

    competitor_entry = next(
        (row for row in data if
         row['Competitor Company'].strip() == competitor_name and
         row['Product'].strip() == product_name and
         row['ADNOC Approved Manufacturer'].strip() == manufacturer),
        None
    )

    if not competitor_entry:
        abort(404)

    other_products = [
        row for row in data
        if row['Competitor Company'].strip() == competitor_name and
        not (row['Product'].strip() == product_name and
             row['ADNOC Approved Manufacturer'].strip() == manufacturer)
    ]

    if request.method == 'POST':
        json_data = request.get_json()
        new_status = json_data.get('status')
        new_remarks = json_data.get('remarks')

        for row in data:
            if row['Competitor Company'].strip() == competitor_name:
                if new_status:
                    save_partnership_update(
                        product_group=row.get('Product Group Number'),
                        product_name=row.get('Product'),
                        manufacturer=row.get('ADNOC Approved Manufacturer'),
                        competitor_name=row.get('Competitor Company'),
                        field='Status',
                        new_value=new_status
                    )
                    row['Status'] = new_status

                if new_remarks:
                    save_partnership_update(
                        product_group=row.get('Product Group Number'),
                        product_name=row.get('Product'),
                        manufacturer=row.get('ADNOC Approved Manufacturer'),
                        competitor_name=row.get('Competitor Company'),
                        field='Remarks',
                        new_value=new_remarks
                    )
                    row['Remarks'] = new_remarks

        return jsonify(success=True)

    return render_template(
        'competitor_profile.html',
        competitor=competitor_entry,
        other_products=other_products,
        user=user
    )

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
            share_link = upload_file_to_sharepoint_folder( folder_path="Shared Documents/QuoteAttachments", file_name=file_name, file_bytes=file_bytes)
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

    attachment_str = quote.get("attachmentlink", "")
    if attachment_str:
        quote["AttachmentLinks"] = [link.strip() for link in attachment_str.split(",") if link.strip()]
    else:
        quote["AttachmentLinks"] = []

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


@app.route("/export_quote/<quote_id>")
def export_quote(quote_id):
    if "user" not in session:
        return redirect(url_for("login"))

    # --- REUSE YOUR EXISTING FETCH LOGIC HERE ---
    # (Copy the logic from quote_details to fetch and parse the quote)
    site_domain = "hamdaz1.sharepoint.com"
    site_path = "/sites/Test"
    list_name = "Quotes"
    quote_items = fetch_sharepoint_list(site_domain, site_path, list_name)
    quote = next((q for q in quote_items if str(q.get("id")) == str(quote_id)), None)
    
    if not quote:
        return "Quote not found", 404

    # Fetch Customer Name
    customer_id = quote.get("CustomerID", "")
    customer_name = get_customer_name_from_zoho(customer_id) or ""
    quote["CustomerName"] = customer_name

    # Parse AllItems (Required for the loop)
    import json, re, html
    all_items_raw = quote.get("AllItems", "")
    try:
        match = re.search(r'\[.*\]', html.unescape(all_items_raw), re.DOTALL)
        if match:
            quote['AllItems_parsed'] = json.loads(match.group(0))
        else:
            quote['AllItems_parsed'] = []
    except:
        quote['AllItems_parsed'] = []
    # ---------------------------------------------

    # Generate the Excel file
    excel_file = generate_quote_excel(quote)
    
    # Return as a download
    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"Quote_{quote_id}.xlsx"
    )


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
# Global chat histories
chat_histories = {}

# Helper to get tasks for a user
def get_tasks_for_user(username):
    now_utc = pd.Timestamp.utcnow()
    user_tasks = [t for t in tasks if t.get("AssignedTo", "").replace(" ", "") == username]
    return user_tasks


# ============================================================================
# HELPER FUNCTIONS FOR RAG
# ============================================================================
def get_embeddings_batch(texts):
    """Batch embed texts (10-100x faster than individual calls)"""
    if not texts:
        return []
    
    all_embeddings = []
    for i in range(0, len(texts), 2048):  # OpenAI limit: 2048 texts/request
        batch = texts[i:i+2048]
        response = openai.Embedding.create(model="text-embedding-ada-002", input=batch)
        all_embeddings.extend([item['embedding'] for item in response['data']])
    
    return all_embeddings


def get_embedding(text):
    """Single text embedding (use batch version when possible)"""
    return openai.Embedding.create(model="text-embedding-ada-002", input=text)['data'][0]['embedding']



def upsert_tasks_to_pinecone(tasks, is_admin, period_type="month"):
    """Index tasks + analytics into Pinecone - only update when data changed."""
    pc = Pinecone(api_key=os.getenv("PINECONE_API_KEY"))
    index = pc.Index("hamdaz")

    all_items = []   # (id, text, metadata)

    # -------------------------------
    # 1. BUILD TASKS + ANALYTICS LIST
    # -------------------------------
    if is_admin:
        overall_analytics, per_user_analytics = get_analytics_data(df, period_type='all')
        period_analytics, period_per_user = get_analytics_data(df, period_type=period_type)

        analytics_text = json.dumps({
            "overall_analytics": overall_analytics,
            "per_user_analytics": per_user_analytics,
            "period_analytics": period_analytics,
            "period_per_user": period_per_user
        }, indent=2)

        all_items.append((
            f"analytics-{period_type}",
            analytics_text,
            {
                "type": "analytics",
                "period_type": str(period_type),
                "text": analytics_text[:40000]
            }
        ))

    for task in tasks:
        task_text = json.dumps(task, default=str)
        assigned_to = task.get("AssignedTo", "").replace(" ", "")
        task_id = task.get("id") or task.get("ID") or task.get("TaskID") or str(uuid4())

        all_items.append((
            f"task-{task_id}",
            task_text,
            {
                "type": "task",
                "assigned_to": str(assigned_to),
                "is_admin": bool(is_admin),
                "full_task": task_text
            }
        ))

    # Gather all ids we need to check
    ids = [item[0] for item in all_items]

    # -------------------------------------
    # 2. FETCH EXISTING RECORDS FROM INDEX
    # -------------------------------------
    existing = index.fetch(ids).vectors
    print(f"Fetched {len(existing)} existing records.", flush=True)

    # -------------------------------------
    # 3. DETERMINE WHICH ITEMS CHANGED
    # -------------------------------------
    to_embed = []
    to_upsert = []

    for item_id, text, metadata in all_items:
        old = existing.get(item_id)

        if not old:
            # New record — must embed and upsert
            to_embed.append((item_id, text, metadata))
            continue

        old_meta = old.metadata or {}

        old_text = old_meta.get("full_task") or old_meta.get("text") or None


        if old_text != text:
            # content changed → regenerate embedding and upsert
            to_embed.append((item_id, text, metadata))
        else:
            # content is identical → SKIP
            pass

    print(f"🔥 Records changed: {len(to_embed)-1}", flush=True)

    if not to_embed:
        print("✅ No changes detected — nothing to upsert.", flush=True)
        return True

    # -------------------------------------
    # 4. BATCH GENERATE EMBEDDINGS
    # -------------------------------------
    texts = [x[1] for x in to_embed]
    new_embeddings = get_embeddings_batch(texts)

    # -------------------------------------
    # 5. BUILD UPSERT PAYLOAD
    # -------------------------------------
    for (item_id, text, metadata), embed in zip(to_embed, new_embeddings):
        to_upsert.append({
            "id": item_id,
            "values": embed,
            "metadata": metadata
        })

    # -------------------------------------
    # 6. BATCH UPSERT ONLY CHANGED ITEMS
    # -------------------------------------
    batch_size = 100
    for i in range(0, len(to_upsert), batch_size):
        index.upsert(vectors=to_upsert[i:i+batch_size])
        print(f"  ↳ Upserted batch {i//batch_size + 1}", flush=True)

    print(f"✅ Completed upsert of {len(to_upsert)} changed items.", flush=True)
    return True

def query_relevant_data(user_query, username, is_admin, top_k=10):
    """Query Pinecone for relevant tasks and analytics"""
    query_embedding = get_embedding(user_query)
    
    # Admin sees all, users see only their tasks
    filter_dict = None if is_admin else {"assigned_to": username, "type": "task"}
    pc = Pinecone(api_key=os.getenv("PINECONE_API_KEY"))
    index = pc.Index("hamdaz")
    results = index.query(
        vector=query_embedding,
        top_k=top_k,
        include_metadata=True,
        filter=filter_dict
    )
    
    tasks, analytics = [], []
    
    for match in results.get('matches', []):
        meta = match.get('metadata', {})
        
        if meta.get('type') == 'task':
            task_info = {
                "score": match['score'],
                "task": json.loads(meta.get('full_task', '{}'))
            }
            tasks.append(task_info)
                
        elif meta.get('type') == 'analytics':
            analytics.append({
                "score": match['score'],
                "analytics": json.loads(meta.get('text', '{}'))
            })
    
    return tasks, analytics 

@app.route("/chatbot")
def chatbot():
    if "user" not in session:
        return redirect(url_for("login"))
    
    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")

    # Initialize chat history if not exists
    if email not in chat_histories:
        chat_histories[email] = [{"role": "system", "content": "You are a helpful assistant."}]
    
    return render_template("pages/chatbot.html", user=user)


def index_tasks():
    """
    Index all tasks to Pinecone (called from background updater).
    """
    try:
        print("[BG] 🔄 Indexing tasks to Pinecone...", flush=True)
        
        print(f"[BG] 📊 Indexing {len(tasks)} tasks...", flush=True)
        
        # Use a system email for background indexing
        # system_email = "system@hamdaz.com"
        
        # Index all tasks as admin
        upsert_tasks_to_pinecone(tasks, is_admin=True)
        
        print(f"[BG] ✅ Successfully indexed {len(tasks)} tasks", flush=True)
        return True
        
    except Exception as e:
        print(f"[BG] ❌ Error indexing tasks: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False
# ...existing code...
@app.route("/chatbot/ask", methods=["POST"])
def ask_chatbot():
    if "user" not in session:
        return jsonify({"error": "Unauthorized"}), 401

    user = session["user"]
    email = user.get("mail") or user.get("userPrincipalName")
    username = user.get("displayName", "").replace(" ", "")
    is_admin_user = is_admin(email)
    
    user_prompt = request.json.get("message", "")
    period_type = request.json.get("period", "month")

    try:
        # 1) Try semantic retrieval from Pinecone
        tasks, analytics = query_relevant_data(
            user_prompt, username, is_admin_user, top_k=200
        )

        # 2) FALLBACK for non-admins: if Pinecone returned no tasks, load directly from SharePoint/global cache
        if not is_admin_user and (not tasks or len(tasks) == 0):
            print("[CHATBOT] Pinecone returned no matches — falling back to SharePoint tasks for user", flush=True)
            sp_tasks = get_tasks_for_user(username)
            tasks = [{"score": 1.0, "task": t} for t in sp_tasks] if sp_tasks else []

        # 3) Build contextual system prompt
        if is_admin_user:
            # Admin: include analytics + all matched tasks
            if analytics:
                analytics_data = analytics[0].get('analytics', {})
                per_user_analytics = analytics_data.get('per_user_analytics', {})
            else:
                _, per_user_analytics = get_analytics_data(df, period_type='all')

            tasks_list = [t['task'] for t in tasks] if tasks else []
            context = (
                f"ADMIN CONTEXT\n\n"
                f"Analytics (per user):\n{json.dumps(per_user_analytics, default=str)}\n\n"
                f"Matched Tasks:\n{json.dumps(tasks_list, default=str)}\n\n"
                "You may use the analytics and tasks above to answer the user's question. asosciate the data properly for better"
            )
        else:
            # Non-admin: MUST answer only from the user's tasks. If no tasks, explicitly state cannot answer.
            if tasks:
                context = (
                    f"USER: {username}\n\n"
                    f"Use ONLY the following tasks to answer the question. Do NOT use any external knowledge.\n\n"
                    f"{json.dumps([t['task'] for t in tasks], default=str)}\n\n"
                    "INSTRUCTIONS: Answer ONLY using the information above. If the question cannot be answered from this data, reply exactly: \"I don't know based on available data.\" Keep the answer concise."
                )
            else:
                context = (
                    f"USER: {username}\n\n"
                    "No task data is available for this user. You must reply: \"I don't know based on available data.\""
                )

        # 4) Compose messages and call model
        messages = [{"role": "system", "content": context}]
        last_summary = session.get(f"{email}_last_summary")
        if last_summary:
            messages.append({"role": "system", "content": f"Previous: {last_summary}"})
        messages.append({"role": "user", "content": user_prompt})

        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.0,   # deterministic answers when relying on data
            max_tokens=8000
        )
        reply = response.choices[0].message.content.strip()

        # 5) Short summary for session memory
        summary_resp = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "Summarize in 1 sentence."},
                {"role": "user", "content": f"User: {user_prompt}\nAssistant: {reply}"}
            ],
            temperature=0.0,
            max_tokens=60
        )
        summary = summary_resp.choices[0].message.content.strip()
        session[f"{email}_last_summary"] = summary

        return jsonify({"reply": reply, "summary": summary})

    except Exception as e:
        print(f"❌ Error in ask_chatbot: {e}", flush=True)
        return jsonify({"error": str(e)}), 500
# ...existing code...
# ==============================================================

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
from PyPDF2 import PdfMerger
import io
DEFAULT_PDF_PATH = 'default.pdf'

@app.route('/merge', methods=['GET', 'POST'])
def merge():
    if request.method == 'POST':
        # Get PDFs from the form
        uploaded_files = request.files.getlist('pdf_files')

        # Prepare in-memory PDFs
        pdf_streams = []
        for f in uploaded_files:
            pdf_streams.append(io.BytesIO(f.read()))

        # Include default PDF if requested
        include_default = request.form.get('include_default')
        if include_default == 'on':
            default_pdf_position = int(request.form.get('default_position', 0))
            with open(DEFAULT_PDF_PATH, 'rb') as df:
                default_pdf_stream = io.BytesIO(df.read())
            pdf_streams.insert(default_pdf_position, default_pdf_stream)

        # Merge PDFs
        merger = PdfMerger()
        for pdf_io in pdf_streams:
            merger.append(pdf_io)

        output_pdf = io.BytesIO()
        merger.write(output_pdf)
        merger.close()
        output_pdf.seek(0)

        return send_file(output_pdf, as_attachment=True, download_name='merged.pdf', mimetype='application/pdf')

    return render_template('merge_tp.html' , user=session.get("user"))

import PyPDF2
# Make sure to set your OpenAI API key
openai.api_key = os.getenv("OPENAI_API_KEY")
CHUNK_SIZE = 3000  # characters per chunk


@app.route("/tp")
def tp():
    try:
        items = get_child_files()
        files = []
        for f in items:
            if "file" in f:
                files.append({
                    "name": f["name"],
                    "url": f["webUrl"],
                    "size": f["size"],
                    "modified": f["lastModifiedDateTime"],
                    "id": f["id"]
                })
        # Default attachments for table
        default_attachments = [
            {"name": f["name"], "id": f["id"], "url": f["url"]}
            for f in files
        ]
    except Exception as e:
        print("Error:", e)
        files = []
        default_attachments = []

    return render_template("tp.html", files=files, default_attachments=default_attachments, user=session.get("user"))
# ddlf

# --- ADD THIS NEW ROUTE ---
@app.route("/download/docx/<file_id>")
def download_source_docx(file_id):
    try:
        # 1. Download the raw DOCX from your source (SharePoint/Graph)
        docx_path = download_docx(file_id)
        
        # 2. Send the DOCX directly (NO PDF CONVERSION)
        return send_file(
            docx_path, 
            as_attachment=True, 
            download_name=os.path.basename(docx_path),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print("Error downloading DOCX:", e)
        return f"Error: {e}", 500

@app.route("/download/file/<file_id>")
def download_file_from_onedrive(file_id):
    try:
        file_path, file_name = download_file(file_id)
        return send_file(
            file_path,
            as_attachment=True,
            download_name=file_name
        )
    except Exception as e:
        print("Error downloading file:", e)
        return f"Error: {e}", 500
    
    
    
@app.route('/analyze', methods=['POST'])
def analyze_document():
    # --- Input Validation ---
    if 'pdf_file' not in request.files:
        return jsonify({"error": "No PDF file uploaded"}), 400

    pdf_file = request.files['pdf_file']
    
    # --- Step 1: Extract text from PDF ---
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
    except Exception as e:
        return jsonify({"error": f"Failed to read PDF: {str(e)}"}), 400

    text_content = ""
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            # Clean extra whitespace & preserve structure
            page_text = "\n".join([line.strip() for line in page_text.splitlines() if line.strip()])
            text_content += page_text + "\n"

    if not text_content.strip():
        return jsonify({"error": "No text could be extracted from PDF"}), 400

    # --- Step 2: Split text into manageable chunks ---
    chunks = [text_content[i:i+CHUNK_SIZE] for i in range(0, len(text_content), CHUNK_SIZE)]

    # --- Step 3: Prepare final result (UPDATED) ---
    final_result = {
        "requirements": [],
        "attachments_needed": [],
        "eligibility_criteria": [],
        "deadlines": [],
        "technical_specifications": [],
        "other_notes": [],
        "required_items_needed": [] 
    }

    # --- Step 4: Analyze each chunk with OpenAI ---
    for chunk in chunks:
        prompt = f"""
You are an expert RFQ/SOW analyst.

Step 1: Read the document chunk below.

Step 2: Extract all relevant information: requirements, attachments, eligibility criteria, deadlines, technical specifications, **required items needed**, and other notes.
# 👆 Instruction updated

Return your answer in JSON exactly like this:

{{
  "requirements": [...],
  "attachments_needed": [...],
  "eligibility_criteria": [...],
  "deadlines": [...],
  "technical_specifications": [...],
  "other_notes": [...],
  "required_items_needed": [...]
  
}}

If a field is missing, return an empty list.

Document Text:
{chunk}
        """

        try:
            # NOTE: Ensure 'openai' library is configured before running.
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are an expert RFQ/SOW analyst."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0,
                max_tokens=2000
            )

            ai_output = response['choices'][0]['message']['content']
            ai_output_clean = ai_output.replace("```json", "").replace("```", "").strip()

            try:
                chunk_data = json.loads(ai_output_clean)
                # Merge chunk data into final result
                # This loop handles all keys, including the new one
                for key in final_result.keys():
                    if key in chunk_data and isinstance(chunk_data[key], list):
                        final_result[key].extend(chunk_data[key])
            except json.JSONDecodeError:
                # fallback for unparsable chunk
                final_result.setdefault("raw_text_chunks", []).append(ai_output_clean)

        except Exception as e:
            # Handle potential API errors (e.g., network, rate limit)
            return jsonify({"error": f"OpenAI API Error: {str(e)}"}), 500

    # --- Step 5: Final Cleanup and Return ---
    # Remove duplicates from lists
    for key in final_result:
        if isinstance(final_result[key], list):
            # Using dict.fromkeys to efficiently remove duplicates while preserving order
            final_result[key] = list(dict.fromkeys(final_result[key]))

    return jsonify({"extracted_data": final_result})



@app.route('/find_distributors', methods=['POST'])
def find_distributors():
    try:
        data = request.get_json()
        item_name = data.get('item_name')
        location = data.get('location') # Expected to be 'UAE'

        if not item_name:
            return jsonify({"error": "Item name is required"}), 400

        # This query uses the OpenAI model to perform the search/analysis
        prompt = f"""
        Find and list at least 3 authorized distributors or major retailers for '{item_name}' in the {location}.
        For each distributor, provide their official name and a direct link to their company profile or a relevant product page.

        Return your answer in JSON exactly like this:

        {{
          "distributors": [
            {{
              "name": "Distributor Name 1",
              "link": "https://profile.link.1"
            }},
            {{
              "name": "Distributor Name 2",
              "link": "https://profile.link.2"
            }}
          ]
        }}
        If no distributors are found, return an empty list for 'distributors'.
        """
        
        response = openai.ChatCompletion.create(
            model="gpt-4o", # Use a model with strong web-browsing capabilities
            messages=[
                {"role": "system", "content": "You are a specialized procurement agent who searches the web for product distributors and returns results in a clean JSON format."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=1000
        )

        ai_output = response['choices'][0]['message']['content']
        ai_output_clean = ai_output.replace("```json", "").replace("```", "").strip()

        try:
            # Parse and return the JSON directly to the frontend
            distributor_data = json.loads(ai_output_clean)
            return jsonify(distributor_data), 200
        except json.JSONDecodeError:
            return jsonify({"error": "AI returned unparsable JSON response.", "raw_output": ai_output_clean}), 500

    except Exception as e:
        # Catch API key errors, network errors, etc.
        return jsonify({"error": f"Failed to find distributors: {str(e)}"}), 500



from werkzeug.datastructures import FileStorage
def get_file_extension(filename):
    """Returns the file extension in lowercase."""
    return filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''


from pypdf import PdfReader # <-- This line is CRITICAL and must be at the top!

@app.route('/rfq_vs_quotes', methods=['GET'])
def load_page():
    """Renders the HTML file uploader page."""
    # Assumes HTML file is at 'templates/pages/rfq_vs_quotes.html'
    return render_template("pages/rfq_vs_quotes.html", user=session.get("user"))

@app.route('/rfq_vs_quotes', methods=['POST'])
def process_files():
    """
    Handles file upload, parses content (PDF, Excel, CSV), sends text to GPT-4o 
    for comparison, and returns a structured JSON result.
    """
    # 1. Input Validation and File Fetching
    if 'rfq_data' not in request.files or 'quote_data' not in request.files:
        return jsonify({'message': 'Missing one or both files in the request'}), 400

    rfq_file: FileStorage = request.files['rfq_data']
    quote_file: FileStorage = request.files['quote_data']
    
    if rfq_file.filename == '' or quote_file.filename == '':
        return jsonify({'message': 'One or more files have no filename'}), 400

    try:
        # --- CONSOLIDATED LOGIC: FILE PARSING (Inside the route) ---
        
        def parse_file_content_internal(file: FileStorage) -> str:
            """Reads and parses file content into a single string for AI analysis."""
            ext = get_file_extension(file.filename)
            file_bytes = file.stream.read()

            if ext in ['xlsx', 'xls', 'csv']:
                # Handle Excel/CSV using Pandas
                file_stream = io.BytesIO(file_bytes)
                
                if ext == 'csv':
                    df = pd.read_csv(file_stream)
                else:
                    df = pd.read_excel(file_stream)
                
                # Convert DataFrame to Markdown format for clear input to the AI model
                return df.to_markdown(index=False) 

            elif ext == 'pdf':
                # --- PDF Text Extraction Implementation ---
                reader = PdfReader(io.BytesIO(file_bytes))
                text = ""
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n\n" # Separate pages with double newline
                
                if not text.strip():
                    # Fallback for scanned/image-only PDFs
                    return f"PDF CONTENT ERROR: Text extraction failed for {file.filename} (might be scanned image). AI will only process the filename."
                    
                return text.strip()

            elif ext in ['txt']:
                # Simple text file
                return file_bytes.decode('utf-8')
                
            else:
                raise ValueError(f"Unsupported file type: .{ext}")

        # Execute parsing
        rfq_text = parse_file_content_internal(rfq_file)
        quote_text = parse_file_content_internal(quote_file)
        
        # --- CONSOLIDATED LOGIC: AI COMPARISON ---
        
        openai.api_key = os.getenv("OPENAI_API_KEY")
        if not openai.api_key:
            return jsonify({'message': "Configuration Error: OPENAI_API_KEY not set."}), 500

        json_structure = {
            "items_requested": "[list of items requested in RFQ]",
            "items_quoted": "[list of items quoted]",
            "differences_in_items": "[list of differences in items, including alternates or EOL items]",
            "discrepancies": "[list of pricing/term discrepancies]",
            "potential_issues": "[list of potential issues, e.g., EOL parts]",
            "summary": "summary of differences"
        }
        
        prompt_content = f"""
        You are an expert in analyzing RFQ and Quote documents.
        Given the following RFQ and Quote texts, identify discrepancies. 
        - Make sure the quoted items are the same as the requested items in the RFQ.
        - If items are different, highlight the differences.
        - If the items are alternatives, mention that.
        - If end-of-life (EOL) items are quoted, mention that.
        - If the RFQ contained an EOL item and an alternative was quoted, mention that.

        Provide the results **strictly** in the following JSON format: {json.dumps(json_structure, indent=2)}

        RFQ Text:
        ---
        {rfq_text}
        ---
        Quote Text:
        ---
        {quote_text}
        ---
        """
        
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", 
                 "content": "You are a specialized JSON output assistant. Your only task is to analyze the provided texts and return the analysis strictly as a single JSON object. DO NOT include any text outside the JSON object."},
                {"role": "user", "content": prompt_content}
            ],
            temperature=0,
            max_tokens=1500,
            response_format={"type": "json_object"}
        )
        
        json_string = response.choices[0].message.content
        comparison_result_dict = json.loads(json_string)
        
        # 4. Return result
        if "error" in comparison_result_dict:
             return jsonify({
                'message': 'AI Comparison Failed',
                'comparison_result': comparison_result_dict
            }), 500

        return jsonify({
            'message': f'Files processed and compared successfully: {rfq_file.filename} vs {quote_file.filename}',
            'comparison_result': comparison_result_dict
        }), 200
        
    except ValueError as ve:
        return jsonify({'message': f'File Processing Error (Unsupported Type/Content): {str(ve)}'}), 400
    except json.JSONDecodeError as e:
        return jsonify({'message': f'AI Model Error: Failed to parse JSON response. Details: {str(e)}'}), 500
    except openai.error.OpenAIError as e:
        return jsonify({'message': f'OpenAI API Error: {str(e)}'}), 500
    except Exception as e:
        print(f"Error during file processing: {e}")
        return jsonify({'message': f'Internal Server Error: {str(e)}'}), 500
    
    
# need to add a route to save the DistributorsData to shgarepoint list -->  get data from html  page form and save to sharepoint list

# @app.route('/assist_rfq', methods=['GET', 'POST'])
# def assist():
#     if "user" not in session:
#         return redirect(url_for('login'))
#     user = session.get("user")
#     if request.method == "POST":
        
#     return render_template("pages/assist.html", user=user)  

# ==============================================================

# @app.route("/assistant" , methods=["GET"])
# def assistant():
#     if "user" not in session:
#         return redirect(url_for('login'))
#     return render_template("pages/assistant.html", user=session.get("user"))



# client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


# @app.route('/process-agent', methods=['POST'])
# def process():
#     data = request.json
#     pdf_text = data.get('text', '')

#     if not pdf_text:
#         return jsonify({"error": "No text provided"}), 400

#     try:
#         # Standard OpenAI Chat Completion call
#         response = client.chat.completions.create(
#             model="gpt-4o",
#             messages=[
#                 {"role": "system", "content": "You are a data extraction expert. Extract the items, quantities, and requirements from the following quotation text and return them as a structured list."},
#                 {"role": "user", "content": pdf_text}
#             ],
#             temperature=0.2 # Lower temperature for more accurate extraction
#         )

#         # Get the text response from the model
#         ai_output = response.choices[0].message.content

#         return jsonify({
#             "status": "success",
#             "output": ai_output
#         })

#     except Exception as e:
#         print(f"OpenAI Error: {e}")
#         return jsonify({"error": "Failed to process text with AI"}), 500

# ==============================================================
# START FLASK + BACKGROUND UPDATER
# ==============================================================
threading.Thread(target=background_updater, daemon=True).start()

if __name__ == "__main__":
    app.run(debug=True)
    


