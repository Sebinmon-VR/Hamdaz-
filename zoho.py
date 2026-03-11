import requests
import pandas as pd
import os
import json
import time
from datetime import datetime, timedelta
from dotenv import load_dotenv, set_key

# =======================
# CONFIGURATION
# =======================
load_dotenv(override=True)

CLIENT_ID = os.getenv("zoho_CLIENT_ID")
CLIENT_SECRET = os.getenv("zoho_CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("zoho_REFRESH_TOKEN")
ORGANIZATION_ID = os.getenv("zoho_ORGANIZATION_ID")
BASE_URL = "https://www.zohoapis.com/books/v3"

ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
TOKEN_EXPIRY = os.getenv("TOKEN_EXPIRY")

# =======================
# ACCESS TOKEN MANAGEMENT
# =======================
def get_access_token():
    global ACCESS_TOKEN, TOKEN_EXPIRY
    if ACCESS_TOKEN and TOKEN_EXPIRY:
        expiry = datetime.fromisoformat(TOKEN_EXPIRY)
        if datetime.utcnow() < expiry:
            return ACCESS_TOKEN

    print("🔑 Refreshing Access Token...")
    url = "https://accounts.zoho.com/oauth/v2/token"
    params = {
        "refresh_token": REFRESH_TOKEN,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "refresh_token"
    }
    response = requests.post(url, params=params)
    data = response.json()

    if "access_token" in data:
        ACCESS_TOKEN = data["access_token"]
        expiry_time = datetime.utcnow() + timedelta(seconds=3500)
        TOKEN_EXPIRY = expiry_time.isoformat()
        set_key(".env", "ACCESS_TOKEN", ACCESS_TOKEN)
        set_key(".env", "TOKEN_EXPIRY", TOKEN_EXPIRY)
        return ACCESS_TOKEN
    else:
        raise Exception(f"❌ Zoho Auth Error: {data}")

# =======================
# PAGINATED FETCH FUNCTIONS
# =======================
def fetch_data_paginated(endpoint, key):
    all_records = []
    page = 1
    has_more = True
    access_token = get_access_token()

    print(f"📡 Initializing scan for module: {endpoint.upper()}")
    while has_more:
        url = f"{BASE_URL}/{endpoint}?organization_id={ORGANIZATION_ID}&page={page}&per_page=200"
        headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
        
        response = requests.get(url, headers=headers).json()
        records = response.get(key, [])
        all_records.extend(records)
        
        print(f"   ∟ Page {page}: Found {len(records)} records. (Total so far: {len(all_records)})")
        
        page_context = response.get("page_context", {})
        has_more = page_context.get("has_more_page", False)
        page += 1
        
    return all_records

# Summary Wrappers
def fetch_items(): return fetch_data_paginated("items", "items")
def fetch_quotes_list(): return fetch_data_paginated("estimates", "estimates")

# =======================
# DEEP DETAIL FETCHERS
# =======================
def get_specific_quote(estimate_id):
    access_token = get_access_token()
    url = f"{BASE_URL}/estimates/{estimate_id}?organization_id={ORGANIZATION_ID}"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("estimate")

def fetch_all_quotes_everything():
    # 1. Discovery Phase
    summary_list = fetch_quotes_list()
    full_detailed_data = []
    total_to_fetch = len(summary_list)
    
    print(f"\n🚀 DISCOVERY COMPLETE: {total_to_fetch} Estimate IDs found.")
    print(f"🛠 Starting Deep Extraction (fetching line items and brands)...")
    
    # 2. Extraction Phase
    for index, summary in enumerate(summary_list, start=1):
        e_id = summary.get("estimate_id")
        e_no = summary.get("estimate_number")
        
        # DEBUG: Progress Indicator
        print(f"   [{index}/{total_to_fetch}] Processing {e_no}...", end="\r")
        
        try:
            detail = get_specific_quote(e_id)
            if detail:
                full_detailed_data.append(detail)
            else:
                print(f"\n⚠️ Warning: No detail data returned for {e_no}")
        except Exception as e:
            print(f"\n❌ Error fetching details for {e_no}: {str(e)}")
            
        # Throttling to prevent 429 Too Many Requests
        time.sleep(0.15)
        
    print(f"\n✨ Extraction finished. {len(full_detailed_data)} full objects collected.")
    return full_detailed_data

# =======================
# DATA STRUCTURING
# =======================
def structure_full_estimates_table(full_quotes):
    print("📊 Formatting data for Excel...")
    rows = []
    for q in full_quotes:
        brands = []
        for item in q.get("line_items", []):
            for cf in item.get("item_custom_fields", []):
                if cf.get("api_name") == "cf_brand":
                    brands.append(cf.get("value"))
        
        unique_brands = ", ".join(list(set(brands)))
        cf_hash = q.get("custom_field_hash", {})

        rows.append({
            "Estimate Number": q.get("estimate_number"),
            "Customer": q.get("customer_name"),
            "Date": q.get("date"),
            "Total": q.get("total"),
            "Brand": unique_brands,
            "Portal": cf_hash.get("cf_portal", ""),
            "Creator": cf_hash.get("cf_quote_creater", ""),
            "Status": q.get("status")
        })
    return pd.DataFrame(rows)





# =======================
# FETCH DATA FUNCTIONS
# =======================
def fetch_data(endpoint, key):
    """Generic fetcher for Zoho Books API"""
    access_token = get_access_token()
    url = f"{BASE_URL}/{endpoint}?organization_id={ORGANIZATION_ID}"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data.get(key, [])

def fetch_items():
    return fetch_data("items", "items")

def fetch_customers():
    return fetch_data("contacts", "contacts")

# def fetch_quotes():
    # return fetch_data("quotes", "quotes")


def fetch_sales_orders():
    return fetch_data("salesorders", "salesorders")

# def fetch_vendors():
#     return fetch_data("vendors", "vendors")



def get_customer_name_from_zoho(customer_id):
    """
    Fetch all customers from Zoho and return the name of the customer
    matching the given customer_id.
    """
    try:
        customers = fetch_customers()
        structured_customers = structure_customers_data(customers)

        for cust in structured_customers:
            if str(cust.get("contact_id")) == str(customer_id):
                # Prefer company name or contact name if available
                return (
                    cust.get("company_name")
                    or cust.get("contact_name")
                    or cust.get("customer_name")
                    or "Unknown Customer"
                )
        return "Unknown Customer"
    except Exception as e:
        print(f"⚠️ Error fetching customer name from Zoho: {e}")
        return "Unknown Customer"




# =======================
# STRUCTURE DATA FUNCTIONS
# =======================
def structure_items_data(items):
    structured_data = []
    for item in items:
        structured_data.append({
            "item_id": item.get("item_id", ""),
            "name": item.get("name", ""),
            "unit": item.get("unit", ""),
            "status": item.get("status", ""),
            "source": item.get("source", ""),
            "rate": item.get("rate", 0),
            "tax_name": item.get("tax_name", ""),
            "tax_percentage": item.get("tax_percentage", 0),
            "purchase_account_name": item.get("purchase_account_name", ""),
            "purchase_rate": item.get("purchase_rate", 0),
            "can_be_sold": item.get("can_be_sold", False),
            "can_be_purchased": item.get("can_be_purchased", False),
            "track_inventory": item.get("track_inventory", False),
            "product_type": item.get("product_type", ""),
            "is_taxable": item.get("is_taxable", False),
            "description": item.get("description", ""),
            "created_time": item.get("created_time", ""),
            "last_modified_time": item.get("last_modified_time", ""),
            "brand": item.get("cf_brand", "")
        })
    return pd.DataFrame(structured_data)

def structure_customers_data(customers):
    """
    Takes a list of raw customer dictionaries and returns a structured
    list of dictionaries ready to be passed to Flask/Jinja templates.
    """
    structured_data = []

    for cust in customers:
        structured_data.append({
            "contact_id": cust.get("contact_id", ""),
            "contact_name": cust.get("contact_name", ""),
            "customer_name": cust.get("customer_name", ""),
            "vendor_name": cust.get("vendor_name", ""),
            "contact_number": cust.get("contact_number", ""),
            "company_name": cust.get("company_name", ""),
            "website": cust.get("website", ""),
            "language": cust.get("language_code_formatted", ""),
            "contact_type": cust.get("contact_type_formatted", ""),
            "status": cust.get("status", ""),
            "customer_sub_type": cust.get("customer_sub_type", ""),
            "source": cust.get("source", ""),
            "is_linked_with_zohocrm": cust.get("is_linked_with_zohocrm", False),
            "payment_terms": cust.get("payment_terms_label", ""),
            "currency_code": cust.get("currency_code", ""),
            "outstanding_receivable_amount": cust.get("outstanding_receivable_amount", 0.0),
            "outstanding_payable_amount": cust.get("outstanding_payable_amount", 0.0),
            "unused_credits_receivable_amount": cust.get("unused_credits_receivable_amount", 0.0),
            "unused_credits_payable_amount": cust.get("unused_credits_payable_amount", 0.0),
            "first_name": cust.get("first_name", ""),
            "last_name": cust.get("last_name", ""),
            "email": cust.get("email", ""),
            "phone": cust.get("phone", ""),
            "mobile": cust.get("mobile", ""),
            "portal_status": cust.get("portal_status_formatted", ""),
            "tax_treatment": cust.get("tax_treatment", ""),
            "has_attachment": cust.get("has_attachment", False),
            "created_time": cust.get("created_time_formatted", ""),
            "last_modified_time": cust.get("last_modified_time_formatted", "")
        })

    # Return as a list of dictionaries (Jinja-friendly)
    return structured_data

def structure_quotes_data(quotes):
    structured_data = []
    for q in quotes:
        structured_data.append({
            "quote_id": q.get("quote_id", ""),
            "quote_number": q.get("quote_number", ""),
            "customer_name": q.get("customer_name", ""),
            "status": q.get("status", ""),
            "date": q.get("date", ""),
            "expiry_date": q.get("expiry_date", ""),
            "total": q.get("total", 0.0),
            "currency_code": q.get("currency_code", ""),
            "notes": q.get("notes", "")
        })
    return pd.DataFrame(structured_data)








# # =======================
# # MAIN EXECUTION
# # =======================
# if __name__ == "__main__":
#     start_time = time.time()
#     try:
#         # Step 1: Deep Fetch
#         detailed_quotes = fetch_all_quotes_everything()
        
#         # Step 2: Save JSON
#         json_file = "all_quotes_full_details.json"
#         with open(json_file, "w", encoding='utf-8') as f:
#             json.dump(detailed_quotes, f, indent=4)
#         print(f"💾 JSON backup created: {json_file}")
            
#         # Step 3: Save Excel
#         excel_file = "Detailed_Estimates_Report.xlsx"
#         df_quotes = structure_full_estimates_table(detailed_quotes)
#         df_quotes.to_excel(excel_file, index=False)
#         print(f"📈 Excel report generated: {excel_file}")

#         duration = round(time.time() - start_time, 2)
#         print(f"\n🏁 FINISHED in {duration} seconds.")

#     except Exception as e:
#         print(f"\n💥 CRITICAL SCRIPT FAILURE: {e}")