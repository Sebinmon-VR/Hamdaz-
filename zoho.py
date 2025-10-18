import requests
import pandas as pd
from dotenv import load_dotenv, set_key
import os
from datetime import datetime, timedelta

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

    # Token missing or expired → request a new one
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
        expiry_time = datetime.utcnow() + timedelta(seconds=3500)  # 1 hour token
        TOKEN_EXPIRY = expiry_time.isoformat()

        # Update .env
        set_key(".env", "ACCESS_TOKEN", ACCESS_TOKEN)
        set_key(".env", "TOKEN_EXPIRY", TOKEN_EXPIRY)

        return ACCESS_TOKEN
    else:
        raise Exception(f"Zoho API: Access token not received: {data}")

def fetch_organizations():
    access_token = get_access_token()
    url = "https://www.zohoapis.com/books/v3/organizations"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data.get("organizations", [])

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

# =======================
# MAIN
# =======================
# if __name__ == "__main__":
#     try:
#         # items = fetch_items()
#         # print("Items DataFrame:")
#         # print(items)

#         # customers = fetch_customers()
#         # print("Customers DataFrame:")
#         # print(customers)

#         # organizations = fetch_organizations()
#         # print("Organizations DataFrame:")
#         # print(organizations)
        
#         data=structure_customers_data(fetch_customers())
#         print(data)
        
        
        
        
        
#     except Exception as e:
#         print("Error:", e)
