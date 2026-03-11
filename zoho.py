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