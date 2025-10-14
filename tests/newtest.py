import os
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

# ----------------------------
# Load environment variables
# ----------------------------
load_dotenv(override=True)

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID]):
    raise ValueError("CLIENT_ID, CLIENT_SECRET, and TENANT_ID must be set in .env")

GRAPH_API = "https://graph.microsoft.com/v1.0"

# ----------------------------
# MSAL Authentication
# ----------------------------
def get_access_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token_response:
        raise Exception(f"Failed to get access token: {token_response}")
    return token_response["access_token"]

# ----------------------------
# Fetch all users in org
# ----------------------------
def get_all_users(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    users = []
    url = f"{GRAPH_API}/users?$select=id,displayName,mail"
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            raise Exception(f"Error fetching users: {resp.text}")
        data = resp.json()
        users.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    # Build cache: user_id -> displayName
    user_cache = {u['id']: u.get('displayName', u.get('mail')) for u in users}
    return user_cache

# ----------------------------
# Get SharePoint site ID
# ----------------------------
def get_site_id(access_token, site_domain, site_path):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_API}/sites/{site_domain}:{site_path}"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error fetching site ID: {resp.text}")
    return resp.json().get("id")

# ----------------------------
# Get SharePoint list ID
# ----------------------------
def get_list_id(access_token, site_id, list_name):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_API}/sites/{site_id}/lists"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error fetching lists: {resp.text}")
    for l in resp.json().get("value", []):
        if l.get("name") == list_name:
            return l.get("id")
    raise Exception(f"List '{list_name}' not found.")

# ----------------------------
# Get list items with fields
# ----------------------------
def get_list_items(access_token, site_id, list_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    items = []
    url = f"{GRAPH_API}/sites/{site_id}/lists/{list_id}/items?expand=fields"
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            raise Exception(f"Error fetching items: {resp.text}")
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items

# ----------------------------
# Flatten fields and replace lookup IDs using user_cache
# ----------------------------
def flatten_fields(item_fields, user_cache=None):
    flat = {}
    for k, v in item_fields.items():
        if isinstance(v, dict):
            # Use lookupValue if available, otherwise replace using user_cache
            lookup_id = v.get("id") or v.get("userId")
            if lookup_id and user_cache:
                flat[k] = user_cache.get(lookup_id, v.get("displayName") or str(v))
            else:
                flat[k] = v.get("lookupValue") or v.get("displayName") or str(v)
        elif isinstance(v, list):
            # For multi-person or multi-lookup fields
            names = []
            for i in v:
                if isinstance(i, dict):
                    lookup_id = i.get("id") or i.get("userId")
                    if lookup_id and user_cache:
                        names.append(user_cache.get(lookup_id, i.get("displayName") or str(i)))
                    else:
                        names.append(i.get("lookupValue") or i.get("displayName") or str(i))
                else:
                    names.append(str(i))
            flat[k] = ", ".join(names)
        else:
            flat[k] = v
    return flat

# ----------------------------
# Main function
# ----------------------------
def main():
    access_token = get_access_token()
    user_cache = get_all_users(access_token)

    SITE_DOMAIN = "hamdaz1.sharepoint.com"
    SITE_PATH = "/sites/ProposalTeam"
    LIST_NAME = "Proposals"

    site_id = get_site_id(access_token, SITE_DOMAIN, SITE_PATH)
    list_id = get_list_id(access_token, site_id, LIST_NAME)
    raw_items = get_list_items(access_token, site_id, list_id)

    # Pass user_cache here
    structured_items = [flatten_fields(item.get("fields", {}), user_cache) for item in raw_items]

    # Print results
    for i, item in enumerate(structured_items, 1):
        print(f"Item {i}:")
        for k, v in item.items():
            print(f"  {k}: {v}")
        print("-----")

if __name__ == "__main__":
    main()
