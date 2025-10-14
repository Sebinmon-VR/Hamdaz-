import requests
from msal import ConfidentialClientApplication
import os
from dotenv import load_dotenv

load_dotenv(override=True)


def get_user_display_name(site_domain, user_id, access_token):
    """
    Fetch the display name of a SharePoint user by their ID
    """
    url = f"https://{site_domain}/_api/web/getuserbyid({user_id})"
    headers = {
        "Accept": "application/json;odata=verbose",
        "Authorization": f"Bearer {access_token}"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return data['d'].get("Title")  # Display name
    else:
        return None


def get_sharepoint_list_items(site_domain, site_path, list_name):
    """
    Fetch all items from a SharePoint list using Microsoft Graph API.
    
    Args:
        site_domain (str): e.g., 'hamdaz1.sharepoint.com'
        site_path (str): e.g., '/sites/ProposalTeam'
        list_name (str): e.g., 'Proposals'
    
    Returns:
        list: A list of dictionaries containing SharePoint list items.
    """
    # Load credentials
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    tenant_id = os.getenv("TENANT_ID")

    if not all([client_id, client_secret, tenant_id]):
        raise ValueError("CLIENT_ID, CLIENT_SECRET, and TENANT_ID must be set in .env")

    # 1. Authenticate and get token
    authority = f'https://login.microsoftonline.com/{tenant_id}'
    scopes = ['https://graph.microsoft.com/.default']

    app = ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    token_response = app.acquire_token_for_client(scopes=scopes)
    if "access_token" not in token_response:
        raise Exception(f"Failed to acquire token: {token_response.get('error_description')}")
    
    access_token = token_response['access_token']
    headers = {'Authorization': f'Bearer {access_token}'}

    # 2. Get Site ID
    site_url = f'https://graph.microsoft.com/v1.0/sites/{site_domain}:{site_path}'
    site_response = requests.get(site_url, headers=headers).json()
    site_id = site_response.get('id')
    if not site_id:
        raise Exception(f"Failed to get site ID: {site_response}")

    # 3. Get List ID
    lists_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists'
    lists_response = requests.get(lists_url, headers=headers).json()
    list_id = next((l['id'] for l in lists_response.get('value', []) if l['name'] == list_name), None)
    if not list_id:
        raise Exception(f"List '{list_name}' not found on site {site_path}")

    # 4. Get List Items
    items_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields'
    items_response = requests.get(items_url, headers=headers).json()

    return [item['fields'] for item in items_response.get('value', [])]

# ==============================================================================
# ==||| FOR BUSINESS CARDS |||==
# ==============================================================================

# --- Configuration for OneDrive Feature ---
# We load the variables again to ensure this section is self-contained.
CLIENT_ID_ONEDRIVE = os.getenv("CLIENT_ID")
CLIENT_SECRET_ONEDRIVE = os.getenv("CLIENT_SECRET")
TENANT_ID_ONEDRIVE = os.getenv("TENANT_ID")
ONEDRIVE_USER_ID = os.getenv("ONEDRIVE_USER_ID")

AUTHORITY_ONEDRIVE = f"https://login.microsoftonline.com/{TENANT_ID_ONEDRIVE}"
SCOPES_ONEDRIVE = ["https://graph.microsoft.com/.default"]
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
FILE_PATH = "Contacts.xlsx"
WORKSHEET_NAME = "Sheet1"

# A separate, dedicated MSAL app instance for our new feature.
onedrive_msal_app = ConfidentialClientApplication(
    client_id=CLIENT_ID_ONEDRIVE,
    authority=AUTHORITY_ONEDRIVE,
    client_credential=CLIENT_SECRET_ONEDRIVE,
)

def get_onedrive_access_token():
    """Acquires an access token specifically for the OneDrive functions."""
    result = onedrive_msal_app.acquire_token_silent(scopes=SCOPES_ONEDRIVE, account=None)
    if not result:
        result = onedrive_msal_app.acquire_token_for_client(scopes=SCOPES_ONEDRIVE)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Failed to acquire OneDrive token: {result.get('error_description')}")

def get_all_contacts_from_onedrive():
    """Fetches all data from the Contacts.xlsx file in the specified user's OneDrive."""
    try:
        access_token = get_onedrive_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = (
            f"{GRAPH_API_ENDPOINT}/users/{ONEDRIVE_USER_ID}/drive/root:/"
            f"{FILE_PATH}:/workbook/worksheets('{WORKSHEET_NAME}')/usedRange"
        )
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        rows = response.json().get("values", [])
        if len(rows) < 2: return []
        header = rows[0]
        return [dict(zip(header + ['row_id'], row + [i+2])) for i, row in enumerate(rows[1:])]
    except Exception as e:
        print(f"Error fetching contacts from OneDrive: {e}")
        return []

def update_contact_in_onedrive_excel(row_id, updated_data_dict):
    """Updates a single row in the Contacts.xlsx file."""
    try:
        access_token = get_onedrive_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        
        # Get header to determine column order
        header_url = (
            f"{GRAPH_API_ENDPOINT}/users/{ONEDRIVE_USER_ID}/drive/root:/"
            f"{FILE_PATH}:/workbook/worksheets('{WORKSHEET_NAME}')/range(address='A1:Z1')"
        )
        header_res = requests.get(header_url, headers=headers)
        header_res.raise_for_status()
        header = header_res.json().get("values", [[]])[0]
        if not header: raise Exception("Could not retrieve header row.")
        
        # Prepare data for update
        values_to_update = [updated_data_dict.get(col_name, "") for col_name in header]
        last_col = chr(ord('A') + len(header) - 1)
        range_address = f"A{row_id}:{last_col}{row_id}"
        
        update_url = (
            f"{GRAPH_API_ENDPOINT}/users/{ONEDRIVE_USER_ID}/drive/root:/"
            f"{FILE_PATH}:/workbook/worksheets('{WORKSHEET_NAME}')/range(address='{range_address}')"
        )
        
        patch_res = requests.patch(update_url, headers=headers, json={"values": [values_to_update]})
        patch_res.raise_for_status()
        return True
    except Exception as e:
        print(f"Error updating contact in OneDrive: {e}")
        return False


# Example usage:
if __name__ == "__main__":
    items = get_sharepoint_list_items('hamdaz1.sharepoint.com', '/sites/ProposalTeam', 'Proposals')
    for i in items:
        print(i)
