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


def get_sharepoint_list_items(site_domain, site_path, list_name, cache_timeout=300):
    """
    Fetch all items from a SharePoint list using Microsoft Graph API.
    Results are cached for the specified timeout period.
    
    Args:
        site_domain (str): e.g., 'hamdaz1.sharepoint.com'
        site_path (str): e.g., '/sites/ProposalTeam'
        list_name (str): e.g., 'Proposals'
        cache_timeout (int): Cache timeout in seconds (default: 5 minutes)
    
    Returns:
        list: A list of dictionaries containing SharePoint list items.
    """
    # Try to get cached data first
    cache_key = f"sharepoint_items_{site_domain}_{site_path}_{list_name}"
    if hasattr(get_sharepoint_list_items, 'cache'):
        cached_data = get_sharepoint_list_items.cache.get(cache_key)
        if cached_data is not None:
            return cached_data
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

