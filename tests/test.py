import requests
from msal import ConfidentialClientApplication
import os
from dotenv import load_dotenv

load_dotenv(override=True)


def get_sharepoint_list_items(site_domain, site_path, list_name):
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    tenant_id = os.getenv("TENANT_ID")

    if not all([client_id, client_secret, tenant_id]):
        raise ValueError("CLIENT_ID, CLIENT_SECRET, and TENANT_ID must be set in .env")

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

    # Get site ID
    site_url = f'https://graph.microsoft.com/v1.0/sites/{site_domain}:{site_path}'
    site_response = requests.get(site_url, headers=headers).json()
    site_id = site_response.get('id')
    if not site_id:
        raise Exception(f"Failed to get site ID: {site_response}")

    # Get list ID
    lists_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists'
    lists_response = requests.get(lists_url, headers=headers).json()
    list_id = next((l['id'] for l in lists_response.get('value', []) if l['name'] == list_name), None)
    if not list_id:
        raise Exception(f"List '{list_name}' not found on site {site_path}")

    # Get list items (all fields)
    items_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields'
    all_items = []
    while items_url:
        response = requests.get(items_url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Error fetching items: {response.text}")
        data = response.json()
        all_items.extend(data.get('value', []))
        items_url = data.get('@odata.nextLink')  # handle pagination

    # Extract fields
    result = [item.get('fields', {}) for item in all_items]
    return result


# Example usage
if __name__ == "__main__":
    items = get_sharepoint_list_items('hamdaz1.sharepoint.com', '/sites/ProposalTeam', 'Proposals')
    for i, item in enumerate(items, 1):
        print(f"Item {i}:")
        for k, v in item.items():
            print(f"  {k}: {v}")
        print("-----")
