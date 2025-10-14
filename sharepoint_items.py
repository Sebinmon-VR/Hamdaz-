# sharepoint_helper.py
import os
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import os
import requests
import pandas as pd
from datetime import datetime
import pytz
from collections import defaultdict
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
    """Returns a Graph API access token using client credentials."""
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
    """Fetch all users and return a cache: user_id -> displayName"""
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
    return {u['id']: u.get('displayName', u.get('mail')) for u in users}


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
# Get list items
# ----------------------------
def get_list_items(access_token, site_id, list_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    items = []
    url = f"{GRAPH_API}/sites/{site_id}/lists/{list_id}/items?expand=fields($expand=AssignedTo,Author,Editor)"
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            raise Exception(f"Error fetching items: {resp.text}")
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


# ----------------------------
# Flatten fields
# ----------------------------
def flatten_fields(item_fields, user_cache=None):
    flat = {}
    for k, v in item_fields.items():
        if isinstance(v, dict):
            lookup_id = v.get("id") or v.get("userId")
            if lookup_id and user_cache:
                flat[k] = user_cache.get(lookup_id, v.get("displayName") or str(v))
            else:
                flat[k] = v.get("lookupValue") or v.get("displayName") or str(v)
        elif isinstance(v, list):
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
# Fetch SharePoint list items as structured data
# ----------------------------
def fetch_sharepoint_list(site_domain, site_path, list_name):
    """
    Returns a list of flattened SharePoint list items with user display names.
    
    Example usage:
        items = fetch_sharepoint_list("hamdaz1.sharepoint.com", "/sites/ProposalTeam", "Proposals")
    """
    access_token = get_access_token()
    user_cache = get_all_users(access_token)

    site_id = get_site_id(access_token, site_domain, site_path)
    list_id = get_list_id(access_token, site_id, list_name)
    raw_items = get_list_items(access_token, site_id, list_id)

    structured_items = [flatten_fields(item.get("fields", {}), user_cache) for item in raw_items]
    return structured_items




import pandas as pd

def items_to_dataframe(items):
    """
    Convert a list of SharePoint item dictionaries into a Pandas DataFrame.
    
    Args:
        items (list): List of dictionaries, each representing a SharePoint item.
        
    Returns:
        pd.DataFrame: DataFrame with all fields as columns.
    """
    if not items:
        return pd.DataFrame()  # Return empty DF if no items
    
    df = pd.DataFrame(items)
    
    # Convert common date columns to datetime, if they exist
    date_cols = ['DueDate', 'BCD', 'Created', 'Modified']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    return df


def compute_overall_analytics(df):
    if df.empty:
        return {"total_users":0,"total_tasks":0,"tasks_completed":0,"tasks_pending":0,"tasks_missed":0,"orders_received":0}
    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)
    df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    total_users = df['AssignedTo'].nunique()
    total_tasks = len(df)
    tasks_completed = len(df[df['SubmissionStatus']=='Submitted'])
    tasks_pending = len(df[(df['SubmissionStatus']!='Submitted') & (df['BCD']>=now_uae)])
    tasks_missed = len(df[(df['SubmissionStatus']!='Submitted') & (df['BCD']<now_uae)])
    orders_received = len(df[df['Status']=='Received']) if 'Status' in df.columns else 0
    return {
        "total_users": total_users,
        "total_tasks": total_tasks,
        "tasks_completed": tasks_completed,
        "tasks_pending": tasks_pending,
        "tasks_missed": tasks_missed,
        "orders_received": orders_received
    }
    
    
    
    
def compute_user_analytics(df):
    if df.empty or 'AssignedTo' not in df.columns:
        return {}

    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    # Convert dates to timezone-aware
    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    if 'Start Date' in df.columns:
        df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    analytics = {}
    for user, user_df in df.groupby('AssignedTo'):
        # Last assigned date = latest Start Date
        last_assigned_date = None
        if 'Start Date' in user_df.columns and not user_df['Start Date'].isna().all():
            last_assigned_date = user_df['Start Date'].max()
            last_assigned_date = last_assigned_date.strftime("%Y-%m-%d %H:%M")

        analytics[user] = {
            "total_tasks": len(user_df),
            "tasks_completed": len(user_df[user_df['SubmissionStatus']=='Submitted']),
            "tasks_pending": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']>=now_uae)]),
            "tasks_missed": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']<now_uae)]),
            "orders_received": len(user_df[user_df['Status']=='Received']) if 'Status' in df.columns else 0,
            "last_assigned_date": last_assigned_date
        }

    return analytics



def compute_user_analytics_with_last_date(df ,EXCLUDED_USERS):
    if df.empty or 'AssignedTo' not in df.columns:
        return {}

    df = df[~df['AssignedTo'].isin(EXCLUDED_USERS)]
    if df.empty:
        return {}

    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    start_col = None
    for col in df.columns:
        if col.lower().replace(" ", "") == "startdate":
            start_col = col
            break

    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    if start_col:
        df[start_col] = pd.to_datetime(df[start_col], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    analytics = {}
    for user, user_df in df.groupby('AssignedTo'):
        last_assigned_date = user_df[start_col].max() if start_col else None
        last_assigned_date_str = last_assigned_date.strftime("%Y-%m-%d %H:%M") if pd.notna(last_assigned_date) else None

        analytics[user] = {
            "total_tasks": len(user_df),
            "tasks_completed": len(user_df[user_df['SubmissionStatus']=='Submitted']),
            "tasks_pending": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']>=now_uae)]),
            "tasks_missed": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']<now_uae)]),
            "orders_received": len(user_df[user_df['Status']=='Received']) if 'Status' in df.columns else 0,
            "last_assigned_date": last_assigned_date_str
        }
    return analytics
