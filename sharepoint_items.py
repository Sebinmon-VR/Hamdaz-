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


def extract_usernames_from_df(df, user_columns=None, exclude_users=None):
    """
    Extract all unique usernames from a DataFrame containing SharePoint items,
    with an option to exclude certain users.
    
    Args:
        df (pd.DataFrame): DataFrame with SharePoint list items.
        user_columns (list, optional): List of column names to extract usernames from. 
                                       Defaults to ['AssignedTo', 'Author', 'Editor'].
        exclude_users (list, optional): List of usernames to exclude. Defaults to None.
    
    Returns:
        set: Set of unique usernames excluding the specified users.
    """
    if user_columns is None:
        user_columns = ['AssignedTo', 'Author', 'Editor']
    if exclude_users is None:
        exclude_users = []

    usernames = set()
    for col in user_columns:
        if col in df.columns:
            df[col].dropna().apply(
                lambda x: [usernames.add(u.strip()) for u in str(x).split(',') if u.strip() not in exclude_users]
            )

    return usernames


def get_user_details(access_token, usernames):
    """
    Fetch details of users from Microsoft Graph API given a list of usernames (UPNs).
    
    Returns a list of dicts with user information.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    user_details = []

    for username in usernames:
        # Microsoft Graph API: Get user by UPN/email
        url = f"https://graph.microsoft.com/v1.0/users/{username}"
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            user_details.append(resp.json())
        else:
            print(f"Failed to fetch {username}: {resp.text}")

    return user_details


def generate_user_analytics(df, user_column='AssignedTo', status_column='Status', 
                            title_column='Title', due_column='DueDate', exclude_users=None):
    """
    Generate per-user analytics with lists of completed, ongoing, and missed tasks,
    excluding specific users.

    Args:
        df (pd.DataFrame): DataFrame containing SharePoint items.
        user_column (str): Column name for assigned user.
        status_column (str): Column name for task status.
        title_column (str): Column name for task title.
        due_column (str): Column name for task due date.
        exclude_users (list): List of users to exclude from analytics.

    Returns:
        pd.DataFrame: Analytics per user with counts and lists of tasks.
    """
    if exclude_users is None:
        exclude_users = []

    analytics = []

    if df.empty:
        return pd.DataFrame(analytics)

    # Ensure DueDate is datetime with UTC timezone
    df[due_column] = pd.to_datetime(df[due_column], utc=True, errors='coerce')

    # Always get a tz-aware UTC timestamp safely
    now_utc = pd.Timestamp.now(tz='UTC')

    # Filter out excluded users
    df_filtered = df[~df[user_column].isin(exclude_users)]

    # Group by user
    grouped = df_filtered.groupby(user_column)

    for user, group in grouped:
        total_tasks = len(group)

        # Completed tasks
        completed_tasks_df = group[group[status_column] == 'Completed']

        # Ongoing tasks: Not submitted + due in future
        ongoing_tasks_df = group[
            (group[status_column] != 'Submitted') & 
            (group[due_column] >= now_utc)
        ]

        # Missed tasks: Not submitted + due in past
        missed_tasks_df = group[
            (group[status_column] != 'Submitted') & 
            (group[due_column] < now_utc)
        ]

        analytics.append({
            'User': user,
            'TotalTasks': total_tasks,
            'CompletedTasksCount': len(completed_tasks_df),
            'OngoingTasksCount': len(ongoing_tasks_df),
            'MissedTasksCount': len(missed_tasks_df),
            'CompletedTasks': completed_tasks_df[title_column].tolist(),
            'OngoingTasks': ongoing_tasks_df[title_column].tolist(),
            'MissedTasks': missed_tasks_df[title_column].tolist()
        })

    analytics_df = pd.DataFrame(analytics)
    return analytics_df

def get_user_analytics_specific(df: pd.DataFrame, username: str) -> dict:
    """
    Returns analytics and task lists for a specific user.
    """
    if df.empty:
        return {
            'Username': username,
            'TotalTasks': 0,
            'OngoingTasksCount': 0,
            'CompletedTasksCount': 0,
            'MissedTasksCount': 0,
            'OngoingTasks': [],
            'CompletedTasks': [],
            'MissedTasks': []
        }

    # Ensure DueDate is datetime with UTC timezone
    df['DueDate'] = pd.to_datetime(df['DueDate'], utc=True)

    # Safe UTC now
    now_utc = pd.Timestamp.utcnow()
    if now_utc.tzinfo is None:
        now_utc = now_utc.tz_localize('UTC')
    else:
        now_utc = now_utc.tz_convert('UTC')

    # Filter tasks for the user
    user_tasks = df[df['AssignedTo'] == username]

    completed_tasks = user_tasks[user_tasks['Status'] == 'Completed']

    ongoing_tasks = user_tasks[
        (user_tasks['Status'] != 'Submitted') &
        (user_tasks['DueDate'] >= now_utc)
    ]

    missed_tasks = user_tasks[
        (user_tasks['Status'] != 'Submitted') &
        (user_tasks['DueDate'] < now_utc)
    ]

    return {
        'Username': username,
        'TotalTasks': len(user_tasks),
        'OngoingTasksCount': len(ongoing_tasks),
        'CompletedTasksCount': len(completed_tasks),
        'MissedTasksCount': len(missed_tasks),
        'OngoingTasks': ongoing_tasks.to_dict('records'),
        'CompletedTasks': completed_tasks.to_dict('records'),
        'MissedTasks': missed_tasks.to_dict('records')
    }





# ==============================================================================
# ==||| FOR BUSINESS CARDS |||==
# ==============================================================================

# --- Configuration for OneDrive Feature ---

CLIENT_ID_ONEDRIVE = os.getenv("CLIENT_ID")
CLIENT_SECRET_ONEDRIVE = os.getenv("CLIENT_SECRET")
TENANT_ID_ONEDRIVE = os.getenv("TENANT_ID")
AUTHORITY_ONEDRIVE = f"https://login.microsoftonline.com/{TENANT_ID_ONEDRIVE}"
SCOPE_ONEDRIVE = ["https://graph.microsoft.com/.default"]
ONEDRIVE_USER_ID = os.getenv("ONEDRIVE_USER_ID")
FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "Contacts.xlsx")
WORKSHEET_NAME = os.getenv("ONEDRIVE_WORKSHEET_NAME", "Sheet1")
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

onedrive_msal_app = ConfidentialClientApplication(
    CLIENT_ID_ONEDRIVE,
    authority=AUTHORITY_ONEDRIVE,
    client_credential=CLIENT_SECRET_ONEDRIVE
)

def get_onedrive_access_token():
    """Acquires an access token for OneDrive operations."""
    result = onedrive_msal_app.acquire_token_silent(SCOPE_ONEDRIVE, account=None)
    if not result:
        result = onedrive_msal_app.acquire_token_for_client(scopes=SCOPE_ONEDRIVE)
    if "access_token" in result:
        return result['access_token']
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
        data = response.json()
        
        rows = data.get('values', [])
        if not rows or len(rows) < 2:
            return [] 

        header = rows[0]
        contacts = []
        for i, row_data in enumerate(rows[1:]):
            contact_dict = {header[j]: row_data[j] if j < len(row_data) else "" for j in range(len(header))}
            contact_dict['row_id'] = i + 2  # Excel rows are 1-based, data starts on row 2
            contacts.append(contact_dict)
        return contacts
    except Exception as e:
        print(f"Error fetching contacts from OneDrive: {e}")
        return []

def update_contact_in_onedrive_excel(row_id, updated_data_dict):
    """Updates a single row in the Contacts.xlsx file."""
    try:
        access_token = get_onedrive_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        
        header_url = (
            f"{GRAPH_API_ENDPOINT}/users/{ONEDRIVE_USER_ID}/drive/root:/"
            f"{FILE_PATH}:/workbook/worksheets('{WORKSHEET_NAME}')/range(address='A1:Z1')"
        )
        header_res = requests.get(header_url, headers=headers)
        header_res.raise_for_status()
        header = header_res.json().get("values", [[]])[0]
        if not header: raise Exception("Could not retrieve header row.")
        
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
