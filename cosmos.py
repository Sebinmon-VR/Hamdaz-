import os
from azure.cosmos import CosmosClient, PartitionKey
from dotenv import load_dotenv
import pandas as pd
import datetime
import uuid

# =======================
# CONFIGURATION
# =======================
load_dotenv(override=True)

ENDPOINT = os.getenv("COSMOS_ENDPOINT")
KEY = os.getenv("COSMOS_KEY")
DATABASE_NAME = "Quotes"
CONTAINER_NAME = "items"

# Initialize Client
try:
    if ENDPOINT and KEY:
        client = CosmosClient(ENDPOINT, KEY)
        database = client.get_database_client(DATABASE_NAME)
        container = database.get_container_client(CONTAINER_NAME)
        
        # --- Dedicated container for Item Distributors ---
        try:
            distributors_container = database.create_container_if_not_exists(
                id="item_distributors",
                partition_key=PartitionKey(path="/id")
            )
            print("[COSMOS INFO] item_distributors container ready.", flush=True)
        except Exception as e_dist:
            print(f"[COSMOS ERROR] Could not create/get item_distributors container: {e_dist}", flush=True)
            distributors_container = None

        # --- Dedicated container for Chat Sessions ---
        # We create it with partition key /id and TTL enabled
        try:
            sessions_container = database.create_container_if_not_exists(
                id="chat_sessions",
                partition_key=PartitionKey(path="/id"),
                default_ttl=-1  # Enable TTL; each doc sets its own ttl value
            )
            print("[COSMOS INFO] chat_sessions container ready.", flush=True)
        except Exception as e_sess:
            print(f"[COSMOS ERROR] Could not create/get chat_sessions container: {e_sess}", flush=True)
            sessions_container = None
    else:
        print("Warning: COSMOS_ENDPOINT or COSMOS_KEY is missing. Cosmos DB features will be disabled.")
        client = None
        database = None
        container = None
        sessions_container = None
except Exception as e:
    print(f"Error initializing CosmosClient: {e}")
    client = None
    database = None
    container = None

# =======================
# DASHBOARD FUNCTIONS
# =======================

def get_all_quotes_for_dashboard():
    """
    Fetches the latest summary of all quotes for the main dashboard table.
    """
    if container is None:
        return pd.DataFrame()
        
    query = "SELECT c.estimate_number, c.customer_name, c.date, c.status, c.total, c.currency_code FROM c"
    
    # query_items returns a generator
    items = list(container.query_items(query=query, enable_cross_partition_query=True))
    return pd.DataFrame(items)

def get_detailed_quote_with_items(estimate_id):
    """
    Fetches EVERYTHING for one specific quote (including line items and brands)
    using the Partition Key (estimate_id).
    """
    if container is None:
        return None
        
    try:
        # read_item is the fastest way to get data if you have the ID and Partition Key
        response = container.read_item(item=estimate_id, partition_key=estimate_id)
        return response
    except Exception as e:
        print(f"❌ Error: Quote {estimate_id} not found. {e}")
        return None


def get_all_data_full():
    """
    Fetches EVERY single field from EVERY quote.
    Warning: If you have thousands of records, this might be slow.
    """
    if container is None:
        return []
        
    print("📡 Fetching Master Data from Cosmos DB...")
    query = "SELECT * FROM c"
    items = list(container.query_items(query=query, enable_cross_partition_query=True))
    return items # Returning raw list of dicts to keep all nested data

def search_quotes_by_item(search_term):
    """
    Searches inside the line_items array for a specific product name.
    Useful for finding: 'Where did we quote this Hard Drive before?'
    """
    if container is None:
        return pd.DataFrame()
        
    print(f"🔍 Searching for items containing: '{search_term}'...")
    
    # Using CONTAINS for a fuzzy search (not case-sensitive in many Cosmos setups)
    query = {
        "query": """
            SELECT 
                c.estimate_number, 
                c.customer_name, 
                c.date, 
                li.name AS item_name, 
                li.rate, 
                li.quantity
            FROM c
            JOIN li IN c.line_items
            WHERE CONTAINS(UPPER(li.name), UPPER(@search))
        """,
        "parameters": [
            {"name": "@search", "value": search_term}
        ]
    }
    
    results = list(container.query_items(query=query, enable_cross_partition_query=True))
    return pd.DataFrame(results)

def deep_search_item_with_quote_context(search_term):
    """
    Returns specific item details PLUS the full parent quote information.
    """
    if container is None:
        return pd.DataFrame()
        
    print(f"🔍 Deep searching for: '{search_term}'...")
    
    query = {
        "query": """
            SELECT 
                li.name AS item_name,
                li.description AS item_description,
                li.rate AS item_rate,
                li.quantity AS item_qty,
                cf["value"] AS item_brand,
                c.estimate_number,
                c.customer_name,
                c.date AS quote_date,
                c.status AS quote_status,
                c.total AS quote_total,
                c.currency_code,
                c.estimate_url,
                c.documents  -- This includes PDF attachment details
            FROM c
            JOIN li IN c.line_items
            JOIN cf IN li.item_custom_fields
            WHERE CONTAINS(UPPER(li.name), UPPER(@search))
               OR CONTAINS(UPPER(li.description), UPPER(@search))
               AND cf.api_name = 'cf_brand'
        """,
        "parameters": [
            {"name": "@search", "value": search_term}
        ]
    }
    
    results = list(container.query_items(query=query, enable_cross_partition_query=True))
    return pd.DataFrame(results)


def search_item_and_get_full_quotes(search_term):
    """
    Finds items matching the search term and returns the 
    ENTIRE parent quote JSON for each match.
    """
    if container is None:
        return []
        
    print(f"🔍 Searching for item: '{search_term}' and retrieving full quotes...")
    
    # We select '*' to get the full document
    # We use EXISTS to check if any line item matches your search
    query = {
        "query": """
            SELECT *
            FROM c
            WHERE EXISTS(
                SELECT VALUE li 
                FROM li IN c.line_items 
                WHERE CONTAINS(UPPER(li.name), UPPER(@search)) 
                   OR CONTAINS(UPPER(li.description), UPPER(@search))
            )
        """,
        "parameters": [
            {"name": "@search", "value": search_term}
        ]
    }
    
    # list() converts the Cosmos result into a standard Python list of dictionaries
    results = list(container.query_items(query=query, enable_cross_partition_query=True))
    
    return results






# =======================
# CHAT SESSION MANAGEMENT
# =======================

def get_user_sessions(user_email):
    """
    Fetches all chat sessions for a specific user.
    """
    if sessions_container is None:
        print("[COSMOS WARN] sessions_container is None. Cannot fetch sessions.", flush=True)
        return []
    
    query = {
        "query": "SELECT c.id, c.session_title, c.updated_at FROM c WHERE c.user_email = @email ORDER BY c.updated_at DESC",
        "parameters": [
            {"name": "@email", "value": user_email}
        ]
    }
    try:
        results = list(sessions_container.query_items(query=query, enable_cross_partition_query=True))
        print(f"[COSMOS INFO] Fetched {len(results)} sessions for {user_email}", flush=True)
        return results
    except Exception as e:
        print(f"[COSMOS ERROR] Error fetching sessions for {user_email}: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return []

def get_session_messages(session_id):
    """
    Fetches the full chat history of a specific session.
    Returns messages stripped of 'timestamp' field so OpenAI API doesn't reject them.
    """
    if sessions_container is None:
        print("[COSMOS WARN] sessions_container is None. Cannot retrieve session.", flush=True)
        return None
        
    try:
        response = sessions_container.read_item(item=session_id, partition_key=session_id)
        messages = response.get("messages", [])
        msg_count = len(messages)
        print(f"[COSMOS INFO] Loaded session {session_id} with {msg_count} messages.", flush=True)
        # Strip timestamp so OpenAI only gets role/content
        clean_messages = [{"role": m["role"], "content": m["content"]} for m in messages]
        return clean_messages
    except Exception as e:
        print(f"[COSMOS ERROR] Error retrieving session {session_id}: {e}", flush=True)
        return None

def save_session_message(session_id, user_email, role, content, title=None):
    """
    Appends a message to a session or creates a new one in chat_sessions container.
    Sets TTL of 5 days (432000 seconds) for automatic deletion.
    """
    if sessions_container is None:
        print("[COSMOS WARN] sessions_container is None. Cannot save session.", flush=True)
        if not session_id:
            session_id = str(uuid.uuid4())
        return session_id

    if not session_id or session_id in ("null", "undefined", ""):
        session_id = str(uuid.uuid4())
        print(f"[COSMOS INFO] Generated new session_id: {session_id}", flush=True)
        
    try:
        try:
            session = sessions_container.read_item(item=session_id, partition_key=session_id)
            print(f"[COSMOS INFO] Found existing session {session_id}", flush=True)
        except Exception:
            print(f"[COSMOS INFO] Creating new session document: {session_id}", flush=True)
            session = {
                "id": session_id,
                "user_email": user_email,
                "session_title": title or "New Chat",
                "messages": [],
                "created_at": datetime.datetime.utcnow().isoformat(),
                "ttl": 432000  # 5 days in seconds
            }

        session["messages"].append({
            "role": role,
            "content": content,
            "timestamp": datetime.datetime.utcnow().isoformat()
        })
        session["updated_at"] = datetime.datetime.utcnow().isoformat()
        
        if title and session.get("session_title") == "New Chat":
            session["session_title"] = title

        sessions_container.upsert_item(body=session)
        print(f"[COSMOS INFO] ✅ Saved {role} message to session {session_id} (title: {session.get('session_title')})", flush=True)
        return session_id
    except Exception as e:
        print(f"[COSMOS ERROR] Error saving session message: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return session_id

def delete_session(session_id):
    """
    Deletes a session document from chat_sessions container.
    """
    if sessions_container is None:
        print("[COSMOS WARN] sessions_container is None. Cannot delete session.", flush=True)
        return False
    try:
        sessions_container.delete_item(item=session_id, partition_key=session_id)
        print(f"[COSMOS INFO] Deleted session {session_id}", flush=True)
        return True
    except Exception as e:
        print(f"[COSMOS ERROR] Error deleting session {session_id}: {e}", flush=True)
        return False

# =======================
# ITEM DISTRIBUTOR MANAGEMENT
# =======================

def upsert_item_distributors(mapping):
    """
    Saves the enriched item mapping to Cosmos DB.
    Each document: { "id": item_id, "name": item_name, "purchase_history": [...] }
    """
    if distributors_container is None:
        print("[COSMOS WARN] distributors_container is None. Cannot save mapping.", flush=True)
        return False
    
    print(f"📡 Syncing {len(mapping)} item-history records to Cosmos DB...")
    try:
        for item_id, details in mapping.items():
            doc = {
                "id": str(item_id),
                "item_name": details.get("name"),
                "purchase_history": details.get("history", []),
                "updated_at": datetime.datetime.utcnow().isoformat()
            }
            distributors_container.upsert_item(body=doc)
        print("✅ Sync complete.", flush=True)
        return True
    except Exception as e:
        print(f"[COSMOS ERROR] Error upserting distributors: {e}", flush=True)
        return False

def get_item_distributors(item_id):
    """
    Retrieves the list of distributors for a specific item_id.
    """
    if distributors_container is None:
        return []
    
    try:
        doc = distributors_container.read_item(item=str(item_id), partition_key=str(item_id))
        return doc.get("distributors", [])
    except Exception:
        # Item not found or error
        return []

def search_item_distributors(search_term):
    """
    Searches for items by name or ID in the item_distributors container.
    Returns a list of matching items with their purchase history.
    """
    if distributors_container is None:
        return []
    
    print(f"🔍 Searching distributors for: '{search_term}'...")
    
    query = {
        "query": """
            SELECT c.id, c.item_name, c.purchase_history, c.updated_at
            FROM c
            WHERE CONTAINS(UPPER(c.item_name), UPPER(@search))
               OR CONTAINS(UPPER(c.id), UPPER(@search))
        """,
        "parameters": [
            {"name": "@search", "value": search_term}
        ]
    }
    
    try:
        results = list(distributors_container.query_items(query=query, enable_cross_partition_query=True))
        return results
    except Exception as e:
        print(f"[COSMOS ERROR] Error searching distributors: {e}", flush=True)
        return []

# =======================
# MAIN EXECUTION
#     print("📊 Loading Dashboard Data...")
    
    # # 1. Main Table
    # df_main = get_all_quotes_for_dashboard()
    # print("\n--- Recent Quotes ---")
    # print(df_main.head())
    
    
    # # 2. Detailed View for a Specific Quote
    # if not df_main.empty:
    #     sample_estimate_id = df_main.iloc[0]['estimate_number']
    #     print(f"\n📋 Fetching details for Quote: {sample_estimate_id}")
    #     detailed_quote = get_detailed_quote_with_items(sample_estimate_id)
    #     print(detailed_quote)
    # else:
    #     print("No quotes found in the database.")
         
    # # 3. Fetching ALL data (for testing)
    # all_data = get_all_data_full()
    # print(f"\n📂 Total Quotes Fetched: {len(all_data)}")
    
    # print("\n--- Sample Quote with All Data ---")
    # print(all_data[0])
    
    
    # # 4. Search for Quotes containing a specific item
    # search_results = search_quotes_by_item("Hard Drive")
    # print("\n--- Search Results ---")
    # print(search_results.head())
    
    # #5. Deep Search for item with full quote context
    # deep_search_results = deep_search_item_with_quote_context("Hard Drive")
    # print("\n--- Deep Search Results ---")
    # print(deep_search_results.head())
    
    # #6. Search for item and get full quote JSON
    # full_quote_results = search_item_and_get_full_quotes("Hard Drive")
    # print("\n--- Full Quote Results for Item Search ---")
    # print(f"Total Quotes Found: {len(full_quote_results)}")
    # if full_quote_results:
    #     print("\n--- Sample Full Quote JSON ---")
    #     print(full_quote_results[0])
        