import os
from azure.cosmos import CosmosClient, PartitionKey
from dotenv import load_dotenv
import pandas as pd
import datetime
import uuid
from logger import log

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
            log.debug("item_distributors container ready.", tag="COSMOS")
        except Exception as e_dist:
            log.error(f"Failed to init item_distributors container: {e_dist}", tag="COSMOS")
            distributors_container = None

        # --- Dedicated container for Chat Sessions ---
        # We create it with partition key /id and TTL enabled
        try:
            sessions_container = database.create_container_if_not_exists(
                id="chat_sessions",
                partition_key=PartitionKey(path="/id"),
                default_ttl=-1  # Enable TTL; each doc sets its own ttl value
            )
            log.debug("chat_sessions container ready.", tag="COSMOS")
        except Exception as e_sess:
            log.error(f"Failed to init chat_sessions container: {e_sess}", tag="COSMOS")
            sessions_container = None
            
        # --- Dedicated container for Shared Projects ---
        try:
            shared_projects_container = database.create_container_if_not_exists(
                id="shared_projects",
                partition_key=PartitionKey(path="/id")
            )
            log.debug("shared_projects container ready.", tag="COSMOS")
        except Exception as e_sp:
            log.error(f"Failed to init shared_projects container: {e_sp}", tag="COSMOS")
            shared_projects_container = None

        # --- Dedicated container for In-App Notifications ---
        try:
            notifications_container = database.create_container_if_not_exists(
                id="in_app_notifications",
                partition_key=PartitionKey(path="/id"),
                default_ttl=2592000  # 30 days
            )
            log.debug("in_app_notifications container ready.", tag="COSMOS")
        except Exception as e_notif:
            log.error(f"Failed to init in_app_notifications container: {e_notif}", tag="COSMOS")
            notifications_container = None
            
        # --- Dedicated container for Procurement Knowledge & Feedback ---
        try:
            procurement_knowledge_container = database.create_container_if_not_exists(
                id="procurement_knowledge",
                partition_key=PartitionKey(path="/id")
            )
            log.debug("procurement_knowledge container ready.", tag="COSMOS")
        except Exception as e_pk:
            log.error(f"Failed to init procurement_knowledge container: {e_pk}", tag="COSMOS")
            procurement_knowledge_container = None
            
        # --- Dedicated container for Tracked Emails ---
        try:
            tracked_emails_container = database.create_container_if_not_exists(
                id="tracked_emails",
                partition_key=PartitionKey(path="/task_id")
            )
            log.debug("tracked_emails container ready.", tag="COSMOS")
        except Exception as e_te:
            log.error(f"Failed to init tracked_emails container: {e_te}", tag="COSMOS")
            tracked_emails_container = None

        # --- Dedicated container for Task Supplier Quotes ---
        try:
            task_supplier_quotes_container = database.create_container_if_not_exists(
                id="task_supplier_quotes",
                partition_key=PartitionKey(path="/task_id")
            )
            log.debug("task_supplier_quotes container ready.", tag="COSMOS")
        except Exception as e_tsq:
            log.error(f"Failed to init task_supplier_quotes container: {e_tsq}", tag="COSMOS")
            task_supplier_quotes_container = None

        # --- Dedicated container for Leave Requests ---
        try:
            leave_requests_container = database.create_container_if_not_exists(
                id="leave_requests",
                partition_key=PartitionKey(path="/user_email")
            )
            log.debug("leave_requests container ready.", tag="COSMOS")
        except Exception as e_lr:
            log.error(f"Failed to init leave_requests container: {e_lr}", tag="COSMOS")
            leave_requests_container = None

        # --- Dedicated container for Leave Settings (limits, holidays, notices) ---
        try:
            leave_settings_container = database.create_container_if_not_exists(
                id="leave_settings",
                partition_key=PartitionKey(path="/setting_type")
            )
            log.debug("leave_settings container ready.", tag="COSMOS")
        except Exception as e_ls:
            log.error(f"Failed to init leave_settings container: {e_ls}", tag="COSMOS")
            leave_settings_container = None

    else:
        log.warn("COSMOS_ENDPOINT or COSMOS_KEY missing — Cosmos DB disabled.", tag="COSMOS")
        client = None
        database = None
        container = None
        sessions_container = None
        tracked_emails_container = None
        task_supplier_quotes_container = None
        leave_requests_container = None
        leave_settings_container = None
except Exception as e:
    log.error("CosmosClient initialization failed", tag="COSMOS", exc=e)
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
        log.error(f"Quote {estimate_id} not found", tag="COSMOS", exc=e)
        return None


def get_all_data_full():
    """
    Fetches EVERY single field from EVERY quote.
    Warning: If you have thousands of records, this might be slow.
    """
    if container is None:
        return []
        
    log.debug("Fetching all master data from Cosmos DB...", tag="COSMOS")
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
        
    log.debug(f"Searching for items: '{search_term}'", tag="COSMOS")
    
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
        
    log.debug(f"Deep searching: '{search_term}'", tag="COSMOS")
    
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
        
    log.debug(f"Full quote search: '{search_term}'", tag="COSMOS")
    
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
        log.warn("sessions_container is None — cannot fetch sessions.", tag="COSMOS")
        return []
    
    query = {
        "query": "SELECT c.id, c.session_title, c.updated_at FROM c WHERE c.user_email = @email ORDER BY c.updated_at DESC",
        "parameters": [
            {"name": "@email", "value": user_email}
        ]
    }
    try:
        results = list(sessions_container.query_items(query=query, enable_cross_partition_query=True))
        log.debug(f"Fetched {len(results)} session(s) for {user_email}", tag="COSMOS")
        return results
    except Exception as e:
        log.error(f"Failed to fetch sessions for {user_email}", tag="COSMOS", exc=e)
        import traceback
        traceback.print_exc()
        return []

def get_session_messages(session_id):
    """
    Fetches the full chat history of a specific session.
    Returns messages stripped of 'timestamp' field so OpenAI API doesn't reject them.
    """
    if sessions_container is None:
        log.warn("sessions_container is None — cannot retrieve session.", tag="COSMOS")
        return None
        
    try:
        response = sessions_container.read_item(item=session_id, partition_key=session_id)
        messages = response.get("messages", [])
        msg_count = len(messages)
        log.debug(f"Session {session_id} loaded ({msg_count} messages)", tag="COSMOS")
        # Strip timestamp so OpenAI only gets role/content
        # Keep full message for UI but clean version for AI (done at the call site if needed)
        return messages
    except Exception as e:
        log.error(f"Failed to retrieve session {session_id}", tag="COSMOS", exc=e)
        return None

def save_session_message(session_id, user_email, role, content, title=None, agent_type="personal", task_id=None):
    """
    Appends a message to a session or creates a new one in chat_sessions container.
    Sets TTL of 5 days (432000 seconds) for automatic deletion.
    """
    if sessions_container is None:
        log.warn("sessions_container is None — cannot save session.", tag="COSMOS")
        if not session_id:
            session_id = str(uuid.uuid4())
        return session_id

    if not session_id or session_id in ("null", "undefined", ""):
        session_id = str(uuid.uuid4())
        log.debug(f"New session created: {session_id}", tag="COSMOS")
        
    try:
        try:
            session = sessions_container.read_item(item=session_id, partition_key=session_id)
            log.debug(f"Existing session found: {session_id}", tag="COSMOS")
        except Exception:
            log.debug(f"Creating new session document: {session_id}", tag="COSMOS")
            session = {
                "id": session_id,
                "user_email": user_email,
                "session_title": title or "New Chat",
                "agent_type": agent_type,
                "task_id": task_id,
                "messages": [],
                "created_at": datetime.datetime.utcnow().isoformat(),
                "ttl": 432000  # 5 days in seconds
            }

        # Ensure agent_type and task_id are updated if switched
        session["agent_type"] = agent_type
        if task_id:
            session["task_id"] = task_id

        session["messages"].append({
            "role": role,
            "content": content,
            "timestamp": datetime.datetime.utcnow().isoformat()
        })
        session["updated_at"] = datetime.datetime.utcnow().isoformat()
        
        if title and session.get("session_title") == "New Chat":
            session["session_title"] = title

        sessions_container.upsert_item(body=session)
        log.debug(f"Saved [{role}] message to session {session_id}", tag="COSMOS")
        return session_id
    except Exception as e:
        log.error("Failed to save session message", tag="COSMOS", exc=e)
        import traceback
        traceback.print_exc()
        return session_id

def delete_session(session_id):
    """
    Deletes a session document from chat_sessions container.
    """
    if sessions_container is None:
        log.warn("sessions_container is None — cannot delete session.", tag="COSMOS")
        return False
    try:
        sessions_container.delete_item(item=session_id, partition_key=session_id)
        log.debug(f"Session deleted: {session_id}", tag="COSMOS")
        return True
    except Exception as e:
        log.error(f"Failed to delete session {session_id}", tag="COSMOS", exc=e)
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
        log.warn("distributors_container is None — cannot save mapping.", tag="COSMOS")
        return False
    
    log.debug(f"Syncing {len(mapping)} item-distributor record(s) to Cosmos DB", tag="COSMOS")
    try:
        for item_id, details in mapping.items():
            doc = {
                "id": str(item_id),
                "item_name": details.get("name"),
                "purchase_history": details.get("history", []),
                "updated_at": datetime.datetime.utcnow().isoformat()
            }
            distributors_container.upsert_item(body=doc)
        log.debug("Distributor sync complete.", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to upsert distributors", tag="COSMOS", exc=e)
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
    
    log.debug(f"Searching distributors for: '{search_term}'", tag="COSMOS")
    
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
        log.error("Failed to search distributors", tag="COSMOS", exc=e)
        return []

# =======================
# PROCUREMENT KNOWLEDGE & FEEDBACK
# =======================

def save_procurement_feedback(user_email, original_items, distributors, is_true_data, notes):
    """
    Saves a curated record of an enquiry, the found distributors, and the user's feedback.
    """
    if procurement_knowledge_container is None:
        log.warn("procurement_knowledge_container is None.", tag="COSMOS")
        return False
        
    doc_id = str(uuid.uuid4())
    doc = {
        "id": doc_id,
        "user_email": user_email,
        "items_enquired": original_items,
        "distributors": distributors,
        "is_true_data": is_true_data,
        "notes": notes,
        "created_at": datetime.datetime.utcnow().isoformat()
    }
    
    try:
        procurement_knowledge_container.upsert_item(body=doc)
        log.debug(f"Procurement feedback saved: {doc_id}", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to save procurement feedback", tag="COSMOS", exc=e)
        return False

def search_procurement_knowledge(query):
    """
    Searches procurement knowledge feedback for the given query. 
    Only returns records where is_true_data is True.
    """
    if procurement_knowledge_container is None:
        return []
        
    log.debug(f"Searching procurement knowledge: '{query}'", tag="COSMOS")
    
    # We look for the query in notes, distributors' names/items, or enquired items
    sql_query = {
        "query": """
            SELECT c.id, c.items_enquired, c.distributors, c.notes, c.user_email, c.created_at
            FROM c
            WHERE c.is_true_data = true
              AND (
                  CONTAINS(UPPER(c.notes), UPPER(@search))
                  OR EXISTS(SELECT VALUE i FROM i IN c.items_enquired WHERE CONTAINS(UPPER(i.name), UPPER(@search)) OR CONTAINS(UPPER(i.type), UPPER(@search)))
                  OR EXISTS(SELECT VALUE d FROM d IN c.distributors WHERE CONTAINS(UPPER(d.name), UPPER(@search)) OR CONTAINS(UPPER(d.item), UPPER(@search)))
              )
            ORDER BY c.created_at DESC
        """,
        "parameters": [
            {"name": "@search", "value": query}
        ]
    }
    
    try:
        results = list(procurement_knowledge_container.query_items(query=sql_query, enable_cross_partition_query=True))
        return results
    except Exception as e:
        log.error("Failed to search procurement knowledge", tag="COSMOS", exc=e)
        return []

# =======================
# COLLABORATION & SHARED PROJECTS
# =======================

def create_shared_project(task_data, creator_email):
    """
    Converts a SharePoint task into a shared project document.
    """
    if not shared_projects_container: return None
    
    project_id = str(uuid.uuid4())
    doc = {
        "id": project_id,
        "type": "shared_project",
        "task_details": task_data,
        "creator": creator_email.lower(),
        "collaborators": [creator_email.lower()],
        "invited": [],
        "messages": [],
        "presence": {},
        "created_at": datetime.datetime.utcnow().isoformat(),
        "updated_at": datetime.datetime.utcnow().isoformat()
    }
    try:
        shared_projects_container.create_item(body=doc)
        return project_id
    except Exception as e:
        log.error("Create shared project failed", tag="COSMOS", exc=e)
        return None

def invite_collaborator(project_id, inviter_email, invitee_email):
    """
    Adds a user to the 'invited' list and saves an in-app notification.
    """
    if not shared_projects_container: return False
    try:
        project = shared_projects_container.read_item(item=project_id, partition_key=project_id)
        if invitee_email in project.get("collaborators", []) or invitee_email in project.get("invited", []):
            return True # Already there
            
        project.setdefault("invited", []).append(invitee_email.lower())
        shared_projects_container.upsert_item(body=project)
        
        # Save notification
        save_user_notification(
            invitee_email.lower(), 
            f"{inviter_email} invited you to collaborate on: {project['task_details'].get('Title', 'Untitled Task')}",
            "collaboration_invite",
            project_id,
            {"inviter": inviter_email, "task_title": project['task_details'].get('Title')}
        )
        return True
    except Exception as e:
        log.error("Invite collaborator failed", tag="COSMOS", exc=e)
        return False

def accept_collaboration_invite(notification_id, user_email):
    """
    Accepts an invite, moves user to collaborators, and marks notification read.
    """
    if not shared_projects_container or not notifications_container: return False
    try:
        # 1. Get notification to find project_id
        notif = notifications_container.read_item(item=notification_id, partition_key=notification_id)
        project_id = notif.get("project_id")
        
        # 2. Update project
        project = shared_projects_container.read_item(item=project_id, partition_key=project_id)
        u_email = user_email.lower()
        if u_email in project.get("invited", []):
            project["invited"].remove(u_email)
            if u_email not in project.get("collaborators", []):
                project["collaborators"].append(u_email)
            shared_projects_container.upsert_item(body=project)
            
        # 3. Mark notification read
        mark_notification_read(notification_id)
        return True
    except Exception as e:
        log.error("Accept invite failed", tag="COSMOS", exc=e)
        return False

def get_shared_projects_for_user(user_email):
    """Fetches all projects where user is creator or collaborator."""
    if not shared_projects_container: return []
    query = {
        "query": "SELECT * FROM c WHERE ARRAY_CONTAINS(c.collaborators, @email) OR ARRAY_CONTAINS(c.collaborators, @email_orig)",
        "parameters": [
            {"name": "@email", "value": user_email.lower()},
            {"name": "@email_orig", "value": user_email}
        ]
    }
    try:
        return list(shared_projects_container.query_items(query=query, enable_cross_partition_query=True))
    except Exception as e:
        log.error("Get shared projects failed", tag="COSMOS", exc=e)
        return []

def get_shared_project_details(project_id):
    if not shared_projects_container: return None
    try:
        return shared_projects_container.read_item(item=project_id, partition_key=project_id)
    except Exception:
        return None

def save_shared_session_message(project_id, role, content, user_email):
    """Saves a message to the shared project chat history."""
    if not shared_projects_container: return False
    try:
        project = shared_projects_container.read_item(item=project_id, partition_key=project_id)
        project.setdefault("messages", []).append({
            "role": role,
            "content": content,
            "user": user_email,
            "timestamp": datetime.datetime.utcnow().isoformat()
        })
        project["updated_at"] = datetime.datetime.utcnow().isoformat()
        shared_projects_container.upsert_item(body=project)
        return True
    except Exception as e:
        log.error("Save shared message failed", tag="COSMOS", exc=e)
        return False

def get_shared_project_activity(project_id):
    """Returns the most recent AI messages and user queries in the project."""
    project = get_shared_project_details(project_id)
    if not project: return []
    msgs = project.get("messages", [])
    # Return last 10 messages for context
    return msgs[-10:]

# =======================
# IN-APP NOTIFICATIONS
# =======================

def save_user_notification(user_email, message, type="info", project_id=None, metadata=None):
    if not notifications_container: return False
    doc = {
        "id": str(uuid.uuid4()),
        "user_email": user_email,
        "message": message,
        "type": type,
        "project_id": project_id,
        "metadata": metadata or {},
        "read": False,
        "created_at": datetime.datetime.utcnow().isoformat()
    }
    try:
        notifications_container.create_item(body=doc)
        return True
    except Exception as e:
        log.error("Save notification failed", tag="COSMOS", exc=e)
        return False

def get_user_notifications(user_email, unread_only=True):
    if not notifications_container: return []
    q_str = "SELECT * FROM c WHERE c.user_email = @email"
    if unread_only: q_str += " AND c.read = false"
    q_str += " ORDER BY c.created_at DESC"
    
    query = {
        "query": q_str,
        "parameters": [{"name": "@email", "value": user_email}]
    }
    try:
        return list(notifications_container.query_items(query=query, enable_cross_partition_query=True))
    except Exception as e:
        log.error("Get notifications failed", tag="COSMOS", exc=e)
        return []

def mark_notification_read(notification_id):
    if not notifications_container: return False
    try:
        notif = notifications_container.read_item(item=notification_id, partition_key=notification_id)
        notif["read"] = True
        notifications_container.upsert_item(body=notif)
        return True
    except Exception as e:
        log.error("Mark notification read failed", tag="COSMOS", exc=e)
        return False

# =======================
# TRACKED EMAILS MANAGEMENT
# =======================

def save_tracked_email(task_id, session_id, to_email, subject, tracking_id, user_email, email_body=""):
    if not tracked_emails_container: return False
    doc = {
        "id": tracking_id,
        "task_id": task_id,
        "session_id": session_id,
        "user_email": user_email.lower(),
        "to_email": to_email,
        "subject": subject,
        "email_body": email_body,
        "status": "Waiting for Reply",
        "reply_content": None,
        "summary": None,
        "quote_doc_path": None,
        "created_at": datetime.datetime.utcnow().isoformat()
    }
    try:
        tracked_emails_container.upsert_item(body=doc)
        return True
    except Exception as e:
        log.error("Save tracked email failed", tag="COSMOS", exc=e)
        return False

def get_tracked_emails_for_task(task_id):
    if not tracked_emails_container: return []
    query = {
        "query": "SELECT * FROM c WHERE c.task_id = @task_id ORDER BY c.created_at DESC",
        "parameters": [
            {"name": "@task_id", "value": task_id}
        ]
    }
    try:
        return list(tracked_emails_container.query_items(query=query, enable_cross_partition_query=True))
    except Exception as e:
        log.error("Get tracked emails failed", tag="COSMOS", exc=e)
        return []

def update_tracked_email_reply(tracking_id, task_id, reply_content, summary, quote_doc_path=None, ai_parsed_data=None):
    if not tracked_emails_container: return False
    try:
        doc = tracked_emails_container.read_item(item=tracking_id, partition_key=task_id)
        doc["status"] = "Reply Received"
        doc["reply_content"] = reply_content
        doc["summary"] = summary
        if quote_doc_path:
            doc["quote_doc_path"] = quote_doc_path
        if ai_parsed_data:
            doc["ai_parsed_data"] = ai_parsed_data
        doc["updated_at"] = datetime.datetime.utcnow().isoformat()
        tracked_emails_container.upsert_item(body=doc)
        return True
    except Exception as e:
        log.error("Update email reply failed", tag="COSMOS", exc=e)
        return False

def save_task_supplier_quote(task_id, tracking_id, supplier_email, summary, parsed_json):
    if not task_supplier_quotes_container: return False
    try:
        quote_id = str(uuid.uuid4())
        doc = {
            "id": quote_id,
            "task_id": task_id,
            "tracking_id": tracking_id,
            "supplier_email": supplier_email,
            "summary": summary,
            "items": parsed_json.get("items", []),
            "parsed_json": parsed_json,
            "collection_status": "pending", # States: pending, collected, discarded
            "created_at": datetime.datetime.utcnow().isoformat()
        }
        task_supplier_quotes_container.upsert_item(body=doc)
        return quote_id
    except Exception as e:
        log.error("Save supplier quote failed", tag="COSMOS", exc=e)
        return False

def update_task_supplier_quote_status(quote_id, task_id, status):
    if not task_supplier_quotes_container: return False
    try:
        doc = task_supplier_quotes_container.read_item(item=quote_id, partition_key=task_id)
        doc["collection_status"] = status
        task_supplier_quotes_container.upsert_item(body=doc)
        return True
    except Exception as e:
        log.error("Update quote status failed", tag="COSMOS", exc=e)
        return False

def get_task_supplier_quotes(task_id, status_filter=None):
    if not task_supplier_quotes_container: return []
    try:
        if status_filter:
            query = "SELECT * FROM c WHERE c.task_id = @taskId AND c.collection_status = @status ORDER BY c.created_at DESC"
            parameters = [{"name": "@taskId", "value": task_id}, {"name": "@status", "value": status_filter}]
        else:
            query = "SELECT * FROM c WHERE c.task_id = @taskId ORDER BY c.created_at DESC"
            parameters = [{"name": "@taskId", "value": task_id}]
        return list(task_supplier_quotes_container.query_items(query=query, parameters=parameters, enable_cross_partition_query=True))
    except Exception as e:
        log.error("Get supplier quotes failed", tag="COSMOS", exc=e)
        return []

def get_pending_tracked_emails(user_email=None):
    if not tracked_emails_container: return []
    q_str = "SELECT * FROM c WHERE c.status = 'Waiting for Reply'"
    parameters = []
    if user_email:
        q_str += " AND c.user_email = @email"
        parameters.append({"name": "@email", "value": user_email.lower()})
    
    query = {
        "query": q_str,
        "parameters": parameters
    }
    try:
        return list(tracked_emails_container.query_items(query=query, enable_cross_partition_query=True))
    except Exception as e:
        log.error("Get pending emails failed", tag="COSMOS", exc=e)
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
        

def update_project_heartbeat(project_id, user_email):
    """Updates the last active timestamp for a user in a project."""
    if not shared_projects_container: return False
    try:
        project = shared_projects_container.read_item(item=project_id, partition_key=project_id)
        presence = project.setdefault("presence", {})
        presence[user_email] = datetime.datetime.utcnow().isoformat()
        project["updated_at"] = datetime.datetime.utcnow().isoformat()
        shared_projects_container.upsert_item(body=project)
        return True
    except Exception as e:
        log.error("Update heartbeat failed", tag="COSMOS", exc=e)
        return False

def get_project_presence(project_id):
    """Returns a list of users who have been active in the last 60 seconds."""
    if not shared_projects_container: return []
    try:
        project = shared_projects_container.read_item(item=project_id, partition_key=project_id)
        presence = project.get("presence", {})
        now = datetime.datetime.utcnow()
        active_users = []
        for email, ts_str in presence.items():
            ts = datetime.datetime.fromisoformat(ts_str)
            if (now - ts).total_seconds() < 60:
                active_users.append(email)
        return active_users
    except Exception:
        return []


# =======================
# LEAVE REQUESTS MANAGEMENT
# =======================

def save_leave_request(user_email, username, leave_start, leave_end,
                       continue_assign=False, handoff_enabled=False,
                       handoff_mode="auto", handoff_to_user=None,
                       proposals_transferred=None, leave_type="full_day",
                       status="active", leave_category="Casual", leave_reason="",
                       reviewed_by=None):
    """
    Saves a new leave request to the leave_requests Cosmos DB container.
    Returns the generated doc_id on success, None on failure.
    
    Schema:
        user_email      - partition key
        username        - display name (e.g. JohnDoe)
        leave_start     - ISO date string "YYYY-MM-DD"
        leave_end       - ISO date string "YYYY-MM-DD"
        continue_assign - bool: if True user stays in task pool
        handoff_enabled - bool: if True proposals were reassigned
        handoff_mode    - "auto" | "manual"
        handoff_to_user - username of recipient (if handoff_enabled)
        proposals_transferred - list of proposal titles that were moved
        leave_type      - "full_day" | "first_half" | "second_half"
        status          - "active" | "completed" | "cancelled"
    """
    if not leave_requests_container:
        log.warn("leave_requests_container is None.", tag="COSMOS")
        return None

    doc_id = str(uuid.uuid4())
    doc = {
        "id": doc_id,
        "user_email": user_email.lower(),
        "username": username,
        "leave_start": leave_start,
        "leave_end": leave_end,
        "continue_assign": continue_assign,
        "handoff_enabled": handoff_enabled,
        "handoff_mode": handoff_mode,
        "handoff_to_user": handoff_to_user,
        "proposals_transferred": proposals_transferred or [],
        "leave_type": leave_type,
        "leave_category": leave_category,
        "leave_reason": leave_reason,
        "status": status,
        "reviewed_by": reviewed_by,
        "reviewed_at": datetime.datetime.utcnow().isoformat() if reviewed_by else None,
        "admin_remarks": None,
        "submitted_at": datetime.datetime.utcnow().isoformat(),
    }
    try:
        leave_requests_container.create_item(body=doc)
        log.info(f"Leave saved for {user_email}: {leave_start} → {leave_end}", tag="COSMOS")
        return doc_id
    except Exception as e:
        log.error("Failed to save leave request", tag="COSMOS", exc=e)
        return None


def get_active_leaves():
    """
    Returns all leave_requests where status='active'.
    Used by the background updater to auto-remove expired excludeusers.
    """
    if not leave_requests_container:
        return []
    query = {
        "query": "SELECT * FROM c WHERE c.status = 'active'",
        "parameters": []
    }
    try:
        return list(leave_requests_container.query_items(
            query=query, enable_cross_partition_query=True
        ))
    except Exception as e:
        log.error("Failed to get active leaves", tag="COSMOS", exc=e)
        return []


def get_leave_history_for_user(user_email):
    """
    Returns all leave records for a specific user, newest first.
    """
    if not leave_requests_container:
        return []
    query = {
        "query": "SELECT * FROM c WHERE LOWER(c.user_email) = @email ORDER BY c.submitted_at DESC",
        "parameters": [{"name": "@email", "value": user_email.lower()}]
    }
    try:
        return list(leave_requests_container.query_items(
            query=query, enable_cross_partition_query=True
        ))
    except Exception as e:
        log.error("Failed to get leave history", tag="COSMOS", exc=e)
        return []


def update_leave_status(doc_id, user_email, new_status):
    """
    Updates the status of a leave request document.
    new_status: "active" | "completed" | "cancelled"
    """
    if not leave_requests_container:
        return False
    try:
        doc = leave_requests_container.read_item(item=doc_id, partition_key=user_email.lower())
        doc["status"] = new_status
        doc["updated_at"] = datetime.datetime.utcnow().isoformat()
        leave_requests_container.upsert_item(body=doc)
        log.info(f"Leave {doc_id} status updated → {new_status}", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to update leave status", tag="COSMOS", exc=e)
        return False


def cancel_leave_request(doc_id, user_email):
    """Convenience wrapper: marks a leave as cancelled."""
    return update_leave_status(doc_id, user_email, "cancelled")


def get_all_leaves():
    """
    Returns ALL leave records across all users. Admin-only function.
    Sorted by submitted_at descending (newest first).
    """
    if not leave_requests_container:
        return []
    query = {
        "query": "SELECT * FROM c ORDER BY c.submitted_at DESC",
        "parameters": []
    }
    try:
        return list(leave_requests_container.query_items(
            query=query, enable_cross_partition_query=True
        ))
    except Exception as e:
        log.error("Failed to get all leaves", tag="COSMOS", exc=e)
        return []


def get_max_concurrent_leave_count(start_date_str, end_date_str):
    """
    Calculates the maximum number of people on leave simultaneously at any point
    within the requested date range [start_date_str, end_date_str].
    Returns the peak count.
    """
    if not leave_requests_container:
        return 0
    
    try:
        # 1. Fetch all active/approved leaves that overlap with the ANY point in the range
        query = {
            "query": "SELECT c.leave_start, c.leave_end FROM c WHERE c.status IN ('active', 'approved') AND c.leave_start <= @end AND c.leave_end >= @start",
            "parameters": [
                {"name": "@start", "value": start_date_str},
                {"name": "@end", "value": end_date_str}
            ]
        }
        overlapping_leaves = list(leave_requests_container.query_items(
            query=query, enable_cross_partition_query=True
        ))
        
        if not overlapping_leaves:
            return 0
        
        # 2. Count daily peak
        import datetime
        start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
        
        peak_count = 0
        current_date = start_date
        while current_date <= end_date:
            curr_str = current_date.strftime("%Y-%m-%d")
            
            # Count how many leaves cover THIS specific day
            day_count = 0
            for l in overlapping_leaves:
                if l['leave_start'] <= curr_str and l['leave_end'] >= curr_str:
                    day_count += 1
            
            if day_count > peak_count:
                peak_count = day_count
            
            current_date += datetime.timedelta(days=1)
            
        return peak_count
    except Exception as e:
        log.error("Failed to get max concurrent leave count", tag="COSMOS", exc=e)
        return 0


# =======================
# LEAVE SETTINGS (Admin)
# =======================

def save_leave_setting(setting_data):
    """Upsert a leave setting document (limits config, etc.)."""
    if not leave_settings_container:
        return None
    try:
        leave_settings_container.upsert_item(body=setting_data)
        log.debug(f"Leave setting saved: {setting_data.get('id')}", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to save leave setting", tag="COSMOS", exc=e)
        return None


def get_leave_settings():
    """Returns the leave limit config doc (id='leave_limit')."""
    if not leave_settings_container:
        return None
    try:
        # The partition key for config is 'config' as per the setting_type path
        return leave_settings_container.read_item(item="leave_limit", partition_key="config")
    except Exception:
        # Return defaults if not found
        return {
            "id": "leave_limit",
            "setting_type": "config",
            "max_concurrent_limit": 3,
            "hr_email": "sebin@hamdaz.com",
            "auto_approval_enabled": True,
            "waitlist_enabled": False
        }


def save_holiday(title, date_str, end_date_str=None, holiday_type="holiday",
                 description="", created_by=""):
    """Create a holiday/event/notice entry."""
    if not leave_settings_container:
        return None
    doc_id = str(uuid.uuid4())
    doc = {
        "id": doc_id,
        "setting_type": "holiday",
        "title": title,
        "date": date_str,
        "end_date": end_date_str or date_str,
        "holiday_type": holiday_type,  # holiday | event | notice
        "description": description,
        "created_by": created_by,
        "created_at": datetime.datetime.utcnow().isoformat(),
    }
    try:
        leave_settings_container.create_item(body=doc)
        log.debug(f"Holiday saved: {title} ({date_str})", tag="COSMOS")
        return doc_id
    except Exception as e:
        log.error("Failed to save holiday", tag="COSMOS", exc=e)
        return None


def get_holidays():
    """Returns all holiday/event/notice entries."""
    if not leave_settings_container:
        return []
    try:
        return list(leave_settings_container.query_items(
            query={"query": "SELECT * FROM c WHERE c.setting_type = 'holiday' ORDER BY c.date ASC"},
            partition_key="holiday"
        ))
    except Exception as e:
        log.error("Failed to get holidays", tag="COSMOS", exc=e)
        return []


def delete_holiday(doc_id):
    """Delete a holiday entry."""
    if not leave_settings_container:
        return False
    try:
        leave_settings_container.delete_item(item=doc_id, partition_key="holiday")
        log.debug(f"Holiday deleted: {doc_id}", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to delete holiday", tag="COSMOS", exc=e)
        return False


def get_user_leave_count(user_email, year=None, month=None):
    """
    Count approved/active leaves for a user in a given year and optionally month.
    Returns dict: {"yearly": int, "monthly": int}
    """
    if not leave_requests_container:
        return {"yearly": 0, "monthly": 0}
    import datetime as dt
    now = dt.datetime.utcnow()
    yr = year or now.year
    mn = month or now.month
    yr_start = f"{yr}-01-01"
    yr_end = f"{yr}-12-31"
    mn_start = f"{yr}-{mn:02d}-01"
    if mn == 12:
        mn_end = f"{yr + 1}-01-01"
    else:
        mn_end = f"{yr}-{mn + 1:02d}-01"
    email_lower = user_email.lower()

    try:
        # Yearly count
        yearly_q = {
            "query": "SELECT VALUE COUNT(1) FROM c WHERE c.user_email = @email AND c.status IN ('active', 'approved', 'completed') AND c.leave_start >= @ys AND c.leave_start <= @ye",
            "parameters": [{"name": "@email", "value": email_lower},
                           {"name": "@ys", "value": yr_start},
                           {"name": "@ye", "value": yr_end}]
        }
        yearly = list(leave_requests_container.query_items(query=yearly_q, partition_key=email_lower))
        yearly_count = yearly[0] if yearly else 0

        # Monthly count
        monthly_q = {
            "query": "SELECT VALUE COUNT(1) FROM c WHERE c.user_email = @email AND c.status IN ('active', 'approved', 'completed') AND c.leave_start >= @ms AND c.leave_start < @me",
            "parameters": [{"name": "@email", "value": email_lower},
                           {"name": "@ms", "value": mn_start},
                           {"name": "@me", "value": mn_end}]
        }
        monthly = list(leave_requests_container.query_items(query=monthly_q, partition_key=email_lower))
        monthly_count = monthly[0] if monthly else 0

        return {"yearly": yearly_count, "monthly": monthly_count}
    except Exception as e:
        log.error("Failed to get user leave count", tag="COSMOS", exc=e)
        return {"yearly": 0, "monthly": 0}


def approve_leave_request(doc_id, user_email, admin_email, remarks=""):
    """Admin approves a leave — sets status to 'active'."""
    if not leave_requests_container:
        return False
    try:
        doc = leave_requests_container.read_item(item=doc_id, partition_key=user_email.lower())
        doc["status"] = "active"
        doc["reviewed_by"] = admin_email
        doc["reviewed_at"] = datetime.datetime.utcnow().isoformat()
        doc["admin_remarks"] = remarks
        leave_requests_container.replace_item(item=doc_id, body=doc)
        log.info(f"Leave {doc_id} approved by {admin_email}", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to approve leave request", tag="COSMOS", exc=e)
        return False


def reject_leave_request(doc_id, user_email, admin_email, remarks=""):
    """Admin rejects a leave — sets status to 'rejected'."""
    if not leave_requests_container:
        return False
    try:
        doc = leave_requests_container.read_item(item=doc_id, partition_key=user_email.lower())
        doc["status"] = "rejected"
        doc["reviewed_by"] = admin_email
        doc["reviewed_at"] = datetime.datetime.utcnow().isoformat()
        doc["admin_remarks"] = remarks
        leave_requests_container.replace_item(item=doc_id, body=doc)
        log.info(f"Leave {doc_id} rejected by {admin_email}", tag="COSMOS")
        return True
    except Exception as e:
        log.error("Failed to reject leave request", tag="COSMOS", exc=e)
        return False

