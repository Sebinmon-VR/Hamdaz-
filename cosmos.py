import os
from azure.cosmos import CosmosClient, PartitionKey
from dotenv import load_dotenv
import pandas as pd

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
    else:
        print("Warning: COSMOS_ENDPOINT or COSMOS_KEY is missing. Cosmos DB features will be disabled.")
        client = None
        database = None
        container = None
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






# EXECUTION
# =======================
# if __name__ == "__main__":
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
        