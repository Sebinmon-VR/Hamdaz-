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
client = CosmosClient(ENDPOINT, KEY)
database = client.get_database_client(DATABASE_NAME)
container = database.get_container_client(CONTAINER_NAME)

# =======================
# DASHBOARD FUNCTIONS
# =======================

def get_all_quotes_for_dashboard():
    """
    Fetches the latest summary of all quotes for the main dashboard table.
    """
    query = "SELECT c.estimate_number, c.customer_name, c.date, c.status, c.total, c.currency_code FROM c"
    
    # query_items returns a generator
    items = list(container.query_items(query=query, enable_cross_partition_query=True))
    return pd.DataFrame(items)

def get_detailed_quote_with_items(estimate_id):
    """
    Fetches EVERYTHING for one specific quote (including line items and brands)
    using the Partition Key (estimate_id).
    """
    try:
        # read_item is the fastest way to get data if you have the ID and Partition Key
        response = container.read_item(item=estimate_id, partition_key=estimate_id)
        return response
    except Exception as e:
        print(f"❌ Error: Quote {estimate_id} not found. {e}")
        return None

# def get_brand_performance():
#     """
#     Advanced Query: Digs into the line_items array across all documents
#     to count which brands are being quoted the most.
#     """
#     query = """
#     SELECT 
#         cf.value AS Brand, 
#         COUNT(1) AS QuoteCount
#     FROM c
#     JOIN li IN c.line_items
#     JOIN cf IN li.item_custom_fields
#     WHERE cf.api_name = 'cf_brand'
#     GROUP BY cf.value
#     """
#     items = list(container.query_items(query=query, enable_cross_partition_query=True))
#     return pd.DataFrame(items)

# =======================
# # EXECUTION
# # =======================
# if __name__ == "__main__":
#     print("📊 Loading Dashboard Data...")
    
#     # 1. Main Table
#     df_main = get_all_quotes_for_dashboard()
#     print("\n--- Recent Quotes ---")
#     print(df_main.head())

