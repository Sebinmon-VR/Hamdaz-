import zoho
import cosmos
import time

def sync():
    start_time = time.time()
    print("🚀 Starting sync: Zoho Purchase Orders -> Cosmos DB (Item Distributors)")
    
    try:
        # 1. Fetch from Zoho
        item_map = zoho.get_item_distributors_map()
        print(f"📦 Found {len(item_map)} items with distributor history in Zoho.")
        
        # 2. Save to Cosmos DB
        success = cosmos.upsert_item_distributors(item_map)
        
        if success:
            duration = round(time.time() - start_time, 2)
            print(f"✨ Sync completed successfully in {duration} seconds.")
        else:
            print("❌ Sync failed during Cosmos DB update.")
            
    except Exception as e:
        print(f"💥 Critical Failure during sync: {e}")

if __name__ == "__main__":
    sync()
