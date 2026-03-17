import zoho
import json

def verify():
    print("🔍 Fetching Item-to-Distributor Mapping...")
    try:
        dist_map = zoho.get_item_distributors_map()
        
        print("\n✅ Successfully retrieved mapping.")
        print(f"Total items with PO history: {len(dist_map)}")
        
        # Display a sample of the mapping
        print("\n--- Sample Mapping (Top 5 items) ---")
        items = zoho.fetch_items()
        item_name_map = {item['item_id']: item['name'] for item in items}
        
        count = 0
        for item_id, vendors in dist_map.items():
            item_name = item_name_map.get(item_id, f"Unknown Item ({item_id})")
            print(f"Item: {item_name}")
            print(f"  ∟ Distributors: {', '.join(vendors)}")
            count += 1
            if count >= 5:
                break
                
    except Exception as e:
        print(f"❌ Verification failed: {e}")

if __name__ == "__main__":
    verify()
