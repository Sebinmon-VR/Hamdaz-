import os
from dotenv import load_dotenv
load_dotenv(override=True)
from sharepoint_items import fetch_sharepoint_list
import json

SITE_DOMAIN = os.getenv("SHAREPOINT_SITE_DOMAIN", "hamdaz1.sharepoint.com")
SITE_PATH = "/sites/ProposalTeam"
LIST_NAME = "Proposals"

try:
    tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
    if tasks:
        keys = list(tasks[0].keys())
        print("KEYS IN PROPOSALS LIST:", keys)
        print("FIRST ITEM:", json.dumps({k: tasks[0][k] for k in keys[:15]}, default=str)) # Print some values
    else:
        print("NO TASKS FOUND")
except Exception as e:
    print("ERROR:", e)
