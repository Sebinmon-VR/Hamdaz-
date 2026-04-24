import os
from dotenv import load_dotenv
import sys

sys.path.append(r"c:\Users\ansha\Hamdaz-")

load_dotenv(override=True)

from sharepoint_items import fetch_sharepoint_list

SITE_DOMAIN = "hamdaz1.sharepoint.com"
TEST_PATH = "/sites/Test"

try:
    superusers_list = fetch_sharepoint_list(SITE_DOMAIN, TEST_PATH, "superusers")
    print("Superusers list:")
    for item in superusers_list:
        print(item)
except Exception as e:
    print("Error fetching superusers:", e)

try:
    approvers_list = fetch_sharepoint_list(SITE_DOMAIN, TEST_PATH, "approvers")
    print("Approvers list:")
    for item in approvers_list:
        print(item)
except Exception as e:
    print("Error fetching approvers:", e)
