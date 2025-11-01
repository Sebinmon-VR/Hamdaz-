import requests
import base64
import json
# ---------------------
# REQUIRED INPUTS
# ---------------------
CLIENT_ID = "4751593e-6e85-4d77-997d-827c7e18ffb3"
CLIENT_SECRET = "Ktd4Vld03AqR2nWmfeWl6j6anP0t4xd5"
APP_KEY = "YqstlnI7zibu8Nkx9jmWVeOBgnK95unO"
BASE_API_URL = "https://openapi.ariba.com/api/sourcing-event/v2/prod"
BASE_API_URL = "https://openapi.ariba.com/api/sourcing-event/v2/prod"
TOKEN_URL = "https://api.ariba.com/v2/oauth/token"

# ---------------------
# STEP 1: Get OAuth2 Token
# ---------------------
credentials = f"{CLIENT_ID}:{CLIENT_SECRET}"
encoded_credentials = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")

token_headers = {
    "Authorization": f"Basic {encoded_credentials}",
    "Content-Type": "application/x-www-form-urlencoded"
}

token_data = {"grant_type": "client_credentials"}

token_response = requests.post(TOKEN_URL, headers=token_headers, data=token_data)

if token_response.status_code != 200:
    print("❌ Failed to get token:", token_response.text)
    exit()

access_token = token_response.json().get("access_token")
print("✅ Access token received:", access_token)

# ---------------------
# STEP 2: Get Event IDs
# ---------------------
event_headers = {
    "Authorization": f"Bearer {access_token}",
    "x-ariba-applicationKey": APP_KEY,  # Must be correct
    "Accept": "application/json"
}

response = requests.get(f"{BASE_API_URL}/events/identifiers", headers=event_headers)

if response.status_code != 200:
    print("❌ Failed to fetch event list:", response.text)
    exit()

event_ids = response.json().get("eventIDs", [])
print(f"✅ Found {len(event_ids)} events:", event_ids)

# ---------------------
# STEP 3: Fetch Full Details for Each Event
# ---------------------
for event_id in event_ids:
    event_url = f"{BASE_API_URL}/events/{event_id}"
    resp = requests.get(event_url, headers=event_headers)
    if resp.status_code == 200:
        event_data = resp.json()
        print(f"\n--- Event ID: {event_id} ---")
        print(json.dumps(event_data, indent=2))
    else:
        print(f"❌ Failed to fetch details for {event_id}: {resp.text}")

# ---------------------
# Optional: Fetch Items or Participants for First Event
# ---------------------
if event_ids:
    first_event_id = event_ids[0]

    # Fetch items
    items_url = f"{BASE_API_URL}/events/{first_event_id}/items"
    items_resp = requests.get(items_url, headers=event_headers)
    if items_resp.status_code == 200:
        print(f"\n--- Items for Event {first_event_id} ---")
        print(json.dumps(items_resp.json(), indent=2))

    # Fetch participants
    participants_url = f"{BASE_API_URL}/events/{first_event_id}/participants"
    participants_resp = requests.get(participants_url, headers=event_headers)
    if participants_resp.status_code == 200:
        print(f"\n--- Participants for Event {first_event_id} ---")
        print(json.dumps(participants_resp.json(), indent=2))
