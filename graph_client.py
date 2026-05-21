"""
Microsoft Graph API Client
Auth flow: Entra ID (MSAL browser login) -> access token -> Graph API

Requirements:
  pip install msal requests

App registration (portal.azure.com):
  - Any app registration with redirect URI: http://localhost
  - API permissions: User.Read (minimum — add more as needed)
  - NO client secret needed
"""

import os
import requests
from msal import PublicClientApplication

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
TENANT_ID = os.getenv("TENANT_ID", "f01e930a-b52e-42b1-b70f-a8882b5d043b")
CLIENT_ID = os.getenv("CLIENT_ID", "your-client-id")   # <-- fill this in

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

SCOPES = [
    "User.Read",
    "Mail.Read",
    "Calendars.Read",
    "Files.Read",
    "Sites.Read.All",
    "Team.ReadBasic.All",
]

# Cache tokens to disk so re-runs don't prompt login again
TOKEN_CACHE_FILE = os.path.join(os.path.dirname(__file__), ".token_cache.json")


# ---------------------------------------------------------------------------
# Token acquisition — browser login
# ---------------------------------------------------------------------------

def _load_cache():
    from msal import SerializableTokenCache
    cache = SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_token() -> str:
    """
    Opens the system browser for Microsoft login on first run.
    Subsequent runs reuse the cached token silently.
    """
    cache = _load_cache()

    app = PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    # 1. Try silent (cached)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # 2. Open browser for interactive login
    print("Opening browser for Microsoft login...")
    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" not in result:
        error = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"Login failed: {error}")

    _save_cache(cache)
    print("Login successful.\n")
    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph API wrapper
# ---------------------------------------------------------------------------

class GraphClient:
    def __init__(self, access_token: str):
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type":  "application/json",
        }

    def get(self, endpoint: str, params: dict = None) -> dict:
        url = endpoint if endpoint.startswith("http") else f"{GRAPH_BASE}{endpoint}"
        resp = requests.get(url, headers=self._headers, params=params)
        resp.raise_for_status()
        return resp.json()

    def me(self) -> dict:
        return self.get("/me")

    def list_emails(self, top: int = 5) -> list:
        return self.get("/me/messages", params={
            "$top": top,
            "$select": "subject,from,receivedDateTime",
            "$orderby": "receivedDateTime desc",
        }).get("value", [])

    def list_calendar_events(self, top: int = 5) -> list:
        return self.get("/me/events", params={
            "$top": top,
            "$select": "subject,start,end",
            "$orderby": "start/dateTime",
        }).get("value", [])

    def list_onedrive_files(self, top: int = 10) -> list:
        return self.get("/me/drive/root/children", params={"$top": top}).get("value", [])

    def list_teams(self) -> list:
        return self.get("/me/joinedTeams").get("value", [])

    def list_sharepoint_sites(self, top: int = 5) -> list:
        return self.get("/sites", params={"search": "*", "$top": top}).get("value", [])

    def list_users(self, top: int = 10) -> list:
        return self.get("/users", params={"$top": top}).get("value", [])

    def list_groups(self, top: int = 10) -> list:
        return self.get("/groups", params={"$top": top}).get("value", [])


# ---------------------------------------------------------------------------
# Demo
# ---------------------------------------------------------------------------

def main():
    print("=== Microsoft Graph API ===\n")

    token = get_token()
    g = GraphClient(token)

    me = g.me()
    print(f"Logged in as: {me.get('displayName')} <{me.get('mail') or me.get('userPrincipalName')}>")

    print("\n--- Recent Emails ---")
    for m in g.list_emails():
        sender = m.get("from", {}).get("emailAddress", {}).get("address", "?")
        print(f"  [{m['receivedDateTime'][:10]}] {m['subject']}  (from: {sender})")

    print("\n--- Calendar Events ---")
    for e in g.list_calendar_events():
        print(f"  [{e['start']['dateTime'][:16]}] {e['subject']}")

    print("\n--- OneDrive Files ---")
    for f in g.list_onedrive_files():
        print(f"  {f['name']}")

    print("\n--- Teams ---")
    for t in g.list_teams():
        print(f"  {t['displayName']}")


if __name__ == "__main__":
    main()
