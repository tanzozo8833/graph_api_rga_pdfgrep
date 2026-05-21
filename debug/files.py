"""
OneDrive debug probe — run via Flask /debug-files route (token passed as argv[1]).
Writes full diagnostic report to debug/files.log.

Probes (in order):
  1. Token JWT claims  — verify Files.Read scope is present
  2. GET /me           — confirm identity
  3. GET /me/drive     — confirm drive exists, get driveType + ID
  4. GET /me/drive/root                  — resolve root item
  5. GET /me/drive/root/children         — primary listing endpoint
  6. GET /me/drive/items/root/children   — alternate path
  7. GET /drives/{driveId}/root/children — explicit drive-ID path
  8. GET /me/drive/root/children (no $orderby, no $select) — stripped params
"""

import sys
import os
import json
import base64
import requests
from datetime import datetime, timezone

LOG_FILE = os.path.join(os.path.dirname(__file__), "files.log")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def log(msg: str):
    print(msg)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def graph(token: str, path: str, params: dict = None):
    url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params=params)
    return resp


def decode_token_claims(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.b64decode(payload))
    except Exception as e:
        return {"error": str(e)}


def probe(token: str, label: str, path: str, params: dict = None):
    log(f"\n{'─'*60}")
    log(f"PROBE: {label}")
    log(f"  URL:    {GRAPH_BASE}{path}")
    if params:
        log(f"  Params: {params}")
    resp = graph(token, path, params)
    log(f"  Status: {resp.status_code}")
    try:
        body = resp.json()
        log(f"  Body:   {json.dumps(body, indent=4)}")
    except Exception:
        log(f"  Body:   {resp.text[:2000]}")
    return resp


def main():
    if len(sys.argv) < 2:
        print("Usage: python debug/files.py <access_token>")
        sys.exit(1)

    token = sys.argv[1]

    # Clear previous log
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("")

    log("=" * 60)
    log(f"OneDrive Debug Probe")
    log(f"Time: {datetime.now(timezone.utc).isoformat()}")
    log("=" * 60)

    # 1. Token claims
    log("\n[1] JWT TOKEN CLAIMS")
    claims = decode_token_claims(token)
    log(f"  aud:  {claims.get('aud', '?')}")
    log(f"  scp:  {claims.get('scp', '— (no scp claim, may be app token)')}")
    log(f"  upn:  {claims.get('upn', claims.get('preferred_username', '?'))}")
    log(f"  tid:  {claims.get('tid', '?')}")
    log(f"  exp:  {claims.get('exp', '?')}")

    scopes = claims.get("scp", "")
    if "Files.Read" not in scopes and "Files.ReadWrite" not in scopes:
        log("\n  *** WARNING: Files.Read not found in token scopes! ***")
        log("  The token may lack OneDrive permission. Re-login may be needed.")

    # 2. /me identity
    probe(token, "/me — identity check", "/me",
          {"$select": "displayName,mail,userPrincipalName"})

    # 3. /me/drive — does drive exist?
    drive_resp = probe(token, "/me/drive — check drive exists", "/me/drive")
    drive_id = None
    if drive_resp.status_code == 200:
        drive_id = drive_resp.json().get("id")
        drive_type = drive_resp.json().get("driveType", "?")
        log(f"\n  >>> Drive ID: {drive_id}  Type: {drive_type}")

    # 4. /me/drive/root
    probe(token, "/me/drive/root — resolve root item", "/me/drive/root")

    # 5. Primary listing endpoint (full params)
    full_params = {
        "$top": 10,
        "$select": "name,size,lastModifiedDateTime,folder,file,webUrl",
        "$orderby": "lastModifiedDateTime desc",
    }
    probe(token, "/me/drive/root/children — full params", "/me/drive/root/children", full_params)

    # 6. Alternate path
    probe(token, "/me/drive/items/root/children — alternate path",
          "/me/drive/items/root/children", full_params)

    # 7. Explicit drive ID path (most reliable)
    if drive_id:
        probe(token, f"/drives/{{driveId}}/root/children — explicit ID",
              f"/drives/{drive_id}/root/children", full_params)
    else:
        log("\nSKIPPED probe 7 — could not get drive ID from probe 3")

    # 8. Stripped params (no $orderby, no $select) — isolate param issue
    probe(token, "/me/drive/root/children — stripped params (no orderby/select)",
          "/me/drive/root/children", {"$top": 10})

    log("\n" + "=" * 60)
    log("Debug complete. Check each probe status above.")
    log(f"Log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
