"""
Search API debug probe — token passed as argv[1].
Writes diagnostic report to debug/search.log.

Probes:
  1. Token JWT claims  — check Files.Read, Files.ReadWrite scopes
  2. GET  /me/drive                     — baseline: does Files.Read work at all?
  3. POST /search/query (driveItem)     — actual search call
  4. POST /search/query (message)       — confirm search API is reachable (Mail.Read already granted)
  5. GET  /me/drive/root/children       — alternative files listing to confirm scope
"""

import sys
import os
import json
import base64
import requests
from datetime import datetime, timezone

LOG_FILE = os.path.join(os.path.dirname(__file__), "search.log")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def log(msg: str):
    print(msg)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def graph_get(token, path, params=None):
    url = f"{GRAPH_BASE}{path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params)
    return resp


def graph_post(token, path, body):
    url = f"{GRAPH_BASE}{path}"
    resp = requests.post(url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body)
    return resp


def decode_claims(token):
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.b64decode(payload))
    except Exception as e:
        return {"error": str(e)}


def report(label, resp):
    log(f"\n{'─'*60}")
    log(f"PROBE: {label}")
    log(f"  Status: {resp.status_code}")
    try:
        body = resp.json()
        log(f"  Body:   {json.dumps(body, indent=4)}")
    except Exception:
        log(f"  Body:   {resp.text[:3000]}")
    return resp


def main():
    if len(sys.argv) < 2:
        print("Usage: python debug/search.py <access_token>")
        sys.exit(1)

    token = sys.argv[1]

    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("")

    log("=" * 60)
    log("Search API Debug Probe")
    log(f"Time: {datetime.now(timezone.utc).isoformat()}")
    log("=" * 60)

    # 1. Token claims
    log("\n[1] JWT TOKEN CLAIMS")
    claims = decode_claims(token)
    scopes = claims.get("scp", "")
    log(f"  scp:  {scopes}")
    log(f"  upn:  {claims.get('upn', claims.get('preferred_username', '?'))}")
    log(f"  tid:  {claims.get('tid', '?')}")
    log(f"  exp:  {claims.get('exp', '?')}")

    has_files = any(s in scopes for s in ["Files.Read", "Files.ReadWrite", "Files.Read.All"])
    if not has_files:
        log("\n  *** WARNING: No Files.Read scope in token! ***")
        log("  Even if you granted it in Azure portal, the current session token")
        log("  was issued BEFORE the permission was added. You MUST logout and")
        log("  login again to get a fresh token that includes Files.Read.")
    else:
        log(f"\n  OK: Files scope present: {[s for s in scopes.split() if 'Files' in s]}")

    # 2. Baseline: /me/drive (requires Files.Read)
    log("\n[2] BASELINE — GET /me/drive (requires Files.Read)")
    r = graph_get(token, "/me/drive")
    report("/me/drive", r)
    if r.status_code == 200:
        log("  >>> Drive accessible — Files.Read IS working in practice")
    elif r.status_code == 403:
        log("  >>> 403 Forbidden — Files.Read permission NOT active in token")
    elif r.status_code == 404:
        log("  >>> 404 — Drive not found (scope may be present but drive not provisioned)")

    # 3. Search: driveItem (the failing one)
    log("\n[3] SEARCH — POST /search/query (entityTypes: driveItem)")
    search_body = {
        "requests": [{
            "entityTypes": ["driveItem"],
            "query": {"queryString": "test"},
            "from": 0,
            "size": 3,
        }]
    }
    r = graph_post(token, "/search/query", search_body)
    report("POST /search/query driveItem", r)

    # 4. Search: message (Mail.Read already granted — confirms search API itself works)
    log("\n[4] SEARCH — POST /search/query (entityTypes: message) — sanity check")
    search_body_msg = {
        "requests": [{
            "entityTypes": ["message"],
            "query": {"queryString": "test"},
            "from": 0,
            "size": 3,
        }]
    }
    r = graph_post(token, "/search/query", search_body_msg)
    report("POST /search/query message", r)
    if r.status_code == 200:
        log("  >>> Search API endpoint is reachable and works for messages")

    # 5. /me/drive/root/children — alternative listing
    log("\n[5] LISTING — GET /me/drive/root/children")
    r = graph_get(token, "/me/drive/root/children", {"$top": 5, "$select": "name,id"})
    report("/me/drive/root/children", r)

    log("\n" + "=" * 60)
    log("SUMMARY")
    log("  If probe [2] returns 404 and token has no Files.Read scope:")
    log("  → Go to /logout, then /login to get a fresh token.")
    log("  If probe [2] returns 200 but probe [3] returns 403:")
    log("  → The search API needs Files.Read.All or admin consent.")
    log("  If probe [4] returns 200 but probe [3] does not:")
    log("  → Search API works but driveItem specifically needs higher permission.")
    log(f"\nLog saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
