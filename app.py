"""
Flask web app — Microsoft Graph API dashboard.

Routes:
  /               — Login page (unauthenticated) or dashboard (authenticated)
  /login          — Redirect to Microsoft OAuth
  /auth/callback  — Exchange auth code for token
  /emails         — Detail: unread emails
  /files          — Detail: OneDrive files
  /teams          — Detail: Teams chat groups
  /logout         — Clear session + Microsoft SSO logout
  /debug          — JSON dump of current account + mailbox status

Azure App Registration:
  - Redirect URI: http://localhost:3000/auth/callback  (type: Web)
  - API permissions: User.Read, Mail.Read, Files.Read, Team.ReadBasic.All (delegated)
  - Client secret in .env as CLIENT_SECRET
"""

import json
import os
import traceback
from flask import Flask, Response, redirect, request, session, stream_with_context, url_for, render_template_string, jsonify
import msal
import requests
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(24))

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = "http://localhost:3000/auth/callback"
SCOPES        = [
    "User.Read",
    "Mail.Read",
    "Files.Read",
    "Files.Read.All",        # Team / SharePoint drives
    "Sites.Read.All",        # SharePoint sites
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All", # list team channels (optional)
]
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"


def _build_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )


def _get_token():
    return session.get("access_token")


def _graph(token, path, params=None):
    headers = {"Authorization": f"Bearer {token}"}
    url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
    resp = requests.get(url, headers=headers, params=params)
    return resp


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    token = _get_token()
    if not token:
        return render_template_string(LOGIN_HTML)

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return render_template_string(LOGIN_HTML)
    me = me_resp.json()

    # Fetch counts in parallel-ish (sequential is fine for a dashboard)
    unread_count     = _fetch_unread_count(token)
    files_count      = _fetch_files_count(token)
    teams_count      = _fetch_teams_count(token)
    sharepoint_count = _fetch_sharepoint_count(token)

    return render_template_string(DASHBOARD_HTML,
        me=me,
        unread_count=unread_count,
        files_count=files_count,
        teams_count=teams_count,
        sharepoint_count=sharepoint_count,
    )


def _fetch_unread_count(token):
    resp = _graph(token, "/me/messages", {"$filter": "isRead eq false", "$top": 1, "$select": "id"})
    if resp.status_code == 404:
        resp = _graph(token, "/me/mailFolders/inbox/messages", {"$filter": "isRead eq false", "$top": 1, "$select": "id"})
    if resp.status_code != 200:
        return None
    # Use @odata.count if available, else fall back to fetching all
    data = resp.json()
    if "@odata.count" in data:
        return data["@odata.count"]
    # Graph doesn't return count without $count=true
    resp2 = _graph(token, "/me/mailFolders/inbox", {"$select": "unreadItemCount"})
    if resp2.status_code == 200:
        return resp2.json().get("unreadItemCount")
    return None


def _fetch_files_count(token):
    resp = _graph(token, "/me/drive/root/children", {"$top": 100, "$select": "id"})
    if resp.status_code == 404:
        _graph(token, "/me/drive", {})
        resp = _graph(token, "/me/drive/items/root/children", {"$top": 100, "$select": "id"})
    if resp.status_code != 200:
        return None
    return len(resp.json().get("value", []))


def _fetch_teams_count(token):
    resp = _graph(token, "/me/joinedTeams", {"$select": "id"})
    if resp.status_code != 200:
        return None
    return len(resp.json().get("value", []))


def _fetch_sharepoint_count(token):
    # /sites?search=* returns sites the user can access
    resp = _graph(token, "/sites", {"search": "*", "$select": "id", "$top": 100})
    if resp.status_code != 200:
        return None
    return len(resp.json().get("value", []))


@app.route("/login")
def login():
    msal_app = _build_msal_app()
    auth_url = msal_app.get_authorization_request_url(
        SCOPES,
        redirect_uri=REDIRECT_URI,
        state=os.urandom(16).hex(),
    )
    return redirect(auth_url)


@app.route("/auth/callback")
def auth_callback():
    code = request.args.get("code")
    if not code:
        return "Login failed: no auth code received.", 400

    msal_app = _build_msal_app()
    result = msal_app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    if "access_token" not in result:
        error = result.get("error_description") or result.get("error", "Unknown error")
        return f"Token exchange failed: {error}", 400

    session["access_token"] = result["access_token"]
    return redirect(url_for("index"))


@app.route("/emails")
def emails():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    filter_mode = request.args.get("filter", "all")  # all | unread | read

    params = {
        "$orderby": "receivedDateTime desc",
        "$top": 50,
        "$select": "subject,from,receivedDateTime,bodyPreview,importance,isRead",
    }
    if filter_mode == "unread":
        params["$filter"] = "isRead eq false"
    elif filter_mode == "read":
        params["$filter"] = "isRead eq true"

    mail_resp = _graph(token, "/me/messages", params)
    if mail_resp.status_code == 404:
        mail_resp = _graph(token, "/me/mailFolders/inbox/messages", params)
        if mail_resp.status_code != 200:
            err = mail_resp.json().get("error", {})
            return render_template_string(ERROR_HTML, me=me,
                message=f"No mailbox found. Error: {err.get('code')} — {err.get('message')}")
    if mail_resp.status_code == 403:
        return render_template_string(ERROR_HTML, me=me,
            message="Permission denied. Ensure Mail.Read is granted.")
    mail_resp.raise_for_status()
    emails_data = mail_resp.json().get("value", [])

    return render_template_string(EMAILS_HTML, me=me, emails=emails_data, filter_mode=filter_mode)


@app.route("/files")
def files():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    query = request.args.get("q", "").strip()

    if query:
        # Drive search — requires only Files.Read (unlike /search/query which needs Files.Read.All)
        import urllib.parse
        escaped = query.replace("'", "''")
        resp = _graph(token, f"/me/drive/root/search(q='{urllib.parse.quote(escaped)}')", {
            "$top": 50,
            "$select": "name,size,lastModifiedDateTime,folder,file,webUrl",
            "$orderby": "lastModifiedDateTime desc",
        })
        if resp.status_code == 403:
            return render_template_string(ERROR_HTML, me=me,
                message="Permission denied for search. Ensure Files.Read is granted.")
        if resp.status_code != 200:
            err = resp.json().get("error", {})
            return render_template_string(ERROR_HTML, me=me,
                message=f"Search failed. {err.get('code')} — {err.get('message')}")
        files_data = resp.json().get("value", [])
    else:
        params = {
            "$top": 100,
            "$select": "name,size,lastModifiedDateTime,folder,file,webUrl",
            "$orderby": "lastModifiedDateTime desc",
        }
        resp = _graph(token, "/me/drive/root/children", params)
        if resp.status_code == 404:
            _graph(token, "/me/drive", {})
            resp = _graph(token, "/me/drive/items/root/children", params)
        if resp.status_code == 403:
            return render_template_string(ERROR_HTML, me=me,
                message="Permission denied. Ensure Files.Read is granted.")
        if resp.status_code != 200:
            err = resp.json().get("error", {})
            return render_template_string(ERROR_HTML, me=me,
                message=f"Could not load files. {err.get('code')} — {err.get('message')}")
        files_data = resp.json().get("value", [])

    return render_template_string(FILES_HTML, me=me, files=files_data, query=query)


@app.route("/teams")
def teams():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    resp = _graph(token, "/me/joinedTeams", {
        "$select": "id,displayName,description,visibility",
    })
    if resp.status_code == 403:
        return render_template_string(ERROR_HTML, me=me,
            message="Permission denied. Ensure Team.ReadBasic.All is granted.")
    if resp.status_code != 200:
        err = resp.json().get("error", {})
        return render_template_string(ERROR_HTML, me=me,
            message=f"Could not load Teams. {err.get('code')} — {err.get('message')}")
    teams_data = resp.json().get("value", [])

    return render_template_string(TEAMS_HTML, me=me, teams=teams_data)


@app.route("/teams/<team_id>")
def team_detail(team_id):
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    # Group/team metadata
    group_resp = _graph(token, f"/groups/{team_id}", {"$select": "id,displayName,description"})
    group = group_resp.json() if group_resp.status_code == 200 else {"displayName": "Team"}

    item_id = request.args.get("item_id")  # optional folder navigation
    select = "id,name,size,lastModifiedDateTime,folder,file,webUrl"
    if item_id:
        resp = _graph(token, f"/groups/{team_id}/drive/items/{item_id}/children",
                      {"$top": 200, "$select": select, "$orderby": "name"})
    else:
        resp = _graph(token, f"/groups/{team_id}/drive/root/children",
                      {"$top": 200, "$select": select, "$orderby": "name"})

    if resp.status_code == 403:
        return render_template_string(ERROR_HTML, me=me,
            message="Permission denied. Ensure Files.Read.All / Group.Read.All is granted.")
    if resp.status_code != 200:
        err = resp.json().get("error", {})
        return render_template_string(ERROR_HTML, me=me,
            message=f"Could not load team files. {err.get('code')} — {err.get('message')}")

    files_data = resp.json().get("value", [])
    return render_template_string(DRIVE_BROWSE_HTML,
        me=me,
        title=group.get("displayName", "Team"),
        subtitle=group.get("description") or "Shared files",
        files=files_data,
        drive_source="team",
        parent_id=team_id,
        current_folder=item_id,
        back_url="/teams",
        back_label="Teams",
        folder_url_base=f"/teams/{team_id}",
    )


@app.route("/sharepoint")
def sharepoint():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    resp = _graph(token, "/sites", {
        "search": "*",
        "$top": 100,
        "$select": "id,name,displayName,webUrl,description",
    })
    if resp.status_code == 403:
        return render_template_string(ERROR_HTML, me=me,
            message="Permission denied. Ensure Sites.Read.All is granted.")
    if resp.status_code != 200:
        err = resp.json().get("error", {})
        return render_template_string(ERROR_HTML, me=me,
            message=f"Could not load SharePoint sites. {err.get('code')} — {err.get('message')}")

    sites_data = resp.json().get("value", [])
    return render_template_string(SHAREPOINT_HTML, me=me, sites=sites_data)


@app.route("/sharepoint/<path:site_id>")
def sharepoint_site(site_id):
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    site_resp = _graph(token, f"/sites/{site_id}", {"$select": "id,displayName,description,webUrl"})
    site = site_resp.json() if site_resp.status_code == 200 else {"displayName": "Site"}

    item_id = request.args.get("item_id")
    select = "id,name,size,lastModifiedDateTime,folder,file,webUrl"
    if item_id:
        resp = _graph(token, f"/sites/{site_id}/drive/items/{item_id}/children",
                      {"$top": 200, "$select": select, "$orderby": "name"})
    else:
        resp = _graph(token, f"/sites/{site_id}/drive/root/children",
                      {"$top": 200, "$select": select, "$orderby": "name"})

    if resp.status_code == 403:
        return render_template_string(ERROR_HTML, me=me,
            message="Permission denied. Ensure Sites.Read.All / Files.Read.All is granted.")
    if resp.status_code != 200:
        err = resp.json().get("error", {})
        return render_template_string(ERROR_HTML, me=me,
            message=f"Could not load site files. {err.get('code')} — {err.get('message')}")

    files_data = resp.json().get("value", [])
    return render_template_string(DRIVE_BROWSE_HTML,
        me=me,
        title=site.get("displayName", "Site"),
        subtitle=site.get("description") or "Documents",
        files=files_data,
        drive_source="site",
        parent_id=site_id,
        current_folder=item_id,
        back_url="/sharepoint",
        back_label="SharePoint",
        folder_url_base=f"/sharepoint/{site_id}",
    )


# Text-like extensions we render inline. Office docs are linked out to webUrl.
_TEXT_EXTS = {
    ".txt", ".md", ".markdown", ".log", ".csv", ".tsv",
    ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf",
    ".py", ".js", ".ts", ".tsx", ".jsx", ".html", ".htm", ".css", ".scss",
    ".java", ".c", ".cpp", ".h", ".hpp", ".rs", ".go", ".rb", ".php", ".sh",
    ".sql", ".env", ".gitignore",
}


def _is_text_filename(name):
    if not name:
        return False
    lower = name.lower()
    for ext in _TEXT_EXTS:
        if lower.endswith(ext):
            return True
    return False


@app.route("/view-content")
def view_content():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))

    headers = {"Authorization": f"Bearer {token}"}
    me_resp = requests.get(f"{GRAPH_BASE}/me", headers=headers)
    if me_resp.status_code == 401:
        session.clear()
        return redirect(url_for("index"))
    me = me_resp.json()

    source    = request.args.get("source", "mydrive")  # mydrive | team | site
    item_id   = request.args.get("item_id")
    parent_id = request.args.get("parent_id", "")
    back_url  = request.args.get("back_url", "/")

    if not item_id:
        return render_template_string(ERROR_HTML, me=me, message="Missing item_id.")

    if source == "mydrive":
        base = f"/me/drive/items/{item_id}"
    elif source == "team":
        base = f"/groups/{parent_id}/drive/items/{item_id}"
    elif source == "site":
        base = f"/sites/{parent_id}/drive/items/{item_id}"
    else:
        return render_template_string(ERROR_HTML, me=me, message="Invalid source.")

    meta_resp = _graph(token, base, {"$select": "id,name,size,webUrl,file,lastModifiedDateTime"})
    if meta_resp.status_code != 200:
        err = meta_resp.json().get("error", {})
        return render_template_string(ERROR_HTML, me=me,
            message=f"Cannot load file metadata. {err.get('code')} — {err.get('message')}")
    meta = meta_resp.json()

    name = meta.get("name", "")
    mime = (meta.get("file") or {}).get("mimeType", "")
    is_text = _is_text_filename(name) or mime.startswith("text/") or mime in (
        "application/json", "application/xml", "application/javascript",
    )

    content_text = None
    truncated = False
    too_big = False

    if is_text:
        size = meta.get("size") or 0
        MAX = 512 * 1024  # 512 KB
        if size > MAX:
            too_big = True
        else:
            c_resp = _graph(token, base + "/content")
            if c_resp.status_code == 200:
                raw = c_resp.content
                if len(raw) > MAX:
                    raw = raw[:MAX]
                    truncated = True
                try:
                    content_text = raw.decode("utf-8")
                except UnicodeDecodeError:
                    content_text = raw.decode("latin-1", errors="replace")
            else:
                content_text = f"(Could not fetch content. HTTP {c_resp.status_code})"

    return render_template_string(VIEW_CONTENT_HTML,
        me=me,
        meta=meta,
        content=content_text,
        is_text=is_text,
        truncated=truncated,
        too_big=too_big,
        back_url=back_url,
    )


@app.route("/debug-search")
def debug_search():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))
    import subprocess, sys
    script = os.path.join(os.path.dirname(__file__), "debug", "search.py")
    subprocess.Popen([sys.executable, script, token])
    return """<html><body style="font-family:sans-serif;padding:32px">
        <p>&#9989; Search debug probe started. Check <code>debug/search.log</code> in a few seconds.</p>
        <a href="/files">&#8592; Back to Files</a>
    </body></html>"""


@app.route("/debug-files")
def debug_files():
    token = _get_token()
    if not token:
        return redirect(url_for("index"))
    import subprocess, sys
    script = os.path.join(os.path.dirname(__file__), "debug", "files.py")
    subprocess.Popen([sys.executable, script, token])
    return """<html><body style="font-family:sans-serif;padding:32px">
        <p>&#9989; Debug probe started. Check <code>debug/files.log</code> in a few seconds.</p>
        <a href="/">&#8592; Back to Dashboard</a>
    </body></html>"""


@app.route("/debug")
def debug():
    token = _get_token()
    if not token:
        return {"error": "no token in session"}, 401
    headers = {"Authorization": f"Bearer {token}"}
    me = requests.get(f"{GRAPH_BASE}/me", headers=headers).json()
    mail_check = requests.get(f"{GRAPH_BASE}/me/mailFolders/inbox", headers=headers)
    return {
        "account": {
            "displayName": me.get("displayName"),
            "mail": me.get("mail"),
            "userPrincipalName": me.get("userPrincipalName"),
        },
        "mailbox_status": mail_check.status_code,
        "mailbox_response": mail_check.json(),
    }


@app.route("/ask", methods=["POST"])
def ask():
    token = _get_token()
    if not token:
        return jsonify({"error": "Not authenticated. Please sign in again."}), 401

    payload = request.get_json(silent=True) or {}
    query = (payload.get("query") or "").strip()
    if not query:
        return jsonify({"error": "Empty query."}), 400

    from agent.agent import run_agent_stream

    def generate():
        # 2 KB padding to flush Werkzeug's internal buffer immediately.
        # Without this, dev server may hold small chunks until the request finishes.
        yield ":" + (" " * 2048) + "\n\n"
        try:
            for event in run_agent_stream(query, token):
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
                # SSE comment line ignored by EventSource but forces socket flush.
                yield ": ping\n\n"
        except Exception as e:
            traceback.print_exc()
            yield f"data: {json.dumps({'event': 'error', 'message': f'{type(e).__name__}: {e}'})}\n\n"
        finally:
            yield "data: {\"event\": \"done\"}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
            "Content-Encoding": "identity",
        },
        direct_passthrough=True,
    )


@app.route("/logout")
def logout():
    session.clear()
    logout_url = (
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri=http://localhost:3000"
    )
    return redirect(logout_url)


# ---------------------------------------------------------------------------
# Shared layout helpers (injected into each template)
# ---------------------------------------------------------------------------

_COMMON_CSS = """
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', sans-serif; background: #f3f4f6; min-height: 100vh; }
  header { background: #0078d4; color: white; padding: 0 24px; height: 56px; display: flex; align-items: center; justify-content: space-between; }
  header h1 { font-size: 18px; font-weight: 600; }
  header a.home { color: white; text-decoration: none; font-size: 18px; font-weight: 600; }
  .user-info { display: flex; align-items: center; gap: 12px; font-size: 14px; }
  .logout { color: rgba(255,255,255,0.85); text-decoration: none; padding: 6px 12px; border: 1px solid rgba(255,255,255,0.4); border-radius: 4px; font-size: 13px; }
  .logout:hover { background: rgba(255,255,255,0.1); }
  main { max-width: 960px; margin: 32px auto; padding: 0 16px; }
  .back { display: inline-flex; align-items: center; gap: 6px; color: #0078d4; text-decoration: none; font-size: 14px; margin-bottom: 20px; }
  .back:hover { text-decoration: underline; }
"""

_HEADER = """
  <header>
    <a href="/" class="home">&#128187; M365 Dashboard</a>
    <div class="user-info">
      <span>{{ me.get('displayName', me.get('userPrincipalName', '')) }}</span>
      <a href="/logout" class="logout">Sign out</a>
    </div>
  </header>
"""

# ---------------------------------------------------------------------------
# HTML Templates
# ---------------------------------------------------------------------------

LOGIN_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>M365 Dashboard</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', sans-serif; background: #f3f4f6; display: flex; align-items: center; justify-content: center; height: 100vh; }
    .card { background: white; border-radius: 12px; padding: 48px; text-align: center; box-shadow: 0 4px 24px rgba(0,0,0,0.08); max-width: 420px; width: 100%; }
    .logo { font-size: 52px; margin-bottom: 16px; }
    h1 { font-size: 24px; color: #111; margin-bottom: 8px; }
    p { color: #666; margin-bottom: 32px; line-height: 1.5; }
    .btn { display: inline-flex; align-items: center; gap: 10px; background: #0078d4; color: white; padding: 12px 28px; border-radius: 6px; text-decoration: none; font-size: 15px; font-weight: 600; transition: background 0.2s; }
    .btn:hover { background: #106ebe; }
    .ms-icon { width: 20px; height: 20px; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">&#128187;</div>
    <h1>M365 Dashboard</h1>
    <p>Sign in with your Microsoft account to view emails, files, and Teams.</p>
    <a href="/login" class="btn">
      <svg class="ms-icon" viewBox="0 0 23 23" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="1" y="1" width="10" height="10" fill="#f25022"/>
        <rect x="12" y="1" width="10" height="10" fill="#7fba00"/>
        <rect x="1" y="12" width="10" height="10" fill="#00a4ef"/>
        <rect x="12" y="12" width="10" height="10" fill="#ffb900"/>
      </svg>
      Sign in with Microsoft
    </a>
  </div>
</body>
</html>
"""

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard — M365</title>
  <style>
    """ + _COMMON_CSS + """
    .welcome { margin-bottom: 24px; }
    .welcome h2 { font-size: 22px; color: #111; }
    .welcome p { color: #666; margin-top: 4px; font-size: 14px; }
    .debug-bar { margin-bottom: 20px; text-align: right; }
    .debug-btn { display: inline-block; padding: 7px 16px; background: #f3f4f6; border: 1px solid #d1d5db; border-radius: 5px; font-size: 13px; color: #555; text-decoration: none; }
    .debug-btn:hover { background: #e5e7eb; }
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
    .card { background: white; border-radius: 10px; padding: 28px 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.07); text-decoration: none; color: inherit; display: flex; align-items: center; gap: 20px; transition: box-shadow 0.15s, transform 0.15s; border-left: 4px solid transparent; }
    .card:hover { box-shadow: 0 6px 20px rgba(0,0,0,0.11); transform: translateY(-2px); }
    .card.email  { border-left-color: #0078d4; }
    .card.files  { border-left-color: #107c41; }
    .card.teams  { border-left-color: #6264a7; }
    .card.sharepoint { border-left-color: #038387; }
    .card-icon { font-size: 36px; flex-shrink: 0; }
    .card-body {}
    .card-count { font-size: 36px; font-weight: 700; line-height: 1; color: #111; }
    .card-label { font-size: 14px; color: #666; margin-top: 4px; }
    .card-count.na { font-size: 18px; color: #999; }

    /* --- Chat panel --- */
    .chat-section { margin-top: 32px; background: white; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.07); overflow: hidden; }
    .chat-header { background: #faf9f8; border-bottom: 1px solid #e5e7eb; padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; }
    .chat-header h3 { font-size: 15px; font-weight: 600; color: #111; }
    .chat-header .hint { font-size: 12px; color: #888; }
    .chat-messages { height: 380px; overflow-y: auto; padding: 16px 20px; background: #fafbfc; }
    .chat-messages:empty::before { content: 'Ask anything about your OneDrive. e.g. "What is the metrics for RAG?"'; color: #999; font-size: 13px; font-style: italic; }
    .msg { margin-bottom: 14px; max-width: 88%; }
    .msg.user { margin-left: auto; text-align: right; }
    .msg.user .bubble { background: #0078d4; color: white; }
    .msg.agent .bubble { background: white; border: 1px solid #e5e7eb; }
    .msg.agent.error .bubble { background: #fdecea; border-color: #f5c2c0; color: #842029; }
    .bubble { display: inline-block; padding: 10px 14px; border-radius: 10px; font-size: 14px; line-height: 1.5; text-align: left; white-space: normal; word-wrap: break-word; max-width: 100%; }
    .msg.user .bubble { white-space: pre-wrap; }
    .agent-status { font-size: 12px; color: #0078d4; margin-bottom: 6px; }
    .agent-steps { display: flex; flex-direction: column; gap: 6px; margin: 6px 0; }
    .live-step { background: #f3f4f6; border-left: 3px solid #0078d4; padding: 8px 10px; border-radius: 4px; font-family: 'Consolas', monospace; font-size: 11px; }
    .live-step .step-head { color: #333; line-height: 1.6; }
    .live-step .step-no { background: #0078d4; color: white; padding: 1px 6px; border-radius: 3px; font-weight: 600; font-size: 10px; }
    .live-step .step-ts { color: #999; font-size: 10px; }
    .live-step .tname { color: #107c41; font-weight: 600; margin-left: 6px; }
    .live-step .tin { color: #6b6b6b; margin-left: 4px; word-break: break-all; }
    .live-step .step-obs { margin-top: 4px; }
    .live-step .step-obs .pending { color: #999; font-style: italic; }
    .live-step pre { white-space: pre-wrap; word-wrap: break-word; max-height: 220px; overflow-y: auto; background: white; padding: 6px; border-radius: 3px; margin: 0; font-size: 11px; }
    .agent-answer { font-size: 14px; line-height: 1.5; }
    .agent-answer:not(:empty) { margin-top: 8px; padding-top: 8px; border-top: 1px dashed #e5e7eb; }
    .summary-box { margin-top: 10px; padding: 10px 12px; background: #fff8e1; border: 1px solid #ffe082; border-radius: 6px; font-size: 12px; }
    .summary-box .sum-title { font-weight: 600; color: #5d4037; margin-bottom: 6px; font-size: 13px; }
    .summary-box .sum-section { margin-top: 6px; color: #333; }
    .summary-box ul { margin: 4px 0 4px 18px; padding: 0; }
    .summary-box li { margin: 2px 0; line-height: 1.4; }
    .summary-box code { background: #fff3cd; padding: 1px 5px; border-radius: 3px; font-size: 11px; color: #6a4f00; word-break: break-all; }
    .summary-box .dim { color: #888; font-size: 11px; }
    .chat-input-row { display: flex; gap: 8px; padding: 12px 20px; border-top: 1px solid #e5e7eb; background: white; }
    .chat-input-row input { flex: 1; padding: 10px 14px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; outline: none; }
    .chat-input-row input:focus { border-color: #0078d4; box-shadow: 0 0 0 2px rgba(0,120,212,0.15); }
    .chat-input-row button { padding: 10px 20px; background: #0078d4; color: white; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; }
    .chat-input-row button:hover:not(:disabled) { background: #106ebe; }
    .chat-input-row button:disabled { background: #94c2e8; cursor: not-allowed; }
    .typing { display: inline-block; }
    .typing span { display: inline-block; width: 6px; height: 6px; background: #888; border-radius: 50%; margin: 0 1px; animation: blink 1.2s infinite; }
    .typing span:nth-child(2) { animation-delay: 0.2s; }
    .typing span:nth-child(3) { animation-delay: 0.4s; }
    @keyframes blink { 0%, 60%, 100% { opacity: 0.2; } 30% { opacity: 1; } }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <div class="welcome">
      <h2>Welcome, {{ me.get('displayName', 'there') }}</h2>
      <p>{{ me.get('mail') or me.get('userPrincipalName', '') }}</p>
    </div>

    <div class="debug-bar">
      <a href="/debug-files" class="debug-btn">&#128027; Debug OneDrive Files</a>
    </div>
    <div class="cards">
      <a href="/emails" class="card email">
        <div class="card-icon">&#128140;</div>
        <div class="card-body">
          {% if unread_count is not none %}
            <div class="card-count">{{ unread_count }}</div>
          {% else %}
            <div class="card-count na">—</div>
          {% endif %}
          <div class="card-label">Unread Emails</div>
        </div>
      </a>

      <a href="/files" class="card files">
        <div class="card-icon">&#128193;</div>
        <div class="card-body">
          {% if files_count is not none %}
            <div class="card-count">{{ files_count }}</div>
          {% else %}
            <div class="card-count na">—</div>
          {% endif %}
          <div class="card-label">OneDrive Files</div>
        </div>
      </a>

      <a href="/teams" class="card teams">
        <div class="card-icon">&#128101;</div>
        <div class="card-body">
          {% if teams_count is not none %}
            <div class="card-count">{{ teams_count }}</div>
          {% else %}
            <div class="card-count na">—</div>
          {% endif %}
          <div class="card-label">Teams Groups</div>
        </div>
      </a>

      <a href="/sharepoint" class="card sharepoint">
        <div class="card-icon">&#127760;</div>
        <div class="card-body">
          {% if sharepoint_count is not none %}
            <div class="card-count">{{ sharepoint_count }}</div>
          {% else %}
            <div class="card-count na">—</div>
          {% endif %}
          <div class="card-label">SharePoint Sites</div>
        </div>
      </a>
    </div>

    <section class="chat-section">
      <div class="chat-header">
        <h3>&#129302; Ask your OneDrive</h3>
        <span class="hint">Keyword-search agent (Microsoft Graph + Claude)</span>
      </div>
      <div id="chat-messages" class="chat-messages"></div>
      <form id="chat-form" class="chat-input-row" autocomplete="off">
        <input id="chat-input" type="text" placeholder="Ask about your files…" />
        <button id="chat-send" type="submit">Send</button>
      </form>
    </section>
  </main>

  <script>
    (function () {
      const form = document.getElementById('chat-form');
      const input = document.getElementById('chat-input');
      const sendBtn = document.getElementById('chat-send');
      const messages = document.getElementById('chat-messages');

      function escapeHtml(s) {
        return String(s)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
      }

      function scroll() { messages.scrollTop = messages.scrollHeight; }

      function appendUser(text) {
        const wrap = document.createElement('div');
        wrap.className = 'msg user';
        wrap.innerHTML = '<div class="bubble">' + escapeHtml(text) + '</div>';
        messages.appendChild(wrap);
        scroll();
      }

      function createAgentContainer() {
        const wrap = document.createElement('div');
        wrap.className = 'msg agent';
        wrap.innerHTML =
          '<div class="bubble">' +
            '<div class="agent-status">' +
              '<span class="typing"><span></span><span></span><span></span></span>' +
              ' <span class="status-text">Thinking...</span>' +
            '</div>' +
            '<div class="agent-steps"></div>' +
            '<div class="agent-answer"></div>' +
          '</div>';
        messages.appendChild(wrap);
        scroll();
        return wrap;
      }

      function addStep(container, ev) {
        const stepsBox = container.querySelector('.agent-steps');
        const div = document.createElement('div');
        div.className = 'live-step';
        div.dataset.step = ev.step;
        div.innerHTML =
          '<div class="step-head">' +
            '<span class="step-no">#' + ev.step + '</span> ' +
            '<span class="step-ts">' + escapeHtml(ev.ts) + '</span> ' +
            '<span class="tname">' + escapeHtml(ev.tool) + '</span>' +
            '<span class="tin">(' + escapeHtml(ev.input) + ')</span>' +
          '</div>' +
          '<div class="step-obs"><em class="pending">...waiting for result</em></div>';
        stepsBox.appendChild(div);
        scroll();
      }

      function fillStepResult(container, ev) {
        const div = container.querySelector('.live-step[data-step="' + ev.step + '"]');
        if (!div) return;
        const obs = div.querySelector('.step-obs');
        obs.innerHTML = '<pre>' + escapeHtml(ev.observation) + '</pre>';
        scroll();
      }

      function setStatus(container, text, hideTyping) {
        const statusBox = container.querySelector('.agent-status');
        if (hideTyping) {
          statusBox.style.display = 'none';
        } else {
          statusBox.querySelector('.status-text').textContent = text;
        }
      }

      function setAnswer(container, text) {
        const box = container.querySelector('.agent-answer');
        box.innerHTML = escapeHtml(text || '(no answer)').replace(/\\n/g, '<br>');
      }

      function setError(container, msg) {
        container.classList.add('error');
        const box = container.querySelector('.agent-answer');
        box.innerHTML = '<strong>Error:</strong> ' + escapeHtml(msg);
      }

      async function consumeStream(resp, container) {
        const reader = resp.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          const parts = buffer.split('\\n\\n');
          buffer = parts.pop();
          for (const raw of parts) {
            const line = raw.trim();
            if (!line.startsWith('data:')) continue;
            const payload = line.slice(5).trim();
            if (!payload) continue;
            let ev;
            try { ev = JSON.parse(payload); } catch (_) { continue; }
            handleEvent(container, ev);
          }
        }
      }

      function renderSummary(container, ev) {
        const wrap = document.createElement('div');
        wrap.className = 'summary-box';
        let h = '<div class="sum-title">📋 Summary</div>';

        h += '<div class="sum-section"><b>🔎 Keywords searched (' + ev.queries.length + ' query/queries):</b><ul>';
        ev.queries.forEach(function (q, i) {
          h += '<li><code>' + escapeHtml(q) + '</code></li>';
        });
        h += '</ul></div>';

        h += '<div class="sum-section"><b>📁 Files returned by Graph (' + ev.files_returned.length + '):</b>';
        if (ev.files_returned.length) {
          h += '<ul>';
          ev.files_returned.forEach(function (f) {
            h += '<li>' + escapeHtml(f.name) + ' <span class="dim">(query: ' + escapeHtml(f.query.slice(0, 60)) + ')</span></li>';
          });
          h += '</ul>';
        }
        h += '</div>';

        h += '<div class="sum-section"><b>📖 Files actually read (' + ev.files_fetched.length + '):</b>';
        if (ev.files_fetched.length) {
          h += '<ul>';
          ev.files_fetched.forEach(function (f) {
            h += '<li><strong>' + escapeHtml(f.name) + '</strong></li>';
          });
          h += '</ul>';
        } else {
          h += ' <em class="dim">none (snippets were enough)</em>';
        }
        h += '</div>';

        if (ev.files_grepped && ev.files_grepped.length) {
          h += '<div class="sum-section"><b>🔬 Grep pattern:</b><ul>';
          ev.files_grepped.forEach(function (f) {
            h += '<li>' + escapeHtml(f.name) + ' ← <code>' + escapeHtml(f.patterns) + '</code></li>';
          });
          h += '</ul></div>';
        }

        wrap.innerHTML = h;
        const stepsBox = container.querySelector('.agent-steps');
        stepsBox.parentNode.insertBefore(wrap, container.querySelector('.agent-answer'));
        scroll();
      }

      function handleEvent(container, ev) {
        switch (ev.event) {
          case 'user_query':
            setStatus(container, 'Searching OneDrive…', false);
            break;
          case 'tool_call':
            setStatus(container, 'Step ' + ev.step + ': calling ' + ev.tool + '…', false);
            addStep(container, ev);
            break;
          case 'tool_result':
            fillStepResult(container, ev);
            setStatus(container, 'Got step ' + ev.step + ' result, continuing…', false);
            break;
          case 'summary':
            renderSummary(container, ev);
            break;
          case 'final_answer':
            setStatus(container, '', true);
            setAnswer(container, ev.answer);
            break;
          case 'error':
            setStatus(container, '', true);
            setError(container, ev.message);
            break;
          case 'done':
            setStatus(container, '', true);
            break;
        }
      }

      form.addEventListener('submit', async function (e) {
        e.preventDefault();
        const q = input.value.trim();
        if (!q) return;
        appendUser(q);
        input.value = '';
        input.disabled = true;
        sendBtn.disabled = true;
        const container = createAgentContainer();

        try {
          const resp = await fetch('/ask', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Accept': 'text/event-stream' },
            body: JSON.stringify({ query: q }),
          });
          if (!resp.ok) {
            const text = await resp.text();
            setStatus(container, '', true);
            setError(container, 'HTTP ' + resp.status + ': ' + text.slice(0, 200));
          } else {
            await consumeStream(resp, container);
          }
        } catch (err) {
          setStatus(container, '', true);
          setError(container, 'Network error: ' + (err.message || String(err)));
        } finally {
          input.disabled = false;
          sendBtn.disabled = false;
          input.focus();
        }
      });
    })();
  </script>
</body>
</html>
"""

EMAILS_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Unread Emails</title>
  <style>
    """ + _COMMON_CSS + """
    .page-title { font-size: 20px; font-weight: 600; color: #111; margin-bottom: 16px; }
    .filter-bar { display: flex; gap: 8px; margin-bottom: 16px; }
    .filter-btn { padding: 7px 16px; border: 1px solid #d1d5db; border-radius: 5px; font-size: 13px; color: #555; text-decoration: none; background: white; }
    .filter-btn:hover { background: #f3f4f6; }
    .filter-btn.active { background: #0078d4; color: white; border-color: #0078d4; }
    .email-list { display: flex; flex-direction: column; gap: 8px; }
    .email-card { background: white; border-radius: 8px; padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 3px solid #0078d4; transition: box-shadow 0.15s; }
    .email-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .email-card.high { border-left-color: #d13438; }
    .email-card.read { border-left-color: #d1d5db; opacity: 0.85; }
    .email-card.read .email-subject { font-weight: 400; }
    .unread-dot { display: inline-block; width: 8px; height: 8px; background: #0078d4; border-radius: 50%; margin-right: 6px; vertical-align: middle; }
    .email-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 6px; }
    .email-subject { font-weight: 600; color: #111; font-size: 15px; flex: 1; }
    .email-date { color: #999; font-size: 12px; white-space: nowrap; }
    .email-from { font-size: 13px; color: #0078d4; margin-bottom: 6px; }
    .email-preview { font-size: 13px; color: #666; line-height: 1.4; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
    .tag { display: inline-block; background: #fde7e9; color: #d13438; font-size: 11px; font-weight: 600; padding: 1px 6px; border-radius: 3px; margin-left: 8px; vertical-align: middle; }
    .empty { text-align: center; padding: 64px 24px; color: #666; }
    .empty .icon { font-size: 56px; margin-bottom: 16px; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&#8592; Dashboard</a>
    <div class="page-title">&#128140; Emails ({{ emails|length }})</div>

    <div class="filter-bar">
      <a href="/emails?filter=all"    class="filter-btn {{ 'active' if filter_mode == 'all'    else '' }}">All</a>
      <a href="/emails?filter=unread" class="filter-btn {{ 'active' if filter_mode == 'unread' else '' }}">Unread</a>
      <a href="/emails?filter=read"   class="filter-btn {{ 'active' if filter_mode == 'read'   else '' }}">Read</a>
    </div>

    {% if emails %}
    <div class="email-list">
      {% for email in emails %}
      {% set high = email.get('importance') == 'high' %}
      {% set is_read = email.get('isRead') %}
      <div class="email-card {{ 'high' if high else '' }} {{ 'read' if is_read else '' }}">
        <div class="email-header">
          <span class="email-subject">
            {% if not is_read %}<span class="unread-dot"></span>{% endif %}
            {{ email.get('subject') or '(No subject)' }}
            {% if high %}<span class="tag">HIGH</span>{% endif %}
          </span>
          <span class="email-date">{{ email.get('receivedDateTime', '')[:10] }}</span>
        </div>
        <div class="email-from">
          {{ email.get('from', {}).get('emailAddress', {}).get('name', '') }}
          &lt;{{ email.get('from', {}).get('emailAddress', {}).get('address', '') }}&gt;
        </div>
        <div class="email-preview">{{ email.get('bodyPreview', '') }}</div>
      </div>
      {% endfor %}
    </div>
    {% else %}
    <div class="empty">
      <div class="icon">&#127881;</div>
      <p>No emails match this filter.</p>
    </div>
    {% endif %}
  </main>
</body>
</html>
"""

FILES_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>OneDrive Files</title>
  <style>
    """ + _COMMON_CSS + """
    .page-title { font-size: 20px; font-weight: 600; color: #111; margin-bottom: 16px; }

    .search-bar { display: flex; gap: 8px; margin-bottom: 20px; align-items: center; }
    .debug-btn { padding: 9px 14px; background: #f3f4f6; border: 1px solid #d1d5db; border-radius: 6px; font-size: 13px; color: #555; text-decoration: none; white-space: nowrap; }
    .debug-btn:hover { background: #e5e7eb; }
    .search-bar input { flex: 1; padding: 10px 14px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; outline: none; }
    .search-bar input:focus { border-color: #0078d4; box-shadow: 0 0 0 2px rgba(0,120,212,0.15); }
    .search-bar button { padding: 10px 20px; background: #0078d4; color: white; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; }
    .search-bar button:hover { background: #106ebe; }
    .clear-search { display: inline-block; margin-bottom: 12px; font-size: 13px; color: #0078d4; text-decoration: none; }
    .clear-search:hover { text-decoration: underline; }
    .search-mode-label { font-size: 13px; color: #666; margin-bottom: 12px; }
    .search-mode-label strong { color: #111; }

    .file-list { display: flex; flex-direction: column; gap: 6px; }
    .file-row { background: white; border-radius: 8px; padding: 14px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); display: flex; align-items: flex-start; gap: 14px; text-decoration: none; color: inherit; transition: box-shadow 0.15s; }
    .file-row:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .file-icon { font-size: 24px; flex-shrink: 0; padding-top: 2px; }
    .file-body { flex: 1; min-width: 0; }
    .file-name { font-weight: 500; color: #111; font-size: 14px; }
    .file-summary { font-size: 12px; color: #666; margin-top: 4px; line-height: 1.4; }
    .file-summary em { font-style: normal; background: #fff3cd; padding: 0 2px; border-radius: 2px; }
    .file-meta { font-size: 12px; color: #999; white-space: nowrap; text-align: right; flex-shrink: 0; }
    .empty { text-align: center; padding: 64px 24px; color: #666; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&#8592; Dashboard</a>

    <form class="search-bar" method="get" action="/files">
      <input type="text" name="q" value="{{ query }}" placeholder="Search file names and content..." autocomplete="off" />
      <button type="submit">&#128269; Search</button>
      <a href="/debug-search" class="debug-btn">&#128027; Debug</a>
    </form>

    {% if query %}
    <div>
      <span class="search-mode-label">Results for <strong>&#34;{{ query }}&#34;</strong> — {{ files|length }} found</span>
      &nbsp;<a href="/files" class="clear-search">&#10005; Clear search</a>
    </div>
    {% else %}
    <div class="page-title">&#128193; OneDrive Files ({{ files|length }})</div>
    {% endif %}

    {% if files %}
    <div class="file-list">
      {% for f in files %}
      <a href="{{ f.get('webUrl', '#') }}" target="_blank" class="file-row">
        <span class="file-icon">{{ '&#128193;' if f.get('folder') else '&#128196;' }}</span>
        <div class="file-body">
          <div class="file-name">{{ f.get('name', '') }}</div>
        </div>
        <div class="file-meta">
          {% if f.get('size') %}{{ (f['size'] / 1024) | round(1) }} KB<br>{% endif %}
          {{ f.get('lastModifiedDateTime', '')[:10] }}
        </div>
      </a>
      {% endfor %}
    </div>
    {% else %}
    <div class="empty">
      {% if query %}
      <p>No files found matching &#34;{{ query }}&#34;.</p>
      {% else %}
      <p>No files found in OneDrive root.</p>
      {% endif %}
    </div>
    {% endif %}
  </main>
</body>
</html>
"""

TEAMS_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Teams Groups</title>
  <style>
    """ + _COMMON_CSS + """
    .page-title { font-size: 20px; font-weight: 600; color: #111; margin-bottom: 16px; }
    .team-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 12px; }
    .team-card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 4px solid #6264a7; text-decoration: none; color: inherit; display: block; transition: box-shadow 0.15s, transform 0.15s; }
    .team-card:hover { box-shadow: 0 6px 20px rgba(0,0,0,0.11); transform: translateY(-2px); }
    .team-name { font-weight: 600; color: #111; font-size: 15px; margin-bottom: 6px; }
    .team-desc { font-size: 13px; color: #666; line-height: 1.4; }
    .team-visibility { display: inline-block; margin-top: 10px; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 3px; background: #f0f0f0; color: #555; text-transform: capitalize; }
    .team-action { margin-top: 12px; font-size: 13px; color: #6264a7; font-weight: 500; }
    .empty { text-align: center; padding: 64px 24px; color: #666; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&#8592; Dashboard</a>
    <div class="page-title">&#128101; Teams Groups ({{ teams|length }})</div>

    {% if teams %}
    <div class="team-list">
      {% for t in teams %}
      <a href="/teams/{{ t.get('id') }}" class="team-card">
        <div class="team-name">{{ t.get('displayName', '') }}</div>
        <div class="team-desc">{{ t.get('description', '') or '—' }}</div>
        <span class="team-visibility">{{ t.get('visibility', 'unknown') }}</span>
        <div class="team-action">&#128193; View shared files &rarr;</div>
      </a>
      {% endfor %}
    </div>
    {% else %}
    <div class="empty"><p>No Teams groups found.</p></div>
    {% endif %}
  </main>
</body>
</html>
"""

DRIVE_BROWSE_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{{ title }} — Files</title>
  <style>
    """ + _COMMON_CSS + """
    .page-title { font-size: 20px; font-weight: 600; color: #111; margin-bottom: 4px; }
    .page-subtitle { color: #666; font-size: 13px; margin-bottom: 16px; }
    .crumb { font-size: 13px; color: #666; margin-bottom: 12px; }
    .crumb a { color: #0078d4; text-decoration: none; }
    .crumb a:hover { text-decoration: underline; }
    .file-list { display: flex; flex-direction: column; gap: 6px; }
    .file-row { background: white; border-radius: 8px; padding: 14px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); display: flex; align-items: center; gap: 14px; text-decoration: none; color: inherit; transition: box-shadow 0.15s; }
    .file-row:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .file-icon { font-size: 22px; flex-shrink: 0; }
    .file-body { flex: 1; min-width: 0; }
    .file-name { font-weight: 500; color: #111; font-size: 14px; }
    .file-meta { font-size: 12px; color: #999; white-space: nowrap; text-align: right; flex-shrink: 0; }
    .file-actions { display: flex; gap: 8px; align-items: center; }
    .action-btn { padding: 5px 10px; background: #f3f4f6; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px; color: #444; text-decoration: none; }
    .action-btn:hover { background: #e5e7eb; }
    .action-btn.primary { background: #0078d4; color: white; border-color: #0078d4; }
    .empty { text-align: center; padding: 64px 24px; color: #666; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="{{ back_url }}" class="back">&#8592; {{ back_label }}</a>
    <div class="page-title">&#128193; {{ title }}</div>
    <div class="page-subtitle">{{ subtitle }}</div>

    {% if current_folder %}
    <div class="crumb"><a href="{{ folder_url_base }}">Root</a> / (folder)</div>
    {% endif %}

    {% if files %}
    <div class="file-list">
      {% for f in files %}
        {% if f.get('folder') %}
        <a href="{{ folder_url_base }}?item_id={{ f.get('id') }}" class="file-row">
          <span class="file-icon">&#128193;</span>
          <div class="file-body">
            <div class="file-name">{{ f.get('name', '') }}</div>
          </div>
          <div class="file-meta">{{ f.get('folder', {}).get('childCount', '') }} items<br>{{ f.get('lastModifiedDateTime', '')[:10] }}</div>
        </a>
        {% else %}
        <div class="file-row" style="cursor:default">
          <span class="file-icon">&#128196;</span>
          <div class="file-body">
            <div class="file-name">{{ f.get('name', '') }}</div>
          </div>
          <div class="file-actions">
            <a class="action-btn primary" href="/view-content?source={{ drive_source }}&item_id={{ f.get('id') }}&parent_id={{ parent_id }}&back_url={{ folder_url_base }}{% if current_folder %}?item_id={{ current_folder }}{% endif %}">View</a>
            <a class="action-btn" href="{{ f.get('webUrl', '#') }}" target="_blank">Open</a>
          </div>
          <div class="file-meta">
            {% if f.get('size') %}{{ (f['size'] / 1024) | round(1) }} KB<br>{% endif %}
            {{ f.get('lastModifiedDateTime', '')[:10] }}
          </div>
        </div>
        {% endif %}
      {% endfor %}
    </div>
    {% else %}
    <div class="empty"><p>This folder is empty.</p></div>
    {% endif %}
  </main>
</body>
</html>
"""

SHAREPOINT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SharePoint Sites</title>
  <style>
    """ + _COMMON_CSS + """
    .page-title { font-size: 20px; font-weight: 600; color: #111; margin-bottom: 16px; }
    .site-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 12px; }
    .site-card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 4px solid #038387; text-decoration: none; color: inherit; display: block; transition: box-shadow 0.15s, transform 0.15s; }
    .site-card:hover { box-shadow: 0 6px 20px rgba(0,0,0,0.11); transform: translateY(-2px); }
    .site-name { font-weight: 600; color: #111; font-size: 15px; margin-bottom: 6px; }
    .site-desc { font-size: 13px; color: #666; line-height: 1.4; }
    .site-action { margin-top: 12px; font-size: 13px; color: #038387; font-weight: 500; }
    .empty { text-align: center; padding: 64px 24px; color: #666; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&#8592; Dashboard</a>
    <div class="page-title">&#127760; SharePoint Sites ({{ sites|length }})</div>

    {% if sites %}
    <div class="site-list">
      {% for s in sites %}
      <a href="/sharepoint/{{ s.get('id') }}" class="site-card">
        <div class="site-name">{{ s.get('displayName') or s.get('name', '(unnamed)') }}</div>
        <div class="site-desc">{{ s.get('description') or '—' }}</div>
        <div class="site-action">&#128196; Browse documents &rarr;</div>
      </a>
      {% endfor %}
    </div>
    {% else %}
    <div class="empty"><p>No SharePoint sites found.</p></div>
    {% endif %}
  </main>
</body>
</html>
"""

VIEW_CONTENT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{{ meta.get('name', 'File') }}</title>
  <style>
    """ + _COMMON_CSS + """
    .file-header { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); margin-bottom: 16px; }
    .file-title { font-size: 18px; font-weight: 600; color: #111; margin-bottom: 6px; word-break: break-all; }
    .file-meta { font-size: 13px; color: #666; }
    .file-meta a { color: #0078d4; text-decoration: none; margin-left: 8px; }
    .file-meta a:hover { text-decoration: underline; }
    .content-box { background: #1e1e1e; color: #d4d4d4; border-radius: 8px; padding: 20px; font-family: 'Consolas', 'Courier New', monospace; font-size: 13px; line-height: 1.5; white-space: pre-wrap; word-wrap: break-word; overflow-x: auto; max-height: 70vh; overflow-y: auto; }
    .notice { background: #fff3cd; border-left: 4px solid #f0ad4e; padding: 12px 16px; border-radius: 4px; margin-bottom: 16px; color: #6a4f00; font-size: 13px; }
    .info { background: white; border-radius: 8px; padding: 32px; text-align: center; color: #555; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
    .info .icon { font-size: 48px; margin-bottom: 12px; }
    .info a.btn { display: inline-block; margin-top: 12px; padding: 10px 22px; background: #0078d4; color: white; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="{{ back_url }}" class="back">&#8592; Back</a>
    <div class="file-header">
      <div class="file-title">&#128196; {{ meta.get('name', '') }}</div>
      <div class="file-meta">
        {% if meta.get('size') %}{{ (meta['size'] / 1024) | round(1) }} KB &middot;{% endif %}
        {{ meta.get('lastModifiedDateTime', '')[:10] }}
        {% if meta.get('webUrl') %}<a href="{{ meta['webUrl'] }}" target="_blank">Open in OneDrive/SharePoint &#8599;</a>{% endif %}
      </div>
    </div>

    {% if too_big %}
    <div class="notice">File is too large to preview inline ({{ (meta.get('size', 0) / 1024) | round(1) }} KB &gt; 512 KB). Use "Open" to view it.</div>
    {% elif truncated %}
    <div class="notice">Showing first 512 KB only. Use "Open" to view the full file.</div>
    {% endif %}

    {% if is_text and content is not none %}
    <div class="content-box">{{ content }}</div>
    {% elif is_text and too_big %}
    {% else %}
    <div class="info">
      <div class="icon">&#128194;</div>
      <p>Inline preview is only available for text files (txt, md, csv, json, code, &hellip;).</p>
      <p>For Office documents, PDFs, or binary files, open them in OneDrive/SharePoint.</p>
      {% if meta.get('webUrl') %}<a class="btn" href="{{ meta['webUrl'] }}" target="_blank">Open file &#8599;</a>{% endif %}
    </div>
    {% endif %}
  </main>
</body>
</html>
"""

ERROR_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Error</title>
  <style>
    """ + _COMMON_CSS + """
    .error-card { background: white; border-radius: 8px; padding: 32px; max-width: 600px; margin: 48px auto; box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #d13438; }
    .error-card h2 { color: #d13438; margin-bottom: 12px; }
    .error-card p { color: #444; line-height: 1.6; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&#8592; Dashboard</a>
    <div class="error-card">
      <h2>Something went wrong</h2>
      <p>{{ message }}</p>
    </div>
  </main>
</body>
</html>
"""

if __name__ == "__main__":
    app.run(port=3000, debug=True, threaded=True)
