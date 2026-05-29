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
        yield (":" + (" " * 2048) + "\n\n").encode("utf-8")
        try:
            for event in run_agent_stream(query, token):
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n".encode("utf-8")
                # SSE comment line ignored by EventSource but forces socket flush.
                yield b": ping\n\n"
        except Exception as e:
            traceback.print_exc()
            err = {'event': 'error', 'message': f'{type(e).__name__}: {e}'}
            yield f"data: {json.dumps(err)}\n\n".encode("utf-8")
        finally:
            yield b'data: {"event": "done"}\n\n'

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
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

  :root {
    --bg-0: #060912;
    --bg-1: #0c1222;
    --surface: rgba(18, 26, 48, 0.55);
    --surface-solid: #0f1729;
    --border: rgba(120, 160, 230, 0.14);
    --border-strong: rgba(120, 160, 230, 0.28);
    --text: #e6ecf7;
    --text-dim: #8a93a8;
    --text-muted: #5a6378;
    --accent: #00d4ff;
    --accent-2: #6c8cff;
    --accent-purple: #b465ff;
    --danger: #ff5572;
    --warning: #ffb648;
    --success: #4ade80;
    --glow-accent: 0 0 24px rgba(0, 212, 255, 0.35);
    --glow-soft: 0 8px 32px rgba(0, 0, 0, 0.35);
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  *::selection { background: rgba(0, 212, 255, 0.3); color: var(--text); }

  html, body { min-height: 100%; }
  body {
    font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif;
    color: var(--text);
    background:
      radial-gradient(1100px 600px at 12% -10%, rgba(108, 140, 255, 0.18), transparent 60%),
      radial-gradient(900px 500px at 95% 0%, rgba(180, 101, 255, 0.14), transparent 55%),
      radial-gradient(700px 700px at 50% 110%, rgba(0, 212, 255, 0.12), transparent 60%),
      linear-gradient(180deg, var(--bg-0), var(--bg-1));
    background-attachment: fixed;
    min-height: 100vh;
    font-feature-settings: 'cv11', 'ss01';
    letter-spacing: -0.005em;
  }

  /* Subtle animated grid overlay */
  body::before {
    content: '';
    position: fixed;
    inset: 0;
    background-image:
      linear-gradient(rgba(120, 160, 230, 0.04) 1px, transparent 1px),
      linear-gradient(90deg, rgba(120, 160, 230, 0.04) 1px, transparent 1px);
    background-size: 48px 48px;
    pointer-events: none;
    z-index: 0;
    mask-image: radial-gradient(ellipse at center, black 30%, transparent 75%);
  }

  header {
    position: sticky; top: 0; z-index: 50;
    height: 64px;
    padding: 0 28px;
    display: flex; align-items: center; justify-content: space-between;
    background: rgba(8, 12, 24, 0.7);
    backdrop-filter: blur(14px) saturate(160%);
    -webkit-backdrop-filter: blur(14px) saturate(160%);
    border-bottom: 1px solid var(--border);
  }
  header a.home {
    color: var(--text); text-decoration: none;
    font-size: 16px; font-weight: 700;
    display: inline-flex; align-items: center; gap: 10px;
    letter-spacing: 0.01em;
  }
  header a.home .logo-mark {
    width: 30px; height: 30px;
    border-radius: 8px;
    background: linear-gradient(135deg, var(--accent), var(--accent-2) 60%, var(--accent-purple));
    display: inline-flex; align-items: center; justify-content: center;
    color: #060912; font-weight: 800; font-size: 14px;
    box-shadow: var(--glow-accent);
    font-family: 'JetBrains Mono', monospace;
  }
  header a.home .brand-text { background: linear-gradient(90deg, #fff, #b8c6e0); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }
  header a.home .brand-text .accent { color: var(--accent); -webkit-text-fill-color: var(--accent); }

  .user-info { display: flex; align-items: center; gap: 14px; font-size: 13px; color: var(--text-dim); }
  .user-info .who { display: flex; align-items: center; gap: 8px; padding: 5px 12px 5px 5px; border: 1px solid var(--border); border-radius: 999px; background: rgba(255,255,255,0.02); }
  .user-info .avatar { width: 24px; height: 24px; border-radius: 50%; background: linear-gradient(135deg, var(--accent-2), var(--accent-purple)); display: inline-flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; color: #060912; }
  .logout {
    color: var(--text); text-decoration: none;
    padding: 7px 14px;
    border: 1px solid var(--border-strong);
    border-radius: 8px; font-size: 12px; font-weight: 500;
    background: transparent;
    transition: all 0.18s ease;
  }
  .logout:hover { background: rgba(255, 85, 114, 0.08); border-color: rgba(255, 85, 114, 0.4); color: var(--danger); }

  main { max-width: 1100px; margin: 36px auto 80px; padding: 0 24px; position: relative; z-index: 1; }

  .back {
    display: inline-flex; align-items: center; gap: 6px;
    color: var(--text-dim); text-decoration: none;
    font-size: 13px; font-weight: 500;
    margin-bottom: 24px;
    padding: 6px 10px 6px 6px; border-radius: 8px;
    transition: all 0.15s ease;
  }
  .back:hover { color: var(--accent); background: rgba(0, 212, 255, 0.06); transform: translateX(-2px); }

  /* Common card surface */
  .surface {
    background: var(--surface);
    backdrop-filter: blur(14px) saturate(140%);
    -webkit-backdrop-filter: blur(14px) saturate(140%);
    border: 1px solid var(--border);
    border-radius: 14px;
    box-shadow: var(--glow-soft);
  }

  /* Scrollbars */
  ::-webkit-scrollbar { width: 10px; height: 10px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: rgba(120, 160, 230, 0.15); border-radius: 10px; }
  ::-webkit-scrollbar-thumb:hover { background: rgba(120, 160, 230, 0.3); }
"""

_HEADER = """
  <header>
    <a href="/" class="home">
      <span class="logo-mark">M</span>
      <span class="brand-text">M365 <span class="accent">/</span> Console</span>
    </a>
    <div class="user-info">
      <div class="who">
        <span class="avatar">{{ (me.get('displayName') or me.get('userPrincipalName') or 'U')[:1] | upper }}</span>
        <span>{{ me.get('displayName', me.get('userPrincipalName', '')) }}</span>
      </div>
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
  <title>M365 Console — Sign in</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
    :root {
      --bg-0: #060912; --bg-1: #0c1222;
      --accent: #00d4ff; --accent-2: #6c8cff; --accent-purple: #b465ff;
      --text: #e6ecf7; --text-dim: #8a93a8;
      --border: rgba(120, 160, 230, 0.16);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', system-ui, sans-serif;
      background:
        radial-gradient(900px 600px at 20% 20%, rgba(108, 140, 255, 0.22), transparent 60%),
        radial-gradient(800px 500px at 80% 80%, rgba(180, 101, 255, 0.16), transparent 60%),
        radial-gradient(600px 400px at 50% 100%, rgba(0, 212, 255, 0.18), transparent 60%),
        linear-gradient(180deg, var(--bg-0), var(--bg-1));
      color: var(--text);
      min-height: 100vh;
      display: flex; align-items: center; justify-content: center;
      padding: 24px;
      overflow: hidden;
      position: relative;
    }
    body::before {
      content: ''; position: fixed; inset: 0; pointer-events: none;
      background-image:
        linear-gradient(rgba(120, 160, 230, 0.05) 1px, transparent 1px),
        linear-gradient(90deg, rgba(120, 160, 230, 0.05) 1px, transparent 1px);
      background-size: 48px 48px;
      mask-image: radial-gradient(ellipse at center, black 20%, transparent 70%);
    }
    /* Floating orbs */
    .orb { position: fixed; border-radius: 50%; filter: blur(60px); opacity: 0.55; pointer-events: none; }
    .orb.a { width: 360px; height: 360px; background: #6c8cff; top: -100px; left: -120px; animation: float 18s ease-in-out infinite; }
    .orb.b { width: 280px; height: 280px; background: #b465ff; bottom: -80px; right: -80px; animation: float 22s ease-in-out infinite reverse; }
    .orb.c { width: 220px; height: 220px; background: #00d4ff; top: 50%; right: 20%; animation: float 26s ease-in-out infinite; opacity: 0.35; }
    @keyframes float {
      0%, 100% { transform: translate(0, 0) scale(1); }
      33% { transform: translate(30px, -40px) scale(1.05); }
      66% { transform: translate(-20px, 30px) scale(0.95); }
    }

    .card {
      position: relative;
      background: rgba(14, 20, 38, 0.6);
      backdrop-filter: blur(20px) saturate(150%);
      -webkit-backdrop-filter: blur(20px) saturate(150%);
      border: 1px solid var(--border);
      border-radius: 20px;
      padding: 56px 48px;
      text-align: center;
      box-shadow: 0 30px 80px rgba(0, 0, 0, 0.45), 0 0 0 1px rgba(255, 255, 255, 0.02) inset;
      max-width: 460px; width: 100%;
      z-index: 1;
    }
    .card::before {
      content: ''; position: absolute; inset: -1px;
      border-radius: 20px;
      padding: 1px;
      background: linear-gradient(135deg, rgba(0, 212, 255, 0.5), transparent 40%, transparent 60%, rgba(180, 101, 255, 0.5));
      -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
      -webkit-mask-composite: xor; mask-composite: exclude;
      pointer-events: none;
    }

    .logo-mark {
      width: 64px; height: 64px;
      border-radius: 18px;
      background: linear-gradient(135deg, var(--accent), var(--accent-2) 60%, var(--accent-purple));
      display: inline-flex; align-items: center; justify-content: center;
      margin: 0 auto 24px;
      font-family: 'JetBrains Mono', monospace;
      font-weight: 800; font-size: 28px; color: #060912;
      box-shadow: 0 0 40px rgba(0, 212, 255, 0.45), 0 8px 24px rgba(0, 0, 0, 0.4);
    }
    .badge {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 4px 12px;
      background: rgba(0, 212, 255, 0.08);
      border: 1px solid rgba(0, 212, 255, 0.25);
      border-radius: 999px;
      font-size: 11px; font-weight: 600;
      color: var(--accent);
      letter-spacing: 0.05em; text-transform: uppercase;
      margin-bottom: 16px;
    }
    .badge .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--accent); box-shadow: 0 0 8px var(--accent); animation: pulse 2s ease-in-out infinite; }
    @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }
    h1 { font-size: 28px; font-weight: 700; margin-bottom: 12px; letter-spacing: -0.02em; background: linear-gradient(180deg, #fff, #b8c6e0); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }
    p.sub { color: var(--text-dim); margin-bottom: 36px; line-height: 1.6; font-size: 14px; }
    .btn {
      position: relative;
      display: inline-flex; align-items: center; gap: 12px;
      background: linear-gradient(135deg, #00d4ff, #6c8cff);
      color: #060912;
      padding: 14px 32px;
      border-radius: 12px;
      text-decoration: none;
      font-size: 14px; font-weight: 700;
      letter-spacing: 0.01em;
      transition: all 0.2s ease;
      box-shadow: 0 0 28px rgba(0, 212, 255, 0.35), 0 10px 24px rgba(0, 0, 0, 0.3);
    }
    .btn:hover { transform: translateY(-2px); box-shadow: 0 0 40px rgba(0, 212, 255, 0.55), 0 14px 30px rgba(0, 0, 0, 0.4); }
    .btn:active { transform: translateY(0); }
    .ms-icon { width: 18px; height: 18px; }

    .features {
      display: flex; justify-content: center; gap: 20px;
      margin-top: 36px; padding-top: 28px;
      border-top: 1px solid var(--border);
      font-size: 12px; color: var(--text-dim);
    }
    .features span { display: inline-flex; align-items: center; gap: 6px; }
    .features svg { width: 14px; height: 14px; color: var(--accent); }
  </style>
</head>
<body>
  <div class="orb a"></div>
  <div class="orb b"></div>
  <div class="orb c"></div>
  <div class="card">
    <div class="logo-mark">M</div>
    <div class="badge"><span class="dot"></span> Microsoft Graph · Live</div>
    <h1>M365 Console</h1>
    <p class="sub">Unified workspace for your Microsoft 365 — emails, OneDrive, Teams &amp; SharePoint, with an AI agent built in.</p>
    <a href="/login" class="btn">
      <svg class="ms-icon" viewBox="0 0 23 23" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="1" y="1" width="10" height="10" fill="#f25022"/>
        <rect x="12" y="1" width="10" height="10" fill="#7fba00"/>
        <rect x="1" y="12" width="10" height="10" fill="#00a4ef"/>
        <rect x="12" y="12" width="10" height="10" fill="#ffb900"/>
      </svg>
      Continue with Microsoft
    </a>
    <div class="features">
      <span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> OAuth 2.0</span>
      <span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> Encrypted</span>
      <span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> Read-only</span>
    </div>
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
    .welcome { margin-bottom: 28px; display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
    .welcome h2 { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; background: linear-gradient(180deg, #fff, #b8c6e0); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }
    .welcome p { color: var(--text-dim); margin-top: 4px; font-size: 13px; font-family: 'JetBrains Mono', monospace; }
    .welcome .badge { display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px; background: rgba(74, 222, 128, 0.08); border: 1px solid rgba(74, 222, 128, 0.25); border-radius: 999px; font-size: 11px; font-weight: 600; color: var(--success); letter-spacing: 0.04em; text-transform: uppercase; }
    .welcome .badge .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--success); box-shadow: 0 0 8px var(--success); animation: pulse 2s ease-in-out infinite; }
    @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }

    .debug-bar { margin-bottom: 16px; text-align: right; }
    .debug-btn { display: inline-flex; align-items: center; gap: 6px; padding: 7px 14px; background: rgba(255,255,255,0.02); border: 1px solid var(--border); border-radius: 8px; font-size: 12px; color: var(--text-dim); text-decoration: none; font-family: 'JetBrains Mono', monospace; transition: all 0.15s ease; }
    .debug-btn:hover { background: rgba(0, 212, 255, 0.06); color: var(--accent); border-color: rgba(0, 212, 255, 0.3); }

    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; }
    .card {
      position: relative; overflow: hidden;
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 24px;
      text-decoration: none; color: inherit;
      display: flex; flex-direction: column; gap: 14px;
      transition: all 0.22s cubic-bezier(0.2, 0.8, 0.2, 1);
    }
    .card::before {
      content: ''; position: absolute; inset: 0;
      background: radial-gradient(400px 200px at var(--mx, 50%) var(--my, 0%), var(--card-glow, rgba(0, 212, 255, 0.12)), transparent 60%);
      opacity: 0; transition: opacity 0.3s ease; pointer-events: none;
    }
    .card:hover { transform: translateY(-4px); border-color: var(--card-border, rgba(0, 212, 255, 0.4)); box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35), 0 0 30px var(--card-glow, rgba(0, 212, 255, 0.15)); }
    .card:hover::before { opacity: 1; }
    .card.email      { --card-glow: rgba(0, 212, 255, 0.18); --card-border: rgba(0, 212, 255, 0.4); }
    .card.files      { --card-glow: rgba(74, 222, 128, 0.18); --card-border: rgba(74, 222, 128, 0.4); }
    .card.teams      { --card-glow: rgba(180, 101, 255, 0.18); --card-border: rgba(180, 101, 255, 0.4); }
    .card.sharepoint { --card-glow: rgba(255, 182, 72, 0.18); --card-border: rgba(255, 182, 72, 0.4); }
    .card-head { display: flex; align-items: center; justify-content: space-between; }
    .card-icon { width: 44px; height: 44px; border-radius: 12px; display: inline-flex; align-items: center; justify-content: center; font-size: 22px; background: rgba(255,255,255,0.04); border: 1px solid var(--border); position: relative; }
    .card.email      .card-icon { color: var(--accent); background: rgba(0, 212, 255, 0.1); border-color: rgba(0, 212, 255, 0.25); }
    .card.files      .card-icon { color: var(--success); background: rgba(74, 222, 128, 0.1); border-color: rgba(74, 222, 128, 0.25); }
    .card.teams      .card-icon { color: var(--accent-purple); background: rgba(180, 101, 255, 0.1); border-color: rgba(180, 101, 255, 0.25); }
    .card.sharepoint .card-icon { color: var(--warning); background: rgba(255, 182, 72, 0.1); border-color: rgba(255, 182, 72, 0.25); }
    .card-arrow { color: var(--text-muted); font-size: 16px; transition: transform 0.2s ease, color 0.2s ease; }
    .card:hover .card-arrow { transform: translateX(4px); color: var(--text); }
    .card-body { display: flex; flex-direction: column; gap: 4px; }
    .card-count { font-size: 38px; font-weight: 800; line-height: 1; color: var(--text); letter-spacing: -0.03em; font-family: 'JetBrains Mono', monospace; }
    .card-label { font-size: 13px; color: var(--text-dim); font-weight: 500; }
    .card-count.na { font-size: 22px; color: var(--text-muted); }

    /* --- Chat panel --- */
    .chat-section {
      margin-top: 36px;
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 16px;
      overflow: hidden;
      box-shadow: var(--glow-soft);
    }
    .chat-header {
      background: linear-gradient(180deg, rgba(0, 212, 255, 0.04), transparent);
      border-bottom: 1px solid var(--border);
      padding: 16px 22px;
      display: flex; align-items: center; justify-content: space-between;
    }
    .chat-header h3 { font-size: 14px; font-weight: 600; color: var(--text); display: inline-flex; align-items: center; gap: 10px; }
    .chat-header h3 .ai-dot { width: 8px; height: 8px; border-radius: 50%; background: linear-gradient(135deg, var(--accent), var(--accent-purple)); box-shadow: 0 0 12px var(--accent); animation: pulse 2s ease-in-out infinite; }
    .chat-header .hint { font-size: 11px; color: var(--text-muted); font-family: 'JetBrains Mono', monospace; padding: 4px 10px; background: rgba(255,255,255,0.02); border: 1px solid var(--border); border-radius: 999px; }
    .chat-messages { height: 420px; overflow-y: auto; padding: 20px 22px; background: rgba(6, 10, 22, 0.4); }
    .chat-messages:empty::before { content: '> Ask anything about your OneDrive. e.g. "What is the metrics for RAG?"'; color: var(--text-muted); font-size: 13px; font-family: 'JetBrains Mono', monospace; }
    .msg { margin-bottom: 16px; max-width: 88%; }
    .msg.user { margin-left: auto; text-align: right; }
    .msg.user .bubble { background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #060912; font-weight: 500; box-shadow: 0 4px 16px rgba(0, 212, 255, 0.25); }
    .msg.agent .bubble { background: rgba(255,255,255,0.03); border: 1px solid var(--border); color: var(--text); }
    .msg.agent.error .bubble { background: rgba(255, 85, 114, 0.08); border-color: rgba(255, 85, 114, 0.3); color: #ffb6c0; }
    .bubble { display: inline-block; padding: 12px 16px; border-radius: 14px; font-size: 14px; line-height: 1.55; text-align: left; white-space: normal; word-wrap: break-word; max-width: 100%; }
    .msg.user .bubble { white-space: pre-wrap; }
    .agent-status { font-size: 11px; color: var(--accent); margin-bottom: 8px; display: inline-flex; align-items: center; gap: 8px; font-family: 'JetBrains Mono', monospace; padding: 3px 10px; background: rgba(0, 212, 255, 0.08); border: 1px solid rgba(0, 212, 255, 0.2); border-radius: 999px; }
    .agent-steps { display: flex; flex-direction: column; gap: 8px; margin: 10px 0; }
    .live-step { background: rgba(6, 10, 22, 0.6); border: 1px solid var(--border); border-left: 3px solid var(--accent); padding: 10px 12px; border-radius: 8px; font-family: 'JetBrains Mono', monospace; font-size: 11px; }
    .live-step .step-head { color: var(--text-dim); line-height: 1.7; }
    .live-step .step-no { background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #060912; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 10px; }
    .live-step .step-ts { color: var(--text-muted); font-size: 10px; }
    .live-step .tname { color: var(--success); font-weight: 600; margin-left: 6px; }
    .live-step .tin { color: var(--text-dim); margin-left: 4px; word-break: break-all; }
    .live-step .step-obs { margin-top: 6px; }
    .live-step .step-obs .pending { color: var(--text-muted); font-style: italic; }
    .live-step pre { white-space: pre-wrap; word-wrap: break-word; max-height: 220px; overflow-y: auto; background: rgba(0,0,0,0.4); border: 1px solid var(--border); color: #b8c6e0; padding: 8px; border-radius: 6px; margin: 0; font-size: 11px; }
    .agent-answer { font-size: 14px; line-height: 1.6; color: var(--text); }
    .agent-answer:not(:empty) { margin-top: 10px; padding-top: 10px; border-top: 1px dashed var(--border); }
    .summary-box { margin-top: 12px; padding: 12px 14px; background: rgba(255, 182, 72, 0.06); border: 1px solid rgba(255, 182, 72, 0.25); border-radius: 10px; font-size: 12px; color: var(--text); }
    .summary-box .sum-title { font-weight: 700; color: var(--warning); margin-bottom: 8px; font-size: 13px; letter-spacing: 0.02em; }
    .summary-box .sum-section { margin-top: 8px; color: var(--text-dim); }
    .summary-box .sum-section b { color: var(--text); font-weight: 600; }
    .summary-box ul { margin: 6px 0 6px 20px; padding: 0; }
    .summary-box li { margin: 3px 0; line-height: 1.5; }
    .summary-box code { background: rgba(255, 182, 72, 0.1); padding: 1px 6px; border-radius: 4px; font-size: 11px; color: var(--warning); word-break: break-all; font-family: 'JetBrains Mono', monospace; }
    .summary-box .dim { color: var(--text-muted); font-size: 11px; }
    .chat-input-row { display: flex; gap: 10px; padding: 14px 22px; border-top: 1px solid var(--border); background: rgba(8, 12, 24, 0.5); }
    .chat-input-row input {
      flex: 1; padding: 12px 16px;
      background: rgba(255,255,255,0.03);
      border: 1px solid var(--border);
      border-radius: 10px; font-size: 14px; outline: none;
      color: var(--text); font-family: inherit;
      transition: all 0.15s ease;
    }
    .chat-input-row input::placeholder { color: var(--text-muted); }
    .chat-input-row input:focus { border-color: var(--accent); background: rgba(0, 212, 255, 0.04); box-shadow: 0 0 0 3px rgba(0, 212, 255, 0.12); }
    .chat-input-row button {
      padding: 12px 24px;
      background: linear-gradient(135deg, var(--accent), var(--accent-2));
      color: #060912; border: none; border-radius: 10px;
      font-size: 14px; font-weight: 700; cursor: pointer;
      transition: all 0.15s ease;
      box-shadow: 0 0 20px rgba(0, 212, 255, 0.25);
    }
    .chat-input-row button:hover:not(:disabled) { transform: translateY(-1px); box-shadow: 0 0 28px rgba(0, 212, 255, 0.4); }
    .chat-input-row button:disabled { opacity: 0.4; cursor: not-allowed; }
    .typing { display: inline-flex; align-items: center; gap: 3px; }
    .typing span { display: inline-block; width: 5px; height: 5px; background: var(--accent); border-radius: 50%; animation: blink 1.2s infinite; }
    .typing span:nth-child(2) { animation-delay: 0.2s; }
    .typing span:nth-child(3) { animation-delay: 0.4s; }
    @keyframes blink { 0%, 60%, 100% { opacity: 0.2; } 30% { opacity: 1; } }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <div class="welcome">
      <div>
        <h2>Welcome back, {{ me.get('displayName', 'there') }}</h2>
        <p>{{ me.get('mail') or me.get('userPrincipalName', '') }}</p>
      </div>
      <div class="badge"><span class="dot"></span> Graph API · Online</div>
    </div>

    <div class="debug-bar">
      <a href="/debug-files" class="debug-btn">&#9881; debug.onedrive</a>
    </div>
    <div class="cards">
      <a href="/emails" class="card email">
        <div class="card-head">
          <div class="card-icon">&#9993;</div>
          <span class="card-arrow">&rarr;</span>
        </div>
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
        <div class="card-head">
          <div class="card-icon">&#128193;</div>
          <span class="card-arrow">&rarr;</span>
        </div>
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
        <div class="card-head">
          <div class="card-icon">&#128101;</div>
          <span class="card-arrow">&rarr;</span>
        </div>
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
        <div class="card-head">
          <div class="card-icon">&#127760;</div>
          <span class="card-arrow">&rarr;</span>
        </div>
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
        <h3><span class="ai-dot"></span> AI Assistant · Ask your OneDrive</h3>
        <span class="hint">graph + claude · keyword agent</span>
      </div>
      <div id="chat-messages" class="chat-messages"></div>
      <form id="chat-form" class="chat-input-row" autocomplete="off">
        <input id="chat-input" type="text" placeholder="Ask anything about your files..." />
        <button id="chat-send" type="submit">Send &rarr;</button>
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
    .page-title { font-size: 22px; font-weight: 700; color: var(--text); margin-bottom: 18px; letter-spacing: -0.02em; display: inline-flex; align-items: center; gap: 10px; }
    .page-title .count { font-family: 'JetBrains Mono', monospace; font-size: 13px; color: var(--accent); padding: 3px 10px; background: rgba(0, 212, 255, 0.08); border: 1px solid rgba(0, 212, 255, 0.25); border-radius: 999px; }
    .filter-bar { display: inline-flex; gap: 4px; margin-bottom: 20px; padding: 4px; background: rgba(255,255,255,0.02); border: 1px solid var(--border); border-radius: 12px; }
    .filter-btn { padding: 7px 18px; border: 1px solid transparent; border-radius: 8px; font-size: 13px; font-weight: 500; color: var(--text-dim); text-decoration: none; background: transparent; transition: all 0.15s ease; }
    .filter-btn:hover { color: var(--text); background: rgba(255,255,255,0.04); }
    .filter-btn.active { background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #060912; font-weight: 700; box-shadow: 0 0 16px rgba(0, 212, 255, 0.3); }
    .email-list { display: flex; flex-direction: column; gap: 10px; }
    .email-card {
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-left: 3px solid var(--accent);
      border-radius: 12px;
      padding: 18px 22px;
      transition: all 0.18s ease;
    }
    .email-card:hover { transform: translateX(2px); border-color: var(--border-strong); box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3); }
    .email-card.high { border-left-color: var(--danger); background: linear-gradient(90deg, rgba(255, 85, 114, 0.05), transparent 40%); }
    .email-card.read { border-left-color: var(--text-muted); opacity: 0.65; }
    .email-card.read .email-subject { font-weight: 500; color: var(--text-dim); }
    .unread-dot { display: inline-block; width: 8px; height: 8px; background: var(--accent); border-radius: 50%; margin-right: 8px; vertical-align: middle; box-shadow: 0 0 8px var(--accent); }
    .email-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 8px; }
    .email-subject { font-weight: 600; color: var(--text); font-size: 15px; flex: 1; letter-spacing: -0.01em; }
    .email-date { color: var(--text-muted); font-size: 11px; white-space: nowrap; font-family: 'JetBrains Mono', monospace; }
    .email-from { font-size: 13px; color: var(--accent); margin-bottom: 8px; font-weight: 500; }
    .email-from .sender-name { color: var(--text); }
    .email-from .sender-email { color: var(--text-muted); font-family: 'JetBrains Mono', monospace; font-size: 11px; }
    .email-preview { font-size: 13px; color: var(--text-dim); line-height: 1.5; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
    .tag { display: inline-block; background: rgba(255, 85, 114, 0.12); color: var(--danger); border: 1px solid rgba(255, 85, 114, 0.3); font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 4px; margin-left: 8px; vertical-align: middle; letter-spacing: 0.05em; text-transform: uppercase; }
    .empty { text-align: center; padding: 80px 24px; color: var(--text-dim); }
    .empty .icon { font-size: 56px; margin-bottom: 16px; opacity: 0.5; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&larr; Dashboard</a>
    <div class="page-title">&#9993; Inbox <span class="count">{{ emails|length }} msg</span></div>

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
            {% if high %}<span class="tag">High</span>{% endif %}
          </span>
          <span class="email-date">{{ email.get('receivedDateTime', '')[:10] }}</span>
        </div>
        <div class="email-from">
          <span class="sender-name">{{ email.get('from', {}).get('emailAddress', {}).get('name', '') }}</span>
          <span class="sender-email">&lt;{{ email.get('from', {}).get('emailAddress', {}).get('address', '') }}&gt;</span>
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
    .page-title { font-size: 22px; font-weight: 700; color: var(--text); margin-bottom: 18px; letter-spacing: -0.02em; display: inline-flex; align-items: center; gap: 10px; }
    .page-title .count { font-family: 'JetBrains Mono', monospace; font-size: 13px; color: var(--success); padding: 3px 10px; background: rgba(74, 222, 128, 0.08); border: 1px solid rgba(74, 222, 128, 0.25); border-radius: 999px; font-weight: 500; }

    .search-bar {
      display: flex; gap: 10px; margin-bottom: 22px; align-items: center;
      padding: 6px;
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 14px;
    }
    .debug-btn { padding: 10px 14px; background: rgba(255,255,255,0.02); border: 1px solid var(--border); border-radius: 10px; font-size: 12px; color: var(--text-dim); text-decoration: none; white-space: nowrap; font-family: 'JetBrains Mono', monospace; transition: all 0.15s ease; }
    .debug-btn:hover { background: rgba(0, 212, 255, 0.06); color: var(--accent); border-color: rgba(0, 212, 255, 0.3); }
    .search-bar input { flex: 1; padding: 11px 16px; background: transparent; border: none; border-radius: 10px; font-size: 14px; outline: none; color: var(--text); font-family: inherit; }
    .search-bar input::placeholder { color: var(--text-muted); }
    .search-bar input:focus { background: rgba(0, 212, 255, 0.04); }
    .search-bar button { padding: 11px 22px; background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #060912; border: none; border-radius: 10px; font-size: 13px; font-weight: 700; cursor: pointer; box-shadow: 0 0 18px rgba(0, 212, 255, 0.25); transition: all 0.15s ease; }
    .search-bar button:hover { transform: translateY(-1px); box-shadow: 0 0 26px rgba(0, 212, 255, 0.4); }
    .clear-search { display: inline-block; margin-bottom: 14px; font-size: 12px; color: var(--accent); text-decoration: none; font-family: 'JetBrains Mono', monospace; }
    .clear-search:hover { text-decoration: underline; }
    .search-mode-label { font-size: 13px; color: var(--text-dim); margin-bottom: 14px; }
    .search-mode-label strong { color: var(--text); }

    .file-list { display: flex; flex-direction: column; gap: 8px; }
    .file-row {
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 14px 20px;
      display: flex; align-items: center; gap: 14px;
      text-decoration: none; color: inherit;
      transition: all 0.15s ease;
    }
    .file-row:hover { transform: translateX(2px); border-color: var(--border-strong); background: rgba(0, 212, 255, 0.04); }
    .file-icon { width: 36px; height: 36px; border-radius: 9px; background: rgba(255,255,255,0.03); border: 1px solid var(--border); display: inline-flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0; color: var(--text-dim); }
    .file-row:hover .file-icon { color: var(--accent); border-color: rgba(0, 212, 255, 0.3); }
    .file-body { flex: 1; min-width: 0; }
    .file-name { font-weight: 500; color: var(--text); font-size: 14px; word-break: break-word; }
    .file-summary { font-size: 12px; color: var(--text-dim); margin-top: 4px; line-height: 1.5; }
    .file-summary em { font-style: normal; background: rgba(255, 182, 72, 0.15); color: var(--warning); padding: 0 4px; border-radius: 3px; }
    .file-meta { font-size: 11px; color: var(--text-muted); white-space: nowrap; text-align: right; flex-shrink: 0; font-family: 'JetBrains Mono', monospace; }
    .empty { text-align: center; padding: 80px 24px; color: var(--text-dim); }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&larr; Dashboard</a>

    <form class="search-bar" method="get" action="/files">
      <input type="text" name="q" value="{{ query }}" placeholder="Search file names and content..." autocomplete="off" />
      <button type="submit">&#128269; Search</button>
      <a href="/debug-search" class="debug-btn">&#9881; debug</a>
    </form>

    {% if query %}
    <div>
      <span class="search-mode-label">Results for <strong>&ldquo;{{ query }}&rdquo;</strong> &mdash; {{ files|length }} found</span>
      &nbsp;<a href="/files" class="clear-search">&times; clear search</a>
    </div>
    {% else %}
    <div class="page-title">&#128193; OneDrive <span class="count">{{ files|length }} items</span></div>
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
      <p>No files found matching &ldquo;{{ query }}&rdquo;.</p>
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
    .page-title { font-size: 22px; font-weight: 700; color: var(--text); margin-bottom: 18px; letter-spacing: -0.02em; display: inline-flex; align-items: center; gap: 10px; }
    .page-title .count { font-family: 'JetBrains Mono', monospace; font-size: 13px; color: var(--accent-purple); padding: 3px 10px; background: rgba(180, 101, 255, 0.08); border: 1px solid rgba(180, 101, 255, 0.25); border-radius: 999px; font-weight: 500; }
    .team-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 14px; }
    .team-card {
      position: relative; overflow: hidden;
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 22px;
      text-decoration: none; color: inherit;
      display: block;
      transition: all 0.22s cubic-bezier(0.2, 0.8, 0.2, 1);
    }
    .team-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 3px; background: linear-gradient(180deg, var(--accent-purple), var(--accent-2)); }
    .team-card::after { content: ''; position: absolute; inset: 0; background: radial-gradient(400px 200px at 0% 0%, rgba(180, 101, 255, 0.12), transparent 60%); opacity: 0; transition: opacity 0.3s ease; pointer-events: none; }
    .team-card:hover { transform: translateY(-3px); border-color: rgba(180, 101, 255, 0.4); box-shadow: 0 20px 40px rgba(0, 0, 0, 0.35), 0 0 30px rgba(180, 101, 255, 0.15); }
    .team-card:hover::after { opacity: 1; }
    .team-name { font-weight: 600; color: var(--text); font-size: 15px; margin-bottom: 6px; letter-spacing: -0.01em; }
    .team-desc { font-size: 13px; color: var(--text-dim); line-height: 1.5; min-height: 20px; }
    .team-visibility { display: inline-block; margin-top: 12px; font-size: 10px; font-weight: 700; padding: 3px 10px; border-radius: 999px; background: rgba(255,255,255,0.04); border: 1px solid var(--border); color: var(--text-dim); text-transform: uppercase; letter-spacing: 0.06em; font-family: 'JetBrains Mono', monospace; }
    .team-action { margin-top: 14px; font-size: 12px; color: var(--accent-purple); font-weight: 600; display: inline-flex; align-items: center; gap: 6px; }
    .team-action .arrow { transition: transform 0.2s ease; }
    .team-card:hover .team-action .arrow { transform: translateX(4px); }
    .empty { text-align: center; padding: 80px 24px; color: var(--text-dim); }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&larr; Dashboard</a>
    <div class="page-title">&#128101; Teams <span class="count">{{ teams|length }} groups</span></div>

    {% if teams %}
    <div class="team-list">
      {% for t in teams %}
      <a href="/teams/{{ t.get('id') }}" class="team-card">
        <div class="team-name">{{ t.get('displayName', '') }}</div>
        <div class="team-desc">{{ t.get('description', '') or '&mdash;' }}</div>
        <span class="team-visibility">{{ t.get('visibility', 'unknown') }}</span>
        <div class="team-action">&#128193; View shared files <span class="arrow">&rarr;</span></div>
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
    .page-title { font-size: 24px; font-weight: 700; color: var(--text); margin-bottom: 6px; letter-spacing: -0.02em; }
    .page-subtitle { color: var(--text-dim); font-size: 13px; margin-bottom: 18px; }
    .crumb { font-size: 12px; color: var(--text-muted); margin-bottom: 14px; font-family: 'JetBrains Mono', monospace; }
    .crumb a { color: var(--accent); text-decoration: none; }
    .crumb a:hover { text-decoration: underline; }
    .file-list { display: flex; flex-direction: column; gap: 8px; }
    .file-row {
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 14px 20px;
      display: flex; align-items: center; gap: 14px;
      text-decoration: none; color: inherit;
      transition: all 0.15s ease;
    }
    .file-row:hover { transform: translateX(2px); border-color: var(--border-strong); background: rgba(0, 212, 255, 0.04); }
    .file-icon { width: 36px; height: 36px; border-radius: 9px; background: rgba(255,255,255,0.03); border: 1px solid var(--border); display: inline-flex; align-items: center; justify-content: center; font-size: 18px; color: var(--text-dim); flex-shrink: 0; }
    .file-body { flex: 1; min-width: 0; }
    .file-name { font-weight: 500; color: var(--text); font-size: 14px; word-break: break-word; }
    .file-meta { font-size: 11px; color: var(--text-muted); white-space: nowrap; text-align: right; flex-shrink: 0; font-family: 'JetBrains Mono', monospace; }
    .file-actions { display: flex; gap: 8px; align-items: center; }
    .action-btn { padding: 6px 12px; background: rgba(255,255,255,0.03); border: 1px solid var(--border); border-radius: 8px; font-size: 12px; font-weight: 500; color: var(--text-dim); text-decoration: none; transition: all 0.15s ease; }
    .action-btn:hover { background: rgba(0, 212, 255, 0.06); color: var(--accent); border-color: rgba(0, 212, 255, 0.3); }
    .action-btn.primary { background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #060912; border-color: transparent; font-weight: 700; box-shadow: 0 0 12px rgba(0, 212, 255, 0.25); }
    .action-btn.primary:hover { transform: translateY(-1px); color: #060912; box-shadow: 0 0 20px rgba(0, 212, 255, 0.4); }
    .empty { text-align: center; padding: 80px 24px; color: var(--text-dim); }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="{{ back_url }}" class="back">&larr; {{ back_label }}</a>
    <div class="page-title">&#128193; {{ title }}</div>
    <div class="page-subtitle">{{ subtitle }}</div>

    {% if current_folder %}
    <div class="crumb"><a href="{{ folder_url_base }}">~/root</a> / (folder)</div>
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
    .page-title { font-size: 22px; font-weight: 700; color: var(--text); margin-bottom: 18px; letter-spacing: -0.02em; display: inline-flex; align-items: center; gap: 10px; }
    .page-title .count { font-family: 'JetBrains Mono', monospace; font-size: 13px; color: var(--warning); padding: 3px 10px; background: rgba(255, 182, 72, 0.08); border: 1px solid rgba(255, 182, 72, 0.25); border-radius: 999px; font-weight: 500; }
    .site-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 14px; }
    .site-card {
      position: relative; overflow: hidden;
      background: var(--surface);
      backdrop-filter: blur(14px) saturate(140%);
      -webkit-backdrop-filter: blur(14px) saturate(140%);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 22px;
      text-decoration: none; color: inherit;
      display: block;
      transition: all 0.22s cubic-bezier(0.2, 0.8, 0.2, 1);
    }
    .site-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 3px; background: linear-gradient(180deg, var(--warning), #ff8a3d); }
    .site-card::after { content: ''; position: absolute; inset: 0; background: radial-gradient(400px 200px at 0% 0%, rgba(255, 182, 72, 0.10), transparent 60%); opacity: 0; transition: opacity 0.3s ease; pointer-events: none; }
    .site-card:hover { transform: translateY(-3px); border-color: rgba(255, 182, 72, 0.4); box-shadow: 0 20px 40px rgba(0, 0, 0, 0.35), 0 0 30px rgba(255, 182, 72, 0.12); }
    .site-card:hover::after { opacity: 1; }
    .site-name { font-weight: 600; color: var(--text); font-size: 15px; margin-bottom: 6px; letter-spacing: -0.01em; }
    .site-desc { font-size: 13px; color: var(--text-dim); line-height: 1.5; min-height: 20px; }
    .site-action { margin-top: 14px; font-size: 12px; color: var(--warning); font-weight: 600; display: inline-flex; align-items: center; gap: 6px; }
    .site-action .arrow { transition: transform 0.2s ease; }
    .site-card:hover .site-action .arrow { transform: translateX(4px); }
    .empty { text-align: center; padding: 80px 24px; color: var(--text-dim); }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&larr; Dashboard</a>
    <div class="page-title">&#127760; SharePoint <span class="count">{{ sites|length }} sites</span></div>

    {% if sites %}
    <div class="site-list">
      {% for s in sites %}
      <a href="/sharepoint/{{ s.get('id') }}" class="site-card">
        <div class="site-name">{{ s.get('displayName') or s.get('name', '(unnamed)') }}</div>
        <div class="site-desc">{{ s.get('description') or '&mdash;' }}</div>
        <div class="site-action">&#128196; Browse documents <span class="arrow">&rarr;</span></div>
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
    .file-header { background: var(--surface); backdrop-filter: blur(14px) saturate(140%); -webkit-backdrop-filter: blur(14px) saturate(140%); border: 1px solid var(--border); border-radius: 14px; padding: 22px; margin-bottom: 18px; box-shadow: var(--glow-soft); }
    .file-title { font-size: 18px; font-weight: 700; color: var(--text); margin-bottom: 8px; word-break: break-all; letter-spacing: -0.01em; }
    .file-meta { font-size: 13px; color: var(--text-dim); font-family: 'JetBrains Mono', monospace; }
    .file-meta a { color: var(--accent); text-decoration: none; margin-left: 8px; }
    .file-meta a:hover { text-decoration: underline; }
    .content-box { background: rgba(0, 0, 0, 0.5); color: #d4dbe8; border: 1px solid var(--border); border-radius: 12px; padding: 20px; font-family: 'JetBrains Mono', 'Consolas', monospace; font-size: 13px; line-height: 1.65; white-space: pre-wrap; word-wrap: break-word; overflow-x: auto; max-height: 70vh; overflow-y: auto; box-shadow: inset 0 1px 0 rgba(255,255,255,0.03); }
    .notice { background: rgba(255, 182, 72, 0.06); border: 1px solid rgba(255, 182, 72, 0.3); border-left: 3px solid var(--warning); padding: 12px 16px; border-radius: 10px; margin-bottom: 16px; color: var(--warning); font-size: 13px; }
    .info { background: var(--surface); backdrop-filter: blur(14px) saturate(140%); -webkit-backdrop-filter: blur(14px) saturate(140%); border: 1px solid var(--border); border-radius: 14px; padding: 48px 32px; text-align: center; color: var(--text-dim); }
    .info .icon { font-size: 56px; margin-bottom: 14px; opacity: 0.6; }
    .info a.btn { display: inline-block; margin-top: 14px; padding: 11px 26px; background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #060912; border-radius: 10px; text-decoration: none; font-weight: 700; font-size: 14px; box-shadow: 0 0 20px rgba(0, 212, 255, 0.25); transition: all 0.15s ease; }
    .info a.btn:hover { transform: translateY(-1px); box-shadow: 0 0 28px rgba(0, 212, 255, 0.4); }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="{{ back_url }}" class="back">&larr; Back</a>
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
    .error-card { background: var(--surface); backdrop-filter: blur(14px) saturate(140%); -webkit-backdrop-filter: blur(14px) saturate(140%); border: 1px solid rgba(255, 85, 114, 0.3); border-left: 3px solid var(--danger); border-radius: 14px; padding: 32px; max-width: 600px; margin: 48px auto; box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35), 0 0 30px rgba(255, 85, 114, 0.1); }
    .error-card .err-tag { display: inline-flex; align-items: center; gap: 6px; font-size: 11px; font-weight: 700; color: var(--danger); padding: 4px 10px; background: rgba(255, 85, 114, 0.08); border: 1px solid rgba(255, 85, 114, 0.3); border-radius: 999px; letter-spacing: 0.06em; text-transform: uppercase; font-family: 'JetBrains Mono', monospace; margin-bottom: 14px; }
    .error-card .err-tag .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--danger); box-shadow: 0 0 8px var(--danger); }
    .error-card h2 { color: var(--text); margin-bottom: 12px; font-size: 20px; font-weight: 700; letter-spacing: -0.01em; }
    .error-card p { color: var(--text-dim); line-height: 1.6; font-size: 14px; }
  </style>
</head>
<body>
  """ + _HEADER + """
  <main>
    <a href="/" class="back">&larr; Dashboard</a>
    <div class="error-card">
      <div class="err-tag"><span class="dot"></span> Error</div>
      <h2>Something went wrong</h2>
      <p>{{ message }}</p>
    </div>
  </main>
</body>
</html>
"""

if __name__ == "__main__":
    app.run(port=3000, debug=True, threaded=True)
