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

import os
from flask import Flask, redirect, request, session, url_for, render_template_string
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
SCOPES        = ["User.Read", "Mail.Read", "Files.Read", "Team.ReadBasic.All"]
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
    unread_count = _fetch_unread_count(token)
    files_count  = _fetch_files_count(token)
    teams_count  = _fetch_teams_count(token)

    return render_template_string(DASHBOARD_HTML,
        me=me,
        unread_count=unread_count,
        files_count=files_count,
        teams_count=teams_count,
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

    params = {
        "$filter": "isRead eq false",
        "$orderby": "receivedDateTime desc",
        "$top": 50,
        "$select": "subject,from,receivedDateTime,bodyPreview,importance",
    }
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

    return render_template_string(EMAILS_HTML, me=me, emails=emails_data)


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
    .card-icon { font-size: 36px; flex-shrink: 0; }
    .card-body {}
    .card-count { font-size: 36px; font-weight: 700; line-height: 1; color: #111; }
    .card-label { font-size: 14px; color: #666; margin-top: 4px; }
    .card-count.na { font-size: 18px; color: #999; }
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
    </div>
  </main>
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
    .email-list { display: flex; flex-direction: column; gap: 8px; }
    .email-card { background: white; border-radius: 8px; padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 3px solid #0078d4; transition: box-shadow 0.15s; }
    .email-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .email-card.high { border-left-color: #d13438; }
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
    <div class="page-title">&#128140; Unread Emails ({{ emails|length }})</div>

    {% if emails %}
    <div class="email-list">
      {% for email in emails %}
      {% set high = email.get('importance') == 'high' %}
      <div class="email-card {{ 'high' if high else '' }}">
        <div class="email-header">
          <span class="email-subject">
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
      <p>No unread emails — inbox zero!</p>
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
    .team-card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 4px solid #6264a7; }
    .team-name { font-weight: 600; color: #111; font-size: 15px; margin-bottom: 6px; }
    .team-desc { font-size: 13px; color: #666; line-height: 1.4; }
    .team-visibility { display: inline-block; margin-top: 10px; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 3px; background: #f0f0f0; color: #555; text-transform: capitalize; }
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
      <div class="team-card">
        <div class="team-name">{{ t.get('displayName', '') }}</div>
        <div class="team-desc">{{ t.get('description', '') or '—' }}</div>
        <span class="team-visibility">{{ t.get('visibility', 'unknown') }}</span>
      </div>
      {% endfor %}
    </div>
    {% else %}
    <div class="empty"><p>No Teams groups found.</p></div>
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
    app.run(port=3000, debug=True)
