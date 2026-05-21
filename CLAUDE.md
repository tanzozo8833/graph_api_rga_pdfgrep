# Microsoft Graph API — Unread Email Viewer
First of all, you don't have permission to read .env file, only .env.example 
## What this app does

A Flask web app that authenticates users via Microsoft Entra ID (OAuth2 Authorization Code flow) and displays their unread emails using the Microsoft Graph API.

Flow:
1. User opens `http://localhost:3000`
2. Clicks **Sign in with Microsoft** → redirected to Microsoft login
3. After login, Microsoft redirects to `/auth/callback` with an auth code
4. App exchanges code for tokens, stores access token in session
5. App queries Graph API for unread emails (`isRead eq false`)
6. Emails rendered in a clean UI with sender, subject, date, and preview

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask app — routes, MSAL auth, Graph API calls, HTML templates |
| `graph_client.py` | Original CLI Graph client (MSAL public client, interactive browser) |
| `.env` | Secrets (CLIENT_ID, TENANT_ID, CLIENT_SECRET, FLASK_SECRET_KEY) |
| `requirements.txt` | Python dependencies |

## Azure App Registration setup

Go to [portal.azure.com](https://portal.azure.com) → Azure Active Directory → App registrations:

1. **Redirect URI**: `http://localhost:3000/auth/callback` (type: **Web**)
2. **API permissions**: `User.Read`, `Mail.Read` (delegated)
3. **Client secret**: Certificates & secrets → New client secret → copy value to `.env`

## Running locally

```bash
pip install -r requirements.txt

# Fill in .env with CLIENT_SECRET and FLASK_SECRET_KEY
python app.py
# → http://localhost:3000
```

## Key implementation details

- Uses `msal.ConfidentialClientApplication` (web app / confidential client) — requires a client secret
- Tokens stored in Flask session (in-memory; not persisted across restarts)
- Graph filter: `$filter=isRead eq false` with `$orderby=receivedDateTime desc`, top 50
- High-importance emails highlighted with red left border and tag
- `graph_client.py` uses `PublicClientApplication` (CLI / desktop flow) — separate from the web app
