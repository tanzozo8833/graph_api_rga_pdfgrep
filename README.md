# Microsoft 365 Keyword-Search Agent

Ask natural-language questions over your OneDrive documents — no vector database required.

Inspired by the paper *"Keyword search is all you need"* (arxiv 2602.23368): instead of chunking + embedding + vector DB, an LLM agent invokes Microsoft Graph's built-in keyword search and then greps for context inside files it downloads.

## Workflow

```
User query -> Flask /ask (SSE)
           -> LangChain agent (langgraph) with 3 tools:

  +------------------------------------------------------+
  | Tool 1: graph_search(query)                          |
  |   POST /search/query on Microsoft Graph              |
  |   -> list of files + snippets matching the keyword   |
  |                                                      |
  | Tool 2: fetch_file_text(item_id)                     |
  |   GET /me/drive/items/{id}/content                   |
  |   -> extract text (PDF/DOCX/XLSX/PPTX) -> cache      |
  |                                                      |
  | Tool 3: grep_context(item_id, regex, +/- N lines)    |
  |   re.compile over the text cache, merge overlapping  |
  +------------------------------------------------------+

The LLM decides which tool to call and loops until it has enough info
-> Final Answer with citation [filename, line N]
-> Frontend receives each step via SSE and renders in real time
```

**How this differs from traditional RAG:** no pre-chunking, no embeddings, no vector DB. Microsoft already indexes OneDrive — the agent just calls the API and greps.

## Structure

```
app.py                  Flask: MSAL OAuth + dashboard + /ask SSE
agent/
|-- agent.py            create_agent (langgraph) + run_agent_stream
|-- llm.py              AzureChatOpenAI factory
|-- prompts.py          SYSTEM_PROMPT enforcing the workflow
|-- cache.py            in-memory text cache
|-- tools/              graph_search, file_fetch, grep_context
`-- extractors/         pdf, docx, xlsx, pptx
```

## Setup

```powershell
pip install -r requirements.txt
```

Fill in `.env` (see `.env.example`):
```
CLIENT_ID=         # Azure App Registration
CLIENT_SECRET=
TENANT_ID=
FLASK_SECRET_KEY=

AZURE_OPENAI_API_KEY=
AZURE_OPENAI_ENDPOINT=
AZURE_OPENAI_DEPLOYMENT=     # must be a tool-calling model: gpt-4o, gpt-4.1
AZURE_OPENAI_API_VERSION=
```

**Azure App Registration** requirements:
- Redirect URI: `http://localhost:3000/auth/callback` (type: Web)
- Permissions (delegated): `User.Read`, `Mail.Read`, `Files.Read`, `Files.Read.All`, `Sites.Read.All`, `Team.ReadBasic.All`

## Run

```powershell
python app.py
```

Open `http://localhost:3000` -> sign in with Microsoft -> use the chat panel on the dashboard.

## Supported

- File types: PDF, DOCX, XLSX, PPTX, plain text
- 25 MB per file limit
- Scanned (image-only) PDFs are not OCR'd yet
