# Microsoft 365 Keyword-Search Agent

Ask natural-language questions over documents scattered across OneDrive, SharePoint, Microsoft Teams and Outlook attachments. One Microsoft login, one search API, no vector database.

## Problems this project solves

**1. Documents are scattered across the enterprise.** A real Microsoft 365 user keeps work artefacts in OneDrive (personal drafts), SharePoint (team libraries), Teams (channel files) and Outlook (attachments). Crawling each surface with its own API and permission model is brittle and expensive.

**2. Classical RAG is heavy and lossy.** Chunking breaks tables and code, embeddings cost money on every change, a vector database is one more service to run, and approximate-vector matching misses exact terms (product codes, acronyms, version numbers).

## Solution

- **Microsoft Graph as the universal entry point.** A single OAuth token, a single `POST /search/query` call, results federated across OneDrive / SharePoint / Teams / Outlook. Permissions inherit from the signed-in user automatically.
- **Agentic keyword search instead of RAG.** An LLM agent generates OR-joined keyword sets, reads the snippets Graph returns, downloads a file only when needed, and runs `grep` with surrounding context inside it. No chunking, no embeddings, no vector store.

Inspired by the paper *"Keyword search is all you need"* (arXiv 2602.23368).

## Goals

- Aggregate documents from multiple Microsoft 365 sources through one API.
- Replace vector-based RAG with fast keyword search plus regex grep.
- Produce cited answers (`[filename]` or `[filename, line N]`) for every response.
- Keep the operational footprint minimal: one Flask process, two SaaS endpoints (Graph + Azure OpenAI), no extra databases.
- Make every reasoning step visible in real time via Server-Sent Events.

## Workflow

```
User query -> Flask /ask (SSE)
           -> LangChain agent (langgraph) with 3 tools:

  +------------------------------------------------------+
  | Tool 1: graph_search(query)                          |
  |   POST /search/query on Microsoft Graph              |
  |   -> ranked files + snippets across                  |
  |      OneDrive / SharePoint / Teams / Outlook         |
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

**How this differs from traditional RAG:** no pre-chunking, no embeddings, no vector DB. Microsoft already indexes the content — the agent just calls the API and greps.

## Repository layout

```
app.py                  Flask: MSAL OAuth + dashboard + /ask SSE
agent/
|-- agent.py            create_agent (langgraph) + run_agent_stream
|-- llm.py              AzureChatOpenAI factory
|-- prompts.py          SYSTEM_PROMPT enforcing the workflow
|-- cache.py            in-memory text cache
|-- tools/              graph_search, file_fetch, grep_context
`-- extractors/         pdf, docx, xlsx, pptx
make_docs_en.py         Generates docs_report_en.docx (product report)
```

## Setup

```powershell
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # macOS / Linux
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

Smoke-test Azure OpenAI credentials before running the app:
```powershell
python test_azure.py
```

## Run

```powershell
python app.py
```

Open `http://localhost:3000` -> sign in with Microsoft -> use the chat panel on the dashboard.

## Supported

- File types: PDF, DOCX, XLSX, PPTX, plain text
- 25 MB per file limit
- Sources reachable via Microsoft Graph: OneDrive, SharePoint, Teams files, Outlook attachments

## Conclusion

The prototype answers both motivating problems with one design: Microsoft Graph solves the document-scatter problem, and an LLM agent with keyword search plus grep solves the RAG-overhead problem. The result is simpler, cheaper to operate and easier to audit than a multi-connector RAG stack. Reserve vector-based RAG for use cases where semantic similarity (paraphrase, cross-language, summarisation across many documents) is demonstrably required.

## Remaining problems

Files whose content is carried by **images** (PNG/JPG, scanned PDFs, image-only slides) have **not yet been validated** end to end. Microsoft 365 runs server-side OCR on such files and merges the extracted text into its full-text index, so in theory `graph_search` should return them. What still needs measurement:

- Recall: do OCR-extracted files actually surface in search results?
- Latency to indexing: how long after upload until they become searchable?
- Snippet quality and ranking against natively-typed text.
- Agent policy when `graph_search` returns an image-only file: `fetch_file_text` cannot OCR locally, so we must decide whether to cite directly from the Graph snippet or skip the file.

Closing this gap is the main follow-up before claiming full coverage of enterprise documents.

## Documentation

A more detailed product report is generated by:

```powershell
python make_docs_en.py
# -> docs_report_en.docx
```
