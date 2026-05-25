# Microsoft 365 Keyword-Search Agent

Hỏi đáp tài liệu trong OneDrive bằng tiếng tự nhiên — không cần vector database.

Lấy cảm hứng từ paper *"Keyword search is all you need"* (arxiv 2602.23368): thay vì chunking + embedding + vector DB, dùng LLM agent gọi keyword search có sẵn của Microsoft Graph, rồi grep ngữ cảnh trên file tải về.

## Workflow

```
User query → Flask /ask (SSE)
           → LangChain agent (langgraph) với 3 tools:

  ┌──────────────────────────────────────────────────────┐
  │ Tool 1: graph_search(query)                          │
  │   POST /search/query của Microsoft Graph             │
  │   → list file + snippet match keyword                │
  │                                                      │
  │ Tool 2: fetch_file_text(item_id)                     │
  │   GET /me/drive/items/{id}/content                   │
  │   → extract text (PDF/DOCX/XLSX/PPTX) → cache        │
  │                                                      │
  │ Tool 3: grep_context(item_id, regex, ±N lines)       │
  │   re.compile trên text cache, merge overlapping      │
  └──────────────────────────────────────────────────────┘

LLM tự quyết định gọi tool nào, lặp đến khi đủ thông tin
→ Final Answer kèm citation [filename, line N]
→ Frontend nhận từng bước qua SSE, hiển thị realtime
```

**Điểm khác RAG truyền thống:** không chunk trước, không embed, không vector DB. Microsoft đã index sẵn OneDrive — agent chỉ gọi API + grep.

## Cấu trúc

```
app.py                  Flask: OAuth MSAL + dashboard + /ask SSE
agent/
├── agent.py            create_agent (langgraph) + run_agent_stream
├── llm.py              AzureChatOpenAI factory
├── prompts.py          SYSTEM_PROMPT bắt buộc workflow
├── cache.py            in-memory text cache
├── tools/              graph_search, file_fetch, grep_context
└── extractors/         pdf, docx, xlsx, pptx
```

## Setup

```powershell
pip install -r requirements.txt
```

Điền `.env` (xem `.env.example`):
```
CLIENT_ID=         # Azure App Registration
CLIENT_SECRET=
TENANT_ID=
FLASK_SECRET_KEY=

AZURE_OPENAI_API_KEY=
AZURE_OPENAI_ENDPOINT=
AZURE_OPENAI_DEPLOYMENT=     # phải là model hỗ trợ tool calling: gpt-4o, gpt-4.1
AZURE_OPENAI_API_VERSION=
```

**Azure App Registration** cần:
- Redirect URI: `http://localhost:3000/auth/callback` (type: Web)
- Permissions (delegated): `User.Read`, `Mail.Read`, `Files.Read`, `Files.Read.All`, `Sites.Read.All`, `Team.ReadBasic.All`

## Chạy

```powershell
python app.py
```

Mở `http://localhost:3000` → đăng nhập Microsoft → dùng chat panel ở dashboard.

## Hỗ trợ

- File types: PDF, DOCX, XLSX, PPTX, plain text
- Giới hạn 25 MB/file
- PDF scan (image-only) chưa OCR
