 └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 1: Flask route /ask                                             │
  │   - Pull access_token từ session (đã có sẵn từ MSAL OAuth flow)      │
  │   - Validate query không rỗng                                        │
  │   - Trả Response(generate(), mimetype="text/event-stream")           │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 2: agent.agent.run_agent_stream(query, token)                   │
  │   yield {"event": "user_query", "query": "...", "ts": "..."}         │
  │   ↓                                                                  │
  │   Build agent:                                                       │
  │     llm = AzureChatOpenAI(deployment=gpt-4o)                         │
  │     tools = [                                                        │
  │       make_graph_search_tool(token),  # closure capture token       │
  │       make_file_fetch_tool(token),                                   │
  │       make_grep_context_tool(),                                      │
  │     ]                                                                │
  │     agent = create_agent(llm, tools, system_prompt=SYSTEM_PROMPT)    │
  │   ↓                                                                  │
  │   agent.stream({"messages": [HumanMessage(query)]}, mode="updates")  │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 3 — Iteration 1 (LLM decision)                                  │
  │                                                                      │
  │   AIMessage tool_calls:                                              │
  │     [{                                                               │
  │       name: "graph_search",                                          │
  │       args: {                                                        │
  │         query: "metric OR evaluation OR benchmark OR RAG OR          │
  │                 retrieval augmented generation"                      │
  │       }                                                              │
  │     }]                                                               │
  │                                                                      │
  │   yield {"event": "tool_call", "step": 1, "tool": "graph_search",    │
  │          "input": "{...}", "ts": "08:35:13"}                         │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 4 — graph_search execution                                      │
  │                                                                      │
  │   POST https://graph.microsoft.com/v1.0/search/query                 │
  │   Headers: { Authorization: Bearer <token> }                         │
  │   Body: {                                                            │
  │     requests: [{                                                     │
  │       entityTypes: ["driveItem"],                                    │
  │       query: { queryString: "metric OR evaluation ..." },            │
  │       from: 0, size: 15,                                             │
  │       fields: ["id","name","webUrl","lastModifiedDateTime","size"]   │
  │     }]                                                               │
  │   }                                                                  │
  │                                                                      │
  │   Response.hits = [                                                  │
  │     { resource: {id: "01ABC...", name: "rag_paper.pdf", ...},        │
  │       summary: "...we use BLEU and ROUGE-L as metrics..." },         │
  │     { resource: {id: "01XYZ...", name: "eval_methods.docx", ...},    │
  │       summary: "Recall@k is the standard metric for retrieval..." },│
  │     ...                                                              │
  │   ]                                                                  │
  │                                                                      │
  │   Format thành string trả về tool result:                            │
  │     "Found 8 files for query 'metric OR ...':                        │
  │      - item_id=01ABC | name=rag_paper.pdf |                          │
  │        snippet="...we use BLEU and ROUGE-L as metrics..." | ...      │
  │      - item_id=01XYZ | name=eval_methods.docx | snippet=..."         │
  │                                                                      │
  │   yield {"event": "tool_result", "step": 1,                          │
  │          "observation": "Found 8 files...", "ts": "08:35:14"}        │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 5 — LLM nhìn snippets, quyết định Iter 2                        │
  │                                                                      │
  │   Trường hợp A — snippet đã đủ trả lời:                              │
  │     LLM → AIMessage(content="The metrics for RAG include BLEU,       │
  │                    ROUGE-L, Recall@k... [rag_paper.pdf,              │
  │                    eval_methods.docx]")                              │
  │     yield {"event": "final_answer", "answer": "..."}                 │
  │     KẾT THÚC.                                                        │
  │                                                                      │
  │   Trường hợp B — cần đọc sâu hơn (giả sử):                           │
  │     AIMessage tool_calls: [                                          │
  │       {name: "fetch_file_text", args: {item_id: "01ABC..."}},        │
  │       {name: "fetch_file_text", args: {item_id: "01XYZ..."}}         │
  │     ]                                                                │
  │     (LLM có thể gọi parallel — gpt-4o hỗ trợ)                        │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 6 — fetch_file_text execution                                   │
  │                                                                      │
  │   Step 6a: GET /me/drive/items/01ABC.../?$select=name,size,file      │
  │            → meta: name="rag_paper.pdf", size=1.2 MB                 │
  │            Check size < 25 MB → OK                                   │
  │            Check ext (.pdf) ∈ SUPPORTED → OK                         │
  │                                                                      │
  │   Step 6b: GET /me/drive/items/01ABC.../content                      │
  │            → binary content (PDF bytes)                              │
  │                                                                      │
  │   Step 6c: extract_text("rag_paper.pdf", bytes)                      │
  │            → dispatch sang pdf.py                                    │
  │            → pdfplumber.open() → iter pages                          │
  │            → return "[PAGE 1]\nText...\n[PAGE 2]\nText..."           │
  │                                                                      │
  │   Step 6d: cache.put("01ABC...", "rag_paper.pdf", text)              │
  │                                                                      │
  │   Return: "Fetched and cached. filename=rag_paper.pdf,               │
  │            size=1.2MB, text_length=45000 chars. Preview: ..."        │
  │                                                                      │
  │   yield {"event": "tool_result", "step": 2, ...}                     │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 7 — LLM quyết định grep                                         │
  │                                                                      │
  │   AIMessage tool_calls: [{                                           │
  │     name: "grep_context",                                            │
  │     args: {                                                          │
  │       item_id: "01ABC...",                                           │
  │       patterns: "(metric|metrics|BLEU|ROUGE|recall|precision|F1|     │
  │                  evaluation|score|benchmark)",                       │
  │       context_lines: 5                                               │
  │     }                                                                │
  │   }]                                                                 │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 8 — grep_context execution                                      │
  │                                                                      │
  │   text = cache.get("01ABC...")["text"]                               │
  │   lines = text.split("\n")                                           │
  │   regex = re.compile(patterns, re.IGNORECASE)                        │
  │   match_indices = [i for i,l in enumerate(lines) if regex.search(l)] │
  │                                                                      │
  │   Merge overlapping windows ±5:                                      │
  │     windows = [(10,21), (45,60), (120,135), ...]                     │
  │                                                                      │
  │   Output:                                                            │
  │     Found 12 matches in rag_paper.pdf (merged into 4 chunks):        │
  │     --- lines 11-21 ---                                              │
  │     >> L13: The primary metrics for retrieval-augmented...           │
  │        L14: ...                                                      │
  │        L15: We report Recall@5 and Recall@10 as the main...          │
  │     --- lines 46-60 ---                                              │
  │     ...                                                              │
  │                                                                      │
  │   yield {"event": "tool_result", "step": 3, ...}                     │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 9 — LLM tổng hợp final answer                                   │
  │                                                                      │
  │   AIMessage (không có tool_calls):                                   │
  │     content: "Based on rag_paper.pdf and eval_methods.docx, the      │
  │               metrics commonly used for RAG include:                 │
  │                                                                      │
  │               1. **Retrieval metrics**:                              │
  │                  - Recall@k (typically k=5, k=10) [rag_paper.pdf, L15]│
  │                  - Precision@k                                       │
  │                                                                      │
  │               2. **Generation metrics**:                             │
  │                  - BLEU score                                        │
  │                  - ROUGE-L [eval_methods.docx]                       │
  │                  - Semantic similarity                               │
  │                                                                      │
  │               3. **End-to-end**:                                     │
  │                  - Faithfulness                                      │
  │                  - Answer relevancy [rag_paper.pdf, page 4]"         │
  │                                                                      │
  │   yield {"event": "final_answer", "answer": "..."}                   │
  │   yield {"event": "done"}                                            │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ BƯỚC 10 — Frontend nhận events                                       │
  │                                                                      │
  │   JS consumeStream() loop:                                           │
  │     reader.read() → decoder.decode()                                 │
  │     split('\n\n') → each "data: {...}" → JSON.parse → handleEvent()  │
  │                                                                      │
  │   handleEvent:                                                       │
  │     - tool_call → addStep() — show step card với status "..."        │
  │     - tool_result → fillStepResult() — fill observation              │
  │     - final_answer → setAnswer() — render below trace                │
  │     - done → hide typing indicator                                   │
  └──────────────────────────────────────────────────────────────────────┘

  Retry strategy khi không tìm thấy

  LLM tự nhận biết kết quả rỗng (snippet không liên quan, grep không match) → tạo
  strategy mới:

  Iter 1: graph_search("metric OR evaluation")  → 0 hits
  Iter 2: Thought: "Có thể từ khóa khác trong corpus tiếng Việt..."
          graph_search("đánh giá OR thước đo OR chỉ số")  → 3 hits
  Iter 3: graph_search("MRR OR NDCG OR hit rate")        → 5 hits (alt names)
  Iter 4: fetch + grep tiếp tục...

  SYSTEM_PROMPT bắt buộc thử ít nhất 2-3 strategy trước khi giving up.

  ---
  3. Cấu trúc code chi tiết

  14_Microsoft_GraphAPI/
  ├── app.py                          # Flask app — thêm POST /ask + chat panel
  ├── graph_client.py                 # CLI client cũ — không đụng
  ├── requirements.txt                # thêm langchain, langchain-openai, pdf/office
  libs
  ├── .env.example                    # thêm AZURE_OPENAI_* vars
  ├── debug/                          # đã có
  │
  ├── agent/                          # 🆕 toàn bộ module agent
  │   ├── __init__.py                 # re-export run_agent
  │   │
  │   ├── llm.py                      # 🆕 factory ChatModel
  │   │   def make_llm():
  │   │     validate 4 env vars
  │   │     return AzureChatOpenAI(
  │   │       azure_endpoint=..., api_key=..., api_version=...,
  │   │       azure_deployment=..., temperature=0, max_tokens=4096)
  │   │
  │   ├── prompts.py                  # 🆕 SYSTEM_PROMPT
  │   │   SYSTEM_PROMPT = """
  │   │     Bạn là assistant ... DO NOT answer from your own knowledge.
  │   │     Có 3 tools: graph_search, fetch_file_text, grep_context.
  │   │     REQUIRED WORKFLOW:
  │   │       Step 1: ALWAYS call graph_search FIRST với OR keywords.
  │   │       Step 2: Đọc snippets — đủ → Final Answer; thiếu → Step 3.
  │   │       Step 3: fetch_file_text top 1-3 file.
  │   │       Step 4: grep_context với multi-pattern regex.
  │   │       Step 5: nếu rỗng → RESTART với strategy khác.
  │   │     RULES:
  │   │       - ALWAYS multi-pattern OR, NEVER single word
  │   │       - ALWAYS cite [filename, line N]
  │   │       - DO NOT invent, nếu không thấy → "I could not find this"
  │   │   """
  │   │
  │   ├── cache.py                    # 🆕 in-memory text cache
  │   │   _TEXT_CACHE: dict[item_id, {filename, text}]
  │   │   def get(item_id), put(item_id, name, text), clear()
  │   │
  │   ├── agent.py                    # 🆕 bootstrap + run loop
  │   │   def _build_agent(token):
  │   │     llm = make_llm()
  │   │     tools = [
  │   │       make_graph_search_tool(token),
  │   │       make_file_fetch_tool(token),
  │   │       make_grep_context_tool(),
  │   │     ]
  │   │     return create_agent(llm, tools, SYSTEM_PROMPT)
  │   │
  │   │   def run_agent_stream(query, token) -> Iterator[dict]:
  │   │     yield {"event": "user_query", ...}
  │   │     agent = _build_agent(token)
  │   │     for chunk in agent.stream({"messages": [...]},
  │   │                                stream_mode="updates"):
  │   │       for node_name, state in chunk.items():
  │   │         for msg in state.get("messages", []):
  │   │           if isinstance(msg, AIMessage) and msg.tool_calls:
  │   │             yield {"event": "tool_call", ...}
  │   │           elif isinstance(msg, AIMessage):
  │   │             yield {"event": "final_answer", ...}
  │   │           elif isinstance(msg, ToolMessage):
  │   │             yield {"event": "tool_result", ...}
  │   │
  │   │   def run_agent(query, token) -> dict:
  │   │     drain stream → return {answer, steps}
  │   │
  │   ├── tools/
  │   │   ├── __init__.py             # re-export 3 factories
  │   │   │
  │   │   ├── graph_search.py         # 🆕 Tool 1
  │   │   │   def make_graph_search_tool(token):
  │   │   │     @tool
  │   │   │     def graph_search(query: str) -> str:
  │   │   │       """Docstring (LLM đọc!) hướng dẫn dùng OR keywords"""
  │   │   │       POST /search/query với entityTypes=['driveItem']
  │   │   │       parse data.value[0].hitsContainers[0].hits
  │   │   │       format "- item_id=... | name=... | snippet=... | webUrl=..."
  │   │   │     return graph_search
  │   │   │
  │   │   ├── file_fetch.py           # 🆕 Tool 2
  │   │   │   MAX_SIZE = 25 MB
  │   │   │   def make_file_fetch_tool(token):
  │   │   │     @tool
  │   │   │     def fetch_file_text(item_id: str) -> str:
  │   │   │       """Docstring: cảnh báo không gọi 2 lần cùng id"""
  │   │   │       if cache.get(item_id): return "ALREADY CACHED"
  │   │   │       GET metadata → check size + ext
  │   │   │       GET content → extract_text(name, bytes)
  │   │   │       cache.put(...)
  │   │   │       return "Fetched. filename=..., text_length=..., Preview: ..."
  │   │   │     return fetch_file_text
  │   │   │
  │   │   └── grep_context.py         # 🆕 Tool 3
  │   │       MAX_CHUNKS = 10
  │   │       MAX_CHARS_PER_LINE = 400
  │   │       def make_grep_context_tool():
  │   │         @tool
  │   │         def grep_context(item_id, patterns, context_lines=5) -> str:
  │   │           """Docstring: hướng dẫn dùng | OR regex"""
  │   │           text = cache.get(item_id)
  │   │           regex = re.compile(patterns, re.IGNORECASE)
  │   │           match_indices = ...
  │   │           merge overlapping ±N windows
  │   │           format ">> L{n}: line" cho hit, "   L{n}: line" cho context
  │   │           cap MAX_CHUNKS
  │   │         return grep_context
  │   │
  │   └── extractors/
  │       ├── __init__.py             # 🆕 dispatcher
  │       │   SUPPORTED = {.pdf, .docx, .xlsx, .pptx,
  │       │                .txt, .md, .csv, .tsv, .log, .json, .xml, .yaml, .yml}
  │       │   def extract_text(filename, bytes):
  │       │     ext = os.path.splitext(filename)[1].lower()
  │       │     if ext == ".pdf": return extract_pdf(bytes)
  │       │     ...
  │       │     if ext in PLAINTEXT_EXTS: return bytes.decode("utf-8")
  │       │
  │       ├── pdf.py                  # 🆕 pdfplumber.open(BytesIO) → iter pages,
  │       │                           #     prepend "[PAGE N]" markers
  │       │
  │       ├── docx.py                 # 🆕 Document(BytesIO) → para.text + tables
  │       │                           #     bảng: "[TABLE n]" + cell1 | cell2 | ...
  │       │
  │       ├── xlsx.py                 # 🆕 load_workbook(read_only=True,
  │       │                           #     data_only=True) — bỏ qua formula
  │       │                           #     "[SHEET name]" + cell1 | cell2 | ...
  │       │                           #     skip blank rows
  │       │
  │       └── pptx.py                 # 🆕 Presentation(BytesIO) → iter slides
  │                                   #     "[SLIDE n]" + shape.text_frame.text
  │                                   #     + "[NOTES] ..." nếu có notes_slide

  ---
  4. Phase implementation chi tiết

  Phase 0 — Setup (30 phút)

  Step 0.1. Append requirements.txt:
  langchain>=0.3.0
  langchain-openai>=0.2.0   # (ban đầu langchain-anthropic, sau đổi)
  langchain-core>=0.3.0
  pdfplumber>=0.11.0
  python-docx>=1.1.0
  openpyxl>=3.1.0
  python-pptx>=1.0.0

  Step 0.2. Append .env.example:
  AZURE_OPENAI_API_KEY=
  AZURE_OPENAI_ENDPOINT=
  AZURE_OPENAI_DEPLOYMENT=
  AZURE_OPENAI_API_VERSION=

  Step 0.3. pip install -r requirements.txt

  Step 0.4. Smoke test:
  from langchain_openai import AzureChatOpenAI
  llm = AzureChatOpenAI(...)
  print(llm.invoke("hello").content)

  ---
  Phase 1 — Text extractors (1 giờ)

  Nguyên tắc thiết kế:
  - Output là plain text line-oriented (split bằng \n) để grep -C N hoạt động
  - Mỗi unit (page/slide/sheet/table) prefix bằng marker [PAGE 3], [SLIDE 2], [SHEET
  Data], [TABLE 1] — để LLM cite chính xác
  - Đầu vào: bytes, đầu ra: str

  1.1. extractors/pdf.py:
  import io, pdfplumber

  def extract_pdf(content: bytes) -> str:
      parts = []
      with pdfplumber.open(io.BytesIO(content)) as pdf:
          for i, page in enumerate(pdf.pages, start=1):
              parts.append(f"[PAGE {i}]")
              parts.append(page.extract_text() or "")
      return "\n".join(parts)

  1.2. extractors/docx.py:
  - paragraphs → mỗi dòng 1 line
  - tables → [TABLE n] rồi mỗi row cell1 | cell2 | cell3
  - skip empty para

  1.3. extractors/xlsx.py:
  - read_only=True, data_only=True — quan trọng để xử lý file lớn + lấy giá trị (không
  formula)
  - mỗi sheet: [SHEET <name>] rồi mỗi row non-empty: cells join |
  - replace \n trong cell bằng space để giữ line-oriented

  1.4. extractors/pptx.py:
  - iter slides, marker [SLIDE n]
  - iter shapes có text_frame, extract paragraph runs
  - notes slide: [NOTES] ...

  1.5. extractors/__init__.py — dispatcher:
  SUPPORTED = OFFICE_EXTS | PLAINTEXT_EXTS

  def extract_text(filename: str, content: bytes) -> str:
      ext = os.path.splitext(filename)[1].lower()
      # dispatch...
      if ext in PLAINTEXT_EXTS:
          try: return content.decode("utf-8")
          except UnicodeDecodeError: return content.decode("latin-1", errors="replace")
      raise ValueError(f"Unsupported: {ext}")

  ---
  Phase 2 — 3 LangChain Tools (2 giờ)

  Nguyên tắc thiết kế:
  - Mỗi tool có factory make_xxx_tool(token) — closure giữ access_token
  - Tool định nghĩa qua @tool decorator → LangChain tự generate schema từ type hints +
  docstring
  - Docstring rất quan trọng — LLM đọc docstring để biết khi nào dùng tool và dùng thế
  nào
  - Return luôn là str (không trả dict/list để tránh format issue)
  - Error → return string "ERROR: ..." thay vì raise exception (để agent tự xử lý)

  2.1. tools/graph_search.py:

  Endpoint: POST https://graph.microsoft.com/v1.0/search/query

  Request body chi tiết:
  {
    "requests": [{
      "entityTypes": ["driveItem"],
      "query": { "queryString": "<keyword OR keyword2>" },
      "from": 0,
      "size": 15,
      "fields": ["id", "name", "webUrl", "lastModifiedDateTime", "size"]
    }]
  }

  Response structure:
  data.value[0].hitsContainers[0].hits = [
    {
      "hitId": "...",
      "rank": 1,
      "summary": "...highlight snippet với <ddd:c0></ddd:c0> markup...",
      "resource": {
        "@odata.type": "#microsoft.graph.driveItem",
        "id": "01ABC...",
        "name": "rag_paper.pdf",
        "webUrl": "https://...",
        ...
      }
    }
  ]

  Tool format output (string mà LLM đọc):
  Found 8 files for query 'metric OR evaluation':
  - item_id=01ABC... | name=rag_paper.pdf | snippet="...we use BLEU and ROUGE..." |
  webUrl=https://...
  - item_id=01XYZ... | name=eval.docx | snippet="..." | webUrl=https://...

  2.2. tools/file_fetch.py:

  3 endpoint Graph API:
  - GET /me/drive/items/{id} với $select=name,size,file → metadata
  - GET /me/drive/items/{id}/content → binary

  Checks tuần tự:
  1. Có cached chưa? → return "ALREADY CACHED"
  2. Get metadata → 401 → return error
  3. size ≤ 25 MB? → quá → return error
  4. ext ∈ SUPPORTED? → không → return error
  5. Download content
  6. extract_text(name, bytes) — wrap trong try/except (bad PDF, corrupted DOCX...)
  7. cache.put(...)
  8. Return preview text 300 chars

  Return format:
  Fetched and cached. filename=rag_paper.pdf, size=1234567 bytes, text_length=45000
  chars.
  Preview: [PAGE 1] Abstract Retrieval-Augmented Generation (RAG) systems combine...

  2.3. tools/grep_context.py:

  Algorithm chi tiết:
  def grep_context(item_id, patterns, context_lines=5):
      text = cache.get(item_id)["text"]
      lines = text.split("\n")

      regex = re.compile(patterns, re.IGNORECASE)

      # Tìm match indices
      match_indices = [i for i, line in enumerate(lines) if regex.search(line)]
      if not match_indices:
          return f"No matches for pattern '{patterns}'"

      # Merge overlapping windows
      windows = []
      for idx in match_indices:
          start = max(0, idx - context_lines)
          end = min(len(lines), idx + context_lines + 1)
          if windows and start <= windows[-1][1]:
              # overlap → extend last window
              windows[-1][1] = max(windows[-1][1], end)
          else:
              windows.append([start, end])

      # Format output
      out = [f"Found {len(match_indices)} matches (merged into {len(windows)} chunks):"]
      for w_start, w_end in windows[:MAX_CHUNKS]:
          out.append(f"--- lines {w_start+1}-{w_end} ---")
          for i in range(w_start, w_end):
              marker = ">>" if regex.search(lines[i]) else "  "
              line = lines[i][:MAX_CHARS_PER_LINE]
              out.append(f"{marker} L{i+1}: {line}")
      return "\n".join(out)

  Cap output:
  - Max 10 chunks (nếu hơn → "+ N more chunks omitted")
  - Max 400 chars/line (nếu hơn → "...[truncated]")
  - → tránh blow context của LLM

  ---
  Phase 3 — Prompts + Agent (1 giờ)

  3.1. prompts.py — SYSTEM_PROMPT đầy đủ:

  Dựa trên cell 10 của notebook gốc, adapt sang context Microsoft Graph:

  Các điểm bắt buộc của prompt:
  1. Cấm trả lời từ kiến thức nội tại — chỉ từ tool results
  2. Mô tả 3 tools với signature + ví dụ
  3. Workflow 5 bước bắt buộc — graph_search trước, fetch nếu cần, grep multi-pattern,
  retry nếu rỗng
  4. Quy tắc cứng:
    - ALWAYS multi-pattern OR, NEVER single word
    - ALWAYS cite [filename, line N] hoặc [filename, PAGE N]
    - DO NOT invent
    - Cache persists → không re-fetch
    - Try 2-3 strategies trước khi giving up

  3.2. agent.py — bootstrap với langchain 1.x:

  from langchain.agents import create_agent
  from langchain_core.messages import AIMessage, ToolMessage

  def _build_agent(access_token):
      llm = make_llm()
      tools = [
          make_graph_search_tool(access_token),
          make_file_fetch_tool(access_token),
          make_grep_context_tool(),
      ]
      return create_agent(model=llm, tools=tools, system_prompt=SYSTEM_PROMPT)

  Streaming với stream_mode="updates":
  - Mỗi event là {node_name: {messages: [delta_messages]}}
  - Node "agent" → AIMessage (có thể có tool_calls)
  - Node "tools" → ToolMessage(s)

  Parse logic:
  for chunk in agent.stream({"messages": [HumanMessage(query)]},
                            stream_mode="updates",
                            config={"recursion_limit": 30}):
      for node_name, state in chunk.items():
          for msg in state.get("messages", []):
              if isinstance(msg, AIMessage):
                  if msg.tool_calls:
                      for tc in msg.tool_calls:
                          yield {"event": "tool_call", "step": step_no,
                                 "tool": tc["name"], "input": json.dumps(tc["args"]),
                                 "ts": now()}
                  else:
                      yield {"event": "final_answer", "answer": msg.content}
              elif isinstance(msg, ToolMessage):
                  yield {"event": "tool_result", "step": step_no,
                         "tool": msg.name, "observation": msg.content[:3000]}

  Console logging (cho dev visibility):
  [08:35:12] USER QUERY: ...
  [08:35:13] [STEP 1] -> graph_search({"query": "..."})
  [08:35:14] [STEP 1] <- graph_search result (1240 chars): Found 8 files...
  [08:35:16] [STEP 2] -> fetch_file_text({"item_id": "01ABC..."})
  ...
  [08:35:22] [FINAL ANSWER]
  The metrics for RAG include...

  ---
  Phase 4 — Flask integration (45 phút)

  4.1. Route /ask với SSE:

  @app.route("/ask", methods=["POST"])
  def ask():
      token = _get_token()  # từ session đã có
      if not token: return jsonify({"error": "Not authenticated"}), 401

      query = request.get_json().get("query", "").strip()
      if not query: return jsonify({"error": "Empty"}), 400

      from agent.agent import run_agent_stream

      def generate():
          try:
              for event in run_agent_stream(query, token):
                  yield f"data: {json.dumps(event)}\n\n"
          except Exception as e:
              yield f"data: {json.dumps({'event': 'error', 'message': str(e)})}\n\n"
          finally:
              yield 'data: {"event": "done"}\n\n'

      return Response(
          stream_with_context(generate()),
          mimetype="text/event-stream",
          headers={
              "Cache-Control": "no-cache",
              "X-Accel-Buffering": "no",   # disable nginx buffering nếu deploy
              "Connection": "keep-alive",
          },
      )

  4.2. Chat panel HTML trong DASHBOARD_HTML:

  Sau .cards div, thêm <section class="chat-section">:
  - Header với title "Ask your OneDrive"
  - Vùng #chat-messages height=380px scroll
  - Form input + send button

  4.3. CSS chi tiết:
  - .msg.user .bubble — xanh, lề phải
  - .msg.agent .bubble — trắng border xám, lề trái
  - .live-step — box màu xám nhạt với border xanh
    - .step-no — chip số bước
    - .step-ts — timestamp nhỏ
    - .tname — tên tool màu xanh
    - .tin — input màu xám
    - .step-obs pre — observation, max-height 220px, scrollable
  - .typing — 3 dots animation
  - .agent-answer — text answer, padding-top + dashed border

  4.4. JS SSE consumer:

  async function consumeStream(resp, container) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const parts = buffer.split('\n\n');
      buffer = parts.pop();  // keep incomplete
      for (const raw of parts) {
        const line = raw.trim();
        if (!line.startsWith('data:')) continue;
        const ev = JSON.parse(line.slice(5).trim());
        handleEvent(container, ev);
      }
    }
  }

  function handleEvent(container, ev) {
    switch (ev.event) {
      case 'user_query':    setStatus("Đang tìm kiếm..."); break;
      case 'tool_call':     setStatus("Bước "+ev.step+": "+ev.tool);
                            addStep(container, ev); break;
      case 'tool_result':   fillStepResult(container, ev); break;
      case 'final_answer':  hideTyping(); setAnswer(container, ev.answer); break;
      case 'error':         setError(container, ev.message); break;
      case 'done':          hideTyping(); break;
    }
  }

  ---
  5. Cảnh báo chi tiết + workaround

  Vấn đề: Permission Files.Read.All
  Chi tiết: /search/query yêu cầu broader scope hơn /me/drive/root/search
  Workaround: Đã có sẵn trong SCOPES của app — OK
  ────────────────────────────────────────
  Vấn đề: PDF scan (image-only)
  Chi tiết: pdfplumber.extract_text() trả "" hoặc rất ít
  Workaround: Phase này chấp nhận; phase 6 thêm Docling/Azure Document Intelligence
  ────────────────────────────────────────
  Vấn đề: File >25 MB
  Chi tiết: Memory + timeout
  Workaround: MAX_SIZE_BYTES = 25 * 1024 * 1024, từ chối với error message
  ────────────────────────────────────────
  Vấn đề: Token expired (1 giờ)
  Chi tiết: Graph API trả 401 mid-conversation
  Workaround: Tool return "ERROR: HTTP 401" → LLM hiểu cần re-login; user phải /login
  lại
  ────────────────────────────────────────
  Vấn đề: Cost mỗi câu hỏi
  Chi tiết: 5-15 LLM calls × gpt-4o input tokens × output tokens
  Workaround: Cache text giúp giảm token; cap MAX_CHUNKS=10; preview-only return từ
  fetch
  ────────────────────────────────────────
  Vấn đề: Search index lag
  Chi tiết: OneDrive index file mới sau vài phút
  Workaround: Không control được; báo user
  ────────────────────────────────────────
  Vấn đề: Privacy
  Chi tiết: Text gửi sang Azure OpenAI
  Workaround: Azure OpenAI region = data residency MS — OK cho enterprise
  ────────────────────────────────────────
  Vấn đề: LangChain 1.0 breaking
  Chi tiết: AgentExecutor → create_agent (langgraph)
  Workaround: Đã handle
  ────────────────────────────────────────
  Vấn đề: Multi-thread Flask
  Chi tiết: Mỗi request build agent mới, share _TEXT_CACHE global
  Workaround: OK cho dev; production cần per-user cache với session id
  ────────────────────────────────────────
  Vấn đề: Streaming buffer
  Chi tiết: Werkzeug dev server stream OK; nginx production cần X-Accel-Buffering: no
  Workaround: Đã thêm header

  ---
  6. Tổng kết thời gian

  ┌─────────────────┬──────────────────────────────────┬───────────┐
  │      Phase      │               Step               │ Thời gian │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 0               │ Setup + smoke test LLM           │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 1               │ 4 extractors + dispatcher        │ 1h        │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 2.1             │ graph_search tool                │ 40m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 2.2             │ file_fetch tool                  │ 40m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 2.3             │ grep_context tool                │ 40m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 3.1             │ SYSTEM_PROMPT chi tiết           │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 3.2             │ agent bootstrap + stream parsing │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 4.1             │ Flask /ask SSE route             │ 15m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 4.2-4.4         │ Chat panel HTML/CSS/JS           │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 5               │ Testing (bạn bỏ qua)             │ 1-2h      │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ Total (đã code) │ ~6 giờ                           │           │
  └─────────────────┴──────────────────────────────────┴───────────┘

  ---
  7. Status hiện tại

  Đã hoàn thành 100% Phase 0→4:
  - Tất cả 4 extractors có [PAGE/SHEET/SLIDE/TABLE N] markers
  - 3 tools với factory pattern + error handling
  - SYSTEM_PROMPT đầy đủ workflow + rules
  - Agent dùng create_agent (langchain 1.x langgraph) + stream_mode="updates"
  - Console logging với [STEP N] timestamp markers
  - Route /ask SSE streaming
  - Chat panel với live step cards, status updates realtime, error inline
  - Đã chuyển từ Anthropic → Azure OpenAI theo .env.example bạn config
  - Đã fix ImportError của langchain 1.3.1

  Còn chờ bạn:
  1. Điền 4 biến AZURE_OPENAI_* thật vào .env
  2. Deployment phải là model hỗ trợ function/tool calling (gpt-4o, gpt-4.1,
  gpt-4-turbo)
  3. Chạy venv\Scripts\python.exe app.py → test



  🌐 Tầng 1 — Browser (JS)

  1. form.addEventListener('submit', ...) — app.py:962
  - Listener gắn vào form chat khi page load. Chặn submit mặc định
  (e.preventDefault()) để không reload page.

  2. appendUser(text) — app.py:843
  - Tạo bubble màu xanh hiển thị câu hỏi user ngay khi gửi.

  3. createAgentContainer() — app.py:851
  - Tạo bubble agent rỗng có 3 vùng: status line (đang typing), khu
  chứa steps, vùng answer cuối.

  4. fetch('/ask', { method:'POST', body:{query} }) — app.py:973
  - Gửi HTTP POST kèm query JSON, Accept: text/event-stream để mở
  kênh SSE.

  5. consumeStream(resp, container) — app.py:913
  - Đọc từng chunk từ response body. Tách theo \n\n. Parse mỗi data:
  {...} thành object event. Gọi handleEvent.

  ---
  🔥 Tầng 2 — Flask backend (Python)

  6. ask() — app.py:643 (route /ask)
  - Lấy access_token từ session Flask (đã có từ MSAL OAuth).
  - Validate query không rỗng.
  - Mở generator SSE, mỗi event yield ra dưới dạng data: {...}\n\n.

  7. run_agent_stream(query, token) — agent/agent.py:53
  - Generator gốc. yield user_query event. Khởi tạo agent rồi loop
  qua các update từ langgraph.

  8. _build_agent(token) — agent/agent.py:18
  - Tạo AzureChatOpenAI (LLM), wrap 3 tools với closure giữ token,
  gọi create_agent (langgraph) với SYSTEM_PROMPT.

  9. make_llm() — agent/llm.py:6
  - Đọc 4 biến AZURE_OPENAI_* từ env, trả về AzureChatOpenAI đã
  config.

  10. make_graph_search_tool(token) / make_file_fetch_tool(token) /
  make_grep_context_tool() — agent/tools/*.py
  - Factory pattern: tạo 3 tool LangChain với access_token đã đóng
  băng trong closure.

  11. agent_graph.stream({"messages": [...]}, stream_mode="updates")
  — agent/agent.py:79
  - Langgraph chạy agent loop. yield mỗi delta state. Mỗi vòng có 2
  node:
    - agent node — gọi LLM, sinh AIMessage (có thể có tool_calls)
    - tools node — chạy tool được LLM yêu cầu, trả ToolMessage

  ---
  🔁 Tầng 3 — Vòng lặp ReAct (lặp nhiều lần)

  Mỗi vòng đi qua các bước:

  12. LLM sinh tool_call (chạy bên trong langgraph, không thấy code
  Python)
  - Đọc SYSTEM_PROMPT (agent/prompts.py:1) + history messages
  - Quyết định gọi tool nào với args nào (vd:
  graph_search(query="metric OR ..."))

  13. _parse_search_result(observation) — agent/agent.py:43
  - Khi tool_result trả về từ graph_search, parse text observation để
   extract [{item_id, name}, ...] — dùng cho summary cuối.

  ---
  Iteration A — gọi graph_search

  14. graph_search(query) — agent/tools/graph_search.py:11 (closure
  trong factory)
  - POST https://graph.microsoft.com/v1.0/search/query với:
    - entityTypes: ["driveItem"]
    - queryString: <LLM's OR-joined keywords>
    - size: 15, fields: id, name, webUrl, lastModifiedDateTime, size,
   parentReference
  - Parse data.value[0].hitsContainers[0].hits
  - Mỗi hit có resource.parentReference.driveId →
  drive_index.put(item_id, drive_id) (agent/drive_index.py:13) lưu
  mapping để fetch sau biết drive nào
  - Trả về string format: - item_id=... | name=... | snippet="..." |
  webUrl=...

  ---
  Iteration B — gọi fetch_file_text (nếu LLM muốn đọc sâu)

  15. fetch_file_text(item_id) — agent/tools/file_fetch.py:30
  - cache.get(item_id) (agent/cache.py:8) — check đã download chưa
  - _base_endpoint(item_id) — agent/tools/file_fetch.py:14
    - Tra drive_index.get(item_id) → có driveId thì dùng
  /drives/{driveId}/items/{id} (file shared/SharePoint/Teams), không
  thì fallback /me/drive/items/{id}
  - GET metadata ($select=name,size,file) → check size ≤ 25MB + ext ∈
   SUPPORTED
  - GET /content → binary bytes

  16. extract_text(name, bytes) — agent/extractors/__init__.py:14
  - Dispatcher theo extension:
    - .pdf → extract_pdf (pdf.py:5) — pdfplumber, prepend [PAGE N]
    - .docx → extract_docx (docx.py:6) — python-docx, paragraphs +
  tables
    - .xlsx → extract_xlsx (xlsx.py:6) — openpyxl, [SHEET name] +
  cells
    - .pptx → extract_pptx (pptx.py:6) — python-pptx, [SLIDE N] +
  notes
    - text formats → decode UTF-8

  17. cache.put(item_id, name, text) — agent/cache.py:12
  - Lưu text vào dict in-memory keyed by item_id, để các grep sau
  khỏi fetch lại.

  ---
  Iteration C — gọi grep_context (sau khi đã có text)

  18. grep_context(item_id, patterns, context_lines) —
  agent/tools/grep_context.py:43
  - Lấy text từ cache.get(item_id) (không re-fetch!)
  - re.compile(patterns, IGNORECASE) — compile regex
  (metric|BLEU|ROUGE|...)
  - Loop qua từng dòng → regex.search(line) → thu thập match_indices
  - Merge overlapping windows ±N lines (algorithm trong code: nếu
  start ≤ last_end thì extend, không thì tạo window mới)
  - _write_debug_log(...) — agent/tools/grep_context.py:18
    - Append vào debug/grep.log toàn bộ context đã extract — phục vụ
  debug
  - Format output: >> L42: hit line và    L41: context line

  ---
  🎯 Tầng 4 — Kết thúc

  19. LLM sinh Final Answer (sau khi đủ thông tin)
  - AIMessage không có tool_calls nữa → là answer cuối kèm citation
  [filename, line N]

  20. Summary — agent/agent.py:177
  - yield event summary chứa: queries đã search, files trả về, files
  đã fetch, files đã grep.
  - In ra console log [SUMMARY] block.

  21. yield done — app.py:653
  - Trong generate() của Flask, sau khi stream xong yield {"event":
  "done"} để JS biết kết thúc.

  ---
  🖼️  Tầng 5 — Browser nhận và render

  22. handleEvent(container, ev) — app.py:990
  - Switch theo ev.event:
    - user_query → setStatus "Đang tìm kiếm..."
    - tool_call → addStep (app.py:868) — tạo step card mới
    - tool_result → fillStepResult (app.py:885) — điền observation
  vào step card
    - summary → renderSummary (app.py:935) — render box màu vàng tổng
   kết
    - final_answer → setAnswer (app.py:902) — render câu trả lời cuối
    - error → setError (app.py:907)
    - done → ẩn typing indicator

  ---
  Sơ đồ tóm tắt

  [Browser]
    form.submit → fetch /ask
         │
         ▼
  [Flask]
    ask() → run_agent_stream()
                │
                ├─→ _build_agent(token)
                │       ├─→ make_llm()
  (AzureChatOpenAI)
                │       ├─→ make_graph_search_tool(token)
                │       ├─→ make_file_fetch_tool(token)
                │       ├─→ make_grep_context_tool()
                │       └─→ create_agent(llm, tools, SYSTEM_PROMPT)
                │
                ▼
           agent_graph.stream(query)   ──────────────┐
                │                                    │
                │   ┌────────────────────────────────┘
                │   │ Loop (langgraph nội bộ):
                │   │
                │   ▼
                │  LLM agent node sinh tool_call
                │   │
                │   ▼
                │  Tools node chạy 1 trong 3:
                │   ├─ graph_search()  → POST /search/query →
  drive_index.put()
                │   ├─ fetch_file_text() → drive_index.get()
                │   │                  → cache.get() (kiểm) → GET
  content
                │   │                  → extract_text() → cache.put()
                │   └─ grep_context() → cache.get() → re.compile()
                │                     → merge windows →
  _write_debug_log()
                │   │
                │   ▼
                │  yield event qua SSE (tool_call, tool_result)
                │   │
                │   ▼ (LLM thấy đủ → no tool_calls)
                │  yield final_answer
                │
                ▼
           yield summary + done
                │
                ▼
  [Flask] data: {...}\n\n
                │
                ▼
  [Browser] consumeStream() → handleEvent() → render UI live

  Cache + index — giữa các vòng

  Lưu trữ: cache._TEXT_CACHE
  File: agent/cache.py
  Mục đích: item_id → {filename, text} để grep nhiều lần không phải
    fetch lại
  ────────────────────────────────────────
  Lưu trữ: drive_index._INDEX
  File: agent/drive_index.py
  Mục đích: item_id → drive_id để fetch dùng đúng
    /drives/{driveId}/...
  ────────────────────────────────────────
  Lưu trữ: debug/grep.log
  File: (file disk)
  Mục đích: Append-only log mỗi lần grep, debug ngoài tiến trình