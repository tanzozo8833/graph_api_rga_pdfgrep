 └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ STEP 1: Flask route /ask                                             │
  │   - Pull access_token from session (already present from MSAL OAuth) │
  │   - Validate query is not empty                                      │
  │   - Return Response(generate(), mimetype="text/event-stream")        │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ STEP 2: agent.agent.run_agent_stream(query, token)                   │
  │   yield {"event": "user_query", "query": "...", "ts": "..."}         │
  │   |                                                                  │
  │   Build agent:                                                       │
  │     llm = AzureChatOpenAI(deployment=gpt-4o)                         │
  │     tools = [                                                        │
  │       make_graph_search_tool(token),  # closure captures token       │
  │       make_file_fetch_tool(token),                                   │
  │       make_grep_context_tool(),                                      │
  │     ]                                                                │
  │     agent = create_agent(llm, tools, system_prompt=SYSTEM_PROMPT)    │
  │   |                                                                  │
  │   agent.stream({"messages": [HumanMessage(query)]}, mode="updates")  │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ STEP 3 - Iteration 1 (LLM decision)                                  │
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
  │ STEP 4 - graph_search execution                                      │
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
  │   Format into a string returned as tool result:                      │
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
  │ STEP 5 - LLM looks at snippets and decides Iter 2                    │
  │                                                                      │
  │   Case A - snippets are enough to answer:                            │
  │     LLM -> AIMessage(content="The metrics for RAG include BLEU,      │
  │                    ROUGE-L, Recall@k... [rag_paper.pdf,              │
  │                    eval_methods.docx]")                              │
  │     yield {"event": "final_answer", "answer": "..."}                 │
  │     DONE.                                                            │
  │                                                                      │
  │   Case B - needs deeper reading (hypothetical):                      │
  │     AIMessage tool_calls: [                                          │
  │       {name: "fetch_file_text", args: {item_id: "01ABC..."}},        │
  │       {name: "fetch_file_text", args: {item_id: "01XYZ..."}}         │
  │     ]                                                                │
  │     (LLM may call in parallel - gpt-4o supports it)                  │
  └────────────────────┬─────────────────────────────────────────────────┘
                       ▼
  ┌──────────────────────────────────────────────────────────────────────┐
  │ STEP 6 - fetch_file_text execution                                   │
  │                                                                      │
  │   Step 6a: GET /me/drive/items/01ABC.../?$select=name,size,file      │
  │            -> meta: name="rag_paper.pdf", size=1.2 MB                │
  │            Check size < 25 MB -> OK                                  │
  │            Check ext (.pdf) in SUPPORTED -> OK                       │
  │                                                                      │
  │   Step 6b: GET /me/drive/items/01ABC.../content                      │
  │            -> binary content (PDF bytes)                             │
  │                                                                      │
  │   Step 6c: extract_text("rag_paper.pdf", bytes)                      │
  │            -> dispatch to pdf.py                                     │
  │            -> pdfplumber.open() -> iter pages                        │
  │            -> return "[PAGE 1]\nText...\n[PAGE 2]\nText..."          │
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
  │ STEP 7 - LLM decides to grep                                         │
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
  │ STEP 8 - grep_context execution                                      │
  │                                                                      │
  │   text = cache.get("01ABC...")["text"]                               │
  │   lines = text.split("\n")                                           │
  │   regex = re.compile(patterns, re.IGNORECASE)                        │
  │   match_indices = [i for i,l in enumerate(lines) if regex.search(l)] │
  │                                                                      │
  │   Merge overlapping windows +/-5:                                    │
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
  │ STEP 9 - LLM composes the final answer                               │
  │                                                                      │
  │   AIMessage (no tool_calls):                                         │
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
  │ STEP 10 - Frontend receives events                                   │
  │                                                                      │
  │   JS consumeStream() loop:                                           │
  │     reader.read() -> decoder.decode()                                │
  │     split('\n\n') -> each "data: {...}" -> JSON.parse -> handleEvent │
  │                                                                      │
  │   handleEvent:                                                       │
  │     - tool_call -> addStep() - show step card with status "..."      │
  │     - tool_result -> fillStepResult() - fill observation             │
  │     - final_answer -> setAnswer() - render below trace               │
  │     - done -> hide typing indicator                                  │
  └──────────────────────────────────────────────────────────────────────┘

  Retry strategy when nothing is found

  The LLM detects empty results (irrelevant snippets, grep matched nothing) -> creates
  a new strategy:

  Iter 1: graph_search("metric OR evaluation")        -> 0 hits
  Iter 2: Thought: "Maybe different keywords in a Vietnamese corpus..."
          graph_search("danh gia OR thuoc do OR chi so") -> 3 hits
  Iter 3: graph_search("MRR OR NDCG OR hit rate")     -> 5 hits (alt names)
  Iter 4: continue with fetch + grep...

  SYSTEM_PROMPT requires trying at least 2-3 strategies before giving up.

  ---
  3. Detailed code structure

  14_Microsoft_GraphAPI/
  ├── app.py                          # Flask app - add POST /ask + chat panel
  ├── graph_client.py                 # legacy CLI client - leave alone
  ├── requirements.txt                # add langchain, langchain-openai, pdf/office
  libs
  ├── .env.example                    # add AZURE_OPENAI_* vars
  ├── debug/                          # already present
  │
  ├── agent/                          # NEW: full agent module
  │   ├── __init__.py                 # re-export run_agent
  │   │
  │   ├── llm.py                      # NEW: ChatModel factory
  │   │   def make_llm():
  │   │     validate 4 env vars
  │   │     return AzureChatOpenAI(
  │   │       azure_endpoint=..., api_key=..., api_version=...,
  │   │       azure_deployment=..., temperature=0, max_tokens=4096)
  │   │
  │   ├── prompts.py                  # NEW: SYSTEM_PROMPT
  │   │   SYSTEM_PROMPT = """
  │   │     You are an assistant ... DO NOT answer from your own knowledge.
  │   │     There are 3 tools: graph_search, fetch_file_text, grep_context.
  │   │     REQUIRED WORKFLOW:
  │   │       Step 1: ALWAYS call graph_search FIRST with OR keywords.
  │   │       Step 2: Read snippets - enough -> Final Answer; not enough -> Step 3.
  │   │       Step 3: fetch_file_text on the top 1-3 files.
  │   │       Step 4: grep_context with multi-pattern regex.
  │   │       Step 5: if empty -> RESTART with a different strategy.
  │   │     RULES:
  │   │       - ALWAYS multi-pattern OR, NEVER single word
  │   │       - ALWAYS cite [filename, line N]
  │   │       - DO NOT invent; if nothing found -> "I could not find this"
  │   │   """
  │   │
  │   ├── cache.py                    # NEW: in-memory text cache
  │   │   _TEXT_CACHE: dict[item_id, {filename, text}]
  │   │   def get(item_id), put(item_id, name, text), clear()
  │   │
  │   ├── agent.py                    # NEW: bootstrap + run loop
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
  │   │     drain stream -> return {answer, steps}
  │   │
  │   ├── tools/
  │   │   ├── __init__.py             # re-export the 3 factories
  │   │   │
  │   │   ├── graph_search.py         # NEW: Tool 1
  │   │   │   def make_graph_search_tool(token):
  │   │   │     @tool
  │   │   │     def graph_search(query: str) -> str:
  │   │   │       """Docstring (the LLM reads it!) explains OR-keyword usage"""
  │   │   │       POST /search/query with entityTypes=['driveItem']
  │   │   │       parse data.value[0].hitsContainers[0].hits
  │   │   │       format "- item_id=... | name=... | snippet=... | webUrl=..."
  │   │   │     return graph_search
  │   │   │
  │   │   ├── file_fetch.py           # NEW: Tool 2
  │   │   │   MAX_SIZE = 25 MB
  │   │   │   def make_file_fetch_tool(token):
  │   │   │     @tool
  │   │   │     def fetch_file_text(item_id: str) -> str:
  │   │   │       """Docstring: warn against calling twice with the same id"""
  │   │   │       if cache.get(item_id): return "ALREADY CACHED"
  │   │   │       GET metadata -> check size + ext
  │   │   │       GET content -> extract_text(name, bytes)
  │   │   │       cache.put(...)
  │   │   │       return "Fetched. filename=..., text_length=..., Preview: ..."
  │   │   │     return fetch_file_text
  │   │   │
  │   │   └── grep_context.py         # NEW: Tool 3
  │   │       MAX_CHUNKS = 10
  │   │       MAX_CHARS_PER_LINE = 400
  │   │       def make_grep_context_tool():
  │   │         @tool
  │   │         def grep_context(item_id, patterns, context_lines=5) -> str:
  │   │           """Docstring: explains how to use | OR regex"""
  │   │           text = cache.get(item_id)
  │   │           regex = re.compile(patterns, re.IGNORECASE)
  │   │           match_indices = ...
  │   │           merge overlapping +/-N windows
  │   │           format ">> L{n}: line" for hits, "   L{n}: line" for context
  │   │           cap MAX_CHUNKS
  │   │         return grep_context
  │   │
  │   └── extractors/
  │       ├── __init__.py             # NEW: dispatcher
  │       │   SUPPORTED = {.pdf, .docx, .xlsx, .pptx,
  │       │                .txt, .md, .csv, .tsv, .log, .json, .xml, .yaml, .yml}
  │       │   def extract_text(filename, bytes):
  │       │     ext = os.path.splitext(filename)[1].lower()
  │       │     if ext == ".pdf": return extract_pdf(bytes)
  │       │     ...
  │       │     if ext in PLAINTEXT_EXTS: return bytes.decode("utf-8")
  │       │
  │       ├── pdf.py                  # NEW: pdfplumber.open(BytesIO) -> iter pages,
  │       │                           #      prepend "[PAGE N]" markers
  │       │
  │       ├── docx.py                 # NEW: Document(BytesIO) -> para.text + tables
  │       │                           #      tables: "[TABLE n]" + cell1 | cell2 | ...
  │       │
  │       ├── xlsx.py                 # NEW: load_workbook(read_only=True,
  │       │                           #      data_only=True) - skip formulas
  │       │                           #      "[SHEET name]" + cell1 | cell2 | ...
  │       │                           #      skip blank rows
  │       │
  │       └── pptx.py                 # NEW: Presentation(BytesIO) -> iter slides
  │                                   #      "[SLIDE n]" + shape.text_frame.text
  │                                   #      + "[NOTES] ..." if notes_slide exists

  ---
  4. Detailed implementation phases

  Phase 0 - Setup (30 minutes)

  Step 0.1. Append to requirements.txt:
  langchain>=0.3.0
  langchain-openai>=0.2.0   # (initially langchain-anthropic, later switched)
  langchain-core>=0.3.0
  pdfplumber>=0.11.0
  python-docx>=1.1.0
  openpyxl>=3.1.0
  python-pptx>=1.0.0

  Step 0.2. Append to .env.example:
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
  Phase 1 - Text extractors (1 hour)

  Design principles:
  - Output is plain text, line-oriented (split by \n) so grep -C N works
  - Each unit (page/slide/sheet/table) is prefixed with a marker [PAGE 3], [SLIDE 2],
  [SHEET Data], [TABLE 1] so the LLM can cite accurately
  - Input: bytes, output: str

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
  - paragraphs -> one line each
  - tables -> [TABLE n] then for each row: cell1 | cell2 | cell3
  - skip empty paragraphs

  1.3. extractors/xlsx.py:
  - read_only=True, data_only=True - important for handling large files and reading
  computed values (not formulas)
  - per sheet: [SHEET <name>] then for each non-empty row: cells joined by |
  - replace \n inside a cell with a space to keep the output line-oriented

  1.4. extractors/pptx.py:
  - iter slides, marker [SLIDE n]
  - iter shapes that have text_frame, extract paragraph runs
  - notes slide: [NOTES] ...

  1.5. extractors/__init__.py - dispatcher:
  SUPPORTED = OFFICE_EXTS | PLAINTEXT_EXTS

  def extract_text(filename: str, content: bytes) -> str:
      ext = os.path.splitext(filename)[1].lower()
      # dispatch...
      if ext in PLAINTEXT_EXTS:
          try: return content.decode("utf-8")
          except UnicodeDecodeError: return content.decode("latin-1", errors="replace")
      raise ValueError(f"Unsupported: {ext}")

  ---
  Phase 2 - 3 LangChain Tools (2 hours)

  Design principles:
  - Each tool has a factory make_xxx_tool(token) - the closure holds the access_token
  - Tools are defined via the @tool decorator -> LangChain auto-generates the schema
  from type hints + docstring
  - The docstring is critical - the LLM reads it to learn when to use the tool and how
  to call it
  - Always return str (do not return dict/list, to avoid formatting issues)
  - On error -> return a string "ERROR: ..." instead of raising an exception (so the
  agent can handle it)

  2.1. tools/graph_search.py:

  Endpoint: POST https://graph.microsoft.com/v1.0/search/query

  Detailed request body:
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
      "summary": "...highlighted snippet with <ddd:c0></ddd:c0> markup...",
      "resource": {
        "@odata.type": "#microsoft.graph.driveItem",
        "id": "01ABC...",
        "name": "rag_paper.pdf",
        "webUrl": "https://...",
        ...
      }
    }
  ]

  Tool output format (the string the LLM reads):
  Found 8 files for query 'metric OR evaluation':
  - item_id=01ABC... | name=rag_paper.pdf | snippet="...we use BLEU and ROUGE..." |
  webUrl=https://...
  - item_id=01XYZ... | name=eval.docx | snippet="..." | webUrl=https://...

  2.2. tools/file_fetch.py:

  3 Graph API endpoints:
  - GET /me/drive/items/{id} with $select=name,size,file -> metadata
  - GET /me/drive/items/{id}/content -> binary

  Sequential checks:
  1. Already cached? -> return "ALREADY CACHED"
  2. Get metadata -> 401 -> return error
  3. size <= 25 MB? -> too large -> return error
  4. ext in SUPPORTED? -> if not -> return error
  5. Download content
  6. extract_text(name, bytes) - wrap in try/except (bad PDF, corrupted DOCX...)
  7. cache.put(...)
  8. Return preview text of 300 chars

  Return format:
  Fetched and cached. filename=rag_paper.pdf, size=1234567 bytes, text_length=45000
  chars.
  Preview: [PAGE 1] Abstract Retrieval-Augmented Generation (RAG) systems combine...

  2.3. tools/grep_context.py:

  Detailed algorithm:
  def grep_context(item_id, patterns, context_lines=5):
      text = cache.get(item_id)["text"]
      lines = text.split("\n")

      regex = re.compile(patterns, re.IGNORECASE)

      # Find match indices
      match_indices = [i for i, line in enumerate(lines) if regex.search(line)]
      if not match_indices:
          return f"No matches for pattern '{patterns}'"

      # Merge overlapping windows
      windows = []
      for idx in match_indices:
          start = max(0, idx - context_lines)
          end = min(len(lines), idx + context_lines + 1)
          if windows and start <= windows[-1][1]:
              # overlap -> extend last window
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

  Output caps:
  - Max 10 chunks (if more -> "+ N more chunks omitted")
  - Max 400 chars/line (if longer -> "...[truncated]")
  - -> avoid blowing the LLM context

  ---
  Phase 3 - Prompts + Agent (1 hour)

  3.1. prompts.py - the full SYSTEM_PROMPT:

  Based on cell 10 of the original notebook, adapted for the Microsoft Graph context:

  Mandatory parts of the prompt:
  1. Forbid answering from internal knowledge - only from tool results
  2. Describe the 3 tools with signature + examples
  3. The required 5-step workflow - graph_search first, fetch if needed, grep
  multi-pattern, retry if empty
  4. Hard rules:
    - ALWAYS multi-pattern OR, NEVER single word
    - ALWAYS cite [filename, line N] or [filename, PAGE N]
    - DO NOT invent
    - Cache persists -> do not re-fetch
    - Try 2-3 strategies before giving up

  3.2. agent.py - bootstrap with langchain 1.x:

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

  Streaming with stream_mode="updates":
  - Each event is {node_name: {messages: [delta_messages]}}
  - Node "agent" -> AIMessage (may contain tool_calls)
  - Node "tools" -> ToolMessage(s)

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

  Console logging (for dev visibility):
  [08:35:12] USER QUERY: ...
  [08:35:13] [STEP 1] -> graph_search({"query": "..."})
  [08:35:14] [STEP 1] <- graph_search result (1240 chars): Found 8 files...
  [08:35:16] [STEP 2] -> fetch_file_text({"item_id": "01ABC..."})
  ...
  [08:35:22] [FINAL ANSWER]
  The metrics for RAG include...

  ---
  Phase 4 - Flask integration (45 minutes)

  4.1. Route /ask with SSE:

  @app.route("/ask", methods=["POST"])
  def ask():
      token = _get_token()  # from the existing session
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
              "X-Accel-Buffering": "no",   # disable nginx buffering when deployed
              "Connection": "keep-alive",
          },
      )

  4.2. Chat panel HTML inside DASHBOARD_HTML:

  After the .cards div, add a <section class="chat-section">:
  - Header with title "Ask your OneDrive"
  - #chat-messages area with height=380px scroll
  - Form input + send button

  4.3. Detailed CSS:
  - .msg.user .bubble - blue, right-aligned
  - .msg.agent .bubble - white with gray border, left-aligned
  - .live-step - light gray box with blue border
    - .step-no - chip showing step number
    - .step-ts - small timestamp
    - .tname - tool name in blue
    - .tin - input in gray
    - .step-obs pre - observation, max-height 220px, scrollable
  - .typing - 3-dot animation
  - .agent-answer - answer text, padding-top + dashed border

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
      case 'user_query':    setStatus("Searching..."); break;
      case 'tool_call':     setStatus("Step "+ev.step+": "+ev.tool);
                            addStep(container, ev); break;
      case 'tool_result':   fillStepResult(container, ev); break;
      case 'final_answer':  hideTyping(); setAnswer(container, ev.answer); break;
      case 'error':         setError(container, ev.message); break;
      case 'done':          hideTyping(); break;
    }
  }

  ---
  5. Detailed warnings + workarounds

  Issue: Permission Files.Read.All
  Details: /search/query requires a broader scope than /me/drive/root/search
  Workaround: Already in the app's SCOPES - OK
  ────────────────────────────────────────
  Issue: Scanned PDF (image-only)
  Details: pdfplumber.extract_text() returns "" or very little
  Workaround: Accepted in this phase; phase 6 will add Docling/Azure Document
  Intelligence
  ────────────────────────────────────────
  Issue: Files >25 MB
  Details: Memory + timeout
  Workaround: MAX_SIZE_BYTES = 25 * 1024 * 1024, reject with an error message
  ────────────────────────────────────────
  Issue: Token expired (1 hour)
  Details: Graph API returns 401 mid-conversation
  Workaround: Tool returns "ERROR: HTTP 401" -> LLM understands re-login is needed;
  user must /login again
  ────────────────────────────────────────
  Issue: Cost per question
  Details: 5-15 LLM calls x gpt-4o input tokens x output tokens
  Workaround: Caching text reduces tokens; cap MAX_CHUNKS=10; preview-only return from
  fetch
  ────────────────────────────────────────
  Issue: Search index lag
  Details: OneDrive indexes new files after a few minutes
  Workaround: Out of our control; notify the user
  ────────────────────────────────────────
  Issue: Privacy
  Details: Text is sent to Azure OpenAI
  Workaround: Azure OpenAI region = MS data residency - OK for enterprise
  ────────────────────────────────────────
  Issue: LangChain 1.0 breaking changes
  Details: AgentExecutor -> create_agent (langgraph)
  Workaround: Already handled
  ────────────────────────────────────────
  Issue: Multi-thread Flask
  Details: Each request builds a new agent and shares the _TEXT_CACHE global
  Workaround: OK for dev; production needs per-user cache keyed by session id
  ────────────────────────────────────────
  Issue: Streaming buffering
  Details: Werkzeug dev server streams fine; nginx in production needs
  X-Accel-Buffering: no
  Workaround: Header already added

  ---
  6. Time summary

  ┌─────────────────┬──────────────────────────────────┬───────────┐
  │      Phase      │               Step               │   Time    │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 0               │ Setup + LLM smoke test           │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 1               │ 4 extractors + dispatcher        │ 1h        │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 2.1             │ graph_search tool                │ 40m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 2.2             │ file_fetch tool                  │ 40m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 2.3             │ grep_context tool                │ 40m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 3.1             │ Detailed SYSTEM_PROMPT           │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 3.2             │ agent bootstrap + stream parsing │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 4.1             │ Flask /ask SSE route             │ 15m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 4.2-4.4         │ Chat panel HTML/CSS/JS           │ 30m       │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ 5               │ Testing (you skip this)          │ 1-2h      │
  ├─────────────────┼──────────────────────────────────┼───────────┤
  │ Total (coded)   │ ~6 hours                         │           │
  └─────────────────┴──────────────────────────────────┴───────────┘

  ---
  7. Current status

  100% complete from Phase 0 through Phase 4:
  - All 4 extractors emit [PAGE/SHEET/SLIDE/TABLE N] markers
  - 3 tools using the factory pattern + error handling
  - SYSTEM_PROMPT with full workflow + rules
  - Agent uses create_agent (langchain 1.x langgraph) + stream_mode="updates"
  - Console logging with [STEP N] timestamp markers
  - SSE streaming on the /ask route
  - Chat panel with live step cards, realtime status updates, inline errors
  - Switched from Anthropic to Azure OpenAI per the .env.example you configured
  - Fixed the ImportError on langchain 1.3.1

  Still pending from you:
  1. Fill in the 4 real AZURE_OPENAI_* variables in .env
  2. The deployment must be a model that supports function/tool calling (gpt-4o,
  gpt-4.1, gpt-4-turbo)
  3. Run venv\Scripts\python.exe app.py -> test



  Tier 1 - Browser (JS)

  1. form.addEventListener('submit', ...) - app.py:962
  - Listener attached to the chat form on page load. Blocks the default submit
  (e.preventDefault()) so the page does not reload.

  2. appendUser(text) - app.py:843
  - Creates a blue bubble showing the user's question as soon as it is sent.

  3. createAgentContainer() - app.py:851
  - Creates an empty agent bubble with 3 regions: a status line (typing indicator),
  a steps container, and a final answer area.

  4. fetch('/ask', { method:'POST', body:{query} }) - app.py:973
  - Sends HTTP POST with the query JSON, Accept: text/event-stream to open the SSE
  channel.

  5. consumeStream(resp, container) - app.py:913
  - Reads chunks from the response body. Splits by \n\n. Parses each data: {...}
  into an event object. Calls handleEvent.

  ---
  Tier 2 - Flask backend (Python)

  6. ask() - app.py:643 (route /ask)
  - Pulls access_token from the Flask session (already present via MSAL OAuth).
  - Validates the query is not empty.
  - Opens an SSE generator; each event is yielded as data: {...}\n\n.

  7. run_agent_stream(query, token) - agent/agent.py:53
  - The root generator. Yields the user_query event. Builds the agent then loops
  through updates from langgraph.

  8. _build_agent(token) - agent/agent.py:18
  - Creates AzureChatOpenAI (LLM), wraps 3 tools with a closure holding the token,
  calls create_agent (langgraph) with SYSTEM_PROMPT.

  9. make_llm() - agent/llm.py:6
  - Reads the 4 AZURE_OPENAI_* env vars and returns the configured
  AzureChatOpenAI.

  10. make_graph_search_tool(token) / make_file_fetch_tool(token) /
  make_grep_context_tool() - agent/tools/*.py
  - Factory pattern: creates the 3 LangChain tools with the access_token frozen in
  the closure.

  11. agent_graph.stream({"messages": [...]}, stream_mode="updates")
  - agent/agent.py:79
  - Langgraph runs the agent loop. Yields each delta state. Each iteration has 2
  nodes:
    - agent node - calls the LLM, produces AIMessage (which may contain tool_calls)
    - tools node - runs the tool the LLM requested, returns ToolMessage

  ---
  Tier 3 - ReAct loop (iterates multiple times)

  Each iteration goes through these steps:

  12. The LLM emits a tool_call (this runs inside langgraph, no Python code visible)
  - Reads SYSTEM_PROMPT (agent/prompts.py:1) + the message history
  - Decides which tool to call and with which args (e.g.,
  graph_search(query="metric OR ..."))

  13. _parse_search_result(observation) - agent/agent.py:43
  - When tool_result comes back from graph_search, parses the text observation to
  extract [{item_id, name}, ...] - used for the final summary.

  ---
  Iteration A - calling graph_search

  14. graph_search(query) - agent/tools/graph_search.py:11 (closure inside the
  factory)
  - POST https://graph.microsoft.com/v1.0/search/query with:
    - entityTypes: ["driveItem"]
    - queryString: <LLM's OR-joined keywords>
    - size: 15, fields: id, name, webUrl, lastModifiedDateTime, size,
  parentReference
  - Parses data.value[0].hitsContainers[0].hits
  - Each hit has resource.parentReference.driveId ->
  drive_index.put(item_id, drive_id) (agent/drive_index.py:13) stores the mapping
  so a later fetch knows which drive to use
  - Returns a formatted string: - item_id=... | name=... | snippet="..." |
  webUrl=...

  ---
  Iteration B - calling fetch_file_text (if the LLM wants to read deeper)

  15. fetch_file_text(item_id) - agent/tools/file_fetch.py:30
  - cache.get(item_id) (agent/cache.py:8) - check if already downloaded
  - _base_endpoint(item_id) - agent/tools/file_fetch.py:14
    - Looks up drive_index.get(item_id) -> if a driveId is found, use
  /drives/{driveId}/items/{id} (shared/SharePoint/Teams files); otherwise fall
  back to /me/drive/items/{id}
  - GET metadata ($select=name,size,file) -> check size <= 25MB + ext in SUPPORTED
  - GET /content -> binary bytes

  16. extract_text(name, bytes) - agent/extractors/__init__.py:14
  - Dispatcher by extension:
    - .pdf -> extract_pdf (pdf.py:5) - pdfplumber, prepend [PAGE N]
    - .docx -> extract_docx (docx.py:6) - python-docx, paragraphs + tables
    - .xlsx -> extract_xlsx (xlsx.py:6) - openpyxl, [SHEET name] + cells
    - .pptx -> extract_pptx (pptx.py:6) - python-pptx, [SLIDE N] + notes
    - text formats -> decode UTF-8

  17. cache.put(item_id, name, text) - agent/cache.py:12
  - Stores the text in an in-memory dict keyed by item_id so subsequent greps do
  not need to refetch.

  ---
  Iteration C - calling grep_context (after text is available)

  18. grep_context(item_id, patterns, context_lines) -
  agent/tools/grep_context.py:43
  - Gets the text from cache.get(item_id) (no refetch!)
  - re.compile(patterns, IGNORECASE) - compiles the regex
  (metric|BLEU|ROUGE|...)
  - Loops over each line -> regex.search(line) -> collects match_indices
  - Merges overlapping windows +/-N lines (algorithm in code: if start <= last_end
  extend, otherwise create a new window)
  - _write_debug_log(...) - agent/tools/grep_context.py:18
    - Appends to debug/grep.log all the context extracted - serves debugging
  - Format output: >> L42: hit line and    L41: context line

  ---
  Tier 4 - Wrap-up

  19. The LLM produces the Final Answer (once it has enough info)
  - AIMessage with no more tool_calls -> the final answer with citations
  [filename, line N]

  20. Summary - agent/agent.py:177
  - Yields a summary event containing: queries searched, files returned, files
  fetched, files grepped.
  - Prints a [SUMMARY] block to the console log.

  21. yield done - app.py:653
  - Inside Flask's generate(), after streaming completes yield
  {"event": "done"} so the JS knows it has ended.

  ---
  Tier 5 - Browser receives and renders

  22. handleEvent(container, ev) - app.py:990
  - Switches on ev.event:
    - user_query -> setStatus "Searching..."
    - tool_call -> addStep (app.py:868) - creates a new step card
    - tool_result -> fillStepResult (app.py:885) - fills the observation into the
  step card
    - summary -> renderSummary (app.py:935) - renders a yellow summary box
    - final_answer -> setAnswer (app.py:902) - renders the final answer
    - error -> setError (app.py:907)
    - done -> hide the typing indicator

  ---
  Summary diagram

  [Browser]
    form.submit -> fetch /ask
         |
         v
  [Flask]
    ask() -> run_agent_stream()
                |
                +-> _build_agent(token)
                |       +-> make_llm()
  (AzureChatOpenAI)
                |       +-> make_graph_search_tool(token)
                |       +-> make_file_fetch_tool(token)
                |       +-> make_grep_context_tool()
                |       +-> create_agent(llm, tools, SYSTEM_PROMPT)
                |
                v
           agent_graph.stream(query)   --------------+
                |                                    |
                |   +--------------------------------+
                |   | Loop (internal to langgraph):
                |   |
                |   v
                |  LLM agent node emits tool_call
                |   |
                |   v
                |  Tools node runs 1 of 3:
                |   +- graph_search()  -> POST /search/query ->
  drive_index.put()
                |   +- fetch_file_text() -> drive_index.get()
                |   |                   -> cache.get() (check) -> GET
  content
                |   |                   -> extract_text() -> cache.put()
                |   +- grep_context() -> cache.get() -> re.compile()
                |                      -> merge windows ->
  _write_debug_log()
                |   |
                |   v
                |  yield event via SSE (tool_call, tool_result)
                |   |
                |   v (LLM is satisfied -> no more tool_calls)
                |  yield final_answer
                |
                v
           yield summary + done
                |
                v
  [Flask] data: {...}\n\n
                |
                v
  [Browser] consumeStream() -> handleEvent() -> render UI live

  Cache + index - shared across iterations

  Storage: cache._TEXT_CACHE
  File: agent/cache.py
  Purpose: item_id -> {filename, text} so repeated greps do not need to
    refetch
  ────────────────────────────────────────
  Storage: drive_index._INDEX
  File: agent/drive_index.py
  Purpose: item_id -> drive_id so fetches use the correct
    /drives/{driveId}/...
  ────────────────────────────────────────
  Storage: debug/grep.log
  File: (disk file)
  Purpose: Append-only log of every grep, for out-of-process debugging
