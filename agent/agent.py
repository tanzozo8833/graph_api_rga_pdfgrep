import json
import re
import traceback
from datetime import datetime
from typing import Iterator

from langchain.agents import create_agent
from langchain_core.messages import AIMessage, ToolMessage

from .llm import make_llm
from .prompts import SYSTEM_PROMPT
from .tools import (
    make_file_fetch_tool,
    make_graph_search_tool,
    make_grep_context_tool,
)


def _build_agent(access_token: str):
    llm = make_llm()
    tools = [
        make_graph_search_tool(access_token),
        make_file_fetch_tool(access_token),
        make_grep_context_tool(),
    ]
    return create_agent(model=llm, tools=tools, system_prompt=SYSTEM_PROMPT)


def _ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _stringify(value) -> str:
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, default=str)
    return str(value)


_FILE_LINE_RE = re.compile(
    r"-\s*item_id=(?P<id>[^|]+?)\s*\|\s*name=(?P<name>[^|]+?)\s*\|"
)


def _parse_search_result(observation: str) -> list[dict]:
    """Extract (item_id, name) pairs from a graph_search observation string."""
    files = []
    for m in _FILE_LINE_RE.finditer(observation):
        files.append({"item_id": m.group("id").strip(), "name": m.group("name").strip()})
    return files


def run_agent_stream(query: str, access_token: str) -> Iterator[dict]:
    """Yield event dicts as the agent runs. Used by the SSE endpoint.

    Event shapes:
        {"event": "user_query", "query": str, "ts": str}
        {"event": "tool_call", "step": int, "tool": str, "input": str, "ts": str}
        {"event": "tool_result", "step": int, "tool": str, "observation": str, "ts": str}
        {"event": "final_answer", "answer": str, "ts": str}
        {"event": "summary", "queries": list, "files_returned": list, "files_fetched": list, "ts": str}
        {"event": "error", "message": str}
    """
    print(f"\n{'=' * 70}\n[{_ts()}] USER QUERY: {query}\n{'=' * 70}")
    yield {"event": "user_query", "query": query, "ts": _ts()}

    try:
        agent_graph = _build_agent(access_token)
    except Exception as e:
        traceback.print_exc()
        yield {"event": "error", "message": f"{type(e).__name__}: {e}"}
        return

    # Track for summary
    queries_sent: list[str] = []
    files_returned: list[dict] = []  # [{query, name, item_id}, ...]
    files_fetched: list[dict] = []   # [{item_id, name (if known)}, ...]
    files_grepped: list[dict] = []   # [{item_id, patterns}, ...]
    last_search_query: str | None = None
    last_fetched_id: str | None = None

    step_no = 0
    try:
        for chunk in agent_graph.stream(
            {"messages": [{"role": "user", "content": query}]},
            stream_mode="updates",
            config={"recursion_limit": 30},
        ):
            for _node_name, node_state in chunk.items():
                if not isinstance(node_state, dict):
                    continue
                for msg in node_state.get("messages", []) or []:
                    if isinstance(msg, AIMessage):
                        if msg.tool_calls:
                            for tc in msg.tool_calls:
                                step_no += 1
                                tool_name = tc.get("name", "?")
                                tool_args = tc.get("args", {}) or {}
                                args_str = _stringify(tool_args)

                                # Track LLM's keyword decisions
                                if tool_name == "graph_search":
                                    q = tool_args.get("query", "")
                                    queries_sent.append(q)
                                    last_search_query = q
                                    print(
                                        f"\n[{_ts()}] [STEP {step_no}] >>> LLM CHOSE SEARCH"
                                        f"\n          keywords: {q!r}"
                                    )
                                elif tool_name == "fetch_file_text":
                                    fid = tool_args.get("item_id", "")
                                    last_fetched_id = fid
                                    name = next(
                                        (
                                            f["name"]
                                            for f in files_returned
                                            if f["item_id"] == fid
                                        ),
                                        "(unknown)",
                                    )
                                    files_fetched.append({"item_id": fid, "name": name})
                                    print(
                                        f"\n[{_ts()}] [STEP {step_no}] >>> LLM CHOSE FETCH FILE"
                                        f"\n          file: {name} (id={fid[:20]}...)"
                                    )
                                elif tool_name == "grep_context":
                                    fid = tool_args.get("item_id", "")
                                    pat = tool_args.get("patterns", "")
                                    name = next(
                                        (
                                            f["name"]
                                            for f in files_fetched
                                            if f["item_id"] == fid
                                        ),
                                        "(unknown)",
                                    )
                                    files_grepped.append(
                                        {"item_id": fid, "name": name, "patterns": pat}
                                    )
                                    print(
                                        f"\n[{_ts()}] [STEP {step_no}] >>> LLM CHOSE GREP"
                                        f"\n          file: {name}"
                                        f"\n          patterns: {pat!r}"
                                    )

                                yield {
                                    "event": "tool_call",
                                    "step": step_no,
                                    "tool": tool_name,
                                    "input": args_str,
                                    "ts": _ts(),
                                }
                        else:
                            text = (
                                msg.content
                                if isinstance(msg.content, str)
                                else _stringify(msg.content)
                            )
                            print(f"\n[{_ts()}] [FINAL ANSWER]\n{text[:500]}")
                            yield {
                                "event": "final_answer",
                                "answer": text,
                                "ts": _ts(),
                            }
                    elif isinstance(msg, ToolMessage):
                        obs = (
                            msg.content
                            if isinstance(msg.content, str)
                            else _stringify(msg.content)
                        )
                        tool_name = getattr(msg, "name", "?")

                        # Parse graph_search results to track file list
                        if tool_name == "graph_search" and last_search_query is not None:
                            parsed = _parse_search_result(obs)
                            for f in parsed:
                                files_returned.append(
                                    {
                                        "query": last_search_query,
                                        "name": f["name"],
                                        "item_id": f["item_id"],
                                    }
                                )
                            print(
                                f"[{_ts()}] [STEP {step_no}] <<< Graph returned {len(parsed)} file(s):"
                            )
                            for i, f in enumerate(parsed[:10], 1):
                                print(f"          {i}. {f['name']}")
                            if len(parsed) > 10:
                                print(f"          ... and {len(parsed) - 10} more file(s)")
                            last_search_query = None
                        else:
                            print(
                                f"[{_ts()}] [STEP {step_no}] <<< {tool_name} result "
                                f"({len(obs)} chars)"
                            )

                        yield {
                            "event": "tool_result",
                            "step": step_no,
                            "tool": tool_name,
                            "observation": obs[:3000],
                            "ts": _ts(),
                        }
    except Exception as e:
        traceback.print_exc()
        yield {"event": "error", "message": f"{type(e).__name__}: {e}"}
        return

    # Summary event
    print(f"\n{'-' * 70}")
    print(f"[{_ts()}] [SUMMARY]")
    print(f"  Search calls: {len(queries_sent)}")
    for i, q in enumerate(queries_sent, 1):
        print(f"    {i}. {q!r}")
    print(f"  Total files returned by Graph (merged): {len(files_returned)}")
    print(f"  Files fetched by LLM (fetch_file_text): {len(files_fetched)}")
    for f in files_fetched:
        print(f"    - {f['name']}")
    print(f"  Files grepped by LLM: {len(files_grepped)}")
    for f in files_grepped:
        print(f"    - {f['name']} ← {f['patterns']!r}")
    print(f"{'-' * 70}\n")

    yield {
        "event": "summary",
        "queries": queries_sent,
        "files_returned": files_returned,
        "files_fetched": files_fetched,
        "files_grepped": files_grepped,
        "ts": _ts(),
    }


def run_agent(query: str, access_token: str) -> dict:
    """Non-streaming wrapper. Drains run_agent_stream and returns {answer, steps}."""
    answer = ""
    steps: list[dict] = []
    current: dict | None = None

    for ev in run_agent_stream(query, access_token):
        et = ev.get("event")
        if et == "tool_call":
            current = {
                "tool": ev["tool"],
                "tool_input": ev["input"],
                "observation": "",
            }
            steps.append(current)
        elif et == "tool_result":
            if current is not None:
                current["observation"] = ev["observation"]
        elif et == "final_answer":
            answer = ev["answer"]
        elif et == "error":
            answer = "ERROR: " + ev["message"]

    return {"answer": answer, "steps": steps}
