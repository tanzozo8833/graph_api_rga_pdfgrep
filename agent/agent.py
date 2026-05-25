import json
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


def run_agent_stream(query: str, access_token: str) -> Iterator[dict]:
    """Yield event dicts as the agent runs. Used by the SSE endpoint.

    Event shapes:
        {"event": "user_query", "query": str, "ts": str}
        {"event": "tool_call", "step": int, "tool": str, "input": str, "ts": str}
        {"event": "tool_result", "step": int, "tool": str, "observation": str, "ts": str}
        {"event": "final_answer", "answer": str, "ts": str}
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
                                tool_args = tc.get("args", {})
                                args_str = _stringify(tool_args)
                                print(
                                    f"[{_ts()}] [STEP {step_no}] -> {tool_name}({args_str[:150]})"
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
                            print(f"[{_ts()}] [FINAL ANSWER]\n{text[:500]}")
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
                        print(
                            f"[{_ts()}] [STEP {step_no}] <- {tool_name} result "
                            f"({len(obs)} chars): {obs[:200].replace(chr(10), ' ')}..."
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
