import re

from langchain_core.tools import tool

from .. import cache

MAX_CHUNKS_RETURNED = 10
MAX_CHARS_PER_LINE = 400


def make_grep_context_tool():
    @tool
    def grep_context(
        item_id: str, patterns: str, context_lines: int = 5
    ) -> str:
        """Regex-search inside a previously-fetched file and return matches with surrounding context.

        Mimics `grep -C N`. The file must have been fetched first via fetch_file_text.

        Args:
            item_id: the item_id from graph_search (also used in fetch_file_text)
            patterns: regex pattern. Use | for OR. Example: "metric|evaluation|F1|BLEU|ROUGE"
            context_lines: number of lines of context before/after each match (default 5)

        Returns matched line numbers + their context. Lines starting with ">>" are the hits.
        Overlapping windows are merged. Up to 10 chunks shown.
        """
        cached = cache.get(item_id)
        if not cached:
            return (
                f"ERROR: item_id '{item_id}' not in cache. "
                f"Call fetch_file_text(item_id) first."
            )

        text = cached["text"]
        filename = cached["filename"]
        lines = text.split("\n")

        try:
            regex = re.compile(patterns, re.IGNORECASE)
        except re.error as e:
            return f"ERROR: invalid regex '{patterns}': {e}"

        match_indices = [i for i, line in enumerate(lines) if regex.search(line)]
        if not match_indices:
            return f"No matches for pattern '{patterns}' in {filename}."

        windows: list[list[int]] = []
        for idx in match_indices:
            start = max(0, idx - context_lines)
            end = min(len(lines), idx + context_lines + 1)
            if windows and start <= windows[-1][1]:
                windows[-1][1] = max(windows[-1][1], end)
            else:
                windows.append([start, end])

        out = [
            f"Found {len(match_indices)} matches in {filename} "
            f"(merged into {len(windows)} chunks):"
        ]
        for w_start, w_end in windows[:MAX_CHUNKS_RETURNED]:
            out.append(f"\n--- lines {w_start + 1}-{w_end} ---")
            for i in range(w_start, w_end):
                marker = ">>" if regex.search(lines[i]) else "  "
                line = lines[i]
                if len(line) > MAX_CHARS_PER_LINE:
                    line = line[:MAX_CHARS_PER_LINE] + "...[truncated]"
                out.append(f"{marker} L{i + 1}: {line}")

        if len(windows) > MAX_CHUNKS_RETURNED:
            out.append(
                f"\n... and {len(windows) - MAX_CHUNKS_RETURNED} more chunks omitted "
                f"(refine your pattern to narrow down)."
            )
        return "\n".join(out)

    return grep_context
