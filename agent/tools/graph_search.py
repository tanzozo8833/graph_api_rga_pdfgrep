import requests
from langchain_core.tools import tool

GRAPH = "https://graph.microsoft.com/v1.0"


def make_graph_search_tool(access_token: str):
    @tool
    def graph_search(query: str) -> str:
        """Search the user's OneDrive via Microsoft Graph full-text search.

        Returns up to 15 matching files. Each line has: item_id, name, snippet, webUrl.
        Use the item_id with fetch_file_text to read the file content.

        Use OR-joined keywords for better recall, e.g. "metric OR evaluation OR benchmark".
        Always call this FIRST before fetch_file_text or grep_context.
        """
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        body = {
            "requests": [
                {
                    "entityTypes": ["driveItem"],
                    "query": {"queryString": query},
                    "from": 0,
                    "size": 15,
                    "fields": [
                        "id",
                        "name",
                        "webUrl",
                        "lastModifiedDateTime",
                        "size",
                    ],
                }
            ]
        }
        resp = requests.post(f"{GRAPH}/search/query", headers=headers, json=body)
        if resp.status_code != 200:
            return f"ERROR: Graph search failed HTTP {resp.status_code}: {resp.text[:300]}"

        data = resp.json()
        try:
            hits = data["value"][0]["hitsContainers"][0].get("hits", [])
        except (KeyError, IndexError):
            return f"No results structure for query '{query}'."

        if not hits:
            return f"No files matched the query '{query}'."

        lines = [f"Found {len(hits)} files for query '{query}':"]
        for h in hits:
            res = h.get("resource", {}) or {}
            item_id = res.get("id", "?")
            name = res.get("name", "(unnamed)")
            summary = (h.get("summary") or "").replace("\n", " ").strip()[:240]
            web_url = res.get("webUrl", "")
            lines.append(
                f"- item_id={item_id} | name={name} | snippet=\"{summary}\" | webUrl={web_url}"
            )
        return "\n".join(lines)

    return graph_search
