import os

import requests
from langchain_core.tools import tool

from .. import cache
from ..extractors import SUPPORTED, extract_text

GRAPH = "https://graph.microsoft.com/v1.0"
MAX_SIZE_BYTES = 25 * 1024 * 1024  # 25 MB


def make_file_fetch_tool(access_token: str):
    @tool
    def fetch_file_text(item_id: str) -> str:
        """Download a OneDrive file by item_id, extract its text, and cache it in memory.

        Supported types: PDF, DOCX, XLSX, PPTX, and plain text (txt/md/csv/json/...).
        Files larger than 25 MB are rejected.

        DO NOT call this tool more than once for the same item_id — the result is cached
        and subsequent grep_context calls read from cache.

        Returns filename, byte size, character length, and a short preview.
        Use grep_context to actually search inside the text.
        """
        cached = cache.get(item_id)
        if cached:
            text = cached["text"]
            return (
                f"ALREADY CACHED. filename={cached['filename']}, "
                f"text_length={len(text)} chars. "
                f"Use grep_context with this item_id to search inside."
            )

        headers = {"Authorization": f"Bearer {access_token}"}

        meta_resp = requests.get(
            f"{GRAPH}/me/drive/items/{item_id}",
            headers=headers,
            params={"$select": "name,size,file"},
        )
        if meta_resp.status_code != 200:
            return f"ERROR: cannot read metadata for {item_id} (HTTP {meta_resp.status_code})."
        meta = meta_resp.json()
        name = meta.get("name", "")
        size = meta.get("size", 0) or 0

        if size > MAX_SIZE_BYTES:
            return f"ERROR: file '{name}' too large ({size/1024/1024:.1f} MB > 25 MB)."

        ext = os.path.splitext(name)[1].lower()
        if ext not in SUPPORTED:
            return (
                f"ERROR: unsupported file type '{ext}' for '{name}'. "
                f"Supported: PDF, DOCX, XLSX, PPTX, and plain text."
            )

        c_resp = requests.get(
            f"{GRAPH}/me/drive/items/{item_id}/content", headers=headers
        )
        if c_resp.status_code != 200:
            return f"ERROR: cannot download '{name}' (HTTP {c_resp.status_code})."

        try:
            text = extract_text(name, c_resp.content)
        except Exception as e:
            return f"ERROR: failed to extract text from '{name}': {type(e).__name__}: {e}"

        cache.put(item_id, name, text)
        preview = text[:300].replace("\n", " ")
        return (
            f"Fetched and cached. filename={name}, size={size} bytes, "
            f"text_length={len(text)} chars.\n"
            f"Preview: {preview}..."
        )

    return fetch_file_text
