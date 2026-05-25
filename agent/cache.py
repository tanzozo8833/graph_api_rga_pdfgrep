"""In-process cache for extracted text, keyed by Graph driveItem id.

Lives for the Flask process lifetime. Simple dict — good enough for dev.
"""

_TEXT_CACHE: dict[str, dict] = {}


def get(item_id: str):
    return _TEXT_CACHE.get(item_id)


def put(item_id: str, filename: str, text: str):
    _TEXT_CACHE[item_id] = {"filename": filename, "text": text}


def clear():
    _TEXT_CACHE.clear()
