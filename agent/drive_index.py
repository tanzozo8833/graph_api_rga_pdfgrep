"""Maps item_id -> drive_id, populated by graph_search and consumed by file_fetch.

Microsoft Graph Search API returns items from multiple drives (user's OneDrive,
shared files, SharePoint, Teams). To download content, we need the driveId — it
is NOT always /me/drive. This index lets file_fetch use the correct drive endpoint.
"""

_INDEX: dict[str, str] = {}  # item_id -> drive_id


def put(item_id: str, drive_id: str | None):
    if item_id and drive_id:
        _INDEX[item_id] = drive_id


def get(item_id: str) -> str | None:
    return _INDEX.get(item_id)


def clear():
    _INDEX.clear()
