"""Viewer for debug/grep.log — prints recent grep_context invocations.

Usage:
    python debug/grep.py            # print last 1 entry
    python debug/grep.py 3          # print last 3 entries
    python debug/grep.py all        # print all entries
    python debug/grep.py --clear    # clear the log
"""

import os
import sys

LOG_FILE = os.path.join(os.path.dirname(__file__), "grep.log")
SEPARATOR = "=" * 70


def main():
    if not os.path.exists(LOG_FILE):
        print(f"No log yet at {LOG_FILE}")
        print("Ask the agent a question that triggers grep_context, then re-run this.")
        return

    if len(sys.argv) > 1 and sys.argv[1] == "--clear":
        open(LOG_FILE, "w", encoding="utf-8").close()
        print(f"Cleared {LOG_FILE}")
        return

    with open(LOG_FILE, encoding="utf-8") as f:
        content = f.read()

    # Split on the separator; each non-empty piece is one entry.
    parts = [p for p in content.split(SEPARATOR) if p.strip()]
    if not parts:
        print(f"{LOG_FILE} is empty.")
        return

    arg = sys.argv[1] if len(sys.argv) > 1 else "1"
    if arg == "all":
        selected = parts
    else:
        try:
            n = int(arg)
        except ValueError:
            print(f"Invalid arg: {arg!r}. Use a number, 'all', or '--clear'.")
            return
        selected = parts[-n:]

    print(f"Showing {len(selected)} of {len(parts)} grep entries from {LOG_FILE}\n")
    for entry in selected:
        print(SEPARATOR + entry.rstrip())
    print(SEPARATOR)


if __name__ == "__main__":
    main()
