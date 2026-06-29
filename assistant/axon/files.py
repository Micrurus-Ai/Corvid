"""Local file reading tool."""
import os

from axon.config import TOOL_TEXT_LIMIT
from axon.util import _result

def _read_file(args):
    path = (args.get("path") or "").strip()
    if not path:
        return _result("Missing file path.", True)
    if not os.path.isfile(path):
        return _result(f"File not found: {path}", True)
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            data = f.read()
    except Exception as e:
        return _result(f"Could not read file: {e}", True)
    return _result(data[:TOOL_TEXT_LIMIT])
