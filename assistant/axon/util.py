"""Small shared helpers used across tool modules."""
import os
import subprocess

# The dot runs as a GUI app (pythonw) with no console of its own, so any child console
# process (powershell, cmd, the Chrome launcher) gets a fresh, visible console window.
# Pass this as `creationflags=NO_WINDOW` to every subprocess call to keep them hidden.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def _result(text, is_error=False):
    """Build an MCP-style tool result dict."""
    return {"content": [{"type": "text", "text": str(text)}], "isError": bool(is_error)}
