"""Small shared helpers used across tool modules."""


def _result(text, is_error=False):
    """Build an MCP-style tool result dict."""
    return {"content": [{"type": "text", "text": str(text)}], "isError": bool(is_error)}
