"""Small shared helpers used across tool modules."""
import os
import re
import subprocess

# The dot runs as a GUI app (pythonw) with no console of its own, so any child console
# process (powershell, cmd, the Chrome launcher) gets a fresh, visible console window.
# Pass this as `creationflags=NO_WINDOW` to every subprocess call to keep them hidden.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


# Axon must never reveal the underlying model/provider. Scrub any brand names the LLM might emit
# ("ChatGPT", "OpenAI", "Mistral", "GPT-4o", ...) from user-facing text and speak only as Axon.
_IDENTITY_RE = re.compile(r"\bchat\s?gpt\b|\bopen\s?ai\b|\bmistral(?:\s?ai)?\b|\bgpt[-\s]?[\w.]+",
                          re.IGNORECASE)


def scrub_identity(text):
    """Replace any leaked model/provider brand names with 'Axon' so responses stay on-brand."""
    if not text:
        return text
    return _IDENTITY_RE.sub("Axon", str(text))


def _result(text, is_error=False):
    """Build an MCP-style tool result dict. Error/warning text is scrubbed of any model/provider
    brand names (OpenAI, ChatGPT, ...) so warnings never reveal what powers Axon. Success results
    are left untouched so legitimate file/email content is never altered."""
    t = str(text)
    if is_error:
        t = scrub_identity(t)
    return {"content": [{"type": "text", "text": t}], "isError": bool(is_error)}
