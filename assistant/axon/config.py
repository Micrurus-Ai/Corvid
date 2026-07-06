"""Configuration: env/key loading, model + path constants, and shared runtime state.

This is a leaf module (it imports nothing else from `axon`), so every other module can depend on
it without circular imports. The dot's window handle is mutated at runtime — always read it as
`config.DOT_HWND` (never `from axon.config import DOT_HWND`, which would copy the value once).
"""
import os
import sys
import platform
import shutil

from dotenv import load_dotenv

# axon/ lives inside the assistant dir, so the app root is one level up.
ASSIST_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

load_dotenv(os.path.join(ASSIST_DIR, ".env"))
# When packaged (PyInstaller), also load the baked-in .env from the bundle / next to the .exe,
# so the app works out of the box with no key entry.
_exe_dir = os.path.dirname(os.path.abspath(sys.executable))
for _base in (getattr(sys, "_MEIPASS", None), _exe_dir, os.path.join(_exe_dir, "_internal")):
    if _base:
        _envp = os.path.join(_base, ".env")
        if os.path.exists(_envp):
            load_dotenv(_envp)

MODEL = os.getenv("ASSISTANT_MODEL", "gpt-4o")
MAX_STEPS = int(os.getenv("ASSISTANT_MAX_STEPS", "30"))

# Provider split: pure-text tasks (chat, tone, notes, doc Q&A, triage, drafts) route to Mistral when
# a key is set — cheaper + EU data residency. Vision (agent/guide), voice, image-gen and embeddings
# stay on OpenAI. See axon/llm.text_llm().
MISTRAL_API_KEY = os.getenv("MISTRAL_API_KEY", "")
MISTRAL_BASE = os.getenv("MISTRAL_BASE", "https://api.mistral.ai/v1")
TEXT_MODEL = os.getenv("ASSISTANT_TEXT_MODEL", "mistral-medium-latest")
BROWSE_MODEL = os.getenv("ASSISTANT_BROWSE_MODEL", "gpt-4o")
BROWSE_FALLBACK_MODEL = os.getenv("ASSISTANT_BROWSE_FALLBACK_MODEL", "gpt-4o-mini")
# Guide mode does vision grounding on a grid; a fast vision model keeps it responsive.
GUIDE_MODEL = os.getenv("ASSISTANT_GUIDE_MODEL", "gpt-4o")
# Second "zoom-in" pass for pixel precision. Set ASSISTANT_GUIDE_REFINE=0 for ~2x faster steps.
GUIDE_REFINE = os.getenv("ASSISTANT_GUIDE_REFINE", "1") != "0"
BROWSE_MAX_STEPS = int(os.getenv("ASSISTANT_BROWSE_MAX_STEPS", "25"))
# Accessibility trees (e.g. Outlook) can be large; keep enough so element indices aren't cut off.
TOOL_TEXT_LIMIT = int(os.getenv("ASSISTANT_TOOL_TEXT_LIMIT", "40000"))

DOWNLOADS_DIR = os.path.join(ASSIST_DIR, ".downloads")

IS_WINDOWS = platform.system() == "Windows"


def _resolve_ocu():
    """Prefer an open-computer-use bundled with the installed app (an `ocu` folder next to the exe /
    in the bundle); otherwise fall back to one on PATH (dev / npm global)."""
    for _base in (os.path.dirname(os.path.abspath(sys.executable)), getattr(sys, "_MEIPASS", None)):
        if _base:
            _cand = os.path.join(_base, "ocu", "open-computer-use.cmd")
            if os.path.exists(_cand):
                return _cand
    return shutil.which("open-computer-use") or "open-computer-use"


_OCU = _resolve_ocu()

# Allow open-computer-use to type into apps whose fields don't expose a keyboard-typable
# accessibility value (e.g. Teams / other WebView2 apps). Without this, type_text errors with
# "UIA ValuePattern text fallback is disabled". Set before any open-computer-use subprocess spawns.
os.environ["OPEN_COMPUTER_USE_WINDOWS_ALLOW_UIA_TEXT_FALLBACK"] = "1"

# The overlay sets these to the floating dot's native window handle + its process id, so apps
# Axon opens are moved onto the SAME monitor as the dot. Mutated at runtime (see module docstring).
DOT_HWND = None
DOT_PID = None

# Read-only names exported via `from axon.config import *`. DOT_HWND/DOT_PID are intentionally
# excluded so callers reference them as config.DOT_HWND and see live updates.
__all__ = [
    "ASSIST_DIR", "MODEL", "MAX_STEPS", "BROWSE_MODEL", "BROWSE_FALLBACK_MODEL",
    "GUIDE_MODEL", "GUIDE_REFINE", "BROWSE_MAX_STEPS", "TOOL_TEXT_LIMIT",
    "DOWNLOADS_DIR", "IS_WINDOWS", "_OCU",
]
