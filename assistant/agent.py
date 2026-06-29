"""
Axon intelligence computer-use agent — public API facade.

Axon intelligence is the brain; `open-computer-use` is the hands. The implementation now lives in
the `axon` package (one responsibility per module):

  axon.config        env/key loading, model + path constants, shared dot-window state
  axon.prompts       system / browse / chat / guide prompt strings
  axon.util          shared result helper
  axon.approval      user-approval gating for state-changing actions
  axon.screen        monitor geometry, screenshots, moving windows onto the dot's monitor
  axon.mcp           the open-computer-use client + open/close/list-app tools
  axon.outlook       Outlook automation (email / folders / calendar / signature / ...)
  axon.messaging     Microsoft Teams messaging
  axon.files         local file reading
  axon.browse        debug Chrome + browser-use navigation + whole-site research
  axon.inbox_filer   inbox watcher + AI folder suggestions (also powers the Outlook add-in)
  axon.tools         the tool schemas (TOOLS) + name->callable map (DISPATCH)
  axon.brain         the LLM agent loop (run_task), chat, and the live guide

This module re-exports the names the rest of the app uses (overlay.py, the Outlook add-in's
--suggest hook, and the CLI below), so those callers don't need to know the package layout.
"""

import sys

from axon import config  # noqa: F401  (sets up env + holds DOT_HWND/DOT_PID)
from axon.config import *  # noqa: F401,F403  (MODEL, IS_WINDOWS, DOWNLOADS_DIR, ... read-only)
from axon.tools import TOOLS, DISPATCH  # noqa: F401
from axon.brain import chat, run_task, guide_live, guide  # noqa: F401
from axon.inbox_filer import (  # noqa: F401
    inbox_watcher_popen, suggest_folders, rank_folders, suggest_filing, move_email_to_folder,
)

if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # Windows console is cp1252 by default
    except Exception:
        pass
    q = " ".join(sys.argv[1:]) or input("Task: ")
    run_task(q, on_status=print)
