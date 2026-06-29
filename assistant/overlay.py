"""Axon intelligence floating dot — entry point.

Single-instance guard, the --suggest CLI hook (used by the Outlook add-in), and main() which
shows the FloatingDot. The UI lives in axon/ui/, the background workers in axon/workers.py.
"""
import socket
import os
import sys
import time

from PySide6 import QtGui, QtWidgets

import agent
from axon.ui.theme import FONT_FAMILY
from axon.ui.dot import FloatingDot

_INSTANCE_LOCK = None
# A healthy dot tags its window with this title (invisible — the dot is frameless) so a later
# launch can tell a *live* dot apart from a crashed instance still holding the lock port.
DOT_WINDOW_TITLE = "AxonIntelligenceDot"


def _bind_lock(port):
    """Try to grab the single-instance lock port. Returns the bound socket, or None if taken."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.listen(1)
        return s
    except OSError:
        s.close()
        return None


def _dot_window_visible():
    """True if a healthy Axon dot window is currently on screen (Windows only)."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        found = {"v": False}
        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def cb(h, _l):
            if u.IsWindowVisible(h):
                n = u.GetWindowTextLengthW(h)
                if n:
                    buf = ctypes.create_unicode_buffer(n + 1)
                    u.GetWindowTextW(h, buf, n + 1)
                    if buf.value == DOT_WINDOW_TITLE:
                        found["v"] = True
                        return False
            return True

        u.EnumWindows(WNDENUMPROC(cb), 0)
        return found["v"]
    except Exception:
        return False


def _kill_lock_owner(port):
    """Terminate whatever process is holding the lock port (a stale/crashed dot)."""
    try:
        import psutil
        for c in psutil.net_connections(kind="inet"):
            if c.laddr and c.laddr.port == port and c.pid and c.pid != os.getpid():
                try:
                    psutil.Process(c.pid).terminate()
                except Exception:
                    pass
    except Exception:
        pass


def _acquire_single_instance(port=49737):
    """Hold a localhost port for the app's lifetime so only one dot runs. Self-healing: if the port
    is taken but no live dot window exists, the holder has crashed — reclaim the lock instead of
    refusing to start (which used to leave the user with no visible dot and no way in)."""
    global _INSTANCE_LOCK
    s = _bind_lock(port)
    if s is not None:
        _INSTANCE_LOCK = s
        return True
    if _dot_window_visible():
        return False  # a real, healthy dot is already up — defer to it
    _kill_lock_owner(port)  # stale holder, no window — take over
    for _ in range(20):
        s = _bind_lock(port)
        if s is not None:
            _INSTANCE_LOCK = s
            return True
        time.sleep(0.25)
    return False










































def main():
    if not _acquire_single_instance():
        print("Axon intelligence is already running - not starting a second dot.")
        return
    app = QtWidgets.QApplication(sys.argv)
    app.setFont(QtGui.QFont(FONT_FAMILY, 9))
    app.setQuitOnLastWindowClosed(False)
    dot = FloatingDot()
    dot.setWindowTitle(DOT_WINDOW_TITLE)  # lets a later launch detect this live dot (see the guard)
    dot.show()
    sys.exit(app.exec())


def _suggest_cli():
    """Headless mode used by the Outlook add-in: read a JSON file of {subject,sender,body,folders},
    emit suggest_filing(...) as JSON to stdout, and exit. Lets the packaged exe serve the add-in
    (no venv). Writes to fd 1 directly so it works even in a windowed (no-console) build."""
    import json
    try:
        with open(sys.argv[2], "r", encoding="utf-8-sig") as f:   # utf-8-sig tolerates a BOM
            d = json.load(f)
        out = json.dumps(agent.suggest_filing(d.get("subject", ""), d.get("sender", ""),
                                              d.get("body", ""), d.get("folders", [])))
    except Exception:
        out = '{"matches": [], "new_folder": ""}'
    try:
        os.write(1, (out + "\n").encode("utf-8"))   # OS stdout handle (works when parent redirects it)
    except Exception:
        try:
            print(out)
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--suggest":
        _suggest_cli()
    else:
        main()
