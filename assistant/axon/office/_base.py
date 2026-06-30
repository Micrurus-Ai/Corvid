"""Shared helpers for the Office/Windows tools: PowerShell runner, temp configs, path
resolution, and placing opened documents on the dot's monitor."""
import os
import json
import tempfile
import subprocess

from axon.util import _result, NO_WINDOW


def _dot_mon_env():
    """The dot's monitor work-rect as DOT_MON_* env vars (so Office windows open on the dot's screen)."""
    try:
        from axon import screen
        d = screen._dot_monitor_env()
        if d:
            return {"DOT_MON_L": str(d["L"]), "DOT_MON_T": str(d["T"]),
                    "DOT_MON_R": str(d["R"]), "DOT_MON_B": str(d["B"])}
    except Exception:
        pass
    return {}


def _move_doc_to_dot(path):
    """Move a just-opened Office document window onto the dot's monitor (by its filename in the title)."""
    try:
        from axon import screen
        screen._move_window_to_dot(os.path.splitext(os.path.basename(path))[0])
    except Exception:
        pass


def _run_ps(script, env=None, timeout=240):
    """Run a PowerShell script (written to a temp .ps1) with extra env vars. Returns output."""
    e = dict(os.environ)
    for k, v in _dot_mon_env().items():
        e[k] = v
    if env:
        for k, v in env.items():
            e[k] = "" if v is None else str(v)
    f = tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig")
    f.write(script)
    f.close()
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", f.name],
            capture_output=True, text=True, env=e, timeout=timeout, creationflags=NO_WINDOW,
        )
        out = (p.stdout or "").strip()
        err = (p.stderr or "").strip()
        return out + (("\n" + err) if err and not out else "")
    except subprocess.TimeoutExpired:
        return "ERR: the operation timed out"
    finally:
        try:
            os.unlink(f.name)
        except Exception:
            pass


def _downloads():
    return os.path.join(os.path.expanduser("~"), "Downloads")


def _resolve_path(path, default_name, default_ext):
    if not path:
        path = default_name
    if not os.path.splitext(path)[1]:
        path += default_ext
    if not os.path.isabs(path):
        path = os.path.join(_downloads(), path)
    return path


def _cfg_file(data):
    f = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8")
    json.dump(data, f)
    f.close()
    return f.name
