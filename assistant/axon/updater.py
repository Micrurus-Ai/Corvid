"""Self-update: check a small JSON manifest for a newer build, then download + run the installer
silently so users update with one click from the composer.

The manifest at AXON_UPDATE_URL is a small JSON:
    {"version": "1.1.0", "url": "<installer location>", "notes": "what's new"}
  - version : the latest available version string (compared to APP_VERSION).
  - url     : where the new Axon-Setup.exe lives. It may be an https URL or a UNC/network/local
              path. Keep it INTERNAL — the installer has baked API keys — e.g. a network share the
              user can already reach with their Windows login.
  - notes   : optional 'what's new' text shown before updating.

Disabled (every call is a no-op returning None) when AXON_UPDATE_URL is unset, so nothing happens
until the hosting location is configured.
"""
import os
import json
import shutil
import tempfile
import subprocess
import urllib.request

from axon.config import APP_VERSION, UPDATE_URL
from axon.util import NO_WINDOW


def _ver_tuple(s):
    """Turn '1.10.2' into (1, 10, 2) for a correct numeric comparison."""
    out = []
    for part in str(s or "").strip().split("."):
        digits = "".join(c for c in part if c.isdigit())
        out.append(int(digits) if digits else 0)
    return tuple(out) or (0,)


def _is_url(s):
    return str(s).lower().startswith(("http://", "https://"))


def _load_manifest():
    if not UPDATE_URL:
        return None
    try:
        if _is_url(UPDATE_URL):
            with urllib.request.urlopen(UPDATE_URL, timeout=8) as r:
                raw = r.read().decode("utf-8", "replace")
        else:
            with open(UPDATE_URL, encoding="utf-8") as f:
                raw = f.read()
        return json.loads(raw)
    except Exception:
        return None


def check_for_update():
    """Return {'version','url','notes'} if a newer version is published, else None."""
    m = _load_manifest()
    if not isinstance(m, dict):
        return None
    ver = str(m.get("version", "")).strip()
    url = str(m.get("url", "")).strip()
    if not ver or not url:
        return None
    if _ver_tuple(ver) > _ver_tuple(APP_VERSION):
        return {"version": ver, "url": url, "notes": str(m.get("notes", "")).strip()}
    return None


def download_installer(url, on_status=None):
    """Fetch the installer (https URL or UNC/local path) to a temp file. Returns the path or ''."""
    def status(msg):
        if on_status:
            on_status(msg)
    dst = os.path.join(tempfile.gettempdir(), "Axon-Setup-update.exe")
    try:
        status("Downloading update...")
        if _is_url(url):
            with urllib.request.urlopen(url, timeout=180) as r, open(dst, "wb") as f:
                shutil.copyfileobj(r, f)
        else:
            shutil.copyfile(url, dst)   # network share / mapped drive / local path
        return dst
    except Exception as e:
        status(f"Couldn't download the update: {e}")
        return ""


def launch_installer(path):
    """Run the downloaded installer silently and ask it to relaunch the app afterwards
    (the installer's [Run] entry keys off /relaunch=1). The caller should quit the app so the
    in-place upgrade can replace files."""
    subprocess.Popen(
        [path, "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/relaunch=1"],
        creationflags=NO_WINDOW)
