"""Real Windows toast notifications (Action Center, DND-aware) via WinRT — no pip dependency.
Falls back to PowerShell's registered AppUserModelID so a toast still shows in dev where Axon has
no Start-menu shortcut yet. Pair with the in-app Toast for a guaranteed-visible alert."""
import os
import html
import tempfile
import itertools
import subprocess

from axon.config import IS_WINDOWS
from axon.util import NO_WINDOW

APP_ID = "AxonIntelligence.Dot"   # matches the installer shortcut's AppUserModelID
_PS_FALLBACK_AUMID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe"
_seq = itertools.count(1)   # unique toast tag so Windows doesn't collapse repeats


def set_app_id():
    """Tag this process with our AppUserModelID so toasts group under 'Axon intelligence'."""
    if not IS_WINDOWS:
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass
    _ensure_shortcut()   # register the AUMID shortcut up front so branded toasts work immediately


_registered = None   # None = not tried yet; True/False after attempting to register the AUMID shortcut


def _ensure_shortcut():
    """Windows only DISPLAYS toasts for an AUMID that has a Start-menu shortcut carrying that
    AppUserModelID. Create/repair one for APP_ID (pointing at the installed exe, or the dev
    launcher) so our branded toasts actually show. Runs once per process; returns True on success."""
    global _registered
    if _registered is not None:
        return _registered
    _registered = False
    if not IS_WINDOWS:
        return False
    try:
        import sys
        import pythoncom
        from win32com.shell import shell
        from win32com.propsys import propsys, pscon

        progs = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs")
        os.makedirs(progs, exist_ok=True)
        path = os.path.join(progs, "Axon intelligence.lnk")

        installed = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Axon", "AxonIntelligence.exe")
        if os.path.isfile(installed):
            target, args, workdir = installed, "", os.path.dirname(installed)
        else:
            target = sys.executable   # dev: pythonw.exe overlay.py
            overlay = os.path.join(os.path.dirname(os.path.dirname(__file__)), "overlay.py")
            args, workdir = '"%s"' % overlay, os.path.dirname(overlay)

        link = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
        link.SetPath(target)
        if args:
            link.SetArguments(args)
        link.SetWorkingDirectory(workdir)
        store = link.QueryInterface(propsys.IID_IPropertyStore)
        store.SetValue(pscon.PKEY_AppUserModel_ID, propsys.PROPVARIANTType(APP_ID, pythoncom.VT_LPWSTR))
        store.Commit()
        link.QueryInterface(pythoncom.IID_IPersistFile).Save(path, True)
        _registered = True
    except Exception:
        _registered = False
    return _registered


def _primary_aumid():
    """Branded AUMID once its shortcut is registered; PowerShell's always-registered AUMID otherwise."""
    return APP_ID if _ensure_shortcut() else _PS_FALLBACK_AUMID


_TOAST_PS = r'''
$ErrorActionPreference = 'SilentlyContinue'
$null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]
$null = [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType=WindowsRuntime]
$null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime]
$xml = @"
<toast><visual><binding template="ToastGeneric"><text>$env:AX_TITLE</text><text>$env:AX_MSG</text></binding></visual></toast>
"@
$doc = New-Object Windows.Data.Xml.Dom.XmlDocument
$doc.LoadXml($xml)
$toast = New-Object Windows.UI.Notifications.ToastNotification $doc
try { $toast.Tag = $env:AX_TAG } catch {}
try {
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($env:AX_APPID).Show($toast)
} catch {
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($env:AX_FALLBACK).Show($toast)
}
'''


def os_notify(title, message):
    """Raise a real Windows toast (best-effort). Blocks ~1-2s, so call from a background thread.
    Uses -STA -File (WinRT needs STA) and a unique tag (so Windows doesn't collapse repeats)."""
    if not IS_WINDOWS:
        return
    env = dict(os.environ)
    env["AX_TITLE"] = html.escape(title or "")
    env["AX_MSG"] = html.escape(message or "")
    env["AX_APPID"] = _primary_aumid()   # branded AUMID if installed, else PowerShell's (dev)
    env["AX_FALLBACK"] = _PS_FALLBACK_AUMID
    env["AX_TAG"] = "axon-%d" % next(_seq)
    tmp = None
    try:
        f = tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig")
        f.write(_TOAST_PS)
        f.close()
        tmp = f.name
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", tmp],
            env=env, capture_output=True, timeout=15, creationflags=NO_WINDOW)
    except Exception:
        pass
    finally:
        if tmp:
            try:
                os.remove(tmp)
            except Exception:
                pass
