"""Real Windows toast notifications (Action Center, DND-aware) via WinRT — no pip dependency.
Falls back to PowerShell's registered AppUserModelID so a toast still shows in dev where Axon has
no Start-menu shortcut yet. Pair with the in-app Toast for a guaranteed-visible alert."""
import os
import html
import subprocess

from axon.config import IS_WINDOWS
from axon.util import NO_WINDOW

APP_ID = "AxonIntelligence.Dot"   # matches the installer shortcut's AppUserModelID
_PS_FALLBACK_AUMID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe"


def set_app_id():
    """Tag this process with our AppUserModelID so toasts group under 'Axon intelligence'."""
    if not IS_WINDOWS:
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


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
try {
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($env:AX_APPID).Show($toast)
} catch {
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($env:AX_FALLBACK).Show($toast)
}
'''


def os_notify(title, message):
    """Raise a real Windows toast (best-effort). Blocks ~1-2s, so call from a background thread."""
    if not IS_WINDOWS:
        return
    env = dict(os.environ)
    env["AX_TITLE"] = html.escape(title or "")
    env["AX_MSG"] = html.escape(message or "")
    env["AX_APPID"] = APP_ID
    env["AX_FALLBACK"] = _PS_FALLBACK_AUMID
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", _TOAST_PS],
            env=env, capture_output=True, timeout=15, creationflags=NO_WINDOW)
    except Exception:
        pass
