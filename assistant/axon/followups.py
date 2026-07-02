"""Follow-up reminders. The Outlook add-in records follow-ups to a shared JSON file; the dot polls
it and pops a reliable tray reminder at the due time — but only if the recipient hasn't replied yet
(Outlook's own mail-item reminders are unreliable, so we don't depend on them)."""
import os
import json
import datetime

from axon.config import IS_WINDOWS
from axon.outlook._base import _run_outlook_ps


def _path():
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonOutlook")
    return os.path.join(base, "followups.json")


def _load():
    try:
        with open(_path(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save(items):
    try:
        os.makedirs(os.path.dirname(_path()), exist_ok=True)
        with open(_path(), "w", encoding="utf-8") as f:
            json.dump(items, f)
    except Exception:
        pass


# Has a message in this conversation arrived (from anyone) since the follow-up was created?
_REPLIED_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6).Items
$topic = $env:AX_TOPIC
$since = [datetime]$env:AX_SINCE
$found = $false
foreach ($m in $inbox) {
  try {
    if ($m.Class -ne 43) { continue }
    if (([string]$m.ConversationTopic) -eq $topic -and ([datetime]$m.ReceivedTime) -gt $since) { $found = $true; break }
  } catch {}
}
if ($found) { "REPLIED" } else { "NONE" }
'''


def _replied(it):
    """Best-effort: True if a reply landed in the Inbox for this conversation after it was set."""
    topic = it.get("topic") or ""
    since = it.get("created") or ""
    if not topic or not since or not IS_WINDOWS:
        return False
    try:
        out = _run_outlook_ps(_REPLIED_PS, {"AX_TOPIC": topic, "AX_SINCE": since}, show=False)
        return "REPLIED" in (out or "")
    except Exception:
        return False


def due_followups():
    """Return follow-ups that are now due and not yet notified (marking them notified so they don't
    repeat). Follow-ups whose conversation already got a reply are dropped silently."""
    items = _load()
    if not items:
        return []
    now = datetime.datetime.now()
    keep, due, changed = [], [], False
    for it in items:
        if it.get("notified"):
            keep.append(it)
            continue
        try:
            when = datetime.datetime.fromisoformat(it.get("due"))
        except Exception:
            when = now
        if when > now:
            keep.append(it)
            continue
        # due now — remind only if there's still no reply
        if _replied(it):
            changed = True          # drop it: they replied, nothing to nag about
            continue
        it["notified"] = True
        changed = True
        due.append(it)
        keep.append(it)
    if changed:
        _save(keep)
    return due
