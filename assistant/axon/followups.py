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


# The add-in OWNS followups.json (append-only). The dot must never write it, or it would clobber a
# follow-up the add-in appended concurrently. Instead the dot tracks which ones it has already
# reminded about in its own file, keyed by id+due.
def _seen_path():
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonOutlook")
    return os.path.join(base, "followups_seen.json")


def _load_seen():
    try:
        with open(_seen_path(), encoding="utf-8") as f:
            return set(json.load(f))
    except Exception:
        return set()


def _save_seen(seen):
    try:
        os.makedirs(os.path.dirname(_seen_path()), exist_ok=True)
        with open(_seen_path(), "w", encoding="utf-8") as f:
            json.dump(sorted(seen), f)
    except Exception:
        pass


def _key(it):
    return (it.get("id") or "") + "|" + (it.get("due") or "")


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
    """Return follow-ups that are now due and not yet reminded about. The dot only writes its own
    'seen' file (never followups.json), so it can't clobber a follow-up the add-in just added.
    Follow-ups whose conversation already got a reply are marked seen and skipped silently."""
    items = _load()
    if not items:
        return []
    now = datetime.datetime.now()
    seen = _load_seen()
    due, changed = [], False
    for it in items:
        k = _key(it)
        if k in seen or it.get("notified"):   # already reminded (seen set, or the legacy flag)
            continue
        try:
            when = datetime.datetime.fromisoformat(it.get("due"))
        except Exception:
            when = now
        if when > now:
            continue
        seen.add(k)
        changed = True
        if _replied(it):
            continue                          # they replied — mark seen, don't nag
        due.append(it)
    if changed:
        _save_seen(seen)
    return due
