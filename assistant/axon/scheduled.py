"""Notify when a scheduled ('Send Later') email's time arrives. Both the Outlook add-in and the dot
record scheduled sends to scheduled.json (append-only); the dot polls and fires a notification at the
due time. Like followups.py, the poller only writes its own 'seen' file, never scheduled.json."""
import os
import json
import datetime


def _path():
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonOutlook")
    return os.path.join(base, "scheduled.json")


def _seen_path():
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonOutlook")
    return os.path.join(base, "scheduled_seen.json")


def _load():
    try:
        with open(_path(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


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
    return (it.get("to") or "") + "|" + (it.get("when") or "")


def record(subject, to, when_iso, remind_before=0):
    """Append a scheduled send (called when a Send Later is set). If remind_before > 0, also store a
    pre-send heads-up time. Best-effort."""
    if not when_iso:
        return
    try:
        entry = {
            "subject": subject or "", "to": to or "", "when": when_iso,
            "created": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        }
        if remind_before:
            try:
                w = datetime.datetime.fromisoformat(when_iso)
                entry["remind_at"] = (w - datetime.timedelta(minutes=int(remind_before))).strftime("%Y-%m-%dT%H:%M:%S")
            except Exception:
                pass
        items = _load()
        items.append(entry)
        os.makedirs(os.path.dirname(_path()), exist_ok=True)
        with open(_path(), "w", encoding="utf-8") as f:
            json.dump(items, f)
    except Exception:
        pass


def due_scheduled():
    """Return due scheduled-send events: a 'remind' heads-up before the send (if the user asked for
    one) and a 'sent' notice when its time arrives. Each entry can produce both, tracked separately."""
    items = _load()
    if not items:
        return []
    now = datetime.datetime.now()
    seen = _load_seen()
    out, changed = [], False
    for it in items:
        base = _key(it)
        remind_at = it.get("remind_at")
        if remind_at:
            rk = base + "|remind"
            if rk not in seen:
                try:
                    t = datetime.datetime.fromisoformat(remind_at)
                except Exception:
                    t = None
                if t and t <= now:
                    seen.add(rk)
                    changed = True
                    out.append(dict(it, kind="remind"))
        sk = base + "|sent"
        if sk not in seen:
            try:
                w = datetime.datetime.fromisoformat(it.get("when"))
            except Exception:
                w = None
            if w and w <= now:
                seen.add(sk)
                changed = True
                out.append(dict(it, kind="sent"))
    if changed:
        _save_seen(seen)
    return out
