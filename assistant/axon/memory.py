"""Axon's long-term memory: durable facts about the user (brand, key people, preferences, ongoing
projects) that persist across sessions and get woven into the system prompt so Axon "remembers"."""
import os
import json

from axon.util import _result

_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "AxonIntelligence")
_PATH = os.path.join(_DIR, "memory.json")


def _load():
    try:
        with open(_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save(items):
    try:
        os.makedirs(_DIR, exist_ok=True)
        with open(_PATH, "w", encoding="utf-8") as f:
            json.dump(items, f, indent=2)
    except Exception:
        pass


def remember(args):
    """Save a durable fact about the user (e.g. 'Our brand color is navy', 'My manager is Sara')."""
    text = (args.get("text") or "").strip()
    if not text:
        return _result("Nothing to remember (provide text).", True)
    cat = (args.get("category") or "general").strip()
    items = _load()
    if not any(i.get("text", "").lower() == text.lower() for i in items):
        items.append({"text": text, "category": cat})
        _save(items)
    return _result(f"Got it — I'll remember: {text}")


def recall(args):
    """List remembered facts, optionally filtered by a query word."""
    q = (args.get("query") or "").lower().strip()
    items = _load()
    if q:
        items = [i for i in items if q in i.get("text", "").lower() or q in i.get("category", "").lower()]
    if not items:
        return _result("I don't have anything saved" + (f" about '{q}'." if q else " yet."))
    return _result("\n".join(f"- ({i.get('category', 'general')}) {i['text']}" for i in items))


def forget(args):
    """Forget remembered facts whose text contains the given query."""
    q = (args.get("query") or "").lower().strip()
    if not q:
        return _result("Tell me what to forget (a query).", True)
    items = _load()
    kept = [i for i in items if q not in i.get("text", "").lower()]
    _save(kept)
    return _result(f"Forgot {len(items) - len(kept)} item(s).")


def context_block():
    """A compact block of everything remembered, for injecting into the system prompt ('' if none)."""
    items = _load()
    if not items:
        return ""
    lines = "\n".join(f"- ({i.get('category', 'general')}) {i['text']}" for i in items[:40])
    return ("\n\nWHAT YOU KNOW ABOUT THIS USER (long-term memory — apply it, and stay consistent "
            "with it):\n" + lines)
