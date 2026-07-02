"""User-approval gating for actions that change state or send messages.

run_task installs a callback into approval._APPROVAL_CB; tools call _ask_approval() before acting,
and run_task checks _needs_approval()/_describe_action() before gated tools run."""

_APPROVAL_CB = None


def _ask_approval(desc, allow_later=False):
    """Ask the user to approve an action. Returns True (do it now), False (skip), or — when
    allow_later is set (email sends) and the user picks 'Send Later' — an ISO datetime string."""
    cb = _APPROVAL_CB
    if cb is None:
        return True
    try:
        return cb(desc, allow_later)
    except Exception:
        return True


_APPROVAL_TOOLS = {
    "send_teams_message", "create_calendar_event", "respond_to_meeting",
    "outlook_move_emails", "outlook_delete_emails", "outlook_categorize", "outlook_mark_read",
    "create_outlook_folders", "delete_outlook_folders", "create_outlook_rule", "set_outlook_signature",
    "outlook_contact", "outlook_task", "schedule_task", "close_app",
}


def _needs_approval(name, args):
    if name in _APPROVAL_TOOLS:
        return True
    if name == "file_op" and (args.get("action") or "").lower() in {
        "delete", "move", "rename", "write_text", "unzip", "copy", "mkdir",
    }:
        return True
    return False


def _count(v):
    return len(v) if isinstance(v, list) else 1


def _describe_action(name, a):
    a = a or {}
    if name == "send_email":
        s = a.get("subject")
        return f"Send email to {a.get('to', '?')}" + (f" — subject: {s}" if s else "")
    if name == "outlook_forward_email":
        return f"Forward this email to {a.get('to', '?')}"
    if name == "outlook_reply_email":
        return "Reply-all to this email" if a.get("reply_all") else "Reply to this email"
    if name == "send_teams_message":
        return f"Send a Teams message to {a.get('to', '?')}: \"{(a.get('message') or '')[:60]}\""
    if name == "create_calendar_event":
        att = a.get("attendees")
        return (f"Create a meeting '{a.get('subject', '')}' with {att}" if att
                else f"Add '{a.get('subject', '')}' to your calendar")
    if name == "respond_to_meeting":
        return f"{a.get('response', 'respond to')} the meeting invite"
    if name == "outlook_move_emails":
        return f"Move {_count(a.get('ids'))} email(s) to '{a.get('to_folder', '')}'"
    if name == "outlook_delete_emails":
        return f"Delete {_count(a.get('ids'))} email(s) (to Deleted Items)"
    if name == "outlook_categorize":
        return f"Categorize {_count(a.get('ids'))} email(s) as '{a.get('category', '')}'"
    if name == "outlook_mark_read":
        return f"Mark {_count(a.get('ids'))} email(s) as {'read' if a.get('read', True) else 'unread'}"
    if name == "create_outlook_folders":
        return "Create folder(s): " + ", ".join(a.get("folders") or [])
    if name == "delete_outlook_folders":
        return "Delete folder(s): " + ", ".join(a.get("folders") or [])
    if name == "create_outlook_rule":
        return f"Create mail rule '{a.get('name', '')}'"
    if name == "set_outlook_signature":
        return "Set/update your email signature"
    if name == "outlook_contact":
        return f"Create contact '{a.get('name', '')}'" if (a.get("action") or "create") == "create" else "Find contacts"
    if name == "outlook_task":
        return f"Create task '{a.get('subject', '')}'" if (a.get("action") or "create") == "create" else "List tasks"
    if name == "schedule_task":
        return f"Create scheduled task '{a.get('name', '')}'"
    if name == "close_app":
        return f"Close {a.get('app', '')}"
    if name == "file_op":
        return f"File: {a.get('action', '')} {a.get('path', '')}".strip()
    return name
