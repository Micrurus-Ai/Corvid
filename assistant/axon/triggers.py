"""Email triggers: "when an email matching X arrives, do Y." Rules persist in settings and are
applied on demand by run_email_triggers (schedule that to run periodically for hands-off automation).
Actions are deliberately safe (move / categorize / mark read / draft a reply) — a reply is only
 DRAFTED and left open for the user to approve; it is never sent automatically."""
import os
import json

from openai import OpenAI

from axon.util import _result
from axon.config import IS_WINDOWS, MODEL
from axon.settings import load_settings, save_settings
from axon.outlook._base import _run_outlook_ps
from axon.outlook import _outlook_move_emails, _outlook_categorize, _outlook_mark_read, my_tone
from axon.outlook.email_send import _OUTLOOK_REPLY_PS

_MAX_DRAFTS_PER_RUN = 5   # don't flood the screen with reply windows in one pass

_UNREAD_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $unread = $ns.GetDefaultFolder(6).Items.Restrict("[UnRead]=true")
    $arr = @()
    foreach ($m in $unread) {
        if ($m.Class -ne 43) { continue }
        $b = [string]$m.Body
        if ($b.Length -gt 1500) { $b = $b.Substring(0, 1500) }
        $arr += [ordered]@{ id = [string]$m.EntryID; from = [string]$m.SenderName; email = [string]$m.SenderEmailAddress; subject = [string]$m.Subject; body = $b }
        if ($arr.Count -ge 100) { break }
    }
    $arr | ConvertTo-Json -Depth 4 -Compress
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _rules():
    return load_settings().get("email_triggers") or []


def add_email_trigger(args):
    """Add a rule: match by sender (from) and/or subject_contains; action =
    move/categorize/mark_read/draft_reply."""
    action = (args.get("action") or "").lower()
    if action not in ("move", "categorize", "mark_read", "draft_reply"):
        return _result("action must be move, categorize, mark_read, or draft_reply.", True)
    rule = {"from": args.get("from") or "", "subject_contains": args.get("subject_contains") or "",
            "action": action, "folder": args.get("folder") or "", "category": args.get("category") or ""}
    if not rule["from"] and not rule["subject_contains"]:
        return _result("Give a match: from (sender name/email contains) and/or subject_contains.", True)
    if action == "move" and not rule["folder"]:
        return _result("The move action needs a folder.", True)
    if action == "categorize" and not rule["category"]:
        return _result("The categorize action needs a category.", True)
    s = load_settings()
    r = s.get("email_triggers") or []
    r.append(rule)
    s["email_triggers"] = r
    save_settings(s)
    return _result(f"Trigger added ({len(r)} total).")


def list_email_triggers(args):
    """List the configured email triggers."""
    r = _rules()
    if not r:
        return _result("No email triggers set.")
    return _result("\n".join(
        f"[{i}] from~'{x['from']}' subj~'{x['subject_contains']}' -> {x['action']} "
        f"{x.get('folder') or x.get('category') or ''}".rstrip() for i, x in enumerate(r)))


def remove_email_trigger(args):
    """Remove an email trigger by its index (from list_email_triggers)."""
    r = _rules()
    try:
        i = int(args.get("index"))
    except Exception:
        return _result("Provide the index to remove (see list_email_triggers).", True)
    if i < 0 or i >= len(r):
        return _result("Index out of range.", True)
    removed = r.pop(i)
    s = load_settings()
    s["email_triggers"] = r
    save_settings(s)
    return _result(f"Removed trigger [{i}] ({removed['action']}).")


def _draft_reply_text(subject, sender, body):
    """Draft a reply body in the user's learned tone (same language as the email). '' on failure."""
    if not os.getenv("OPENAI_API_KEY"):
        return ""
    tone = my_tone()
    toneline = ("\nMatch the user's personal writing style:\n" + tone) if tone else ""
    prompt = (
        "Draft a reply to this email. Begin with a greeting to the sender by first name and end with a "
        "courteous sign-off. Reply in the SAME language as the email. Output ONLY the reply text "
        "(no subject, no quoted original)." + toneline +
        f"\n\nFrom: {sender}\nSubject: {subject}\n\n{body[:2000]}")
    try:
        from axon.llm import text_llm
        client, model = text_llm()
        r = client.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}], temperature=0.4)
        return (r.choices[0].message.content or "").strip()
    except Exception:
        return ""


def _match(rule, m):
    f = (rule.get("from") or "").lower()
    sc = (rule.get("subject_contains") or "").lower()
    if f and f not in (m.get("from", "") + " " + m.get("email", "")).lower():
        return False
    if sc and sc not in m.get("subject", "").lower():
        return False
    return True


def run_email_triggers(args):
    """Apply the configured triggers to current unread inbox emails now (move/categorize/mark read)."""
    if not IS_WINDOWS:
        return _result("Email triggers need the Windows app with Outlook.", True)
    rules = _rules()
    if not rules:
        return _result("No email triggers configured (use add_email_trigger).")
    out = _run_outlook_ps(_UNREAD_PS, {}, show=False)
    if not out or out.startswith("OL_ERROR"):
        return _result("Could not read the inbox. " + (out or ""), True)
    try:
        data = json.loads(out)
    except Exception:
        data = []
    if isinstance(data, dict):
        data = [data]
    actions, report, drafts = 0, [], 0
    for m in data:
        for rule in rules:
            if _match(rule, m):
                act, eid = rule["action"], m["id"]
                if act == "move":
                    _outlook_move_emails({"ids": [eid], "to_folder": rule["folder"]})
                elif act == "categorize":
                    _outlook_categorize({"ids": [eid], "category": rule["category"]})
                elif act == "mark_read":
                    _outlook_mark_read({"ids": [eid]})
                elif act == "draft_reply":
                    if drafts >= _MAX_DRAFTS_PER_RUN:
                        continue
                    body = _draft_reply_text(m.get("subject", ""), m.get("from", ""), m.get("body", ""))
                    if not body:
                        report.append(f"draft_reply (couldn't draft): {m.get('subject', '')[:50]}")
                        break
                    _run_outlook_ps(_OUTLOOK_REPLY_PS,
                                    {"OL_ID": eid, "OL_BODY": body, "OL_REPLYALL": ""}, show=True)
                    drafts += 1
                actions += 1
                report.append(f"{act}: {m.get('subject', '')[:50]}")
                break
    if not actions:
        return _result("No unread emails matched your triggers.")
    return _result(f"Applied {actions} trigger action(s):\n" + "\n".join(report))
