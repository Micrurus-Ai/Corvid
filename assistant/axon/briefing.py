"""Daily briefing + inbox triage: read Outlook (calendar, unread, tasks) and summarize / prioritize."""
import os
import json

from openai import OpenAI

from axon.util import _result
from axon.config import IS_WINDOWS, MODEL
from axon.outlook._base import _run_outlook_ps

_BRIEFING_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $today = (Get-Date).Date
    $tom = $today.AddDays(1)
    $cal = $ns.GetDefaultFolder(9).Items
    $cal.IncludeRecurrences = $true
    $cal.Sort("[Start]")
    $f = "[Start] >= '" + $today.ToString("MM/dd/yyyy") + "' AND [Start] < '" + $tom.ToString("MM/dd/yyyy") + "'"
    $appts = @()
    foreach ($a in $cal.Restrict($f)) {
        $appts += [ordered]@{ time = ([datetime]$a.Start).ToString("HH:mm"); subject = [string]$a.Subject; location = [string]$a.Location }
        if ($appts.Count -ge 20) { break }
    }
    $inbox = $ns.GetDefaultFolder(6).Items
    $unreadItems = $inbox.Restrict("[UnRead]=true")
    $unread = @(); $uc = 0
    foreach ($m in $unreadItems) { $uc++; if ($unread.Count -lt 10) { $unread += [ordered]@{ from = [string]$m.SenderName; subject = [string]$m.Subject } } }
    $tasks = @()
    try {
        $tk = $ns.GetDefaultFolder(13).Items.Restrict("[Complete]=false")
        foreach ($t in $tk) {
            if ($tasks.Count -lt 15) {
                $due = ""
                try { if ($t.DueDate.Year -lt 4000) { $due = ([datetime]$t.DueDate).ToString("yyyy-MM-dd") } } catch {}
                $tasks += [ordered]@{ subject = [string]$t.Subject; due = $due }
            }
        }
    } catch {}
    ([ordered]@{ appointments = $appts; unread_count = $uc; unread = $unread; tasks = $tasks }) | ConvertTo-Json -Depth 5 -Compress
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


_NEXT_EVENT_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $now = Get-Date
    $cal = $ns.GetDefaultFolder(9).Items
    $cal.IncludeRecurrences = $true
    $cal.Sort("[Start]")
    $f = "[Start] >= '" + $now.ToString("MM/dd/yyyy") + "' AND [Start] <= '" + $now.AddDays(1).ToString("MM/dd/yyyy") + "'"
    foreach ($a in $cal.Restrict($f)) {
        $mins = [int]([datetime]$a.Start - $now).TotalMinutes
        if ($mins -ge 0 -and $mins -le [int]$env:AX_WIN) {
            [ordered]@{ id = [string]$a.EntryID; subject = [string]$a.Subject; minutes = $mins; time = ([datetime]$a.Start).ToString("HH:mm") } | ConvertTo-Json -Compress
            break
        }
    }
} catch {}
'''


def upcoming_event(within_min=10):
    """Return the next calendar event starting within `within_min` minutes, or None. Used by the
    dot's proactive nudge — kept lightweight (no model call, Outlook not surfaced)."""
    if not IS_WINDOWS:
        return None
    out = _run_outlook_ps(_NEXT_EVENT_PS, {"AX_WIN": str(int(within_min))}, show=False)
    if not out or not out.lstrip().startswith("{"):
        return None
    try:
        return json.loads(out)
    except Exception:
        return None


def daily_briefing(args):
    """Summarize today: calendar events, unread email (count + top senders), and open tasks."""
    if not IS_WINDOWS:
        return _result("The daily briefing needs the Windows app with Outlook.", True)
    out = _run_outlook_ps(_BRIEFING_PS, {}, show=False)
    if not out or out.startswith("OL_ERROR") or not out.lstrip().startswith("{"):
        return _result("Could not read Outlook for the briefing. " + (out or ""), True)
    try:
        d = json.loads(out)
    except Exception:
        return _result("Could not parse the briefing data.", True)
    lines = ["Here's your day:"]
    appts = d.get("appointments") or []
    lines.append(f"\nCalendar ({len(appts)} today):")
    if appts:
        for a in appts:
            loc = f"  @ {a['location']}" if a.get("location") else ""
            lines.append(f"  {a.get('time', '')}  {a.get('subject', '')}{loc}")
    else:
        lines.append("  (nothing scheduled)")
    lines.append(f"\nUnread email: {d.get('unread_count', 0)}")
    for m in (d.get("unread") or []):
        lines.append(f"  - {m.get('from', '')}: {m.get('subject', '')}")
    tasks = d.get("tasks") or []
    if tasks:
        lines.append(f"\nOpen tasks ({len(tasks)}):")
        for t in tasks:
            due = f"  (due {t['due']})" if t.get("due") else ""
            lines.append(f"  - {t.get('subject', '')}{due}")
    return _result("\n".join(lines))


_TRIAGE_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
try {
    $ol = New-Object -ComObject Outlook.Application
    $unread = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items.Restrict("[UnRead]=true")
    $arr = @()
    foreach ($m in $unread) {
        if ($m.Class -ne 43) { continue }
        $b = [string]$m.Body
        if ($b.Length -gt 400) { $b = $b.Substring(0, 400) }
        $b = $b -replace "[\r\n]+", " "
        $arr += [ordered]@{ from = [string]$m.SenderName; subject = [string]$m.Subject; snippet = $b }
        if ($arr.Count -ge 30) { break }
    }
    $arr | ConvertTo-Json -Depth 4 -Compress
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def inbox_triage(args):
    """Triage the unread inbox: order by importance and give each a priority, a one-line summary,
    and a suggested action. Use for 'what needs my attention / triage my inbox'."""
    if not IS_WINDOWS:
        return _result("Inbox triage needs the Windows app with Outlook.", True)
    out = _run_outlook_ps(_TRIAGE_PS, {}, show=False)
    if not out or out.startswith("OL_ERROR"):
        return _result("Couldn't read the inbox. " + (out or ""), True)
    try:
        data = json.loads(out)
    except Exception:
        data = []
    if isinstance(data, dict):
        data = [data]
    if not data:
        return _result("No unread emails — inbox zero!")
    if not os.getenv("OPENAI_API_KEY"):
        return _result("Unread:\n" + "\n".join(f"- {m.get('from', '')}: {m.get('subject', '')}" for m in data))
    items = "\n".join(
        f"{i + 1}. From: {m.get('from', '')} | Subject: {m.get('subject', '')} | {m.get('snippet', '')}"
        for i, m in enumerate(data))
    prompt = (
        "Here are my unread emails. Triage them: order them most-important/urgent first, and for each give "
        "a priority (High/Medium/Low), a one-line summary, and a short suggested action. Be concise.\n\n"
        + items + "\n\nReturn a clean numbered list, most important first.")
    try:
        from axon.llm import text_llm
        client, model = text_llm()
        r = client.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}], temperature=0.2)
        return _result((r.choices[0].message.content or "").strip() or "(no triage produced)")
    except Exception as e:
        return _result(f"Triage failed: {e}", True)
