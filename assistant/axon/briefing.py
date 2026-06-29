"""Daily briefing: today's calendar, unread email, and open tasks, gathered from Outlook and
formatted into a short summary the user can read or have spoken."""
import json

from axon.util import _result
from axon.config import IS_WINDOWS
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
