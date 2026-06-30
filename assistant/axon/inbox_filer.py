"""Inbox auto-filer: watch for opened inbox emails, suggest the best subfolder with the LLM,
and move messages. Also powers the Outlook add-in's --suggest hook."""
import os
import re
import json
import subprocess

from openai import OpenAI

from axon.config import IS_WINDOWS, MODEL
from axon.util import _result, NO_WINDOW
from axon.outlook import _run_outlook_ps

_INBOX_WATCH_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
$prev = ""
$prompted = @{}
while ($true) {
  # Exit if Axon (the parent) is gone, so this watcher never orphans.
  if ($env:AXON_PARENT_PID) {
    if (-not (Get-Process -Id ([int]$env:AXON_PARENT_PID) -ErrorAction SilentlyContinue)) { break }
  }
  $ol = $null
  try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch {}
  if ($ol -eq $null) { Start-Sleep -Seconds 5; $prev = ""; continue }   # don't force-start Outlook
  $cur = ""; $subj = ""; $sender = ""
  try {
    $ns = $ol.GetNamespace("MAPI")
    $inboxId = $ns.GetDefaultFolder(6).EntryID
    $item = $null
    $insp = $ol.ActiveInspector()
    if ($insp -ne $null) { $item = $insp.CurrentItem }
    if ($item -eq $null) {
      $exp = $ol.ActiveExplorer()
      if ($exp -ne $null -and $exp.Selection.Count -eq 1) { $item = $exp.Selection.Item(1) }
    }
    if ($item -ne $null -and $item.Class -eq 43) {            # 43 = olMail
      $parent = $item.Parent
      if ($parent -ne $null -and $parent.EntryID -eq $inboxId) {   # still in the Inbox itself
        $cur = [string]$item.EntryID
        $subj = ([string]$item.Subject) -replace "[\r\n|]", " "
        $sender = ([string]$item.SenderName) -replace "[\r\n|]", " "
      }
    }
  } catch {}
  # Require the same email to be in view for two polls (~2-4s dwell) so quick scrolling doesn't fire.
  if ($cur -ne "" -and $cur -eq $prev -and -not $prompted.ContainsKey($cur)) {
    $prompted[$cur] = $true
    [Console]::Out.WriteLine("OPENED|" + $cur + "|" + $subj + "|" + $sender)
    [Console]::Out.Flush()
  }
  $prev = $cur
  Start-Sleep -Seconds 2
}
'''


def inbox_watcher_popen():
    """Start the background Outlook inbox watcher. Returns a Popen that prints 'OPENED|eid|subj|sender'
    lines whenever an unfiled Inbox email is opened. None if not supported."""
    if not IS_WINDOWS:
        return None
    try:
        env = dict(os.environ)
        env["AXON_PARENT_PID"] = str(os.getpid())
        return subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", _INBOX_WATCH_PS],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, bufsize=1, env=env,
            creationflags=NO_WINDOW,
        )
    except Exception:
        return None


_SUGGEST_INFO_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $m = $ns.GetItemFromID($env:OL_ID)
    $body = [string]$m.Body
    if ($body.Length -gt 1500) { $body = $body.Substring(0, 1500) }
    $folders = @()
    foreach ($f in $inbox.Folders) { $folders += [string]$f.Name }
    $obj = [ordered]@{
        subject = [string]$m.Subject
        sender = [string]$m.SenderName
        senderEmail = [string]$m.SenderEmailAddress
        body = $body
        folders = $folders
    }
    $obj | ConvertTo-Json -Compress -Depth 4
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def suggest_folders(eid):
    """For an Inbox email, return {subject, sender, suggestions:[top folders], folders:[all subfolders]}.
    The AI ranks the user's existing subfolders; we never invent new folder names."""
    out = _run_outlook_ps(_SUGGEST_INFO_PS, {"OL_ID": eid}, show=False)
    if not out or out.startswith("OL_ERROR") or not out.lstrip().startswith("{"):
        return None
    try:
        info = json.loads(out)
    except Exception:
        return None
    folders = [f for f in (info.get("folders") or []) if f]
    result = {"subject": info.get("subject", ""), "sender": info.get("sender", ""),
              "suggestions": [], "folders": folders}
    if not folders or not os.getenv("OPENAI_API_KEY"):
        return result
    try:
        client = OpenAI()
        prompt = (
            f"Email subject: {info.get('subject')}\n"
            f"From: {info.get('sender')} <{info.get('senderEmail')}>\n"
            f"Body (truncated):\n{info.get('body')}\n\n"
            f"The user's existing Inbox subfolders: {folders}\n\n"
            "Choose the 5 folders from that list that best fit this email, best first. "
            "Reply with ONLY a JSON array of folder names taken EXACTLY from the list — no new names."
        )
        resp = client.chat.completions.create(
            model=MODEL, messages=[{"role": "user", "content": prompt}], temperature=0)
        txt = resp.choices[0].message.content or ""
        m = re.search(r"\[.*\]", txt, re.S)
        sug = json.loads(m.group(0)) if m else []
        result["suggestions"] = [s for s in sug if s in folders][:5]
    except Exception:
        pass
    return result


def rank_folders(subject, sender, body, folders):
    """Return up to 5 of the GIVEN folder names that best fit the email (AI ranking only, no COM).
    Used by the Outlook add-in's Move button, which enumerates folders itself and just needs ranking."""
    folders = [f for f in (folders or []) if f]
    if not folders or not os.getenv("OPENAI_API_KEY"):
        return []
    try:
        client = OpenAI()
        prompt = (
            f"Email subject: {subject}\nFrom: {sender}\nBody (truncated):\n{(body or '')[:1500]}\n\n"
            f"The user's folders: {folders}\n\n"
            "Choose the 5 folders from that list that best fit this email, best first. "
            "Reply with ONLY a JSON array of folder names taken EXACTLY from the list — no new names."
        )
        resp = client.chat.completions.create(
            model=os.getenv("ASSISTANT_RANK_MODEL", "gpt-4o-mini"),  # fast model — ranking is simple
            messages=[{"role": "user", "content": prompt}], temperature=0)
        txt = resp.choices[0].message.content or ""
        m = re.search(r"\[.*\]", txt, re.S)
        sug = json.loads(m.group(0)) if m else []
        return [s for s in sug if s in folders][:5]
    except Exception:
        return []


def _folder_name_from_subject(subject):
    """Last-resort folder name derived from the email subject (strip Re:/Fwd:, Title-Case a few words)."""
    s = re.sub(r"(?i)^\s*(re|fw|fwd)\s*:\s*", "", (subject or "").strip())
    s = re.sub(r"[^A-Za-z0-9 ]+", " ", s)
    words = [w for w in s.split() if len(w) > 2][:3]
    return " ".join(w.capitalize() for w in words) or "Filed Email"


def suggest_filing(subject, sender, body, folders):
    """For the Outlook add-in: return {'matches': [up to 5 fitting existing folders, best first],
    'new_folder': 'Name'} — new_folder is a short name to CREATE when none of the existing folders
    fit well (otherwise empty, so the UI can pre-fill the create box only when it's needed)."""
    folders = [f for f in (folders or []) if f]
    result = {"matches": [], "new_folder": ""}
    if not os.getenv("OPENAI_API_KEY"):
        return result
    try:
        client = OpenAI()
        prompt = (
            f"Email subject: {subject}\nFrom: {sender}\nBody (truncated):\n{(body or '')[:1500]}\n\n"
            f"The user's existing folders:\n{folders}\n\n"
            "Decide where to file this email. Reply with ONLY JSON:\n"
            '{"matches": [up to 5 folder names taken EXACTLY from the list that fit well, best first], '
            '"new_folder": "if NONE of the existing folders is a good fit, a short clear name (1-3 words) '
            'for a NEW folder to create (this must NOT be empty in that case); if an existing folder fits, '
            'an empty string"}\n'
            "Never invent names in matches — they must come from the list."
        )
        resp = client.chat.completions.create(
            model=os.getenv("ASSISTANT_RANK_MODEL", "gpt-4o-mini"),
            messages=[{"role": "user", "content": prompt}], temperature=0)
        txt = resp.choices[0].message.content or ""
        m = re.search(r"\{.*\}", txt, re.S)
        if m:
            d = json.loads(m.group(0))
            by_lower = {f.lower(): f for f in folders}  # match case-insensitively, keep the exact name
            seen = []
            for s in (d.get("matches") or []):
                f = by_lower.get(str(s).strip().lower())
                if f and f not in seen:
                    seen.append(f)
            result["matches"] = seen[:5]
            nf = (d.get("new_folder") or "").strip()
            if nf:
                existing = by_lower.get(nf.lower())
                if existing:  # the "new" name is really an existing folder -> use it as a match
                    if existing not in result["matches"]:
                        result["matches"] = (result["matches"] + [existing])[:5]
                else:
                    result["new_folder"] = nf
            # Nothing fit and no name proposed: always give the create box a sensible default.
            if not result["matches"] and not result["new_folder"]:
                result["new_folder"] = _folder_name_from_subject(subject)
    except Exception:
        pass
    return result


_MOVE_TO_FOLDER_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $m = $ns.GetItemFromID($env:OL_ID)
    $dest = $null
    foreach ($f in $inbox.Folders) { if ($f.Name -eq $env:OL_FOLDER) { $dest = $f; break } }
    if ($dest -eq $null) { Write-Output ("OL_ERROR: folder not found: " + $env:OL_FOLDER); exit }
    [void]$m.Move($dest)
    Write-Output ("MOVED_OK|" + $env:OL_FOLDER)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


_ACTIVE_EMAIL_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    $item = $null
    $insp = $ol.ActiveInspector(); if ($insp -ne $null) { $item = $insp.CurrentItem }
    if ($item -eq $null) { $exp = $ol.ActiveExplorer(); if ($exp -ne $null -and $exp.Selection.Count -ge 1) { $item = $exp.Selection.Item(1) } }
    if ($item -ne $null -and $item.Class -eq 43) {
        $subj = ([string]$item.Subject) -replace "[\r\n|]", " "
        $sender = ([string]$item.SenderName) -replace "[\r\n|]", " "
        [Console]::Out.WriteLine("OK|" + [string]$item.EntryID + "|" + $subj + "|" + $sender)
    } else { [Console]::Out.WriteLine("NONE") }
} catch { [Console]::Out.WriteLine("NONE") }
'''


def active_email():
    """The email currently open or selected in Outlook as (entry_id, subject, sender), or None.
    Used by the global hotkey to file whatever the user is looking at right now."""
    if not IS_WINDOWS:
        return None
    out = _run_outlook_ps(_ACTIVE_EMAIL_PS, {}, show=False)
    for line in (out or "").splitlines():
        line = line.strip()
        if line.startswith("OK|"):
            parts = line.split("|", 3)
            if len(parts) == 4:
                return (parts[1], parts[2], parts[3])
    return None


def move_email_to_folder(eid, folder):
    """Move the Inbox email with this EntryID into the named Inbox subfolder."""
    return _run_outlook_ps(_MOVE_TO_FOLDER_PS, {"OL_ID": eid, "OL_FOLDER": folder}, show=False)
