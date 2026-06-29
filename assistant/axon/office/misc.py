"""Misc Windows/Office tools: file ops, Outlook contacts/tasks/attachments/meetings, scheduled tasks, system query, text-to-speech."""
import os
import json

from axon.util import _result
from axon.office._base import _run_ps, _resolve_path, _cfg_file, _move_doc_to_dot, _downloads

_FILE_OP_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $a = $env:FOP_ACTION
    $p = $env:FOP_PATH
    $dest = $env:FOP_DEST
    $pattern = $env:FOP_PATTERN
    $recurse = ($env:FOP_RECURSE -eq "1")
    switch ($a) {
        "list" {
            $items = Get-ChildItem -LiteralPath $p -ErrorAction Stop | Select-Object -First 300
            ($items | ForEach-Object { ($(if ($_.PSIsContainer){"[DIR] "}else{"      "})) + $_.Name }) -join "`n"
        }
        "search" {
            $items = Get-ChildItem -Path $p -Recurse:$recurse -Filter $pattern -ErrorAction SilentlyContinue | Select-Object -First 300
            ($items | ForEach-Object { $_.FullName }) -join "`n"
        }
        "move"   { Move-Item -LiteralPath $p -Destination $dest -Force; "MOVED $p -> $dest" }
        "copy"   { Copy-Item -LiteralPath $p -Destination $dest -Recurse -Force; "COPIED $p -> $dest" }
        "rename" { Rename-Item -LiteralPath $p -NewName $dest -Force; "RENAMED -> $dest" }
        "delete" { Remove-Item -LiteralPath $p -Recurse -Force; "DELETED $p" }
        "mkdir"  { New-Item -ItemType Directory -Force -Path $p | Out-Null; "CREATED FOLDER $p" }
        "zip"    { Compress-Archive -Path $p -DestinationPath $dest -Force; "ZIPPED -> $dest" }
        "unzip"  { Expand-Archive -LiteralPath $p -DestinationPath $dest -Force; "UNZIPPED -> $dest" }
        "exists" { if (Test-Path -LiteralPath $p) { "EXISTS" } else { "NOT FOUND" } }
        "read_text"  { Get-Content -Raw -LiteralPath $p }
        "write_text" { Set-Content -LiteralPath $p -Value $env:FOP_CONTENT -Encoding UTF8; "WROTE $p" }
        "open"   { Start-Process $p; "OPENED $p" }
        default  { "ERR: unknown action '$a'" }
    }
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def file_op(args):
    action = (args.get("action") or "").lower()
    if not action:
        return _result("Need an action (list/search/move/copy/rename/delete/mkdir/zip/unzip/exists/read_text/write_text/open).", True)
    out = _run_ps(_FILE_OP_PS, {
        "FOP_ACTION": action, "FOP_PATH": args.get("path") or "", "FOP_DEST": args.get("dest") or "",
        "FOP_PATTERN": args.get("pattern") or "*", "FOP_RECURSE": "1" if args.get("recursive") else "",
        "FOP_CONTENT": args.get("content") or "",
    })
    return _result(out[:40000] if out else "(no output)", out.startswith("ERR"))


_CONTACT_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    if ($env:CT_ACTION -eq "find") {
        $folder = $ns.GetDefaultFolder(10)  # olFolderContacts
        $q = ($env:CT_QUERY).ToLower()
        $out = New-Object System.Collections.ArrayList
        foreach ($c in $folder.Items) {
            try {
                $hay = (("" + $c.FullName + " " + $c.Email1Address + " " + $c.CompanyName)).ToLower()
                if (-not $q -or $hay.Contains($q)) {
                    [void]$out.Add($c.FullName + " <" + $c.Email1Address + ">" + $(if ($c.CompanyName){" - "+$c.CompanyName}else{""}))
                }
            } catch {}
            if ($out.Count -ge 50) { break }
        }
        if ($out.Count -eq 0) { Write-Output "No matching contacts." } else { Write-Output ($out -join "`n") }
    } else {
        $c = $ol.CreateItem(2)  # olContactItem
        if ($env:CT_NAME) { $c.FullName = $env:CT_NAME }
        if ($env:CT_EMAIL) { $c.Email1Address = $env:CT_EMAIL }
        if ($env:CT_COMPANY) { $c.CompanyName = $env:CT_COMPANY }
        if ($env:CT_PHONE) { $c.MobileTelephoneNumber = $env:CT_PHONE }
        if ($env:CT_TITLE) { $c.JobTitle = $env:CT_TITLE }
        $c.Save()
        Write-Output ("CONTACT_OK: saved " + $env:CT_NAME)
    }
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def outlook_contact(args):
    action = (args.get("action") or "create").lower()
    if action == "create" and not (args.get("name") or "").strip():
        return _result("Need a contact name to create.", True)
    out = _run_ps(_CONTACT_PS, {
        "CT_ACTION": action, "CT_QUERY": args.get("query") or "",
        "CT_NAME": args.get("name") or "", "CT_EMAIL": args.get("email") or "",
        "CT_COMPANY": args.get("company") or "", "CT_PHONE": args.get("phone") or "",
        "CT_TITLE": args.get("title") or "",
    })
    return _result(out, out.startswith("ERR"))


_TASK_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    if ($env:TK_ACTION -eq "list") {
        $folder = $ns.GetDefaultFolder(13)  # olFolderTasks
        $out = New-Object System.Collections.ArrayList
        foreach ($t in $folder.Items) {
            try {
                if (-not $t.Complete) {
                    $due = if ($t.DueDate -and $t.DueDate.Year -lt 4000) { $t.DueDate.ToString("yyyy-MM-dd") } else { "no due date" }
                    [void]$out.Add("[ ] " + $t.Subject + " (due " + $due + ")")
                }
            } catch {}
            if ($out.Count -ge 100) { break }
        }
        if ($out.Count -eq 0) { Write-Output "No open tasks." } else { Write-Output ($out -join "`n") }
    } else {
        $t = $ol.CreateItem(3)  # olTaskItem
        $t.Subject = $env:TK_SUBJECT
        if ($env:TK_BODY) { $t.Body = $env:TK_BODY }
        if ($env:TK_DUE) { $t.DueDate = [datetime]::Parse($env:TK_DUE) }
        if ($env:TK_REMIND) { $t.ReminderSet = $true; $t.ReminderTime = [datetime]::Parse($env:TK_REMIND) }
        $t.Save()
        Write-Output ("TASK_OK: created '" + $env:TK_SUBJECT + "'")
    }
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def outlook_task(args):
    action = (args.get("action") or "create").lower()
    if action == "create" and not (args.get("subject") or "").strip():
        return _result("Need a task subject.", True)
    out = _run_ps(_TASK_PS, {
        "TK_ACTION": action, "TK_SUBJECT": args.get("subject") or "", "TK_BODY": args.get("body") or "",
        "TK_DUE": args.get("due") or "", "TK_REMIND": args.get("reminder") or "",
    })
    return _result(out, out.startswith("ERR"))


_SAVE_ATTACH_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:AT_ID)
    $folder = $env:AT_FOLDER
    if (-not $folder) { $folder = [Environment]::GetFolderPath("MyDocuments") }
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    $n = 0; $names = New-Object System.Collections.ArrayList
    foreach ($a in $m.Attachments) {
        $p = Join-Path $folder $a.FileName
        $a.SaveAsFile($p); $n++; [void]$names.Add($a.FileName)
    }
    if ($n -eq 0) { Write-Output "No attachments on that email." }
    else { Write-Output ("ATTACH_OK: saved $n file(s) to $folder -> " + ($names -join ", ")) }
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def save_email_attachments(args):
    if not (args.get("id") or ""):
        return _result("Need the email id (from outlook_list_emails).", True)
    out = _run_ps(_SAVE_ATTACH_PS, {"AT_ID": args["id"], "AT_FOLDER": args.get("folder") or ""})
    return _result(out, out.startswith("ERR"))


_MEETING_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $item = $ns.GetItemFromID($env:MR_ID)
    $map = @{ "accept" = 3; "tentative" = 2; "decline" = 4 }   # olMeeting* response codes
    $code = $map[$env:MR_RESPONSE]
    if (-not $code) { Write-Output "ERR: response must be accept, tentative, or decline"; exit }
    $appt = $item
    try { if ($item.GetAssociatedAppointment) { $appt = $item.GetAssociatedAppointment($false) } } catch {}
    $resp = $appt.Respond($code, $true)
    if ($env:MR_SEND -eq "1") { $resp.Send() }
    Write-Output ("MEETING_OK: " + $env:MR_RESPONSE + "ed")
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def respond_to_meeting(args):
    if not (args.get("id") or ""):
        return _result("Need the meeting/email id.", True)
    resp = (args.get("response") or "").lower()
    if resp not in ("accept", "tentative", "decline"):
        return _result("response must be accept, tentative, or decline.", True)
    out = _run_ps(_MEETING_PS, {"MR_ID": args["id"], "MR_RESPONSE": resp,
                                "MR_SEND": "1" if args.get("send_response", True) else ""})
    return _result(out, out.startswith("ERR"))


_SCHEDULE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $name = $env:ST_NAME
    $cmd = $env:ST_CMD
    $argstr = $env:ST_ARGS
    if ($argstr) { $action = New-ScheduledTaskAction -Execute $cmd -Argument $argstr }
    else { $action = New-ScheduledTaskAction -Execute $cmd }
    $freq = $env:ST_FREQ
    $at = $env:ST_TIME
    switch ($freq) {
        "daily"  { $trigger = New-ScheduledTaskTrigger -Daily -At $at }
        "weekly" { $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $env:ST_DOW -At $at }
        default  { $trigger = New-ScheduledTaskTrigger -Once -At $at }
    }
    Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Force | Out-Null
    Write-Output ("SCHEDULE_OK: '" + $name + "' (" + $freq + " at " + $at + ")")
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def schedule_task(args):
    if not (args.get("name") or "").strip() or not (args.get("command") or "").strip() or not (args.get("time") or "").strip():
        return _result("Need name, command, and time (e.g. '08:00' or '2026-07-01 09:00').", True)
    out = _run_ps(_SCHEDULE_PS, {
        "ST_NAME": args["name"], "ST_CMD": args["command"], "ST_ARGS": args.get("arguments") or "",
        "ST_FREQ": (args.get("frequency") or "once").lower(), "ST_TIME": args["time"],
        "ST_DOW": args.get("day_of_week") or "Monday",
    })
    return _result(out, out.startswith("ERR"))


_SYSTEM_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $what = $env:SY_WHAT
    $f = $env:SY_FILTER
    switch ($what) {
        "processes" {
            $ps = Get-Process | Sort-Object -Property WS -Descending
            if ($f) { $ps = $ps | Where-Object { $_.Name -like "*$f*" } }
            ($ps | Select-Object -First 30 | ForEach-Object { "{0,-30} {1,8:N0} MB" -f $_.Name, ($_.WS/1MB) }) -join "`n"
        }
        "services" {
            $sv = Get-Service
            if ($f) { $sv = $sv | Where-Object { $_.Name -like "*$f*" -or $_.DisplayName -like "*$f*" } }
            ($sv | Select-Object -First 50 | ForEach-Object { "{0,-10} {1}" -f $_.Status, $_.DisplayName }) -join "`n"
        }
        "disks" {
            (Get-PSDrive -PSProvider FileSystem | ForEach-Object {
                "{0}: {1:N1} GB free of {2:N1} GB" -f $_.Name, ($_.Free/1GB), (($_.Free + $_.Used)/1GB) }) -join "`n"
        }
        default {
            $os = Get-CimInstance Win32_OperatingSystem
            $cs = Get-CimInstance Win32_ComputerSystem
            "Computer: $($cs.Name)`nOS: $($os.Caption) $($os.Version)`nRAM: {0:N1} GB`nUser: $($cs.UserName)`nUptime since: $($os.LastBootUpTime)" -f ($cs.TotalPhysicalMemory/1GB)
        }
    }
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def system_query(args):
    what = (args.get("what") or "system").lower()
    out = _run_ps(_SYSTEM_PS, {"SY_WHAT": what, "SY_FILTER": args.get("filter") or ""})
    return _result(out[:40000] if out else "(no output)", out.startswith("ERR"))


_SPEAK_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $v = New-Object -ComObject SAPI.SpVoice
    if ($env:TTS_RATE) { $v.Rate = [int]$env:TTS_RATE }
    $v.Speak($env:TTS_TEXT) | Out-Null
    Write-Output "SPOKE_OK"
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def speak(args):
    text = (args.get("text") or "").strip()
    if not text:
        return _result("Need text to speak.", True)
    out = _run_ps(_SPEAK_PS, {"TTS_TEXT": text, "TTS_RATE": args.get("rate") or ""})
    return _result("Spoke the text aloud." if out.startswith("SPOKE_OK") else out, out.startswith("ERR"))
