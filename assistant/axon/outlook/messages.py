"""Working with received mail: list, move, delete, save, open, mark read, categorize."""
from axon.config import TOOL_TEXT_LIMIT
from axon.util import _result
from axon.screen import _move_window_to_dot
from axon.outlook._base import _run_outlook_ps, _OL_WIN_CLASS


_OUTLOOK_LIST_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $name = $env:OL_FOLDER
    $folder = $inbox
    if ($name -and $name -ne "Inbox") {
        $folder = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } }
        if (-not $folder) { foreach ($f in $inbox.Parent.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } } }
        if (-not $folder) { Write-Output "OL_ERROR: folder '$name' not found"; exit }
    }
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    $q = $env:OL_QUERY
    $unread = ($env:OL_UNREAD -eq "1")
    $limit = [int]$env:OL_LIMIT; if ($limit -le 0) { $limit = 25 }
    $out = @(); $n = 0
    foreach ($m in $items) {
        if ($m.Class -ne 43) { continue }
        if ($unread -and -not $m.UnRead) { continue }
        if ($q) { if (-not (($m.Subject -like "*$q*") -or ($m.SenderName -like "*$q*"))) { continue } }
        $out += [pscustomobject]@{ id=$m.EntryID; from=$m.SenderName; subject=$m.Subject; received=$m.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); unread=[bool]$m.UnRead }
        $n++; if ($n -ge $limit) { break }
    }
    if ($out.Count -eq 0) { Write-Output "[]" } else { $out | ConvertTo-Json -Depth 3 -Compress }
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


_OUTLOOK_MOVE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $name = $env:OL_FOLDER
    $folder = $null
    foreach ($f in $inbox.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } }
    if (-not $folder) { foreach ($f in $inbox.Parent.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } } }
    if (-not $folder) { Write-Output "OL_ERROR: destination folder '$name' not found"; exit }
    $moved = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); [void]$m.Move($folder); $moved++ } catch {}
    }
    Write-Output ("MOVED: $moved to '$name'")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


_OUTLOOK_DELETE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $del = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); $m.Delete(); $del++ } catch {}
    }
    Write-Output ("DELETED: $del (moved to Deleted Items)")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


_SHOW_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
''' + _OL_WIN_CLASS + r'''
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    $m.Display($false)
    Start-Sleep -Milliseconds 500
    [OcuWinF]::Front([string]$m.Subject)
    Start-Sleep -Milliseconds 1100
    Write-Output ("SHOWN|" + $m.Subject)
} catch {}
'''


def _show_email(eid):
    """Open the email in its own window and bring it to the front, so the user SEES which email
    is about to be acted on (move/delete/categorize/mark)."""
    if eid:
        out = _run_outlook_ps(_SHOW_EMAIL_PS, {"OL_ID": eid}, show=False)
        if "SHOWN|" in out:
            subj = out.split("SHOWN|", 1)[1].strip().splitlines()[0]
            _move_window_to_dot(subj)


def _outlook_list_emails(args):
    out = _run_outlook_ps(_OUTLOOK_LIST_PS, {
        "OL_FOLDER": args.get("folder") or "Inbox",
        "OL_QUERY": args.get("query") or "",
        "OL_UNREAD": "1" if args.get("unread_only") else "",
        "OL_LIMIT": args.get("limit") or 25,
    })
    return _result(out[:TOOL_TEXT_LIMIT], out.startswith("OL_ERROR"))


def _outlook_move_emails(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids:
        return _result("No email ids provided.", True)
    if not (args.get("to_folder") or "").strip():
        return _result("No destination folder provided.", True)
    _show_email(ids[0])  # show the email being filed
    out = _run_outlook_ps(_OUTLOOK_MOVE_PS, {"OL_IDS": "|".join(ids), "OL_FOLDER": args["to_folder"]}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


def _outlook_delete_emails(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids:
        return _result("No email ids provided.", True)
    _show_email(ids[0])  # show the email before deleting it
    out = _run_outlook_ps(_OUTLOOK_DELETE_PS, {"OL_IDS": "|".join(ids)}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


_OUTLOOK_SAVE_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    $fmt = 9; $ext = ".msg"          # olMSGUnicode
    switch ($env:OL_FORMAT) {
        "txt"  { $fmt = 0; $ext = ".txt" }
        "html" { $fmt = 5; $ext = ".html" }
        "rtf"  { $fmt = 1; $ext = ".rtf" }
        default { $fmt = 9; $ext = ".msg" }
    }
    $dir = $env:OL_DIR
    if (-not $dir) { $dir = Join-Path $env:USERPROFILE "Downloads" }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $subj = $m.Subject; if (-not $subj) { $subj = "email" }
    $safe = ($subj -replace '[\\/:*?"<>|]', '_').Trim()
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    $path = Join-Path $dir ($safe + $ext)
    $m.SaveAs($path, $fmt)
    Write-Output ("SAVED_OK: " + $path)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _outlook_save_email(args):
    if not (args.get("id") or ""):
        return _result("Need the email id (from outlook_list_emails).", True)
    fmt = (args.get("format") or "msg").lower()
    if fmt not in ("msg", "txt", "html", "rtf"):
        fmt = "msg"
    out = _run_outlook_ps(_OUTLOOK_SAVE_EMAIL_PS, {
        "OL_ID": args["id"], "OL_FORMAT": fmt, "OL_DIR": args.get("folder") or "",
    })
    return _result(out, out.startswith("OL_ERROR"))


_OUTLOOK_ACTIVE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $item = $null
    try { $insp = $ol.ActiveInspector(); if ($insp) { $item = $insp.CurrentItem } } catch {}
    if (-not $item) {
        try { $ex = $ol.ActiveExplorer(); if ($ex -and $ex.Selection.Count -gt 0) { $item = $ex.Selection.Item(1) } } catch {}
    }
    if (-not $item) { Write-Output "NONE: no email is open or selected in Outlook"; exit }
    $from = ""; try { $from = $item.SenderName } catch {}
    $addr = ""; try { $addr = $item.SenderEmailAddress } catch {}
    $body = ""; try { $body = [string]$item.Body } catch {}
    $body = ($body -replace '\s+', ' ').Trim()
    if ($body.Length -gt 400) { $body = $body.Substring(0, 400) }
    $info = @{ id = $item.EntryID; subject = $item.Subject; from = $from; sender_email = $addr; body_preview = $body } | ConvertTo-Json -Compress
    Write-Output ("ACTIVE: " + $info)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _get_open_email(args):
    # Read-only: do NOT surface/refocus Outlook (the user is already looking at the email).
    out = _run_outlook_ps(_OUTLOOK_ACTIVE_PS, {}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


_OUTLOOK_MARK_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $read = ($env:OL_READ -eq "1")
    $n = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); $m.UnRead = (-not $read); $m.Save(); $n++ } catch {}
    }
    Write-Output ("MARKED: $n as " + $(if ($read) {"read"} else {"unread"}))
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


_OUTLOOK_CATEGORIZE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $n = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); $m.Categories = $env:OL_CATEGORY; $m.Save(); $n++ } catch {}
    }
    Write-Output ("CATEGORIZED: $n as '" + $env:OL_CATEGORY + "'")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _outlook_mark_read(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids:
        return _result("No email ids provided.", True)
    read = args.get("read", True)
    _show_email(ids[0])  # show the email being marked
    out = _run_outlook_ps(_OUTLOOK_MARK_PS, {"OL_IDS": "|".join(ids), "OL_READ": "1" if read else ""}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


def _outlook_categorize(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids or not (args.get("category") or "").strip():
        return _result("Need email ids and a category.", True)
    _show_email(ids[0])  # show the email being categorized
    out = _run_outlook_ps(_OUTLOOK_CATEGORIZE_PS, {"OL_IDS": "|".join(ids), "OL_CATEGORY": args["category"]}, show=False)
    return _result(out, out.startswith("OL_ERROR"))
