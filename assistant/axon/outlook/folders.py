"""Outlook folder management (create / delete / list) and inbox rules."""
from axon.config import IS_WINDOWS
from axon.util import _result
from axon.outlook._base import _run_outlook_ps


_CREATE_FOLDERS_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
    $created = @(); $existing = @()
    foreach ($name in ($env:OCU_FOLDERS -split [regex]::Escape("|"))) {
        if (-not $name) { continue }
        $found = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -eq $name) { $found = $f; break } }
        if ($found) { $existing += $name } else { [void]$inbox.Folders.Add($name); $created += $name }
    }
    Write-Output ("CREATED: " + ($created -join ", "))
    Write-Output ("ALREADY_EXISTED: " + ($existing -join ", "))
} catch {
    Write-Output ("FOLDER_ERROR: " + $_.Exception.Message)
}
'''


def _create_outlook_folders(args):
    folders = args.get("folders") or []
    if isinstance(folders, str):
        folders = [folders]
    folders = [f.strip() for f in folders if f and f.strip()]
    if not folders:
        return _result("No folder names provided.", True)
    if not IS_WINDOWS:
        return _result("Outlook folders are supported on the Windows app only.", True)
    out = _run_outlook_ps(_CREATE_FOLDERS_PS, {"OCU_FOLDERS": "|".join(folders)})
    err = ("FOLDER_ERROR" in out) or out.startswith("OL_ERROR")
    return _result(out or "No output from Outlook.", err)


_DELETE_FOLDERS_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
    $root = $inbox.Parent
    $keep = ($env:OCU_KEEP -eq "1")    # keep the emails (move them out) before deleting the folder?
    # Where kept emails go (chosen by the user; defaults to the Inbox).
    $dest = $inbox; $destName = "Inbox"
    if ($keep -and $env:OCU_MOVE_TO) {
        $d = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $env:OCU_MOVE_TO) { $d = $f; break } }
        if (-not $d) { foreach ($f in $root.Folders) { if ($f.Name -ieq $env:OCU_MOVE_TO) { $d = $f; break } } }
        if (-not $d) { Write-Output ("FOLDER_ERROR: destination folder '" + $env:OCU_MOVE_TO + "' not found"); exit }
        $dest = $d; $destName = $env:OCU_MOVE_TO
    }
    $deleted = @(); $notfound = @()
    $targets = $env:OCU_FOLDERS -split [regex]::Escape("|")
    if ($targets -contains "*") {   # "*" = every subfolder directly under the Inbox
        $targets = @(); foreach ($f in $inbox.Folders) { $targets += $f.Name }
    }
    foreach ($name in $targets) {
        if (-not $name) { continue }
        $found = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $name) { $found = $f; break } }
        if (-not $found) { foreach ($f in $root.Folders) { if ($f.Name -ieq $name) { $found = $f; break } } }
        if (-not $found) { $notfound += $name; continue }
        $moved = 0
        if ($keep) {
            for ($i = $found.Items.Count; $i -ge 1; $i--) {
                try { $found.Items.Item($i).Move($dest) | Out-Null; $moved++ } catch {}
            }
        }
        $found.Delete()
        if ($keep) { $deleted += ($name + " (" + $moved + " emails moved to " + $destName + ")") } else { $deleted += ($name + " (with contents)") }
    }
    Write-Output ("DELETED: " + ($deleted -join "; "))
    if ($notfound.Count -gt 0) { Write-Output ("NOT_FOUND: " + ($notfound -join ", ")) }
} catch {
    Write-Output ("FOLDER_ERROR: " + $_.Exception.Message)
}
'''


def _delete_outlook_folders(args):
    folders = args.get("folders") or []
    if isinstance(folders, str):
        folders = [folders]
    folders = [f.strip() for f in folders if f and f.strip()]
    if not folders:
        return _result("No folder names provided.", True)
    keep = args.get("keep_contents")          # the agent sets this from the user's request
    keep = True if keep is None else bool(keep)  # safe fallback only when unspecified
    out = _run_outlook_ps(_DELETE_FOLDERS_PS, {
        "OCU_FOLDERS": "|".join(folders),
        "OCU_KEEP": "1" if keep else "",
        "OCU_MOVE_TO": args.get("move_to") or "",
    })
    err = ("FOLDER_ERROR" in out) or out.startswith("OL_ERROR")
    return _result(out or "No output from Outlook.", err)


_LIST_FOLDERS_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
    $out = New-Object System.Collections.ArrayList
    function Walk($folder, $indent) {
        foreach ($f in $folder.Folders) {
            [void]$out.Add($indent + $f.Name + " (" + $f.Items.Count + " emails)")
            if ($f.Folders.Count -gt 0) { Walk $f ($indent + "    ") }
        }
    }
    Walk $inbox ""
    if ($out.Count -eq 0) { Write-Output "No subfolders under the Inbox." }
    else { Write-Output ($out -join "`n") }
} catch { Write-Output ("FOLDER_ERROR: " + $_.Exception.Message) }
'''


def _list_outlook_folders(args):
    out = _run_outlook_ps(_LIST_FOLDERS_PS, {})
    err = ("FOLDER_ERROR" in out) or out.startswith("OL_ERROR")
    return _result(out or "No output from Outlook.", err)


_OUTLOOK_RULE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $store = $ns.DefaultStore
    $rules = $store.GetRules()
    $rule = $rules.Create($env:OL_RULE_NAME, 0)   # 0 = olRuleReceive
    $hasCond = $false; $hasAct = $false
    if ($env:OL_FROM) { $c = $rule.Conditions.SenderAddress; $c.Enabled = $true; $c.Address = @($env:OL_FROM -split [regex]::Escape("|")); $hasCond = $true }
    if ($env:OL_SUBJECT) { $c = $rule.Conditions.Subject; $c.Enabled = $true; $c.Text = @($env:OL_SUBJECT -split [regex]::Escape("|")); $hasCond = $true }
    if ($env:OL_MOVE) {
        $inbox = $ns.GetDefaultFolder(6); $dest = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $env:OL_MOVE) { $dest = $f; break } }
        if (-not $dest) { foreach ($f in $inbox.Parent.Folders) { if ($f.Name -ieq $env:OL_MOVE) { $dest = $f; break } } }
        if (-not $dest) { Write-Output "OL_ERROR: folder '$($env:OL_MOVE)' not found"; exit }
        $a = $rule.Actions.MoveToFolder; $a.Enabled = $true; $a.Folder = $dest; $hasAct = $true
    }
    if ($env:OL_CATEGORY) { $a = $rule.Actions.AssignToCategory; $a.Enabled = $true; $a.Categories = @($env:OL_CATEGORY); $hasAct = $true }
    if ($env:OL_DELETE -eq "1") { $a = $rule.Actions.Delete; $a.Enabled = $true; $hasAct = $true }
    if (-not $hasCond) { Write-Output "OL_ERROR: a rule needs at least one condition (from or subject)"; exit }
    if (-not $hasAct) { Write-Output "OL_ERROR: a rule needs at least one action (move, category, or delete)"; exit }
    $rules.Save()
    Write-Output ("RULE_OK: created '" + $env:OL_RULE_NAME + "'")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _create_outlook_rule(args):
    if not (args.get("name") or "").strip():
        return _result("Rule needs a name.", True)
    out = _run_outlook_ps(_OUTLOOK_RULE_PS, {
        "OL_RULE_NAME": args["name"],
        "OL_FROM": "|".join(args["from_contains"]) if isinstance(args.get("from_contains"), list) else (args.get("from_contains") or ""),
        "OL_SUBJECT": "|".join(args["subject_contains"]) if isinstance(args.get("subject_contains"), list) else (args.get("subject_contains") or ""),
        "OL_MOVE": args.get("move_to_folder") or "",
        "OL_CATEGORY": args.get("category") or "",
        "OL_DELETE": "1" if args.get("delete") else "",
    })
    return _result(out, out.startswith("OL_ERROR"))
