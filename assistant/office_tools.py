"""Office / Windows automation via COM and PowerShell — reliable, no UI clicking.

Covers: Excel (read/write/charts/PDF), Word (write/read/PDF), PowerPoint (decks/PDF),
convert-to-PDF, file operations (move/copy/search/zip...), Outlook contacts & tasks,
saving email attachments, responding to meeting invites, Windows scheduled tasks,
system queries (processes/services/info), and text-to-speech.

Each tool returns the MCP-style {"content":[{"type":"text","text":...}], "isError":bool}
dict that agent.py already understands. The module exports TOOLS (OpenAI tool schemas)
and DISPATCH ({name: function}) for agent.py to merge in.
"""
import os
import json
import tempfile
import subprocess


# --------------------------------------------------------------------------- helpers
def _result(text, is_error=False):
    return {"content": [{"type": "text", "text": str(text)}], "isError": bool(is_error)}


def _dot_mon_env():
    """The dot's monitor work-rect as DOT_MON_* env vars (so Office windows open on the dot's screen)."""
    try:
        import agent
        d = agent._dot_monitor_env()
        if d:
            return {"DOT_MON_L": str(d["L"]), "DOT_MON_T": str(d["T"]),
                    "DOT_MON_R": str(d["R"]), "DOT_MON_B": str(d["B"])}
    except Exception:
        pass
    return {}


def _move_doc_to_dot(path):
    """Move a just-opened Office document window onto the dot's monitor (by its filename in the title)."""
    try:
        import agent
        agent._move_window_to_dot(os.path.splitext(os.path.basename(path))[0])
    except Exception:
        pass


def _run_ps(script, env=None, timeout=240):
    """Run a PowerShell script (written to a temp .ps1) with extra env vars. Returns output."""
    e = dict(os.environ)
    for k, v in _dot_mon_env().items():
        e[k] = v
    if env:
        for k, v in env.items():
            e[k] = "" if v is None else str(v)
    f = tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig")
    f.write(script)
    f.close()
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", f.name],
            capture_output=True, text=True, env=e, timeout=timeout,
        )
        out = (p.stdout or "").strip()
        err = (p.stderr or "").strip()
        return out + (("\n" + err) if err and not out else "")
    except subprocess.TimeoutExpired:
        return "ERR: the operation timed out"
    finally:
        try:
            os.unlink(f.name)
        except Exception:
            pass


def _downloads():
    return os.path.join(os.path.expanduser("~"), "Downloads")


def _resolve_path(path, default_name, default_ext):
    if not path:
        path = default_name
    if not os.path.splitext(path)[1]:
        path += default_ext
    if not os.path.isabs(path):
        path = os.path.join(_downloads(), path)
    return path


def _cfg_file(data):
    f = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8")
    json.dump(data, f)
    f.close()
    return f.name


# =========================================================================== EXCEL
_EXCEL_WRITE_PS = r'''
$ErrorActionPreference = "Stop"
Add-Type @"
using System; using System.Runtime.InteropServices;
public class XlMove {
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool rp);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
  public static void Move(IntPtr h, int ml, int mt, int mr, int mb){
    try { SetThreadDpiAwarenessContext((IntPtr)(-4)); } catch {}
    int mw=mr-ml, mh=mb-mt; int w=(int)(mw*0.82), ht=(int)(mh*0.86);
    MoveWindow(h, ml+((mw-w)/2), mt+((mh-ht)/2), w, ht, true);
  }
}
"@
try {
    $cfg = Get-Content -Raw -LiteralPath $env:CFG_PATH | ConvertFrom-Json
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false   # build hidden, then show on the dot's monitor
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    if ($cfg.sheet_name) { $ws.Name = [string]$cfg.sheet_name }
    $r = 1
    if ($cfg.title) {
        $ws.Cells.Item($r, 1) = [string]$cfg.title
        $ws.Cells.Item($r, 1).Font.Bold = $true
        $ws.Cells.Item($r, 1).Font.Size = 14
        $r += 2
    }
    $headerRow = 0
    $cols = 1
    if ($cfg.headers) {
        $cols = $cfg.headers.Count
        for ($c = 0; $c -lt $cfg.headers.Count; $c++) {
            $ws.Cells.Item($r, $c + 1) = [string]$cfg.headers[$c]
            $cell = $ws.Cells.Item($r, $c + 1)
            $cell.Font.Bold = $true
            $cell.Interior.Color = 15917529
        }
        $headerRow = $r; $r++
    }
    $firstData = $r
    if ($cfg.rows) {
        foreach ($row in $cfg.rows) {
            for ($c = 0; $c -lt $row.Count; $c++) { $ws.Cells.Item($r, $c + 1) = $row[$c] }
            if ($row.Count -gt $cols) { $cols = $row.Count }
            $r++
        }
    }
    $lastRow = $r - 1
    $ws.Columns.AutoFit() | Out-Null
    if ($cfg.chart -and $lastRow -ge $firstData -and $headerRow -gt 0) {
        $co = $ws.ChartObjects().Add(420, 20, 380, 240)
        $ch = $co.Chart
        $rng = $ws.Range($ws.Cells.Item($headerRow, 1), $ws.Cells.Item($lastRow, $cols))
        $ch.SetSourceData($rng)
        switch ([string]$cfg.chart) {
            "bar"  { $ch.ChartType = 57 }
            "line" { $ch.ChartType = 4 }
            "pie"  { $ch.ChartType = 5 }
            default { $ch.ChartType = 51 }
        }
    }
    # Position the (still-hidden) Excel window on the dot's monitor, THEN show it -> no flash.
    if ($env:DOT_MON_R) {
        try { [XlMove]::Move([IntPtr]$xl.Hwnd, [int]$env:DOT_MON_L, [int]$env:DOT_MON_T, [int]$env:DOT_MON_R, [int]$env:DOT_MON_B) } catch {}
    }
    $xl.Visible = $true
    $path = $cfg.path
    $wb.SaveAs($path)
    $msg = "EXCEL_OK: saved " + $path
    if ($cfg.pdf) {
        $pdf = [System.IO.Path]::ChangeExtension($path, ".pdf")
        $wb.ExportAsFixedFormat(0, $pdf)
        $msg += " | PDF: " + $pdf
    }
    $wb.Save()   # leave the workbook OPEN and visible so the user can see / review it
    Write-Output $msg
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''

_EXCEL_READ_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($env:XL_PATH, $false, $true)
    if ($env:XL_SHEET) { $ws = $wb.Worksheets.Item($env:XL_SHEET) } else { $ws = $wb.Worksheets.Item(1) }
    $used = $ws.UsedRange
    $rows = [int]$used.Rows.Count
    $colsN = [int]$used.Columns.Count
    $maxr = [int]$env:XL_MAXROWS; if ($maxr -le 0) { $maxr = 200 }
    if ($rows -lt $maxr) { $maxr = $rows }
    $data = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $maxr; $i++) {
        $line = New-Object System.Collections.ArrayList
        for ($j = 1; $j -le $colsN; $j++) { [void]$line.Add([string]$used.Cells.Item($i, $j).Text) }
        [void]$data.Add($line)
    }
    $wb.Close($false); $xl.Quit()
    Write-Output ($data | ConvertTo-Json -Depth 4 -Compress)
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def excel(args):
    action = (args.get("action") or "write").lower()
    if action == "read":
        path = args.get("path") or ""
        if not path or not os.path.isabs(path):
            path = _resolve_path(path, "", "")
        if not os.path.exists(path):
            return _result(f"File not found: {path}", True)
        out = _run_ps(_EXCEL_READ_PS, {"XL_PATH": path, "XL_SHEET": args.get("sheet_name") or "",
                                       "XL_MAXROWS": args.get("max_rows") or 200})
        return _result(out[:40000], out.startswith("ERR"))
    path = _resolve_path(args.get("path"), "report", ".xlsx")
    cfg = {
        "path": path, "sheet_name": args.get("sheet_name"), "title": args.get("title"),
        "headers": args.get("headers"), "rows": args.get("rows") or [],
        "chart": args.get("chart"), "pdf": bool(args.get("pdf")),
    }
    cf = _cfg_file(cfg)
    try:
        out = _run_ps(_EXCEL_WRITE_PS, {"CFG_PATH": cf})
    finally:
        try:
            os.unlink(cf)
        except Exception:
            pass
    return _result(out, out.startswith("ERR"))


# =========================================================================== WORD
_WORD_WRITE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $cfg = Get-Content -Raw -LiteralPath $env:CFG_PATH | ConvertFrom-Json
    $w = New-Object -ComObject Word.Application
    $w.Visible = $true   # show it — nothing happens behind closed doors
    $doc = $w.Documents.Add()
    $sel = $w.Selection
    if ($cfg.title) { $sel.Style = "Title"; $sel.TypeText([string]$cfg.title); $sel.TypeParagraph() }
    foreach ($p in $cfg.paragraphs) {
        $style = "Normal"
        if ($p.style) {
            switch (([string]$p.style).ToLower()) {
                "heading1" { $style = "Heading 1" }
                "heading2" { $style = "Heading 2" }
                "heading3" { $style = "Heading 3" }
                "bullet"   { $style = "List Bullet" }
                "number"   { $style = "List Number" }
                default    { $style = "Normal" }
            }
        }
        try { $sel.Style = $style } catch { $sel.Style = "Normal" }
        $sel.TypeText([string]$p.text)
        $sel.TypeParagraph()
    }
    $path = $cfg.path
    $doc.SaveAs2($path)
    $msg = "WORD_OK: saved " + $path
    if ($cfg.pdf) {
        $pdf = [System.IO.Path]::ChangeExtension($path, ".pdf")
        $doc.ExportAsFixedFormat($pdf, 17)
        $msg += " | PDF: " + $pdf
    }
    $doc.Save()   # leave the document OPEN and visible for the user
    Write-Output $msg
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''

_WORD_READ_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $doc = $w.Documents.Open($env:DOC_PATH, $false, $true)
    $text = $doc.Content.Text
    $doc.Close($false); $w.Quit()
    Write-Output $text
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def word(args):
    action = (args.get("action") or "write").lower()
    if action == "read":
        path = args.get("path") or ""
        if not os.path.isabs(path):
            path = _resolve_path(path, "", "")
        if not os.path.exists(path):
            return _result(f"File not found: {path}", True)
        out = _run_ps(_WORD_READ_PS, {"DOC_PATH": path})
        return _result(out[:40000], out.startswith("ERR"))
    path = _resolve_path(args.get("path"), "document", ".docx")
    paragraphs = args.get("paragraphs")
    if not paragraphs and args.get("body"):
        paragraphs = [{"text": ln, "style": "normal"} for ln in str(args["body"]).split("\n")]
    cfg = {"path": path, "title": args.get("title"), "paragraphs": paragraphs or [], "pdf": bool(args.get("pdf"))}
    cf = _cfg_file(cfg)
    try:
        out = _run_ps(_WORD_WRITE_PS, {"CFG_PATH": cf})
    finally:
        try:
            os.unlink(cf)
        except Exception:
            pass
    _move_doc_to_dot(path)  # put the Word window on the dot's monitor
    return _result(out, out.startswith("ERR"))


# =========================================================================== POWERPOINT
_PPT_WRITE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $cfg = Get-Content -Raw -LiteralPath $env:CFG_PATH | ConvertFrom-Json
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Add($true)
    $i = 1
    foreach ($s in $cfg.slides) {
        $slide = $pres.Slides.Add($i, 2)  # 2 = ppLayoutText (title + body)
        try { $slide.Shapes.Title.TextFrame.TextRange.Text = [string]$s.title } catch {}
        if ($s.bullets) {
            $body = ($s.bullets -join "`r`n")
            try { $slide.Shapes.Item(2).TextFrame.TextRange.Text = $body } catch {}
        }
        $i++
    }
    $path = $cfg.path
    $pres.SaveAs($path)
    $msg = "PPT_OK: saved " + $path
    if ($cfg.pdf) {
        $pdf = [System.IO.Path]::ChangeExtension($path, ".pdf")
        $pres.SaveAs($pdf, 32)  # 32 = ppSaveAsPDF
        $msg += " | PDF: " + $pdf
    }
    $pp.Activate()   # leave the deck OPEN and visible for the user
    Write-Output $msg
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def powerpoint(args):
    path = _resolve_path(args.get("path"), "presentation", ".pptx")
    slides = args.get("slides") or []
    if not slides:
        return _result("Need at least one slide (slides: [{title, bullets:[...]}]).", True)
    cfg = {"path": path, "slides": slides, "pdf": bool(args.get("pdf"))}
    cf = _cfg_file(cfg)
    try:
        out = _run_ps(_PPT_WRITE_PS, {"CFG_PATH": cf})
    finally:
        try:
            os.unlink(cf)
        except Exception:
            pass
    _move_doc_to_dot(path)  # put the PowerPoint window on the dot's monitor
    return _result(out, out.startswith("ERR"))


# =========================================================================== CONVERT TO PDF
_TO_PDF_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $src = $env:SRC_PATH
    $pdf = [System.IO.Path]::ChangeExtension($src, ".pdf")
    $ext = [System.IO.Path]::GetExtension($src).ToLower()
    if ($ext -in ".doc", ".docx", ".rtf", ".txt", ".odt") {
        $w = New-Object -ComObject Word.Application; $w.Visible = $false
        $doc = $w.Documents.Open($src, $false, $true)
        $doc.ExportAsFixedFormat($pdf, 17)
        $doc.Close($false); $w.Quit()
    } elseif ($ext -in ".xls", ".xlsx", ".csv") {
        $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($src, $false, $true)
        $wb.ExportAsFixedFormat(0, $pdf)
        $wb.Close($false); $xl.Quit()
    } elseif ($ext -in ".ppt", ".pptx") {
        $pp = New-Object -ComObject PowerPoint.Application
        $pres = $pp.Presentations.Open($src, $true, $false, $false)
        $pres.SaveAs($pdf, 32)
        $pres.Close(); $pp.Quit()
    } else { Write-Output ("ERR: unsupported file type for PDF: " + $ext); exit }
    Write-Output ("PDF_OK: " + $pdf)
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def convert_to_pdf(args):
    path = args.get("path") or ""
    if not os.path.isabs(path):
        path = _resolve_path(path, "", "")
    if not os.path.exists(path):
        return _result(f"File not found: {path}", True)
    out = _run_ps(_TO_PDF_PS, {"SRC_PATH": path})
    return _result(out, out.startswith("ERR"))


# =========================================================================== FILE OPERATIONS
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


# =========================================================================== OUTLOOK CONTACTS
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


# =========================================================================== OUTLOOK TASKS
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


# =========================================================================== SAVE ATTACHMENTS
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


# =========================================================================== RESPOND TO MEETING
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


# =========================================================================== SCHEDULED TASK
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


# =========================================================================== SYSTEM QUERY
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


# =========================================================================== TEXT TO SPEECH
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


# =========================================================================== TOOL SCHEMAS
def _fn(name, description, properties, required):
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties, "required": required},
    }}


TOOLS = [
    _fn("excel",
        "Create or read an Excel workbook (no UI). action='write' builds a .xlsx from headers+rows with optional title, chart (bar/line/pie/column) and PDF export; action='read' returns a workbook's cell data. Great for professional data reports.",
        {"action": {"type": "string", "enum": ["write", "read"]},
         "path": {"type": "string", "description": "File path or name (defaults to Downloads, .xlsx)"},
         "title": {"type": "string"},
         "sheet_name": {"type": "string"},
         "headers": {"type": "array", "items": {"type": "string"}},
         "rows": {"type": "array", "items": {"type": "array", "items": {}}, "description": "Rows of cell values"},
         "chart": {"type": "string", "enum": ["bar", "line", "pie", "column"]},
         "pdf": {"type": "boolean", "description": "Also export a PDF"},
         "max_rows": {"type": "integer", "description": "read: max rows to return"}},
        ["action"]),
    _fn("word",
        "Create or read a Word document (no UI). action='write' builds a .docx from a title + paragraphs (each {text, style: normal/heading1/heading2/bullet/number}) with optional PDF export; action='read' extracts text from a .docx.",
        {"action": {"type": "string", "enum": ["write", "read"]},
         "path": {"type": "string", "description": "File path or name (defaults to Downloads, .docx)"},
         "title": {"type": "string"},
         "paragraphs": {"type": "array", "items": {"type": "object", "properties": {
             "text": {"type": "string"}, "style": {"type": "string"}}}},
         "body": {"type": "string", "description": "Alternative to paragraphs: plain text (one paragraph per line)"},
         "pdf": {"type": "boolean"}},
        ["action"]),
    _fn("powerpoint",
        "Create a PowerPoint deck (no UI) from slides (each {title, bullets:[...]}). Optional PDF export.",
        {"path": {"type": "string", "description": "File path or name (defaults to Downloads, .pptx)"},
         "slides": {"type": "array", "items": {"type": "object", "properties": {
             "title": {"type": "string"}, "bullets": {"type": "array", "items": {"type": "string"}}}}},
         "pdf": {"type": "boolean"}},
        ["slides"]),
    _fn("convert_to_pdf",
        "Convert an existing Office file (Word/Excel/PowerPoint/CSV/txt) to PDF in the same folder.",
        {"path": {"type": "string", "description": "Path to the file to convert"}},
        ["path"]),
    _fn("file_op",
        "Manage files and folders (no UI). actions: list, search (pattern + recursive), move, copy, rename, delete, mkdir, zip, unzip, exists, read_text, write_text, open.",
        {"action": {"type": "string", "enum": ["list", "search", "move", "copy", "rename", "delete", "mkdir", "zip", "unzip", "exists", "read_text", "write_text", "open"]},
         "path": {"type": "string", "description": "Target path (source for move/copy/zip)"},
         "dest": {"type": "string", "description": "Destination path / new name / zip output"},
         "pattern": {"type": "string", "description": "search: filename pattern e.g. *.pdf"},
         "recursive": {"type": "boolean", "description": "search: include subfolders"},
         "content": {"type": "string", "description": "write_text: the text to write"}},
        ["action"]),
    _fn("outlook_contact",
        "Create or find an Outlook contact. action='create' (name + optional email/phone/company/title); action='find' (query matches name/email/company).",
        {"action": {"type": "string", "enum": ["create", "find"]},
         "name": {"type": "string"}, "email": {"type": "string"}, "phone": {"type": "string"},
         "company": {"type": "string"}, "title": {"type": "string"},
         "query": {"type": "string", "description": "find: search text"}},
        ["action"]),
    _fn("outlook_task",
        "Create or list Outlook tasks. action='create' (subject + optional body/due/reminder, dates like '2026-07-01 09:00'); action='list' (open tasks).",
        {"action": {"type": "string", "enum": ["create", "list"]},
         "subject": {"type": "string"}, "body": {"type": "string"},
         "due": {"type": "string"}, "reminder": {"type": "string"}},
        ["action"]),
    _fn("save_email_attachments",
        "Save all attachments from an Outlook email (by its id from outlook_list_emails) to a folder (defaults to Documents).",
        {"id": {"type": "string"}, "folder": {"type": "string"}},
        ["id"]),
    _fn("respond_to_meeting",
        "Accept, tentatively accept, or decline a meeting invitation (by the email/meeting id).",
        {"id": {"type": "string"}, "response": {"type": "string", "enum": ["accept", "tentative", "decline"]},
         "send_response": {"type": "boolean", "description": "Send the response to the organizer (default true)"}},
        ["id", "response"]),
    _fn("schedule_task",
        "Create a Windows scheduled task to run a command on a schedule. frequency: once/daily/weekly; time like '08:00' or '2026-07-01 09:00'.",
        {"name": {"type": "string"}, "command": {"type": "string", "description": "Program/exe to run"},
         "arguments": {"type": "string"}, "time": {"type": "string"},
         "frequency": {"type": "string", "enum": ["once", "daily", "weekly"]},
         "day_of_week": {"type": "string", "description": "weekly: e.g. Monday"}},
        ["name", "command", "time"]),
    _fn("system_query",
        "Query the computer (no UI). what: 'processes' (top by memory), 'services', 'disks', or 'system' (OS/RAM/uptime). Optional filter.",
        {"what": {"type": "string", "enum": ["processes", "services", "disks", "system"]},
         "filter": {"type": "string"}},
        []),
    _fn("speak",
        "Speak text aloud through the computer's speakers (text-to-speech).",
        {"text": {"type": "string"}, "rate": {"type": "integer", "description": "-10 (slow) to 10 (fast)"}},
        ["text"]),
]


DISPATCH = {
    "excel": excel,
    "word": word,
    "powerpoint": powerpoint,
    "convert_to_pdf": convert_to_pdf,
    "file_op": file_op,
    "outlook_contact": outlook_contact,
    "outlook_task": outlook_task,
    "save_email_attachments": save_email_attachments,
    "respond_to_meeting": respond_to_meeting,
    "schedule_task": schedule_task,
    "system_query": system_query,
    "speak": speak,
}
