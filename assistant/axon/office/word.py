"""Word automation: create/write documents and read text."""
import os
import json

from axon.util import _result
from axon.office._base import _run_ps, _resolve_path, _cfg_file, _move_doc_to_dot, _downloads

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
