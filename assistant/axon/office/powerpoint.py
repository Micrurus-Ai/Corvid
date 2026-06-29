"""PowerPoint automation: build slide decks (title + bullets)."""
import os
import json

from axon.util import _result
from axon.office._base import _run_ps, _resolve_path, _cfg_file, _move_doc_to_dot, _downloads

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
