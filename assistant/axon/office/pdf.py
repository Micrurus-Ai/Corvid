"""Convert Office documents to PDF."""
import os
import json

from axon.util import _result
from axon.office._base import _run_ps, _resolve_path, _cfg_file, _move_doc_to_dot, _downloads

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
