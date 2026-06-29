"""Excel automation: create/write workbooks (optional chart) and read sheets."""
import os
import json

from axon.util import _result
from axon.office._base import _run_ps, _resolve_path, _cfg_file, _move_doc_to_dot, _downloads

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
