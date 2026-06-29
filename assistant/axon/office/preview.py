"""Document lifecycle for the structured (python-pptx/docx/openpyxl) editing engine.

Office holds an exclusive lock on a document while it's open, so before we edit a file on disk we
must close it in the app; after editing we (re)open it on the dot's monitor so the user sees the
result update — the "preview" half of the read -> edit -> see -> refine loop.
"""
import os
import time

from axon.config import IS_WINDOWS
from axon.office._base import _run_ps, _move_doc_to_dot

# Close the document in whichever Office app has it open so the file unlocks. We mark it Saved (to
# suppress the save prompt without writing — the on-disk file is the source of truth our editor
# reads), and also close any Protected View window (files opened from Downloads land there, NOT in
# the normal Documents/Workbooks collection — which was why the file stayed locked).
_CLOSE_DOC_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
$target = $env:DOC_PATH
$leaf = [System.IO.Path]::GetFileName($target)
foreach ($prog in @("Word.Application","Excel.Application","PowerPoint.Application")) {
    $app = $null
    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject($prog) } catch {}
    if ($app -eq $null) { continue }
    try {
        if ($prog -like "Word*")        { $docs = $app.Documents }
        elseif ($prog -like "Excel*")   { $docs = $app.Workbooks }
        else                            { $docs = $app.Presentations }
        for ($i = $docs.Count; $i -ge 1; $i--) {
            $d = $docs.Item($i)
            $full = ""; try { $full = [string]$d.FullName } catch {}
            if ($full -and ($full -ieq $target -or [System.IO.Path]::GetFileName($full) -ieq $leaf)) {
                try { $d.Saved = $true } catch {}
                try { $d.Close() } catch {}
                Write-Output "CLOSED"
            }
        }
    } catch {}
    try {
        $pvws = $app.ProtectedViewWindows
        if ($pvws -ne $null) {
            for ($i = $pvws.Count; $i -ge 1; $i--) {
                $pv = $pvws.Item($i)
                $sn = ""; try { $sn = [string]$pv.SourceName } catch {}
                $wf = ""; try { $wf = [string]$pv.Workbook.FullName } catch {}
                if (($sn -and $sn -ieq $leaf) -or ($wf -and $wf -ieq $target)) {
                    try { $pv.Close() } catch {}
                    Write-Output "CLOSED_PV"
                }
            }
        }
    } catch {}
}
'''


def ensure_closed(path):
    """If the document is open in an Office app, save + close it so the file can be edited on disk."""
    if not IS_WINDOWS:
        return
    try:
        _run_ps(_CLOSE_DOC_PS, {"DOC_PATH": os.path.abspath(path)})
    except Exception:
        pass


def save_with_retry(path, save_to, attempts=5):
    """Save the document robustly. `save_to(target)` writes it to a path. We try the original path,
    closing the doc in Office between attempts if it's locked; if it stays locked (e.g. several
    stray Office instances hold it), we fall back to an '<name> (edited).<ext>' copy so the user
    always gets the result. Returns (final_path, error)."""
    for i in range(attempts):
        try:
            save_to(path)
            return path, None
        except PermissionError:
            ensure_closed(path)        # Office grabbed the file again — release it
            time.sleep(0.5 * (i + 1))  # give the handle time to drop
        except Exception as e:
            return path, e
    base, ext = os.path.splitext(path)  # original is still locked — write a copy instead
    alt = base + " (edited)" + ext
    try:
        save_to(alt)
        return alt, None
    except Exception as e:
        return path, e


def open_doc(path):
    """Open the document in its Office app and move it onto the dot's monitor (the live preview).
    Set AXON_NO_PREVIEW=1 to skip (used by automated tests so Office windows don't pop up)."""
    if not IS_WINDOWS or os.getenv("AXON_NO_PREVIEW"):
        return
    try:
        os.startfile(os.path.abspath(path))  # open in the default app for the extension
        time.sleep(1.2)
        _move_doc_to_dot(path)
    except Exception:
        pass
