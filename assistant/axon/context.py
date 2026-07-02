"""Detect what the user currently has open (active Office document) so phrases like
"this file / this sheet / the document" resolve to a real path without asking."""
import subprocess

from axon.config import IS_WINDOWS
from axon.util import NO_WINDOW

# Probe the running Office apps via COM for their active document's full path. COM's ActiveWorkbook/
# ActiveDocument/ActivePresentation don't depend on window focus, so this still works once the
# composer has taken focus.
_ACTIVE_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
$out = @()
try {
  $x = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  if ($x -and $x.ActiveWorkbook) { $out += "Excel: " + $x.ActiveWorkbook.FullName + " (active sheet: " + $x.ActiveSheet.Name + ")" }
} catch {}
try {
  $w = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
  if ($w -and $w.ActiveDocument) { $out += "Word: " + $w.ActiveDocument.FullName }
} catch {}
try {
  $p = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
  if ($p -and $p.ActivePresentation) { $out += "PowerPoint: " + $p.ActivePresentation.FullName }
} catch {}
$out -join "`n"
'''

# Only spend the COM probe when the request actually refers to "what's open".
_TRIGGERS = ("this file", "this sheet", "this document", "this workbook", "this spreadsheet",
             "this deck", "this presentation", "the file", "the document", "the sheet",
             "the workbook", "the spreadsheet", "open file", "current file", "the open",
             "what i have open", "on my screen", "active file", "this csv", "this excel",
             "analyse this", "analyze this", "summarize this", "summarise this",
             "the data", "this data", "these numbers", "the numbers", "this table",
             "the table", "this report", "these figures")


def refers_to_open_doc(text):
    t = (text or "").lower()
    return any(k in t for k in _TRIGGERS)


def active_documents():
    """Return a description of the Office documents currently open, or '' if none/unsupported."""
    if not IS_WINDOWS:
        return ""
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", _ACTIVE_PS],
            capture_output=True, text=True, timeout=15, creationflags=NO_WINDOW)
        return (p.stdout or "").strip()
    except Exception:
        return ""


def active_context_note(question):
    """A context note to inject when the user refers to 'this/the file' etc., naming the open
    document(s) so the agent uses the real path instead of asking. '' when not relevant."""
    if not refers_to_open_doc(question):
        return ""
    docs = active_documents()
    if not docs:
        return ""
    return ("\n\n[CONTEXT — files the user currently has open right now. When they say "
            "\"this file/sheet/document\" or \"the file\", they mean one of these — use its FULL "
            "path directly instead of asking. If more than one is listed, pick the one that best "
            "fits the request (e.g. a spreadsheet/CSV for data analysis, a document for writing) "
            "and only ask if it is genuinely ambiguous:\n" + docs + "]")
