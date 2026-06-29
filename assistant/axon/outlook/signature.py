"""Set the Outlook email signature."""
from axon.util import _result
from axon.outlook._base import _run_outlook_ps


_OUTLOOK_SIGNATURE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $name = $env:OL_SIG_NAME; if (-not $name) { $name = "Maia" }
    $dir = Join-Path $env:APPDATA "Microsoft\Signatures"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $html = $env:OL_SIG_HTML
    $txt = (($html -replace '<br\s*/?>', "`r`n") -replace '<[^>]+>', '') -replace '[ \t]+', ' '
    Set-Content -Path (Join-Path $dir ($name + ".htm")) -Value $html -Encoding UTF8
    Set-Content -Path (Join-Path $dir ($name + ".txt")) -Value $txt -Encoding UTF8
    $msg = "SIGNATURE_OK: saved signature '$name'"
    if ($env:OL_SIG_DEFAULT -eq "1") {
        foreach ($b in @("HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings", "HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings")) {
            if (Test-Path $b) {
                Set-ItemProperty -Path $b -Name "NewSignature" -Value $name -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $b -Name "ReplySignature" -Value $name -ErrorAction SilentlyContinue
            }
        }
        $msg += " and set as default (restart Outlook; if it doesn't apply, pick it once in File > Options > Mail > Signatures)"
    }
    try { Start-Process explorer.exe $dir } catch {}   # show the saved signature file
    Write-Output $msg
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _set_outlook_signature(args):
    if not (args.get("html") or "").strip():
        return _result("Need signature content (html).", True)
    out = _run_outlook_ps(_OUTLOOK_SIGNATURE_PS, {
        "OL_SIG_NAME": args.get("name") or "Maia",
        "OL_SIG_HTML": args["html"],
        "OL_SIG_DEFAULT": "1" if args.get("set_default", True) else "",
    })
    return _result(out, out.startswith("OL_ERROR"))
