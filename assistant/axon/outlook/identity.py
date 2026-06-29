"""Resolve the user's own email address (and look up others) via Outlook."""
from axon.outlook._base import _run_outlook_ps


_SELF_EMAIL_CACHE = {}


_SELF_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    Write-Output $ns.Accounts.Item(1).SmtpAddress
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


_RESOLVE_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $r = $ns.CreateRecipient($env:OL_NAME)
    $r.Resolve() | Out-Null
    if ($r.Resolved) {
        $ae = $r.AddressEntry; $smtp = $null
        try { $eu = $ae.GetExchangeUser(); if ($eu) { $smtp = $eu.PrimarySmtpAddress } } catch {}
        if (-not $smtp) { $smtp = $ae.Address }
        Write-Output ("SMTP:" + $smtp)
    } else { Write-Output "UNRESOLVED" }
} catch { Write-Output ("ERR: " + $_.Exception.Message) }
'''


def _self_email():
    if not _SELF_EMAIL_CACHE.get("v"):
        out = _run_outlook_ps(_SELF_EMAIL_PS, {}).strip()
        _SELF_EMAIL_CACHE["v"] = "" if (out.startswith("ERR") or "@" not in out) else out
    return _SELF_EMAIL_CACHE["v"]
