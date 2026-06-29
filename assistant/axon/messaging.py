"""Microsoft Teams messaging (resolves a recipient via Outlook, then sends a chat)."""
import os
import subprocess

from axon.config import IS_WINDOWS
from axon.util import _result
from axon.outlook import _run_outlook_ps, _self_email, _RESOLVE_EMAIL_PS

_TEAMS_SEND_PS = r'''
$ErrorActionPreference = "Stop"
try {
    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class TeamsWin {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder t, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, IntPtr e);
  public static string Needle = "";
  public static IntPtr Found = IntPtr.Zero;
  public static bool Cb(IntPtr h, IntPtr p){
    if(!IsWindowVisible(h)) return true;
    var sb=new StringBuilder(1024); GetWindowText(h,sb,1024);
    var t=sb.ToString();
    if(t.Length>0 && t.Contains(Needle)){ Found=h; return false; }
    return true;
  }
  public static IntPtr Find(string needle){ Needle=needle; Found=IntPtr.Zero; EnumWindows(Cb,IntPtr.Zero); return Found; }
  public static void Front(IntPtr h){
    keybd_event(0x12,0,0,IntPtr.Zero); keybd_event(0x12,0,2,IntPtr.Zero);  // tap ALT to unlock focus
    ShowWindow(h,9); SetForegroundWindow(h);
  }
  public static void Enter(){ keybd_event(0x0D,0,0,IntPtr.Zero); keybd_event(0x0D,0,2,IntPtr.Zero); }
}
"@
    Start-Process $env:TEAMS_LINK
    Start-Sleep -Seconds 6
    $h = [IntPtr]::Zero
    for ($i=0; ($i -lt 16) -and ($h -eq [IntPtr]::Zero); $i++) { Start-Sleep -Milliseconds 400; $h = [TeamsWin]::Find("Microsoft Teams") }
    if ($h -eq [IntPtr]::Zero) { Write-Output "TEAMS_ERROR: Teams window not found after opening the chat link"; exit }
    [TeamsWin]::Front($h)
    Start-Sleep -Milliseconds 1500
    if ($env:TEAMS_SEND -eq "1") {
        [TeamsWin]::Enter()
        Start-Sleep -Milliseconds 500
        Write-Output "TEAMS_SENT"
    } else {
        Write-Output "TEAMS_DRAFTED"
    }
} catch { Write-Output ("TEAMS_ERROR: " + $_.Exception.Message) }
'''


def _send_teams_message(args):
    import urllib.parse as _up
    if not IS_WINDOWS:
        return _result("Teams messaging is supported on Windows only.", True)
    msg = (args.get("message") or "").strip()
    if not msg:
        return _result("Need a message to send.", True)
    recips = args.get("to") or []
    if isinstance(recips, str):
        recips = [recips]
    emails, problems = [], []
    for r in recips:
        r = (r or "").strip()
        if not r:
            continue
        if r.lower() in ("me", "myself", "self"):
            e = _self_email()
            (emails.append(e) if e else problems.append("could not determine your own email"))
            continue
        if "@" in r:
            emails.append(r)
            continue
        out = _run_outlook_ps(_RESOLVE_EMAIL_PS, {"OL_NAME": r}).strip()
        if out.startswith("SMTP:") and "@" in out:
            emails.append(out[5:].strip())
        else:
            problems.append(f"could not resolve '{r}' to an email")
    emails = [e for e in emails if e]
    if not emails:
        return _result("No valid Teams recipients (" + "; ".join(problems) + "). Provide an email address.", True)
    users = ",".join(emails)
    link = "msteams:/l/chat/0/0?users=" + _up.quote(users, safe="@,") + "&message=" + _up.quote(msg, safe="")
    send = args.get("send", True)
    out = _run_outlook_ps(_TEAMS_SEND_PS, {"TEAMS_LINK": link, "TEAMS_SEND": "1" if send else ""})
    note = (" (note: " + "; ".join(problems) + ")") if problems else ""
    if out.startswith("TEAMS_SENT"):
        return _result(f'Sent Teams message to {users}: "{msg}"' + note)
    if out.startswith("TEAMS_DRAFTED"):
        return _result(f"Opened the Teams chat to {users} with the message drafted (not sent)." + note)
    return _result("Teams send failed: " + out + note, True)
