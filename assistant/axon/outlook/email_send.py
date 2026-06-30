"""Sending mail: compose/send, drafts (draft -> approve -> send), reply, and forward."""
import os
import subprocess

from axon.config import IS_WINDOWS
from axon.util import _result, NO_WINDOW
from axon.approval import _ask_approval
from axon.screen import _move_window_to_dot
from axon.outlook._base import _run_outlook_ps, _OL_WIN_CLASS


_SEND_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class OcuWin {
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
    if(t.Length>0 && t.Contains(Needle) && t.Contains("Message")){ Found=h; return false; }
    return true;
  }
  public static IntPtr Find(string needle){ Needle=needle; Found=IntPtr.Zero; EnumWindows(Cb,IntPtr.Zero); return Found; }
  public static void Front(IntPtr h){
    keybd_event(0x12,0,0,IntPtr.Zero); keybd_event(0x12,0,2,IntPtr.Zero); // tap ALT to unlock focus
    ShowWindow(h,9); SetForegroundWindow(h);                              // 9 = SW_RESTORE
  }
}
"@
    $ol = New-Object -ComObject Outlook.Application
    $mail = $ol.CreateItem(0)
    $mail.To = $env:OCU_MAIL_TO
    if ($env:OCU_MAIL_CC) { $mail.CC = $env:OCU_MAIL_CC }
    $mail.Subject = $env:OCU_MAIL_SUBJECT
    # An HTML report is set directly; a PLAIN body is typed LIVE into the window below (visibly),
    # in word-count-aware chunks so even a long report types in a few seconds.
    if ($env:OCU_MAIL_HTML) { $mail.HTMLBody = $env:OCU_MAIL_HTML }
    if ($env:OCU_MAIL_ATTACH) {
        foreach ($p in ($env:OCU_MAIL_ATTACH -split [regex]::Escape("|"))) {
            if ($p) { $mail.Attachments.Add($p) | Out-Null }
        }
    }
    [void]$mail.Recipients.ResolveAll()
    $insp = $mail.GetInspector
    $insp.Display($false)
    # Bring the compose window to the foreground (Windows blocks plain .Activate from a bg process).
    $needle = $env:OCU_MAIL_SUBJECT; if (-not $needle) { $needle = "Message" }
    $h = [IntPtr]::Zero
    for ($i=0; ($i -lt 12) -and ($h -eq [IntPtr]::Zero); $i++) { Start-Sleep -Milliseconds 350; $h = [OcuWin]::Find($needle) }
    if ($h -ne [IntPtr]::Zero) { [OcuWin]::Front($h) }
    Start-Sleep -Milliseconds 600
    if ($env:OCU_MAIL_HTML) {
        # Rich HTML report already set as HTMLBody; show the rendered window briefly.
        $secs = 4
        if ($env:OCU_MAIL_SHOW_SECONDS) { [int]::TryParse($env:OCU_MAIL_SHOW_SECONDS, [ref]$secs) | Out-Null }
        Start-Sleep -Seconds $secs
    } elseif ($env:OCU_MAIL_BODY) {
        # Type the plain body LIVE and visibly, but in chunks sized by the body length so the total
        # typing time stays a few seconds no matter how long the report is. Short emails -> 1 char at
        # a time (classic typing feel); long emails -> bigger bursts (faster), so it never times out.
        try {
            $b = $env:OCU_MAIL_BODY
            $sel = $insp.WordEditor.Application.Selection
            $strokes = 50
            if ($env:OCU_MAIL_TYPE_STROKES) { [int]::TryParse($env:OCU_MAIL_TYPE_STROKES, [ref]$strokes) | Out-Null }
            $delay = 28
            if ($env:OCU_MAIL_TYPE_DELAY_MS) { [int]::TryParse($env:OCU_MAIL_TYPE_DELAY_MS, [ref]$delay) | Out-Null }
            $chunk = [Math]::Max(1, [Math]::Ceiling($b.Length / [double]$strokes))
            $lines = $b -split "`r`n|`n|`r"
            for ($li = 0; $li -lt $lines.Count; $li++) {
                $line = $lines[$li]; $j = 0
                while ($j -lt $line.Length) {
                    $n = [Math]::Min($chunk, $line.Length - $j)
                    $sel.TypeText($line.Substring($j, $n)); $j += $n
                    Start-Sleep -Milliseconds $delay
                }
                if ($li -lt ($lines.Count - 1)) { $sel.TypeParagraph() }
            }
        } catch {
            $mail.Body = $env:OCU_MAIL_BODY   # fallback: set directly if live typing fails
        }
        Start-Sleep -Milliseconds 600
    }
    $mail.Save()
    Write-Output ("DRAFT_OK|" + $mail.EntryID + "|" + $mail.Subject)
} catch {
    Write-Output ("SEND_ERROR: " + $_.Exception.Message)
}
'''


_SEND_DRAFT_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    $m.Send()
    Write-Output "SENT_OK"
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _send_draft(eid):
    return _run_outlook_ps(_SEND_DRAFT_PS, {"OL_ID": eid}, show=False)


def _draft_then_send(out_text, approval_desc, sent_msg, draft_msg):
    """Shared flow: a draft PS just ran (output_text = DRAFT_OK|<id>|<subject>). Move the draft
    window onto the dot's monitor, ask the user, then send or leave as a draft."""
    if "DRAFT_OK" not in out_text:
        return _result(out_text or "No output from Outlook.", True)
    payload = out_text.split("DRAFT_OK|", 1)[1].strip().splitlines()[0]
    parts = payload.split("|", 1)
    eid = parts[0].strip()
    subject = parts[1].strip() if len(parts) > 1 else ""
    _move_window_to_dot(subject)  # bring the open draft window onto the dot's monitor
    if not _ask_approval(approval_desc):
        return _result(draft_msg, False)
    s = _send_draft(eid)
    if "SENT_OK" in s:
        return _result(sent_msg, False)
    return _result("Tried to send but failed: " + s, True)


def _send_email(args):
    to = (args.get("to") or "").strip()
    if not to:
        return _result("Missing 'to' address.", True)
    if not IS_WINDOWS:
        return _result("send_email currently supports the Windows Outlook app only.", True)
    attachments = args.get("attachments") or []
    if isinstance(attachments, str):
        attachments = [attachments]
    missing = [p for p in attachments if not os.path.isfile(p)]
    if missing:
        return _result("Attachment file(s) not found: " + "; ".join(missing), True)
    env = dict(os.environ)
    env["OCU_MAIL_TO"] = to
    env["OCU_MAIL_CC"] = args.get("cc") or ""
    env["OCU_MAIL_SUBJECT"] = args.get("subject") or ""
    env["OCU_MAIL_BODY"] = args.get("body") or ""
    env["OCU_MAIL_HTML"] = args.get("html") or ""
    env["OCU_MAIL_ATTACH"] = "|".join(attachments)
    # 1) DRAFT it visibly (open compose window, fill recipient/subject, live-type the body) — no send.
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", _SEND_EMAIL_PS],
            capture_output=True, text=True, env=env, timeout=150, creationflags=NO_WINDOW,
        )
    except subprocess.TimeoutExpired:
        return _result("Drafting timed out (Outlook may have shown a security prompt).", True)
    out = ((proc.stdout or "") + (proc.stderr or "")).strip()
    # 2) Ask to approve the SEND (after the user can see the drafted email), then send or keep draft.
    subj = args.get("subject") or ""
    return _draft_then_send(
        out,
        f"Send this email to {to}" + (f" — subject: {subj}" if subj else "") + "?",
        f"Sent email to {to}.",
        f"Drafted the email to {to} and left it open in Outlook for you to review/send (NOT sent).")


_OUTLOOK_FORWARD_PS = r'''
$ErrorActionPreference = "Stop"
try {
''' + _OL_WIN_CLASS + r'''
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    $fwd = $m.Forward()
    $fwd.To = $env:OL_TO
    if ($env:OL_NOTE) { $fwd.Body = $env:OL_NOTE + "`r`n`r`n" + $fwd.Body }
    [void]$fwd.Recipients.ResolveAll()
    $insp = $fwd.GetInspector; $insp.Display($false)   # show the forward window
    Start-Sleep -Milliseconds 500
    [OcuWinF]::Front([string]$fwd.Subject)
    Start-Sleep -Milliseconds 800
    $fwd.Save()
    Write-Output ("DRAFT_OK|" + $fwd.EntryID + "|" + $fwd.Subject)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _outlook_forward_email(args):
    if not (args.get("id") or ""):
        return _result("No email id provided.", True)
    if not (args.get("to") or ""):
        return _result("No recipient provided.", True)
    out = _run_outlook_ps(_OUTLOOK_FORWARD_PS, {"OL_ID": args["id"], "OL_TO": args["to"], "OL_NOTE": args.get("note") or ""}, show=False)
    return _draft_then_send(
        out, f"Forward this email to {args['to']}?",
        f"Forwarded the email to {args['to']}.",
        f"Drafted the forward to {args['to']} and left it open in Outlook for you to review/send (NOT sent).")


_OUTLOOK_REPLY_PS = r'''
$ErrorActionPreference = "Stop"
try {
''' + _OL_WIN_CLASS + r'''
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    if ($env:OL_REPLYALL -eq "1") { $r = $m.ReplyAll() } else { $r = $m.Reply() }
    if ($env:OL_BODY) { $r.Body = $env:OL_BODY + "`r`n`r`n" + $r.Body }
    [void]$r.Recipients.ResolveAll()
    $insp = $r.GetInspector; $insp.Display($false)   # show the reply window
    Start-Sleep -Milliseconds 500
    [OcuWinF]::Front([string]$r.Subject)
    Start-Sleep -Milliseconds 800
    $r.Save()
    Write-Output ("DRAFT_OK|" + $r.EntryID + "|" + $r.Subject)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _outlook_reply_email(args):
    if not (args.get("id") or ""):
        return _result("No email id provided.", True)
    out = _run_outlook_ps(_OUTLOOK_REPLY_PS, {
        "OL_ID": args["id"], "OL_BODY": args.get("body") or "",
        "OL_REPLYALL": "1" if args.get("reply_all") else "",
    }, show=False)
    return _draft_then_send(
        out, f"Send this reply?",
        "Reply sent.",
        f"Drafted the reply and left it open in Outlook for you to review/send (NOT sent).")
