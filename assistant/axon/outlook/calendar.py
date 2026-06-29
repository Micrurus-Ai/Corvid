"""Create Outlook calendar events / meetings."""
from axon.util import _result
from axon.outlook._base import _run_outlook_ps


_OUTLOOK_CALENDAR_PS = r'''
$ErrorActionPreference = "Stop"
try {
    Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class OcuAppt {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder t, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, IntPtr e);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out MRECT r);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool rp);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
  [StructLayout(LayoutKind.Sequential)] public struct MRECT { public int L,T,R,B; }
  public static string Needle = ""; public static IntPtr Found = IntPtr.Zero;
  static int EI(string n){ int v; return int.TryParse(Environment.GetEnvironmentVariable(n), out v) ? v : 0; }
  public static void MoveToDot(IntPtr h){
    if(string.IsNullOrEmpty(Environment.GetEnvironmentVariable("DOT_MON_R"))) return;
    try { SetThreadDpiAwarenessContext((IntPtr)(-4)); } catch {}
    int ml=EI("DOT_MON_L"), mt=EI("DOT_MON_T"), mr=EI("DOT_MON_R"), mb=EI("DOT_MON_B");
    MRECT wr; if(!GetWindowRect(h, out wr)) return; int w=wr.R-wr.L, ht=wr.B-wr.T; if(w<=0||ht<=0) return;
    int mw=mr-ml, mh=mb-mt; if(w>mw){w=mw;} if(ht>mh){ht=mh;}
    MoveWindow(h, ml+((mw-w)/2), mt+((mh-ht)/2), w, ht, true);
  }
  public static bool Cb(IntPtr h, IntPtr p){
    if(!IsWindowVisible(h)) return true;
    var sb=new StringBuilder(512); GetWindowText(h,sb,512); var t=sb.ToString();
    if(t.Length>0 && t.Contains(Needle)){ Found=h; return false; } return true;
  }
  public static void Front(string needle){
    Needle=needle; Found=IntPtr.Zero; EnumWindows(Cb,IntPtr.Zero);
    if(Found!=IntPtr.Zero){ keybd_event(0x12,0,0,IntPtr.Zero);keybd_event(0x12,0,2,IntPtr.Zero); ShowWindow(Found,9); SetForegroundWindow(Found); MoveToDot(Found); }
  }
}
"@
    $ol = New-Object -ComObject Outlook.Application
    $appt = $ol.CreateItem(1)   # olAppointmentItem
    $appt.Subject = $env:OL_SUBJECT
    $appt.Start = [datetime]::Parse($env:OL_START)
    if ($env:OL_END) { $appt.End = [datetime]::Parse($env:OL_END) } else { $appt.Duration = 60 }
    if ($env:OL_LOCATION) { $appt.Location = $env:OL_LOCATION }
    if ($env:OL_BODY) { $appt.Body = $env:OL_BODY }
    if ($env:OL_RECUR) {
        $rp = $appt.GetRecurrencePattern()
        switch ($env:OL_RECUR) {
            "daily"   { $rp.RecurrenceType = 0 }
            "weekly"  { $rp.RecurrenceType = 1 }
            "monthly" { $rp.RecurrenceType = 2 }
            "yearly"  { $rp.RecurrenceType = 5 }
        }
        if ($env:OL_RECUR_INTERVAL) { $rp.Interval = [int]$env:OL_RECUR_INTERVAL }
        if ($env:OL_RECUR_COUNT) { $rp.Occurrences = [int]$env:OL_RECUR_COUNT }
        elseif ($env:OL_RECUR_UNTIL) { $rp.PatternEndDate = [datetime]::Parse($env:OL_RECUR_UNTIL) }
    }
    $isMeeting = $false
    if ($env:OL_ATTENDEES) {
        foreach ($a in ($env:OL_ATTENDEES -split [regex]::Escape("|"))) { if ($a) { [void]$appt.Recipients.Add($a) } }
        $appt.MeetingStatus = 1   # olMeeting
        [void]$appt.Recipients.ResolveAll()
        $isMeeting = $true
    }
    # Show the appointment/meeting window so the user SEES it (transparency, like email compose).
    $insp = $appt.GetInspector
    $insp.Display($false)
    $needle = $env:OL_SUBJECT; if (-not $needle) { $needle = "Appointment" }
    Start-Sleep -Milliseconds 500
    [OcuAppt]::Front($needle)
    Start-Sleep -Milliseconds 2200
    if ($isMeeting) {
        $appt.Send()
        Write-Output ("EVENT_OK: meeting '" + $appt.Subject + "' shown and invite sent")
    } else {
        $appt.Save()
        Write-Output ("EVENT_OK: '" + $appt.Subject + "' shown and added to your calendar")
    }
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _create_calendar_event(args):
    if not (args.get("subject") or "").strip() or not (args.get("start") or "").strip():
        return _result("Need at least a subject and a start time.", True)
    attendees = args.get("attendees") or []
    if isinstance(attendees, str):
        attendees = [attendees]
    out = _run_outlook_ps(_OUTLOOK_CALENDAR_PS, {
        "OL_SUBJECT": args["subject"], "OL_START": args["start"], "OL_END": args.get("end") or "",
        "OL_LOCATION": args.get("location") or "", "OL_BODY": args.get("body") or "",
        "OL_ATTENDEES": "|".join(attendees),
        "OL_RECUR": (args.get("recurrence") or "").lower(),
        "OL_RECUR_INTERVAL": args.get("recur_interval") or "",
        "OL_RECUR_COUNT": args.get("recur_count") or "",
        "OL_RECUR_UNTIL": args.get("recur_until") or "",
    }, show=False)  # the script shows its own meeting window; don't also flash the Inbox
    return _result(out, out.startswith("OL_ERROR"))
