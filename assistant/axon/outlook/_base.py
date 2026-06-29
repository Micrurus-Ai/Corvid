"""Shared Outlook plumbing: the PowerShell runner that drives Outlook via COM, the
"make Outlook visible" snippet, and a window-locator helper class used by templates."""
import os
import subprocess

from axon.config import IS_WINDOWS
from axon.screen import _dot_monitor_env


_OL_WIN_CLASS = r'''
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class OcuWinF {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder t, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, IntPtr e);
  public static string Needle = ""; public static IntPtr Found = IntPtr.Zero;
  public static bool Cb(IntPtr h, IntPtr p){
    if(!IsWindowVisible(h)) return true;
    var sb=new StringBuilder(512); GetWindowText(h,sb,512); var t=sb.ToString();
    if(t.Length>0 && Needle.Length>0 && t.Contains(Needle)){ Found=h; return false; } return true;
  }
  public static void Front(string needle){
    Needle=needle; Found=IntPtr.Zero; EnumWindows(Cb,IntPtr.Zero);
    if(Found!=IntPtr.Zero){ keybd_event(0x12,0,0,IntPtr.Zero);keybd_event(0x12,0,2,IntPtr.Zero); ShowWindow(Found,9); SetForegroundWindow(Found); }
  }
}
"@
'''


_SHOW_OUTLOOK_SNIPPET = '''
try {
    Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class OcuOL {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder t, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, IntPtr e);
  public static IntPtr Found = IntPtr.Zero;
  public static bool Cb(IntPtr h, IntPtr p){
    if(!IsWindowVisible(h)) return true;
    var sb = new StringBuilder(512); GetWindowText(h, sb, 512); var t = sb.ToString();
    if(t.EndsWith("Outlook")){ Found = h; return false; }
    return true;
  }
  public static void Front(){
    Found = IntPtr.Zero; EnumWindows(Cb, IntPtr.Zero);
    if(Found != IntPtr.Zero){
      keybd_event(0x12,0,0,IntPtr.Zero); keybd_event(0x12,0,2,IntPtr.Zero);  // tap ALT to unlock focus
      ShowWindow(Found, 9); SetForegroundWindow(Found);
    }
  }
}
"@
    $__o = New-Object -ComObject Outlook.Application
    if (-not $__o.ActiveExplorer()) { $__o.GetNamespace("MAPI").GetDefaultFolder(6).Display() }
    Start-Sleep -Milliseconds 400
    [OcuOL]::Front()
} catch {}
'''


def _run_outlook_ps(ps, extra_env, show=True):
    if not IS_WINDOWS:
        return "OL_ERROR: Outlook automation is supported on the Windows app only."
    env = dict(os.environ)
    _m = _dot_monitor_env()
    if _m:
        for _k in ("L", "T", "R", "B"):
            env["DOT_MON_" + _k] = str(_m[_k])
    for k, v in extra_env.items():
        env[k] = "" if v is None else str(v)
    # If this script drives Outlook, first make Outlook visible so nothing happens behind the scenes.
    if show and "Outlook.Application" in ps:
        ps = _SHOW_OUTLOOK_SNIPPET + "\n" + ps
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True, text=True, env=env, timeout=90,
        )
    except subprocess.TimeoutExpired:
        return "OL_ERROR: Outlook operation timed out."
    return ((proc.stdout or "") + (proc.stderr or "")).strip()
