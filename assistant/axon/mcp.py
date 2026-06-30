"""The 'hands': a persistent open-computer-use MCP client, plus desktop control tools
(open / close / list apps). Apps Axon opens are placed on the dot's monitor."""
import os
import json
import time
import re
import subprocess

from axon import config
from axon.config import IS_WINDOWS, _OCU, TOOL_TEXT_LIMIT
from axon.util import _result, NO_WINDOW
from axon.screen import _move_window_to_dot, _dot_monitor_env
from axon.outlook import _run_outlook_ps

def _extract(result):
    """Split an MCP-style result into (text, base64_png_or_None)."""
    text_parts, image = [], None
    for block in result.get("content", []) or []:
        if block.get("type") == "text":
            text_parts.append(block.get("text", ""))
        elif block.get("type") == "image" and block.get("data"):
            image = block["data"]
    text = "\n".join(p for p in text_parts if p) or "(no text)"
    if result.get("isError"):
        text = "ERROR: " + text
    return text, image


class MCPClient:
    """A persistent open-computer-use MCP server process (keeps per-session state)."""

    def __init__(self):
        inner = [_OCU, "mcp"]
        cmd = (["cmd", "/c"] + inner) if IS_WINDOWS else inner
        self.proc = subprocess.Popen(
            cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, text=True, bufsize=1, creationflags=NO_WINDOW,
        )
        self._id = 0
        self._handshake()

    def _send(self, obj):
        self.proc.stdin.write(json.dumps(obj) + "\n")
        self.proc.stdin.flush()

    def _read(self, want_id):
        for line in self.proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                m = json.loads(line)
            except Exception:
                continue
            if m.get("id") == want_id:
                if "error" in m:
                    return _result(str(m["error"]), True)
                return m.get("result", {})
        return _result("MCP server closed unexpectedly.", True)

    def _handshake(self):
        self._id += 1
        i = self._id
        self._send({
            "jsonrpc": "2.0", "id": i, "method": "initialize",
            "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                       "clientInfo": {"name": "desktop-assistant", "version": "1"}},
        })
        self._read(i)
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    def call(self, tool, args):
        self._id += 1
        i = self._id
        self._send({"jsonrpc": "2.0", "id": i, "method": "tools/call",
                    "params": {"name": tool, "arguments": args or {}}})
        return self._read(i)

    def close(self):
        for fn in (lambda: self.proc.stdin.close(), self.proc.terminate):
            try:
                fn()
            except Exception:
                pass


_LIST_INSTALLED_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $apps = Get-StartApps | Select-Object -ExpandProperty Name -Unique | Sort-Object
    $apps -join "`n"
} catch {
    try {
        $paths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        $names = foreach ($p in $paths) {
            Get-ItemProperty $p -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | Select-Object -ExpandProperty DisplayName
        }
        ($names | Sort-Object -Unique) -join "`n"
    } catch { "INSTALLED_ERROR: " + $_.Exception.Message }
}
'''


def _list_installed_apps(args):
    if not IS_WINDOWS:
        return _result("Listing installed apps is supported on Windows only.", True)
    out = _run_outlook_ps(_LIST_INSTALLED_PS, {})
    err = out.startswith("OL_ERROR") or out.startswith("INSTALLED_ERROR")
    return _result(out[:TOOL_TEXT_LIMIT], err)


_OPEN_APP_PS = r'''
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class WinMove {
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr h);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool repaint);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr h, int flags);
  [DllImport("user32.dll")] public static extern bool GetMonitorInfo(IntPtr hMon, ref MONITORINFO mi);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
  [StructLayout(LayoutKind.Sequential)] public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public int dwFlags; }
  public static void Dpi(){ SetThreadDpiAwarenessContext((IntPtr)(-4)); }
  public static RECT DotMon(IntPtr dot){
    IntPtr m = MonitorFromWindow(dot, 2);
    MONITORINFO mi = new MONITORINFO(); mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
    GetMonitorInfo(m, ref mi); return mi.rcWork;
  }
  public static IntPtr[] TopWindows(){
    List<IntPtr> wins = new List<IntPtr>();
    EnumWindows(delegate(IntPtr h, IntPtr l){ wins.Add(h); return true; }, IntPtr.Zero);
    return wins.ToArray();
  }
  public static string Title(IntPtr h){
    int n = GetWindowTextLength(h);
    if (n <= 0) return "";
    StringBuilder s = new StringBuilder(n + 1);
    GetWindowText(h, s, s.Capacity);
    return s.ToString();
  }
}
"@
[WinMove]::Dpi() | Out-Null
$name = $env:OCU_OPEN_NAME
$disp = $name
$before = @{}
foreach ($h in [WinMove]::TopWindows()) { $before[$h.ToInt64()] = $true }
try {
    $app = Get-StartApps | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
    if (-not $app) { $app = Get-StartApps | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1 }
    if ($app) { Start-Process ("shell:AppsFolder\" + $app.AppID); $disp = $app.Name }
    else { Start-Process $name }
} catch {
    try { Start-Process $name } catch { Write-Output ("OPEN_ERROR: could not launch '" + $name + "': " + $_.Exception.Message); exit }
}
# Place ONLY the brand-new window onto the dot's monitor — never touch the user's existing windows,
# and keep the window's own size (just reposition it) so nothing is disturbed on the other screen.
$dot = [IntPtr]::Zero
if ($env:DOT_HWND) { try { $dot = [IntPtr][int64]$env:DOT_HWND } catch {} }
$dotpid = 0; if ($env:DOT_PID) { [int]::TryParse($env:DOT_PID, [ref]$dotpid) | Out-Null }
if ($dot -ne [IntPtr]::Zero) {
    try {
        $r = [WinMove]::DotMon($dot)
        $mw = $r.R - $r.L; $mh = $r.B - $r.T
        for ($i = 0; $i -lt 30; $i++) {
            Start-Sleep -Milliseconds 250
            $newWin = [IntPtr]::Zero
            foreach ($h in [WinMove]::TopWindows()) {
                if ($before.ContainsKey($h.ToInt64())) { continue }       # existing window -> leave it
                if ($h -eq $dot) { continue }
                if (-not [WinMove]::IsWindowVisible($h)) { continue }
                if ([WinMove]::IsIconic($h)) { continue }
                if ([WinMove]::GetWindowTextLength($h) -eq 0) { continue }
                $wpid = 0; [WinMove]::GetWindowThreadProcessId($h, [ref]$wpid) | Out-Null
                if ($wpid -eq $dotpid) { continue }
                $newWin = $h; break
            }
            if ($newWin -ne [IntPtr]::Zero) {
                $wr = New-Object WinMove+RECT
                [void][WinMove]::GetWindowRect($newWin, [ref]$wr)
                $w = $wr.R - $wr.L; $wh = $wr.B - $wr.T
                if ($w -le 0 -or $w -gt $mw) { $w = [int]($mw * 0.8) }
                if ($wh -le 0 -or $wh -gt $mh) { $wh = [int]($mh * 0.85) }
                $wx = $r.L + [int](($mw - $w) / 2); $wy = $r.T + [int](($mh - $wh) / 2)
                [WinMove]::MoveWindow($newWin, $wx, $wy, $w, $wh, $true) | Out-Null
                Write-Output ("PLACED_ON_DOT_MONITOR: " + [WinMove]::Title($newWin))
                break
            }
        }
    } catch {}
}
Write-Output ("LAUNCHED: " + $disp)
'''


def _open_app(args):
    name = (args.get("name") or "").strip()
    if not name:
        return _result("Missing app name.", True)
    if not IS_WINDOWS:
        try:
            subprocess.Popen([name])
            return _result(f"Launched '{name}'.")
        except Exception as e:
            return _result(f"Failed to launch '{name}': {e}", True)
    # On Windows, resolve via the Start menu (handles packaged apps like Teams reliably) and move
    # the new window onto the dot's monitor.
    env = {"OCU_OPEN_NAME": name}
    if config.DOT_HWND:
        env["DOT_HWND"] = str(config.DOT_HWND)
        env["DOT_PID"] = str(config.DOT_PID or 0)
    out = _run_outlook_ps(_OPEN_APP_PS, env)
    if out.startswith("OPEN_ERROR"):
        return _result(out, True)
    return _result(out + ". Wait ~2-4s (some apps are slow), then call get_app_state to confirm it opened.")


def _close_app(client, args):
    app = (args.get("app") or "").strip()
    if not app:
        return _result("Missing app name.", True)
    if bool(args.get("force")):
        if IS_WINDOWS:
            image = app if app.lower().endswith(".exe") else app + ".exe"
            proc = subprocess.run(["cmd", "/c", "taskkill", "/F", "/IM", image],
                                  capture_output=True, text=True, creationflags=NO_WINDOW)
        else:
            proc = subprocess.run(["pkill", "-9", "-f", app], capture_output=True, text=True)
        out = ((proc.stdout or "") + (proc.stderr or "")).strip()
        return _result(out or "Force-killed.", proc.returncode != 0)

    # Graceful: read the window, find its Close button, click it.
    state = client.call("get_app_state", {"app": app})
    text, _ = _extract(state)
    if state.get("isError"):
        return state
    idx = None
    for ln in text.splitlines():
        m = re.match(r"\s*(\d+)\s+[\w ]*button\s+Close\b", ln)
        if m:
            idx = m.group(1)
            break
    if idx is None:
        return _result(f"No Close button found for '{app}'. Try clicking it via get_app_state, or use force=true.", True)
    res = client.call("click", {"app": app, "element_index": idx})
    t, _ = _extract(res)
    # Clicking Close makes the window vanish, so the tool's post-action state capture
    # often reports "appNotFound" — that means success, not failure.
    if not res.get("isError") or "notfound" in t.lower().replace(" ", "") or "not found" in t.lower():
        return _result(f"Closed '{app}' (clicked Close button, element {idx}).")
    return _result(f"Tried to close '{app}' (element {idx}); it may still be open or showing a dialog: {t[:160]}", True)
