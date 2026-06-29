"""Screen, monitor, and window helpers: DPI-aware monitor geometry, moving newly-opened
windows onto the dot's monitor, screenshots, grids, and image encoding for vision/guide."""
import os
import io
import time
import base64

from axon import config
from axon.config import IS_WINDOWS, DOWNLOADS_DIR
from axon.util import _result

_VSCREEN_CACHE = {}  # cache of the virtual-screen geometry

def _dot_monitor_env():
    """The dot's monitor work-area (physical pixels) as a dict, or {} if unknown."""
    if not config.DOT_HWND:
        return {}
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        u.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))  # physical pixels

        class RECT(ctypes.Structure):
            _fields_ = [("L", ctypes.c_int), ("T", ctypes.c_int), ("R", ctypes.c_int), ("B", ctypes.c_int)]

        class MI(ctypes.Structure):
            _fields_ = [("cb", ctypes.c_int), ("m", RECT), ("w", RECT), ("f", ctypes.c_int)]

        mon = u.MonitorFromWindow(ctypes.c_void_p(int(config.DOT_HWND)), 2)  # NEAREST
        mi = MI()
        mi.cb = ctypes.sizeof(MI)
        if not u.GetMonitorInfoW(mon, ctypes.byref(mi)):
            return {}
        return {"L": mi.w.L, "T": mi.w.T, "R": mi.w.R, "B": mi.w.B}
    except Exception:
        return {}


def _move_window_to_dot(title_substr):
    """Find the visible window whose title CONTAINS title_substr and move it onto the dot's monitor
    (keeping its size, centred). Matching by title is reliable — the foreground window often is not
    the one we just opened."""
    if not IS_WINDOWS or not title_substr:
        return
    env = _dot_monitor_env()
    if not env:
        return
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        u.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))
        needle = title_substr.lower()[:40].strip()
        if not needle:
            return
        found = {"h": None}
        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def cb(h, _lparam):
            if not u.IsWindowVisible(h):
                return True
            n = u.GetWindowTextLengthW(h)
            if n == 0:
                return True
            buf = ctypes.create_unicode_buffer(n + 1)
            u.GetWindowTextW(h, buf, n + 1)
            if needle in buf.value.lower():
                found["h"] = h
                return False
            return True

        u.EnumWindows(WNDENUMPROC(cb), 0)
        h = found["h"]
        if not h:
            return
        rect = wintypes.RECT()
        u.GetWindowRect(h, ctypes.byref(rect))
        w, ht = rect.right - rect.left, rect.bottom - rect.top
        if w <= 0 or ht <= 0:
            return
        mw, mh = env["R"] - env["L"], env["B"] - env["T"]
        w, ht = min(w, mw), min(ht, mh)
        u.MoveWindow(h, env["L"] + (mw - w) // 2, env["T"] + (mh - ht) // 2, w, ht, True)
    except Exception:
        pass


def _grab_screen_b64():
    """Capture the whole (virtual) desktop and return a base64 PNG."""
    from PIL import ImageGrab
    import io as _io
    import base64 as _b64
    img = ImageGrab.grab(all_screens=True)
    w = img.size[0]
    max_w = 1700
    if w > max_w:
        r = max_w / w
        img = img.resize((int(w * r), int(img.size[1] * r)))
    buf = _io.BytesIO()
    img.save(buf, format="PNG")
    return _b64.b64encode(buf.getvalue()).decode()


def _dot_monitor_crop(full_size):
    """Pixel box (l,t,r,b) of the dot's monitor within a full-desktop screenshot, or None.
    Working on a single monitor keeps the 0-1000 grid a clean uniform rectangle (robust on any
    multi-monitor layout)."""
    if not config.DOT_HWND:
        return None
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        u.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        u.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        u.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))  # physical pixels
        vox = u.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN
        voy = u.GetSystemMetrics(77)  # SM_YVIRTUALSCREEN

        class RECT(ctypes.Structure):
            _fields_ = [("L", ctypes.c_int), ("T", ctypes.c_int), ("R", ctypes.c_int), ("B", ctypes.c_int)]

        class MI(ctypes.Structure):
            _fields_ = [("cb", ctypes.c_int), ("m", RECT), ("w", RECT), ("f", ctypes.c_int)]

        mon = u.MonitorFromWindow(ctypes.c_void_p(int(config.DOT_HWND)), 2)  # NEAREST
        mi = MI()
        mi.cb = ctypes.sizeof(MI)
        if not u.GetMonitorInfoW(mon, ctypes.byref(mi)):
            return None
        W, H = full_size
        l = max(0, min(mi.m.L - vox, W))
        t = max(0, min(mi.m.T - voy, H))
        r = max(0, min(mi.m.R - vox, W))
        b = max(0, min(mi.m.B - voy, H))
        if r - l < 10 or b - t < 10:
            return None
        return (l, t, r, b)
    except Exception:
        return None


def _grab_dot_monitor_img():
    """PIL screenshot of just the dot's monitor (no grid)."""
    from PIL import ImageGrab
    img = ImageGrab.grab(all_screens=True).convert("RGB")
    crop = _dot_monitor_crop(img.size)
    return img.crop(crop) if crop else img


def _take_screenshot(args):
    """Capture the screen to a PNG file and return its path (so it can be emailed/attached/saved)."""
    if not IS_WINDOWS:
        return _result("Taking a screenshot is supported on the Windows app only.", True)
    try:
        from PIL import ImageGrab
        scope = (args.get("scope") or "dot").strip().lower()
        if scope in ("all", "full", "everything", "all_screens", "screens"):
            img = ImageGrab.grab(all_screens=True).convert("RGB")
        else:
            img = _grab_dot_monitor_img()
        os.makedirs(DOWNLOADS_DIR, exist_ok=True)
        name = (args.get("filename") or "").strip()
        if not name:
            name = "screenshot_%d.png" % int(time.time())
        if not name.lower().endswith((".png", ".jpg", ".jpeg")):
            name += ".png"
        path = os.path.join(DOWNLOADS_DIR, name)
        img.save(path)
        return _result(f"Screenshot saved: {path}\n(Pass this path to send_email's attachments to send it.)")
    except Exception as e:
        return _result(f"Failed to take screenshot: {e}", True)


def _grid_b64(img, max_w=1700, step=10):
    """Draw a 0-1000 red coordinate grid on a copy of img and return a base64 PNG.
    `step` is the minor-line spacing in grid units (smaller = finer). The full screen uses a
    readable spacing; the zoomed refine pass uses a finer one (the target is magnified there)."""
    from PIL import ImageDraw, ImageFont
    import io as _io
    import base64 as _b64
    img = img.copy()
    W, H = img.size
    if W > max_w:
        r = max_w / W
        img = img.resize((max_w, int(H * r)))
        W, H = img.size
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except Exception:
        font = ImageFont.load_default()
    # All lines are THIN (width 1) so the model can read exact positions; tiers shown by COLOUR.
    # Minor line every `step`, labelled line every 5% (50) and bold every 10% (100).
    for kk in range(0, 1001, step):
        x = int(kk / 1000.0 * (W - 1))
        y = int(kk / 1000.0 * (H - 1))
        if kk % 100 == 0:
            col, labeled = (235, 0, 0), True
        elif kk % 50 == 0:
            col, labeled = (255, 95, 95), True
        else:
            col, labeled = (255, 200, 200), False
        d.line([(x, 0), (x, H)], fill=col, width=1)
        d.line([(0, y), (W, y)], fill=col, width=1)
        if labeled:
            d.text((min(x + 1, W - 26), 1), str(kk), fill=(190, 0, 0), font=font)
            d.text((1, min(y + 1, H - 15)), str(kk), fill=(190, 0, 0), font=font)
    buf = _io.BytesIO()
    img.save(buf, format="PNG")
    return _b64.b64encode(buf.getvalue()).decode()


def _grab_grid_b64():
    return _grid_b64(_grab_dot_monitor_img())


def _img_b64(img, max_w=1700):
    """Plain (no-grid) base64 PNG of a PIL image, downscaled for the model."""
    import io as _io
    import base64 as _b64
    W, H = img.size
    if W > max_w:
        r = max_w / W
        img = img.resize((max_w, int(H * r)))
    buf = _io.BytesIO()
    img.save(buf, format="PNG")
    return _b64.b64encode(buf.getvalue()).decode()


def _screen_signature():
    """A tiny grayscale thumbnail of the desktop, for detecting that the user acted (screen changed)."""
    from PIL import ImageGrab
    img = ImageGrab.grab(all_screens=True).convert("L").resize((64, 36))
    return list(img.getdata())


def _screens_differ(a, b, thresh=7):
    if not a or not b or len(a) != len(b):
        return True
    diff = sum(abs(x - y) for x, y in zip(a, b)) / len(a)
    return diff > thresh


def _proc_name(u, k, hwnd):
    from ctypes import wintypes
    import ctypes
    pid = wintypes.DWORD()
    u.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    h = k.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return None, pid.value
    buf = ctypes.create_unicode_buffer(512)
    size = wintypes.DWORD(512)
    k.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size))
    k.CloseHandle(h)
    return (os.path.splitext(os.path.basename(buf.value))[0] or None), pid.value


def _foreground_window():
    """(process_name, left, top) of the window the USER is looking at, in PHYSICAL pixels.

    Accessibility frames are WINDOW-RELATIVE, so we add this origin to get screen coords. We skip
    Axon's own windows (the dot/overlay/composer) and pick the topmost real app window, because
    GetForegroundWindow can return our own window right after the composer is dismissed.
    """
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        k = ctypes.windll.kernel32
        u.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        u.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        u.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2 -> physical
        mypid = os.getpid()
        found = {"hwnd": None}

        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def cb(hwnd, _lparam):
            if not u.IsWindowVisible(hwnd) or u.IsIconic(hwnd):
                return True
            if u.GetWindowTextLengthW(hwnd) == 0:
                return True
            _name, pid = _proc_name(u, k, hwnd)
            if pid == mypid:
                return True  # skip Axon's own dot / overlay / composer
            found["hwnd"] = hwnd
            return False  # topmost qualifying window — stop

        u.EnumWindows(WNDENUMPROC(cb), 0)
        hwnd = found["hwnd"] or u.GetForegroundWindow()
        rect = wintypes.RECT()
        u.GetWindowRect(hwnd, ctypes.byref(rect))
        name, _pid = _proc_name(u, k, hwnd)
        return name, rect.left, rect.top
    except Exception:
        return None, 0, 0


def _virtual_screen():
    """(originX, originY, width, height) of the virtual desktop in PHYSICAL pixels — the same
    coordinate space the accessibility frames use (so box fractions are DPI-correct)."""
    if not _VSCREEN_CACHE:
        try:
            from PIL import ImageGrab
            img = ImageGrab.grab(all_screens=True)
            _VSCREEN_CACHE["v"] = (0, 0, img.size[0], img.size[1])
        except Exception:
            _VSCREEN_CACHE["v"] = (0, 0, 0, 0)
    return _VSCREEN_CACHE["v"]
