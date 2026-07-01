"""Web browsing: a persistent debug Chrome on the dot's monitor, a browser-use agent for
navigation, and a deterministic whole-site research crawler."""
import os
import time
import shutil
import subprocess
import urllib.request

from axon import config
from axon.config import BROWSE_MODEL, BROWSE_FALLBACK_MODEL, BROWSE_MAX_STEPS, DOWNLOADS_DIR
from axon.util import _result, NO_WINDOW
from axon.screen import _dot_monitor_env
from axon.mcp import _extract
from axon.prompts import BROWSE_STRATEGY

# Page-name keywords that mark a site page worth visiting during research.
_RESEARCH_KEYWORDS = ["product", "solution", "service", "about", "industr", "application",
                      "catalog", "catalogue", "range", "portfolio", "technolog", "sector", "what-we"]

CDP_PORT = int(os.getenv("BROWSE_CDP_PORT", "9222"))


CHROME_DEBUG_PROFILE = os.getenv(
    "BROWSE_CHROME_PROFILE", os.path.join(config.ASSIST_DIR, ".chrome-debug-profile")
)


def _cdp_reachable(port):
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2) as r:
            return r.status == 200
    except Exception:
        return False


def _find_chrome():
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return shutil.which("chrome")


def _ensure_debug_chrome(port=CDP_PORT, profile=CHROME_DEBUG_PROFILE):
    """Reuse an agent Chrome already running on this port; otherwise launch one. Returns a CDP url or None."""
    if _cdp_reachable(port):  # check first — reuse the existing (possibly logged-in) browser
        return f"http://127.0.0.1:{port}"
    chrome = _find_chrome()
    if not chrome:
        return None
    try:
        flags = [chrome, f"--remote-debugging-port={port}", f"--user-data-dir={profile}"]
        m = _dot_monitor_env()  # open Chrome directly on the dot's monitor (no reposition flash)
        if m:
            mw, mh = m["R"] - m["L"], m["B"] - m["T"]
            w, h = int(mw * 0.86), int(mh * 0.88)
            x, y = m["L"] + (mw - w) // 2, m["T"] + (mh - h) // 2
            flags += [f"--window-position={x},{y}", f"--window-size={w},{h}"]
        flags.append("about:blank")
        subprocess.Popen(flags, creationflags=NO_WINDOW)
    except Exception:
        return None
    for _ in range(40):
        if _cdp_reachable(port):
            return f"http://127.0.0.1:{port}"
        time.sleep(0.5)
    return None


def setup_browser(args=None):
    """One-time onboarding: open Axon's own Chrome so the user can sign in to the accounts they want
    Axon to use. The logins are saved in Axon's browser profile and reused for all future browsing."""
    url = (args or {}).get("url") or "https://accounts.google.com/"
    cdp = _ensure_debug_chrome()   # launch (or reuse) the automation Chrome on the dot's monitor
    if not cdp:
        return _result("Couldn't start Axon's browser — is Google Chrome installed?", True)
    chrome = _find_chrome()
    if chrome:
        try:
            # A second launch with the same profile forwards the URL to the running window and exits.
            subprocess.Popen([chrome, f"--user-data-dir={CHROME_DEBUG_PROFILE}", url],
                             creationflags=NO_WINDOW)
        except Exception:
            pass
    return _result(
        "Opened Axon's browser. Sign in to the accounts you want Axon to use — your Google account, "
        "your intranet/portals, analytics, etc. (complete any 2-step verification). You only need to "
        "do this ONCE: Axon saves these logins in its own browser profile and reuses them for every "
        "future browsing task. Open more sites in the same window to add them, then you can close it.")


def _browse(args):
    """Hand a web task to browser-use (its own Axon intelligence + CDP browser) and return the result."""
    task = (args.get("task") or "").strip()
    if not task:
        return _result("Missing web task.", True)
    try:
        max_pages = int(args.get("max_pages") or 0)  # >0 = multi-page crawl with a hard loop guard
    except (TypeError, ValueError):
        max_pages = 0
    try:
        import asyncio
        from browser_use import Agent, ChatOpenAI, BrowserSession
    except Exception as e:
        return _result(f"browser-use is not available: {e}", True)

    # Check for / start the dedicated agent Chrome and attach to it (persistent login).
    cdp = os.getenv("BROWSE_CDP_URL") or _ensure_debug_chrome()
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)
    before = set(os.listdir(DOWNLOADS_DIR))

    async def _run():
        opts = {
            "downloads_path": DOWNLOADS_DIR,
            "accept_downloads": True,
            # GA4 and other heavy apps reload slowly; wait longer so a slow reload isn't
            # mistaken for a failed click (which caused re-click loops).
            "minimum_wait_page_load_time": float(os.getenv("BROWSE_MIN_WAIT", "1.0")),
            "wait_for_network_idle_page_load_time": float(os.getenv("BROWSE_IDLE_WAIT", "4.0")),
            "wait_between_actions": float(os.getenv("BROWSE_ACTION_WAIT", "1.0")),
        }
        if cdp:
            session = BrowserSession(cdp_url=cdp, **opts)
        else:
            session = BrowserSession(headless=False, **opts)
        agent_kwargs = dict(
            task=task,
            llm=ChatOpenAI(model=BROWSE_MODEL),
            browser_session=session,
            extend_system_message=BROWSE_STRATEGY,
            use_vision=True,                 # let it SEE the page, not just the DOM
            max_failures=int(os.getenv("BROWSE_MAX_FAILURES", "8")),  # tolerate slow-app hiccups
        )
        # If a step's model call comes back empty/unparseable, fall back to a second model for that
        # step instead of looping forever. (Older browser-use versions ignore this kwarg.)
        try:
            agent = Agent(fallback_llm=ChatOpenAI(model=BROWSE_FALLBACK_MODEL), **agent_kwargs)
        except TypeError:
            agent = Agent(**agent_kwargs)

        # Hard loop guard for multi-page crawls (website research). Does NOT depend on the model
        # choosing to stop: if it reloads any page 3+ times (looping) or has already read max_pages
        # distinct pages, force the agent to stop. Only active when max_pages > 0, so single-page
        # apps (dashboards that legitimately reload the same URL) are unaffected.
        from collections import Counter

        def _norm(u):
            return (u or "").split("#")[0].split("?")[0].rstrip("/").lower()

        async def _on_step_end(ag):
            if max_pages <= 0:
                return
            try:
                hist = getattr(ag, "history", None) or getattr(getattr(ag, "state", None), "history", None)
                urls = hist.urls() if (hist is not None and hasattr(hist, "urls")) else None
                if not urls:
                    return
                norm = [n for n in (_norm(u) for u in urls) if n and "about:blank" not in n]
                if not norm:
                    return
                counts = Counter(norm)
                if max(counts.values()) >= 3 or len(counts) >= max_pages:
                    ag.stop()
            except Exception:
                pass

        try:
            history = await agent.run(max_steps=BROWSE_MAX_STEPS, on_step_end=_on_step_end)
        except TypeError:  # older browser-use without on_step_end
            history = await agent.run(max_steps=BROWSE_MAX_STEPS)
        out = history.final_result()
        if not out:
            # Forced-stop or no explicit "done": salvage the notes the agent extracted along the way.
            try:
                chunks = [c for c in (history.extracted_content() or []) if c and str(c).strip()]
                if chunks:
                    out = "\n\n".join(str(c) for c in chunks[-10:])
            except Exception:
                pass
        if not out:
            errs = [e for e in (history.errors() or []) if e]
            out = "Errors: " + "; ".join(errs[-2:]) if errs else "Browser task finished without an explicit result."
        return out

    try:
        out = asyncio.run(_run())
    except Exception as e:
        return _result(f"Browser task failed: {e}", True)
    # Report any files downloaded during this run so the model can read them with read_file.
    new_files = [os.path.join(DOWNLOADS_DIR, f) for f in os.listdir(DOWNLOADS_DIR) if f not in before]
    if new_files:
        out = (out or "") + "\n\nDownloaded files (use read_file to read them):\n" + "\n".join(new_files)
    return _result(out)


def _norm_url(u):
    return (u or "").split("#")[0].split("?")[0].rstrip("/").lower()


def _site_links(start_url):
    """Fetch the homepage HTML and return same-domain page URLs found in it (deterministic, no model)."""
    import urllib.request
    import urllib.parse
    from html.parser import HTMLParser
    req = urllib.request.Request(start_url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"})
    html = urllib.request.urlopen(req, timeout=20).read().decode("utf-8", "ignore")

    class _LinkParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.hrefs = []

        def handle_starttag(self, tag, attrs):
            if tag == "a":
                for k, v in attrs:
                    if k == "href" and v:
                        self.hrefs.append(v)

    p = _LinkParser()
    p.feed(html)
    domain = urllib.parse.urlparse(start_url).netloc.lower().replace("www.", "")
    skip_ext = (".jpg", ".jpeg", ".png", ".gif", ".pdf", ".svg", ".zip", ".webp", ".css", ".js", ".ico", ".mp4")
    out, seen = [], set()
    for h in p.hrefs:
        absu = urllib.parse.urljoin(start_url, h.strip())
        pu = urllib.parse.urlparse(absu)
        if pu.scheme not in ("http", "https"):
            continue
        if pu.netloc.lower().replace("www.", "") != domain:
            continue
        if any(pu.path.lower().endswith(e) for e in skip_ext):
            continue
        key = _norm_url(absu)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(absu)
    return out


def _research_website(args):
    """Deterministic website research: LIST the site's pages up front, then visit each ONCE, in order,
    marking each visited — so the browser never re-crawls the same pages."""
    start = (args.get("url") or args.get("task") or "").strip()
    if not start:
        return _result("Missing website url.", True)
    if not start.startswith("http"):
        start = "https://" + start
    try:
        cap = int(args.get("max_pages") or 6)
    except (TypeError, ValueError):
        cap = 6
    cap = max(2, min(cap, 10))

    # 1) DISCOVER the pages to visit — build the whole worklist before visiting anything.
    try:
        links = _site_links(start)
    except Exception:
        links = []

    def _score(u):
        ul = u.lower()
        for i, k in enumerate(_RESEARCH_KEYWORDS):
            if k in ul:
                return i
        return len(_RESEARCH_KEYWORDS) + 1

    links.sort(key=_score)  # product/about/etc. pages first
    worklist = [start]
    for u in links:
        if all(_norm_url(u) != _norm_url(w) for w in worklist):
            worklist.append(u)
        if len(worklist) >= cap:
            break

    # If discovery found nothing (JS-only nav or blocked), fall back to one guarded browse crawl.
    if len(worklist) <= 1:
        return _browse({
            "task": f"Research the website {start}. Open the homepage, then its main Products/Services, "
                    f"About and Industries pages, scrolling each fully. Visit each page only once and do not "
                    f"re-open pages. Extract detailed notes on the company's products, then stop.",
            "max_pages": cap,
        })

    # 2) VISIT each page ONCE, in order; mark it visited; move to the next.
    visited, notes = [], []
    for i, u in enumerate(worklist, 1):
        r = _browse({
            "task": f"Open this exact page and nothing else: {u}\n"
                    f"Scroll from the top to the very bottom so all content loads. Then extract concise, "
                    f"factual notes about the company's PRODUCTS/SERVICES on THIS page (names, types, features, "
                    f"specifications, applications, industries served). Do NOT click through to other pages — "
                    f"when you've read this page, return the notes.",
            "max_pages": 2,   # allow this one page; hard-stop if it wanders off it
        })
        visited.append(u)
        notes.append(f"### Page {i}: {u}\n{_extract(r)[0].strip()}")

    header = (f"Researched {start} — planned {len(worklist)} pages, visited each once:\n"
              + "\n".join(f"- {u}" for u in visited))
    return _result(header + "\n\n" + "\n\n".join(notes))
