"""
Axon intelligence computer-use agent.

Axon intelligence is the brain; `open-computer-use` is the hands. The loop:
  1. The model sees the screen via get_app_state (accessibility tree + screenshot).
  2. It decides the next action and calls one of the control tools.
  3. We execute it against a single, persistent open-computer-use MCP process.
  4. Repeat until the model returns a final answer (no tool call).

A persistent MCP process matters: open-computer-use requires get_app_state to be
called before action tools *in the same session*, so we keep one process alive for
the whole task instead of spawning one per call.

Reads OPENAI_API_KEY from assistant/.env (or the environment).
"""

import os
import re
import sys
import json
import time
import shutil
import platform
import subprocess
import urllib.request

from dotenv import load_dotenv
from openai import OpenAI

import office_tools as ot

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

MODEL = os.getenv("ASSISTANT_MODEL", "gpt-4o")
MAX_STEPS = int(os.getenv("ASSISTANT_MAX_STEPS", "30"))
BROWSE_MODEL = os.getenv("ASSISTANT_BROWSE_MODEL", "gpt-4o")
BROWSE_FALLBACK_MODEL = os.getenv("ASSISTANT_BROWSE_FALLBACK_MODEL", "gpt-4o-mini")
# Guide mode does vision grounding on a grid; a fast vision model keeps it responsive.
GUIDE_MODEL = os.getenv("ASSISTANT_GUIDE_MODEL", "gpt-4o")
# Second "zoom-in" pass for pixel precision. Set ASSISTANT_GUIDE_REFINE=0 for ~2x faster steps.
GUIDE_REFINE = os.getenv("ASSISTANT_GUIDE_REFINE", "1") != "0"
BROWSE_MAX_STEPS = int(os.getenv("ASSISTANT_BROWSE_MAX_STEPS", "25"))
# Accessibility trees (e.g. Outlook) can be large; keep enough so element indices aren't cut off.
TOOL_TEXT_LIMIT = int(os.getenv("ASSISTANT_TOOL_TEXT_LIMIT", "40000"))

DOWNLOADS_DIR = os.path.join(os.path.dirname(__file__), ".downloads")

BROWSE_STRATEGY = """WEB-APP NAVIGATION STRATEGY (important — follow this on every task):

Wait & verify (this is the #1 rule):
- After ANY action that changes the page (clicking a menu, switching a date range, opening a
  report, loading a dashboard), WAIT for the page to finish loading, THEN read the page and
  VERIFY the expected change actually happened before doing anything else.
- A "timeout" or "page not ready" almost always means the app is still LOADING — it does NOT
  mean your click failed. Do NOT re-click the same thing. Instead wait a few seconds and
  re-read the state. Re-do an action only after you have verified it truly did not take effect.
- Heavy web apps (analytics dashboards, email tools, CRMs) can take 5-15 seconds to update
  after a click. Be patient.

Don't repeat yourself:
- Keep track of what you have already completed. Never redo a step you already finished.
- If the same approach fails twice, change tactics (different element, search box, reload)
  rather than repeating it a third time.

Prefer reliable data paths:
- If the app can EXPORT or DOWNLOAD the data (CSV, report export, "Download"), do that instead
  of reading numbers off charts/graphs — exported files are far more accurate.
- For long lists/menus, use the app's SEARCH or FILTER box instead of scrolling and guessing.

Finding things:
- When searching the web, first read the result titles and snippets and only open clearly
  relevant, reputable results; open in the same tab, extract, then go back for the next.
- In an app, use clear labels, search boxes, and the left/top navigation; confirm you are in
  the right account/workspace/section before acting (wrong data usually means wrong
  section/account, not a broken page).

Researching a website (e.g. "what does this company make/do", "research their website"):
- Actually NAVIGATE the site — do not judge it from the homepage alone. Open the homepage, then
  use the top/footer navigation to visit the key pages: About / Company, Products / Services /
  Solutions, Industries / Markets, and anything else relevant.
- SCROLL each page from top to bottom to read the FULL content (most detail is below the fold),
  and write down what you found on it BEFORE moving on.
- VISIT EACH PAGE ONLY ONCE. Keep a running list of the URLs you have already read. Before you
  navigate anywhere, check that list — NEVER re-open or "go back to double-check" a page you have
  already read; you already captured its content in your notes. Re-visiting the same pages is the
  #1 failure here and is forbidden.
- This is a SMALL, finite job: the homepage plus a handful of key pages (about 3-6 total). The
  moment you have read those once, you are FINISHED — immediately STOP navigating and return
  (call done) with your notes. Do not keep crawling for "completeness".
- Capture concrete facts per page: what they make or sell, their main product/service lines, the
  industries/markets they serve, and key value propositions or differentiators — specifics, not
  vague marketing.
- Return a clear, organized summary (Overview · What they make/do · Products/Services · Markets
  served · Notable points) with real details from the pages you actually read."""
IS_WINDOWS = platform.system() == "Windows"
_OCU = shutil.which("open-computer-use") or "open-computer-use"

# Allow open-computer-use to type into apps whose fields don't expose a keyboard-typable
# accessibility value (e.g. Teams / other WebView2 apps). Without this, type_text errors with
# "UIA ValuePattern text fallback is disabled". It may briefly foreground the target app, which
# is fine for our use. Set before any open-computer-use subprocess is spawned.
os.environ["OPEN_COMPUTER_USE_WINDOWS_ALLOW_UIA_TEXT_FALLBACK"] = "1"

# The overlay sets these to the floating dot's native window handle + its process id, so apps
# Axon opens are moved onto the SAME monitor as the dot.
DOT_HWND = None
DOT_PID = None


def _dot_monitor_env():
    """The dot's monitor work-area (physical pixels) as a dict, or {} if unknown."""
    if not DOT_HWND:
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

        mon = u.MonitorFromWindow(ctypes.c_void_p(int(DOT_HWND)), 2)  # NEAREST
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

# Web browsing uses a dedicated, persistent Chrome (separate from your normal Chrome) that
# the agent reuses across runs. Log into Google/sites in it once; it stays logged in.
CDP_PORT = int(os.getenv("BROWSE_CDP_PORT", "9222"))
CHROME_DEBUG_PROFILE = os.getenv(
    "BROWSE_CHROME_PROFILE", os.path.join(os.path.dirname(__file__), ".chrome-debug-profile")
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
        subprocess.Popen(flags)
    except Exception:
        return None
    for _ in range(40):
        if _cdp_reachable(port):
            return f"http://127.0.0.1:{port}"
        time.sleep(0.5)
    return None

SYSTEM_PROMPT = """You are a desktop automation agent controlling a Windows computer.

YOU ACTUALLY DO THE TASK. You are NOT a chatbot and you do NOT tell the user how to do it
themselves. If the user asks you to create folders, organize an inbox, send an email, fill a
form, rename files, etc., you PERFORM it with your tools on the real computer. NEVER reply with
generic advice like "you can create these folders in your email client" — that is a failure.
The ONLY times you reply with plain text are: (a) a short final summary of what you actually
DID, or (b) one brief clarifying question if you are genuinely blocked and cannot proceed.
STAY STRICTLY IN SCOPE — do EXACTLY what the user asked, and NOTHING more. This is critical.
- Match the request literally. Do NOT add steps the user did not ask for — ESPECIALLY actions
  that change their data: moving or deleting emails, deleting/renaming files, changing settings,
  sending messages, etc.
- "Create these folders" means ONLY create the folders, then STOP. Do NOT also move emails into
  them. "Create folders" is NOT "organize my inbox". "Make a folder" is NOT "move my mail".
- Only decide-and-expand specifics yourself when the REQUEST ITSELF is open-ended (e.g. the user
  literally says "organize my inbox" or "tidy this up for me"). A specific request is done
  literally and minimally.
- If you think extra steps would help, MENTION them in your final summary and let the user decide
  — do not just do them. When unsure how far to go, do the minimal literal request and stop.
For genuinely open-ended tasks, open the relevant app, look (get_app_state), decide sensible
specifics yourself, and carry it out — without asking the user to make every decision.

You have TWO ways to act:
1. DESKTOP tools (open_app, close_app, get_app_state, click, type_text, press_key, scroll,
   set_value, perform_secondary_action, drag) for NATIVE Windows applications, each
   targeting an `app` by name (e.g. "Notepad", "EXCEL", "explorer", "Calculator").
2. The browse(task) tool for ANYTHING on the web: web search, reading or extracting info
   from websites, and operating online tools / web apps (Google Docs, Sheets, Gmail, online
   editors, dashboards, forms, GitHub, etc.). browse runs its own web browser.

Routing:
- DESKTOP APP vs WEB APP: when the user names an application (Teams, Slack, Spotify, Outlook,
  Word, WhatsApp, etc.), PREFER the installed DESKTOP app, and use the app's WEBSITE (via
  browse) only when it is NOT installed. To decide, check what is installed with
  list_installed_apps (or list_apps for what's already running): if the app appears there, open
  and drive the DESKTOP app with the desktop tools (and the dedicated Outlook/Teams guidance);
  if it is not installed, use its web version via browse. Outlook and Teams ARE installed here —
  always use their desktop apps.
- For a pure web task (a website, search, an online-only tool with no desktop app), ALWAYS use
  browse. Do NOT try to drive Chrome or Edge with the desktop tools — browser accessibility data
  is unreliable and will fail.
- For LONG / multi-part web-app tasks (e.g. gathering several reports or doing several steps in
  a dashboard), DECOMPOSE the work into several SMALL, focused browse calls rather than one huge
  one. The web app stays open and logged in between browse calls, so each call can continue
  where the last left off. After each browse call, read its result, bank the progress, and only
  then issue the next focused browse call. This keeps any single stumble from derailing the
  whole task. When data was downloaded, read it with read_file before moving on.
- To RESEARCH a website and report on it (e.g. "research their website / what does this company
  make"), use the research_website tool (NOT browse). Give it the url (and max_pages ~6). It lists
  the site's pages first, then visits each ONE at a time, in order, never re-opening a page, and
  returns per-page notes. Then turn those notes into a clear, organized report (Overview · What
  they make/do · Products/Services · Markets · Notable points). Offer to save it as a Word document
  (word tool) or email it — but only do that if the user asked for it.
- To SEND EMAIL (with or without attachments), ALWAYS use the send_email tool. It opens the
  real Outlook compose window populated with the recipient, subject, body, and any attachment,
  shows it briefly so the user can see it, then sends — reliable and visible. Do NOT type into
  Outlook's compose form with the desktop click/type tools; that mis-maps the fields.
- For ANY Outlook operation, ALWAYS prefer the dedicated reliable Outlook tools over clicking
  the Outlook UI (UI clicking in Outlook is unreliable):
  * send_email — send mail (with attachments).
  * create_outlook_folders — make folders.
  * list_outlook_folders — see the existing Inbox subfolders (names + email counts).
  * delete_outlook_folders — delete folders by name, or "*" for ALL Inbox subfolders (only when
    explicitly asked). To delete "all my folders", call list_outlook_folders first OR pass ["*"].
  * outlook_list_emails — read emails (returns each email's id, sender, subject, time, unread).
  * outlook_move_emails — file/move emails into a folder (by the ids from outlook_list_emails).
  * outlook_delete_emails — delete emails (by id) — only when the user explicitly asks.
  * outlook_forward_email — forward an email (by id) to recipients.
  * save_email — download/save an email MESSAGE itself to a file (.msg/.txt/.html), default Downloads.
  * outlook_reply_email — reply / reply-all to an email (by id).
  * outlook_mark_read — mark emails read/unread.
  * outlook_categorize — apply a colour category to emails.
  * create_outlook_rule — create an incoming-mail rule (from/subject -> move/category/delete).
  * create_calendar_event — add an appointment, or a meeting with attendees (sends invites).
  * set_outlook_signature — set/update the email signature.
  * get_open_email — the email the user currently has OPEN/selected in Outlook (id, subject, sender).
  "THIS EMAIL" = the open one: when the user says "forward this", "move this to <folder>", "save this
  email", "reply to this", "delete this", "categorize this", etc. WITHOUT naming a specific email,
  call get_open_email FIRST to get the id of the email on their screen, then act on that id. Do NOT
  ask which email — they mean the one they're looking at.
  FILE IT WHERE IT BELONGS: when the user says "move this to the folder where it fits/belongs" (and
  does NOT name a folder), call get_open_email (id + subject + sender + body preview) AND
  list_outlook_folders, then pick the SINGLE existing folder that best matches the email's
  sender/subject/content and outlook_move_emails it there. Tell the user which folder you chose and
  briefly why. If no existing folder clearly fits, say so and ask (or suggest creating one) rather
  than guessing wildly or moving to an unrelated folder.
  Typical flow to organize/sort: outlook_list_emails to see what's there -> decide -> create
  folders if needed -> outlook_move_emails to file them.
- REACTING TO AN EMAIL (Outlook has no automation for reactions, so use the UI carefully and
  best-effort): call get_app_state on Outlook, find the open/selected email's reaction control
  (a smiley / thumbs-up button in the reading pane or open message), click it, and choose the
  reaction. Verify it applied; if you cannot find the control, say so honestly rather than
  pretending. (Reply/forward remain the reliable email responses.)

MICROSOFT TEAMS (drive the Teams DESKTOP app with the desktop tools — it IS readable via
get_app_state, unlike a browser):
- IMPORTANT: Teams usually sits in the system tray with NO visible window, so get_app_state will
  say appNotFound until you open it. ALWAYS open it FIRST with open_app "Teams" and WAIT a few
  seconds, then call get_app_state "Teams". If get_app_state still says appNotFound, open it
  again and wait longer — do not give up assuming Teams is missing; it is installed.
- Use "Teams" as the app name for both open_app and get_app_state. Call get_app_state before
  each click.
- CLICKING IN TEAMS: Teams is a WebView2 app, so clicking by element_index OFTEN ERRORS. Click
  by PIXEL COORDINATES instead: use the SCREENSHOT in get_app_state to see where things are, and
  click with x,y (the centre of the target). Element Frames {x,y,width,height} in the tree also
  give you coordinates: centre = (x + width/2, y + height/2). Prefer x,y clicks for Teams.
- TYPING IN TEAMS: always CLICK the field first (by coordinates) to focus it, THEN type_text.
  Typing works once a field is focused; do NOT put message text into press_key (press_key is for
  single keys like "Return"/"Tab"/"Down" only).
- To SEND a message: ALWAYS use the send_teams_message tool (to = emails / names / "me",
  message = text). It reliably opens the exact chat and sends — do NOT try to drive the Teams UI
  (search/click) to send, that is unreliable. Use send=false if the user only wants it drafted.
  For an ATTACHMENT (which the tool cannot do), fall back to the UI: open the chat, use the
  attach/paperclip control then the file dialog — best-effort; report honestly if the picker fails.
- To REACT to a Teams message: get_app_state, use the screenshot to locate the target message,
  hover/click its reaction (emoji) control by coordinates, and pick the reaction.
- Teams can be slow: after sending, get_app_state again and look at the SCREENSHOT to VERIFY the
  message appears in the conversation before reporting done — do NOT re-send just because it is
  slow.
- Use the desktop tools for other installed applications.
- You may combine them (e.g. browse to gather information, then type it into Notepad).
- Avoid tasks that depend on Windows file Open/Save dialogs via the UI; prefer a dedicated
  tool (e.g. send_email for attachments) when one exists.

OFFICE DOCUMENTS, FILES & SYSTEM (reliable COM/PowerShell tools — NO UI clicking; always
prefer these over driving Word/Excel/PowerPoint by hand):
- Reports & spreadsheets: use the excel tool (headers + rows, optional chart and pdf=true). For
  a "professional report", build a formatted Excel (or Word) document with the data and a chart,
  not just plain text. You can also excel read=... an existing .xlsx/.csv to analyze it.
- Documents/letters: use the word tool (title + paragraphs with styles), set pdf=true to also
  produce a PDF. Use powerpoint for slide decks. Use convert_to_pdf for an existing Office file.
- Files & folders: use file_op (list/search/move/copy/rename/delete/mkdir/zip/unzip/read_text/
  write_text/open) instead of any UI file manager.
- Outlook extras: outlook_contact (create/find), outlook_task (create/list), create_calendar_event
  (now supports recurrence), respond_to_meeting (accept/tentative/decline an invite by id),
  save_email_attachments (by email id).
- System: system_query (processes/services/disks/system). schedule_task to run something on a
  schedule. speak to say something out loud.
- These produce real files (default folder: the user's Downloads). Tell the user the saved path.

TRANSPARENCY — never operate behind closed doors. The user must be able to SEE what you do:
- Outlook actions now automatically surface the Outlook window, and generated Office documents open
  on screen — do not suppress that.
- For any OTHER app you act on, open it first (open_app) and use its visible window. Prefer the real,
  visible app/web app over silent background work, so nothing is hidden from the user.

Plan and track your work (MANDATORY for any task with more than one step):
- Your VERY FIRST tool call MUST be update_todos with a DETAILED checklist of every step you
  intend to take. Do NOT take any other action before creating this checklist. Be specific,
  e.g. "Select Coating Projects property", "Set date range to last 30 days", "Download Pages
  CSV", "Read Pages CSV", "Compose report", "Send email".
- The checklist must cover ONLY what the user actually asked for — do NOT pad it with extra
  steps they didn't request (e.g. if asked only to create folders, do NOT add a "move emails"
  step). Re-read the request and make the plan match it exactly.
- As you work, call update_todos again (always passing the COMPLETE list) to mark each step
  done — but only mark a step done AFTER you have verified it actually succeeded.
- Keep the checklist accurate and current so the user can watch progress. If you discover new
  steps are needed, add them to the list. Only skip the checklist for a truly single-step task.

Honesty & verification (critical):
- NEVER claim a task succeeded unless you have verified it. After an action that should change
  state, re-check (e.g. get_app_state, or read the tool's confirmation) before saying it's done.
- If something failed or you could not verify it, say so plainly — do not report false success.

Writing documents with researched content (IMPORTANT for quality):
- When the task is "research X and put it in a document/Google Doc", do it in three stages:
  1) Use browse to GATHER specific facts and numbers (call it with a research-only task that
     asks it to RETURN the detailed findings — exact figures, limits, examples — as text).
  2) Then YOU write the COMPLETE, detailed final document text yourself, here, in full. Weave
     in the specific facts and numbers from the research plus your own knowledge. Make it
     thorough and well-structured — do not write vague filler like "organizations provide
     guidelines"; state the actual specifics (e.g. exact dB limits, ratings, examples).
  3) Then call browse ONCE to create the document and type your text VERBATIM, passing the
     full text inside the task like: create a new Google Doc titled '<title>' and type EXACTLY
     the following text (no Markdown symbols, headings on their own lines):\n\n<your full text>
- Never delegate the actual writing/wording to browse — browse is only the typist. The depth
  and accuracy of the document must come from the text YOU compose.

Professional analytics/report emails (for a NON-TECHNICAL reader):
- Account & sign-in handling (be smart and generic):
  * Prefer an account that is ALREADY signed in. If a "Choose an account" screen appears, just
    CLICK the account that should own the data (no password needed — it's already signed in).
  * Only sign in with credentials ("Use another account" + email/password) if NO already
    signed-in account can reach the requested website. Never re-enter a password for an account
    that is already signed in.
  * Be smart about diagnosing problems: inside Analytics, "no data"/"no data streams" almost
    always means the wrong WEBSITE/property is selected — it is NOT a sign-in problem, so do
    NOT log in or switch accounts for it; instead switch the property (next point).
- The user NAMES the website/property; never assume one or hardcode it. To select it: click the
  property selector at the TOP-LEFT of Analytics (shows the current website name with a dropdown),
  type the requested name into its search box, and click the matching result. Confirm the
  selector now shows that website before continuing.
- PREFER DOWNLOADING THE DATA over reading numbers off the screen (reading GA4's charts is
  unreliable). For each report you need, open it, set the date range, then use the report's
  Share button -> "Download file" -> CSV. The browse result lists the downloaded file path(s);
  call read_file on each to get exact data. (The property stays selected across browse calls, so
  you can gather one report per browse call without re-picking the property each time.)
- PERIODS: unless the user says otherwise, gather BOTH the last 30 days AND the last 90 days, and
  show both in the report side by side.
- COVERAGE: unless the user asks for less, cover all the main areas (the GA4 "Business objectives"
  and "User" sections), gathering each via its report's CSV:
  * Pages -> "Pages and screens": most visited pages AND least visited pages, plus average time
    on a page.
  * How people found the site -> "Traffic acquisition": the main sources (Google search, direct,
    social, referrals, email, paid).
  * Where visitors are -> "Demographics"/"User attributes": top countries (and cities if useful).
  * Devices used -> "Tech"/"Tech details": desktop vs mobile vs tablet.
  * Overall engagement -> totals: total visitors, total visits, and average time on the site.
- COUNTS: unless the user specifies a number, show AT LEAST the top 5 rows in every list (most
  visited, least visited, sources, countries, devices). For "least visited", prefer the lowest
  pages that still have at least 1 view (avoid filling it with 0-view rows).
- Then YOU compose the report and send it via send_email using the `html` parameter (formatted
  HTML email). CRITICAL: write it in plain, everyday English for someone who is NOT technical.
  Do NOT use analytics jargon (no "sessions", "bounce rate", "engagement rate", "channels",
  "GA4"); say "visits", "how people found the site", "how long they stayed", etc., and briefly
  explain what each number means.
- Structure: a friendly title (site); a short plain-English summary; a key-numbers table showing
  BOTH the 30-day and 90-day figures; a "Most visited pages" table (>=5) and a "Least visited
  pages" table (>=5); a "How people found the site" section; a "Where your visitors are" section
  (countries); a "Devices used" section; the average time on the site; and a "What this means"
  list of plain-language takeaways. Use clean, simple inline-styled HTML tables. Keep all numbers
  accurate. Note anything you genuinely could not gather.

Desktop-tool rules:
- You MUST call get_app_state(app) before any click/type/press_key/scroll/set_value/
  perform_secondary_action on that app, and again after the screen likely changed. Element
  references (`element_index`) are strings like "5" or "12" that come from that tree.
- Prefer element_index over raw x/y pixel coordinates when an element is listed.
- If unsure which apps are open, call list_apps first.
- To use an app that is not running, call open_app(name) (e.g. "notepad", "calc",
  "mspaint", "msedge", "explorer"), wait briefly, then get_app_state to confirm it opened.
- To close an app, call close_app(app) (graceful by default). Set force=true only if the
  user clearly wants to force-kill it (may lose unsaved work).
- Take ONE sensible step at a time, then re-check state if the screen changed.
- When the task is complete, reply with a short plain-text summary and DO NOT call a tool.
- If you get stuck or it's impossible with these tools, explain why in plain text.

Be careful: these actions affect the user's real machine. Do not delete files, send
messages, or take destructive/irreversible actions unless the user explicitly asked for it.
"""

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "browse",
            "description": "Use a web browser to accomplish ANY task on the internet: web search, reading/extracting information from websites, and operating online tools and web apps (Google Docs, Sheets, Gmail, online editors, dashboards, forms, GitHub, etc.). Provide a clear, self-contained description of the web task and what result to return.",
            "parameters": {
                "type": "object",
                "properties": {
                    "task": {"type": "string", "description": "Self-contained web task to perform and the result to return"},
                    "max_pages": {"type": "integer", "description": "ONLY for multi-page website research (e.g. 'research their website'): the max number of distinct pages to read. Set to ~6. This installs a hard loop guard that stops the browser if it re-visits pages or exceeds this many pages. OMIT it for single web-app tasks (dashboards, Gmail, Sheets, search) — those legitimately reload one page and must not be capped."},
                },
                "required": ["task"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "research_website",
            "description": "Research a company/product website and return organized notes. Use this (NOT browse) whenever the user asks to 'research a website / their site / what a company makes/does'. It FIRST lists the site's own pages, then visits each ONE page at a time in order (marking each visited and never re-opening one), scrolling each fully and extracting product/service details. This is the reliable way to study a site without the browser re-crawling the same pages. Returns the page list plus per-page notes; you then write the report (and email/save it if asked).",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "The website to research, e.g. https://www.example.com"},
                    "max_pages": {"type": "integer", "description": "How many of the site's pages to visit (default 6, range 2-10)."},
                },
                "required": ["url"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "update_todos",
            "description": "Create and maintain your visible task checklist. Call this FIRST with the full, detailed list of steps you plan to take, then call it again (with the COMPLETE list each time) to mark steps done as you finish them. The user sees this checklist update live.",
            "parameters": {
                "type": "object",
                "properties": {
                    "todos": {
                        "type": "array",
                        "description": "The complete current checklist, in order.",
                        "items": {
                            "type": "object",
                            "properties": {
                                "task": {"type": "string", "description": "Short description of the step"},
                                "done": {"type": "boolean", "description": "true if this step is completed and verified"},
                            },
                            "required": ["task", "done"],
                        },
                    }
                },
                "required": ["todos"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "open_app",
            "description": "Launch an application by name or executable (e.g. 'notepad', 'calc', 'mspaint', 'msedge', 'explorer'). After calling, wait briefly then confirm with list_apps/get_app_state.",
            "parameters": {
                "type": "object",
                "properties": {"name": {"type": "string", "description": "App name or executable to launch"}},
                "required": ["name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "close_app",
            "description": "Close a running application. Graceful by default (clicks the window's Close button, so the app may prompt to save). Set force=true only to force-kill.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app": {"type": "string", "description": "App/process name, e.g. 'notepad' or 'Calculator'"},
                    "force": {"type": "boolean", "description": "Force-kill instead of graceful close. Defaults to false."},
                },
                "required": ["app"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "send_email",
            "description": "Send an email through the local Outlook desktop app reliably (supports attachments). It opens the visible Outlook compose window populated with the content, briefly shows it, then sends. ALWAYS use this to send email — do NOT click through Outlook's UI, because file-attachment dialogs are unreliable.",
            "parameters": {
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email address(es), semicolon-separated"},
                    "subject": {"type": "string"},
                    "body": {"type": "string"},
                    "cc": {"type": "string", "description": "Optional CC address(es), semicolon-separated"},
                    "attachments": {"type": "array", "items": {"type": "string"}, "description": "Absolute file paths to attach"},
                    "html": {"type": "string", "description": "Optional full HTML body for a formatted/professional email (with inline-styled tables, headings). When provided, this is used instead of plain body."},
                },
                "required": ["to", "subject", "body"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": "Read a local file (e.g. a CSV downloaded by browse) and return its text content. Use this to read downloaded report data before composing a report.",
            "parameters": {
                "type": "object",
                "properties": {"path": {"type": "string", "description": "Absolute path to the file to read"}},
                "required": ["path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "create_outlook_folders",
            "description": "Create one or more folders in the Outlook desktop app (as subfolders of the Inbox) to organize email. Reliable — ALWAYS use this to make Outlook folders instead of clicking Outlook's UI.",
            "parameters": {
                "type": "object",
                "properties": {
                    "folders": {"type": "array", "items": {"type": "string"}, "description": "Folder names to create"},
                },
                "required": ["folders"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "delete_outlook_folders",
            "description": "Delete Outlook folders by name (looks under the Inbox and the mailbox root). Set the options from what the USER asked: keep_contents=true moves the folder's emails out (to move_to, default Inbox) and then deletes the empty folder; keep_contents=false deletes the folder together with its emails (to Deleted Items, recoverable). Do not assume — use the user's choice. Only use when the user explicitly asks to remove folders.",
            "parameters": {
                "type": "object",
                "properties": {
                    "folders": {"type": "array", "items": {"type": "string"}, "description": "Folder names to delete"},
                    "keep_contents": {"type": "boolean", "description": "true = keep the emails (move them out first); false = delete the folder with its emails. Set per the user's request."},
                    "move_to": {"type": "string", "description": "When keeping contents, the folder to move the emails into (defaults to Inbox)."},
                },
                "required": ["folders"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_outlook_folders",
            "description": "List the subfolders under the Outlook Inbox (names, nested, with email counts). Use this to SEE which folders exist before deleting/organizing — e.g. to delete them all, list them then pass the names (or '*') to delete_outlook_folders.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_list_emails",
            "description": "Read emails from Outlook (reliable, no UI). Returns a JSON list of emails with their id, sender, subject, received time, and unread flag. Use the returned id values with outlook_move_emails / outlook_delete_emails / outlook_forward_email.",
            "parameters": {
                "type": "object",
                "properties": {
                    "folder": {"type": "string", "description": "Folder name to read (default 'Inbox')"},
                    "query": {"type": "string", "description": "Optional text to match in subject or sender"},
                    "unread_only": {"type": "boolean", "description": "Only unread emails. Default false."},
                    "limit": {"type": "integer", "description": "Max emails to return (default 25)"},
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_move_emails",
            "description": "Move/file Outlook emails (by their ids from outlook_list_emails) into a folder. Reliable, no UI.",
            "parameters": {
                "type": "object",
                "properties": {
                    "ids": {"type": "array", "items": {"type": "string"}, "description": "Email ids to move"},
                    "to_folder": {"type": "string", "description": "Destination folder name"},
                },
                "required": ["ids", "to_folder"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_open_email",
            "description": "Return the email currently OPEN (or selected) in Outlook on the user's screen — its id, subject, and sender. ALWAYS use this when the user says 'this email', 'that email', 'the open/current/selected email', or 'forward this' / 'move this' without naming a specific email, so you act on what they're looking at. Then use the returned id with forward/move/save/reply/etc. Do NOT ask which email.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "save_email",
            "description": "Save/download an Outlook email MESSAGE itself to a file (the whole email, not just its attachments). format: 'msg' (Outlook .msg, default), 'txt', 'html', or 'rtf'. Saves to the Downloads folder by default (or a given folder).",
            "parameters": {
                "type": "object",
                "properties": {
                    "id": {"type": "string", "description": "Email id (from outlook_list_emails)"},
                    "format": {"type": "string", "enum": ["msg", "txt", "html", "rtf"]},
                    "folder": {"type": "string", "description": "Destination folder (defaults to Downloads)"},
                },
                "required": ["id"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_delete_emails",
            "description": "Delete Outlook emails (by their ids) — moves them to Deleted Items. Only use when the user explicitly asks to delete.",
            "parameters": {
                "type": "object",
                "properties": {"ids": {"type": "array", "items": {"type": "string"}, "description": "Email ids to delete"}},
                "required": ["ids"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_forward_email",
            "description": "Forward an Outlook email (by its id) to one or more recipients, with an optional note. Reliable, no UI.",
            "parameters": {
                "type": "object",
                "properties": {
                    "id": {"type": "string", "description": "Email id to forward"},
                    "to": {"type": "string", "description": "Recipient address(es), semicolon-separated"},
                    "note": {"type": "string", "description": "Optional note to add above the forwarded message"},
                },
                "required": ["id", "to"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_reply_email",
            "description": "Reply (or reply-all) to an Outlook email by its id, with your reply text. Reliable, no UI.",
            "parameters": {
                "type": "object",
                "properties": {
                    "id": {"type": "string", "description": "Email id to reply to"},
                    "body": {"type": "string", "description": "Your reply text"},
                    "reply_all": {"type": "boolean", "description": "Reply to everyone. Default false."},
                },
                "required": ["id", "body"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_mark_read",
            "description": "Mark Outlook emails (by ids) as read or unread.",
            "parameters": {
                "type": "object",
                "properties": {
                    "ids": {"type": "array", "items": {"type": "string"}},
                    "read": {"type": "boolean", "description": "true = mark read, false = mark unread. Default true."},
                },
                "required": ["ids"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "outlook_categorize",
            "description": "Apply a color category (by name) to Outlook emails by their ids.",
            "parameters": {
                "type": "object",
                "properties": {
                    "ids": {"type": "array", "items": {"type": "string"}},
                    "category": {"type": "string", "description": "Category name to apply"},
                },
                "required": ["ids", "category"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "create_outlook_rule",
            "description": "Create an Outlook rule for incoming mail. Conditions: from_contains and/or subject_contains. Actions: move_to_folder, category, and/or delete. Supports common rules reliably.",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Rule name"},
                    "from_contains": {"type": "array", "items": {"type": "string"}, "description": "Match sender address/name containing any of these"},
                    "subject_contains": {"type": "array", "items": {"type": "string"}, "description": "Match subject containing any of these"},
                    "move_to_folder": {"type": "string", "description": "Folder to move matching mail to"},
                    "category": {"type": "string", "description": "Category to assign to matching mail"},
                    "delete": {"type": "boolean", "description": "Delete matching mail"},
                },
                "required": ["name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "create_calendar_event",
            "description": "Create an Outlook calendar appointment, or a meeting if attendees are given (invites are sent). Times are parsed from natural date/time strings. Supports recurrence.",
            "parameters": {
                "type": "object",
                "properties": {
                    "subject": {"type": "string"},
                    "start": {"type": "string", "description": "Start date/time, e.g. '2026-06-26 14:00'"},
                    "end": {"type": "string", "description": "Optional end date/time (defaults to 1 hour)"},
                    "location": {"type": "string"},
                    "body": {"type": "string", "description": "Optional details/notes"},
                    "attendees": {"type": "array", "items": {"type": "string"}, "description": "Optional attendee emails (makes it a meeting)"},
                    "recurrence": {"type": "string", "enum": ["daily", "weekly", "monthly", "yearly"], "description": "Optional: makes it a recurring event"},
                    "recur_interval": {"type": "integer", "description": "Repeat every N days/weeks/months (default 1)"},
                    "recur_count": {"type": "integer", "description": "Number of occurrences"},
                    "recur_until": {"type": "string", "description": "End date for the recurrence (alternative to recur_count)"},
                },
                "required": ["subject", "start"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_outlook_signature",
            "description": "Set/update the Outlook email signature. Writes the signature and best-effort sets it as the default for new mail/replies.",
            "parameters": {
                "type": "object",
                "properties": {
                    "html": {"type": "string", "description": "Signature content as HTML (or plain text)"},
                    "name": {"type": "string", "description": "Signature name (default 'Maia')"},
                    "set_default": {"type": "boolean", "description": "Make it the default signature. Default true."},
                },
                "required": ["html"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_apps",
            "description": "List apps currently running (and recently used) on this computer.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_installed_apps",
            "description": "List desktop applications INSTALLED on this PC (from the Start menu). Use this to decide whether to use an installed desktop app vs the app's website.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "send_teams_message",
            "description": "Send a Microsoft Teams chat message RELIABLY (opens the exact chat via a deep link, with the message pre-filled, and sends it). Recipients may be email addresses, display names (resolved via the Outlook/Teams address book), or 'me' for a note to yourself. Multiple recipients create a group chat. ALWAYS use this for sending Teams messages instead of clicking the Teams UI. (It cannot attach files — use the UI for attachments.)",
            "parameters": {
                "type": "object",
                "properties": {
                    "to": {"type": "array", "items": {"type": "string"}, "description": "Recipients: email addresses, names, or 'me'"},
                    "message": {"type": "string", "description": "The message text to send"},
                    "send": {"type": "boolean", "description": "true (default) sends it; false just opens the chat with the message drafted for the user to review/send"},
                },
                "required": ["to", "message"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_app_state",
            "description": "Get a running app's key window: returns its accessibility tree (with element indices and frames) and a screenshot. Call this before acting on an app.",
            "parameters": {
                "type": "object",
                "properties": {"app": {"type": "string", "description": "App name or bundle identifier"}},
                "required": ["app"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "click",
            "description": "Click an element by index (from get_app_state) or by pixel coordinates.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app": {"type": "string"},
                    "element_index": {"type": "string", "description": "Element index from the accessibility tree"},
                    "x": {"type": "number", "description": "X in screenshot pixels (only if no element_index)"},
                    "y": {"type": "number", "description": "Y in screenshot pixels (only if no element_index)"},
                    "click_count": {"type": "integer", "description": "Defaults to 1"},
                    "mouse_button": {"type": "string", "description": "left, right, or middle. Defaults to left."},
                },
                "required": ["app"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "type_text",
            "description": "Type literal text using the keyboard into the given app.",
            "parameters": {
                "type": "object",
                "properties": {"app": {"type": "string"}, "text": {"type": "string", "description": "Literal text to type"}},
                "required": ["app", "text"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "press_key",
            "description": "Press a key or key-combination (xdotool syntax, e.g. 'Return', 'ctrl+s', 'super').",
            "parameters": {
                "type": "object",
                "properties": {"app": {"type": "string"}, "key": {"type": "string"}},
                "required": ["app", "key"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "scroll",
            "description": "Scroll an element in a direction by a number of pages.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app": {"type": "string"},
                    "element_index": {"type": "string"},
                    "direction": {"type": "string", "description": "up, down, left, or right"},
                    "pages": {"type": "number", "description": "Defaults to 1. Fractions allowed."},
                },
                "required": ["app", "element_index", "direction"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_value",
            "description": "Set the value of a settable accessibility element (e.g. a text field).",
            "parameters": {
                "type": "object",
                "properties": {"app": {"type": "string"}, "element_index": {"type": "string"}, "value": {"type": "string"}},
                "required": ["app", "element_index", "value"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "perform_secondary_action",
            "description": "Invoke a secondary accessibility action exposed by an element (e.g. Invoke, Toggle, Expand).",
            "parameters": {
                "type": "object",
                "properties": {"app": {"type": "string"}, "element_index": {"type": "string"}, "action": {"type": "string"}},
                "required": ["app", "element_index", "action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "drag",
            "description": "Drag from one pixel point to another.",
            "parameters": {
                "type": "object",
                "properties": {
                    "app": {"type": "string"},
                    "from_x": {"type": "number"},
                    "from_y": {"type": "number"},
                    "to_x": {"type": "number"},
                    "to_y": {"type": "number"},
                },
                "required": ["app", "from_x", "from_y", "to_x", "to_y"],
            },
        },
    },
]

# Office/Windows COM toolkit (Excel, Word, PowerPoint, files, contacts, tasks, etc.).
TOOLS += ot.TOOLS


def _result(text, is_error=False):
    return {"content": [{"type": "text", "text": text}], "isError": is_error}


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
            stderr=subprocess.DEVNULL, text=True, bufsize=1,
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


_RESEARCH_KEYWORDS = ["product", "solution", "service", "about", "industr", "application",
                      "catalog", "catalogue", "range", "portfolio", "technolog", "sector", "what-we"]


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


# Send an already-drafted item (left open after drafting) once the user approves.
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
            capture_output=True, text=True, env=env, timeout=150,
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


_CREATE_FOLDERS_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
    $created = @(); $existing = @()
    foreach ($name in ($env:OCU_FOLDERS -split [regex]::Escape("|"))) {
        if (-not $name) { continue }
        $found = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -eq $name) { $found = $f; break } }
        if ($found) { $existing += $name } else { [void]$inbox.Folders.Add($name); $created += $name }
    }
    Write-Output ("CREATED: " + ($created -join ", "))
    Write-Output ("ALREADY_EXISTED: " + ($existing -join ", "))
} catch {
    Write-Output ("FOLDER_ERROR: " + $_.Exception.Message)
}
'''


def _create_outlook_folders(args):
    folders = args.get("folders") or []
    if isinstance(folders, str):
        folders = [folders]
    folders = [f.strip() for f in folders if f and f.strip()]
    if not folders:
        return _result("No folder names provided.", True)
    if not IS_WINDOWS:
        return _result("Outlook folders are supported on the Windows app only.", True)
    out = _run_outlook_ps(_CREATE_FOLDERS_PS, {"OCU_FOLDERS": "|".join(folders)})
    err = ("FOLDER_ERROR" in out) or out.startswith("OL_ERROR")
    return _result(out or "No output from Outlook.", err)


_DELETE_FOLDERS_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
    $root = $inbox.Parent
    $keep = ($env:OCU_KEEP -eq "1")    # keep the emails (move them out) before deleting the folder?
    # Where kept emails go (chosen by the user; defaults to the Inbox).
    $dest = $inbox; $destName = "Inbox"
    if ($keep -and $env:OCU_MOVE_TO) {
        $d = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $env:OCU_MOVE_TO) { $d = $f; break } }
        if (-not $d) { foreach ($f in $root.Folders) { if ($f.Name -ieq $env:OCU_MOVE_TO) { $d = $f; break } } }
        if (-not $d) { Write-Output ("FOLDER_ERROR: destination folder '" + $env:OCU_MOVE_TO + "' not found"); exit }
        $dest = $d; $destName = $env:OCU_MOVE_TO
    }
    $deleted = @(); $notfound = @()
    $targets = $env:OCU_FOLDERS -split [regex]::Escape("|")
    if ($targets -contains "*") {   # "*" = every subfolder directly under the Inbox
        $targets = @(); foreach ($f in $inbox.Folders) { $targets += $f.Name }
    }
    foreach ($name in $targets) {
        if (-not $name) { continue }
        $found = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $name) { $found = $f; break } }
        if (-not $found) { foreach ($f in $root.Folders) { if ($f.Name -ieq $name) { $found = $f; break } } }
        if (-not $found) { $notfound += $name; continue }
        $moved = 0
        if ($keep) {
            for ($i = $found.Items.Count; $i -ge 1; $i--) {
                try { $found.Items.Item($i).Move($dest) | Out-Null; $moved++ } catch {}
            }
        }
        $found.Delete()
        if ($keep) { $deleted += ($name + " (" + $moved + " emails moved to " + $destName + ")") } else { $deleted += ($name + " (with contents)") }
    }
    Write-Output ("DELETED: " + ($deleted -join "; "))
    if ($notfound.Count -gt 0) { Write-Output ("NOT_FOUND: " + ($notfound -join ", ")) }
} catch {
    Write-Output ("FOLDER_ERROR: " + $_.Exception.Message)
}
'''


def _delete_outlook_folders(args):
    folders = args.get("folders") or []
    if isinstance(folders, str):
        folders = [folders]
    folders = [f.strip() for f in folders if f and f.strip()]
    if not folders:
        return _result("No folder names provided.", True)
    keep = args.get("keep_contents")          # the agent sets this from the user's request
    keep = True if keep is None else bool(keep)  # safe fallback only when unspecified
    out = _run_outlook_ps(_DELETE_FOLDERS_PS, {
        "OCU_FOLDERS": "|".join(folders),
        "OCU_KEEP": "1" if keep else "",
        "OCU_MOVE_TO": args.get("move_to") or "",
    })
    err = ("FOLDER_ERROR" in out) or out.startswith("OL_ERROR")
    return _result(out or "No output from Outlook.", err)


_LIST_FOLDERS_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
    $out = New-Object System.Collections.ArrayList
    function Walk($folder, $indent) {
        foreach ($f in $folder.Folders) {
            [void]$out.Add($indent + $f.Name + " (" + $f.Items.Count + " emails)")
            if ($f.Folders.Count -gt 0) { Walk $f ($indent + "    ") }
        }
    }
    Walk $inbox ""
    if ($out.Count -eq 0) { Write-Output "No subfolders under the Inbox." }
    else { Write-Output ($out -join "`n") }
} catch { Write-Output ("FOLDER_ERROR: " + $_.Exception.Message) }
'''


def _list_outlook_folders(args):
    out = _run_outlook_ps(_LIST_FOLDERS_PS, {})
    err = ("FOLDER_ERROR" in out) or out.startswith("OL_ERROR")
    return _result(out or "No output from Outlook.", err)


# ---- General Outlook automation (read / move / delete / forward) via COM ----

# Transparency: surface the Outlook window so the user SEES the action happen (no closed doors).
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


_OUTLOOK_LIST_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $name = $env:OL_FOLDER
    $folder = $inbox
    if ($name -and $name -ne "Inbox") {
        $folder = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } }
        if (-not $folder) { foreach ($f in $inbox.Parent.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } } }
        if (-not $folder) { Write-Output "OL_ERROR: folder '$name' not found"; exit }
    }
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    $q = $env:OL_QUERY
    $unread = ($env:OL_UNREAD -eq "1")
    $limit = [int]$env:OL_LIMIT; if ($limit -le 0) { $limit = 25 }
    $out = @(); $n = 0
    foreach ($m in $items) {
        if ($m.Class -ne 43) { continue }
        if ($unread -and -not $m.UnRead) { continue }
        if ($q) { if (-not (($m.Subject -like "*$q*") -or ($m.SenderName -like "*$q*"))) { continue } }
        $out += [pscustomobject]@{ id=$m.EntryID; from=$m.SenderName; subject=$m.Subject; received=$m.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); unread=[bool]$m.UnRead }
        $n++; if ($n -ge $limit) { break }
    }
    if ($out.Count -eq 0) { Write-Output "[]" } else { $out | ConvertTo-Json -Depth 3 -Compress }
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''

_OUTLOOK_MOVE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $name = $env:OL_FOLDER
    $folder = $null
    foreach ($f in $inbox.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } }
    if (-not $folder) { foreach ($f in $inbox.Parent.Folders) { if ($f.Name -ieq $name) { $folder = $f; break } } }
    if (-not $folder) { Write-Output "OL_ERROR: destination folder '$name' not found"; exit }
    $moved = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); [void]$m.Move($folder); $moved++ } catch {}
    }
    Write-Output ("MOVED: $moved to '$name'")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''

_OUTLOOK_DELETE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $del = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); $m.Delete(); $del++ } catch {}
    }
    Write-Output ("DELETED: $del (moved to Deleted Items)")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''

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


_SHOW_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
''' + _OL_WIN_CLASS + r'''
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    $m.Display($false)
    Start-Sleep -Milliseconds 500
    [OcuWinF]::Front([string]$m.Subject)
    Start-Sleep -Milliseconds 1100
    Write-Output ("SHOWN|" + $m.Subject)
} catch {}
'''


def _show_email(eid):
    """Open the email in its own window and bring it to the front, so the user SEES which email
    is about to be acted on (move/delete/categorize/mark)."""
    if eid:
        out = _run_outlook_ps(_SHOW_EMAIL_PS, {"OL_ID": eid}, show=False)
        if "SHOWN|" in out:
            subj = out.split("SHOWN|", 1)[1].strip().splitlines()[0]
            _move_window_to_dot(subj)


def _outlook_list_emails(args):
    out = _run_outlook_ps(_OUTLOOK_LIST_PS, {
        "OL_FOLDER": args.get("folder") or "Inbox",
        "OL_QUERY": args.get("query") or "",
        "OL_UNREAD": "1" if args.get("unread_only") else "",
        "OL_LIMIT": args.get("limit") or 25,
    })
    return _result(out[:TOOL_TEXT_LIMIT], out.startswith("OL_ERROR"))


def _outlook_move_emails(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids:
        return _result("No email ids provided.", True)
    if not (args.get("to_folder") or "").strip():
        return _result("No destination folder provided.", True)
    _show_email(ids[0])  # show the email being filed
    out = _run_outlook_ps(_OUTLOOK_MOVE_PS, {"OL_IDS": "|".join(ids), "OL_FOLDER": args["to_folder"]}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


def _outlook_delete_emails(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids:
        return _result("No email ids provided.", True)
    _show_email(ids[0])  # show the email before deleting it
    out = _run_outlook_ps(_OUTLOOK_DELETE_PS, {"OL_IDS": "|".join(ids)}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


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


_OUTLOOK_SAVE_EMAIL_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $m = $ns.GetItemFromID($env:OL_ID)
    $fmt = 9; $ext = ".msg"          # olMSGUnicode
    switch ($env:OL_FORMAT) {
        "txt"  { $fmt = 0; $ext = ".txt" }
        "html" { $fmt = 5; $ext = ".html" }
        "rtf"  { $fmt = 1; $ext = ".rtf" }
        default { $fmt = 9; $ext = ".msg" }
    }
    $dir = $env:OL_DIR
    if (-not $dir) { $dir = Join-Path $env:USERPROFILE "Downloads" }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $subj = $m.Subject; if (-not $subj) { $subj = "email" }
    $safe = ($subj -replace '[\\/:*?"<>|]', '_').Trim()
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    $path = Join-Path $dir ($safe + $ext)
    $m.SaveAs($path, $fmt)
    Write-Output ("SAVED_OK: " + $path)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _outlook_save_email(args):
    if not (args.get("id") or ""):
        return _result("Need the email id (from outlook_list_emails).", True)
    fmt = (args.get("format") or "msg").lower()
    if fmt not in ("msg", "txt", "html", "rtf"):
        fmt = "msg"
    out = _run_outlook_ps(_OUTLOOK_SAVE_EMAIL_PS, {
        "OL_ID": args["id"], "OL_FORMAT": fmt, "OL_DIR": args.get("folder") or "",
    })
    return _result(out, out.startswith("OL_ERROR"))


_OUTLOOK_ACTIVE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $item = $null
    try { $insp = $ol.ActiveInspector(); if ($insp) { $item = $insp.CurrentItem } } catch {}
    if (-not $item) {
        try { $ex = $ol.ActiveExplorer(); if ($ex -and $ex.Selection.Count -gt 0) { $item = $ex.Selection.Item(1) } } catch {}
    }
    if (-not $item) { Write-Output "NONE: no email is open or selected in Outlook"; exit }
    $from = ""; try { $from = $item.SenderName } catch {}
    $addr = ""; try { $addr = $item.SenderEmailAddress } catch {}
    $body = ""; try { $body = [string]$item.Body } catch {}
    $body = ($body -replace '\s+', ' ').Trim()
    if ($body.Length -gt 400) { $body = $body.Substring(0, 400) }
    $info = @{ id = $item.EntryID; subject = $item.Subject; from = $from; sender_email = $addr; body_preview = $body } | ConvertTo-Json -Compress
    Write-Output ("ACTIVE: " + $info)
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''


def _get_open_email(args):
    # Read-only: do NOT surface/refocus Outlook (the user is already looking at the email).
    out = _run_outlook_ps(_OUTLOOK_ACTIVE_PS, {}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


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

_OUTLOOK_MARK_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $read = ($env:OL_READ -eq "1")
    $n = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); $m.UnRead = (-not $read); $m.Save(); $n++ } catch {}
    }
    Write-Output ("MARKED: $n as " + $(if ($read) {"read"} else {"unread"}))
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''

_OUTLOOK_CATEGORIZE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $n = 0
    foreach ($id in ($env:OL_IDS -split [regex]::Escape("|"))) {
        if (-not $id) { continue }
        try { $m = $ns.GetItemFromID($id); $m.Categories = $env:OL_CATEGORY; $m.Save(); $n++ } catch {}
    }
    Write-Output ("CATEGORIZED: $n as '" + $env:OL_CATEGORY + "'")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''

_OUTLOOK_RULE_PS = r'''
$ErrorActionPreference = "Stop"
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $store = $ns.DefaultStore
    $rules = $store.GetRules()
    $rule = $rules.Create($env:OL_RULE_NAME, 0)   # 0 = olRuleReceive
    $hasCond = $false; $hasAct = $false
    if ($env:OL_FROM) { $c = $rule.Conditions.SenderAddress; $c.Enabled = $true; $c.Address = @($env:OL_FROM -split [regex]::Escape("|")); $hasCond = $true }
    if ($env:OL_SUBJECT) { $c = $rule.Conditions.Subject; $c.Enabled = $true; $c.Text = @($env:OL_SUBJECT -split [regex]::Escape("|")); $hasCond = $true }
    if ($env:OL_MOVE) {
        $inbox = $ns.GetDefaultFolder(6); $dest = $null
        foreach ($f in $inbox.Folders) { if ($f.Name -ieq $env:OL_MOVE) { $dest = $f; break } }
        if (-not $dest) { foreach ($f in $inbox.Parent.Folders) { if ($f.Name -ieq $env:OL_MOVE) { $dest = $f; break } } }
        if (-not $dest) { Write-Output "OL_ERROR: folder '$($env:OL_MOVE)' not found"; exit }
        $a = $rule.Actions.MoveToFolder; $a.Enabled = $true; $a.Folder = $dest; $hasAct = $true
    }
    if ($env:OL_CATEGORY) { $a = $rule.Actions.AssignToCategory; $a.Enabled = $true; $a.Categories = @($env:OL_CATEGORY); $hasAct = $true }
    if ($env:OL_DELETE -eq "1") { $a = $rule.Actions.Delete; $a.Enabled = $true; $hasAct = $true }
    if (-not $hasCond) { Write-Output "OL_ERROR: a rule needs at least one condition (from or subject)"; exit }
    if (-not $hasAct) { Write-Output "OL_ERROR: a rule needs at least one action (move, category, or delete)"; exit }
    $rules.Save()
    Write-Output ("RULE_OK: created '" + $env:OL_RULE_NAME + "'")
} catch { Write-Output ("OL_ERROR: " + $_.Exception.Message) }
'''

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


def _outlook_mark_read(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids:
        return _result("No email ids provided.", True)
    read = args.get("read", True)
    _show_email(ids[0])  # show the email being marked
    out = _run_outlook_ps(_OUTLOOK_MARK_PS, {"OL_IDS": "|".join(ids), "OL_READ": "1" if read else ""}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


def _outlook_categorize(args):
    ids = args.get("ids") or []
    if isinstance(ids, str):
        ids = [ids]
    if not ids or not (args.get("category") or "").strip():
        return _result("Need email ids and a category.", True)
    _show_email(ids[0])  # show the email being categorized
    out = _run_outlook_ps(_OUTLOOK_CATEGORIZE_PS, {"OL_IDS": "|".join(ids), "OL_CATEGORY": args["category"]}, show=False)
    return _result(out, out.startswith("OL_ERROR"))


def _create_outlook_rule(args):
    if not (args.get("name") or "").strip():
        return _result("Rule needs a name.", True)
    out = _run_outlook_ps(_OUTLOOK_RULE_PS, {
        "OL_RULE_NAME": args["name"],
        "OL_FROM": "|".join(args["from_contains"]) if isinstance(args.get("from_contains"), list) else (args.get("from_contains") or ""),
        "OL_SUBJECT": "|".join(args["subject_contains"]) if isinstance(args.get("subject_contains"), list) else (args.get("subject_contains") or ""),
        "OL_MOVE": args.get("move_to_folder") or "",
        "OL_CATEGORY": args.get("category") or "",
        "OL_DELETE": "1" if args.get("delete") else "",
    })
    return _result(out, out.startswith("OL_ERROR"))


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


def _set_outlook_signature(args):
    if not (args.get("html") or "").strip():
        return _result("Need signature content (html).", True)
    out = _run_outlook_ps(_OUTLOOK_SIGNATURE_PS, {
        "OL_SIG_NAME": args.get("name") or "Maia",
        "OL_SIG_HTML": args["html"],
        "OL_SIG_DEFAULT": "1" if args.get("set_default", True) else "",
    })
    return _result(out, out.startswith("OL_ERROR"))


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


# --- Microsoft Teams messaging via deep link (reliable, no fragile UI navigation) ---
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


def _self_email():
    if not _SELF_EMAIL_CACHE.get("v"):
        out = _run_outlook_ps(_SELF_EMAIL_PS, {}).strip()
        _SELF_EMAIL_CACHE["v"] = "" if (out.startswith("ERR") or "@" not in out) else out
    return _SELF_EMAIL_CACHE["v"]


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


def _read_file(args):
    path = (args.get("path") or "").strip()
    if not path:
        return _result("Missing file path.", True)
    if not os.path.isfile(path):
        return _result(f"File not found: {path}", True)
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            data = f.read()
    except Exception as e:
        return _result(f"Could not read file: {e}", True)
    return _result(data[:TOOL_TEXT_LIMIT])


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
    if DOT_HWND:
        env["DOT_HWND"] = str(DOT_HWND)
        env["DOT_PID"] = str(DOT_PID or 0)
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
                                  capture_output=True, text=True)
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


# Approval callback for tools that DRAFT first, then ask before the final send (email send/reply/forward).
_APPROVAL_CB = None


def _ask_approval(desc):
    cb = _APPROVAL_CB
    if cb is None:
        return True
    try:
        return bool(cb(desc))
    except Exception:
        return True


# Actions gated BEFORE they run (atomic state changes / sends with no draft step).
# Email send/reply/forward are NOT here — they draft visibly first, then ask before sending.
_APPROVAL_TOOLS = {
    "send_teams_message", "create_calendar_event", "respond_to_meeting",
    "outlook_move_emails", "outlook_delete_emails", "outlook_categorize", "outlook_mark_read",
    "create_outlook_folders", "delete_outlook_folders", "create_outlook_rule", "set_outlook_signature",
    "outlook_contact", "outlook_task", "schedule_task", "close_app",
}


def _needs_approval(name, args):
    if name in _APPROVAL_TOOLS:
        return True
    if name == "file_op" and (args.get("action") or "").lower() in {
        "delete", "move", "rename", "write_text", "unzip", "copy", "mkdir",
    }:
        return True
    return False


def _count(v):
    return len(v) if isinstance(v, list) else 1


def _describe_action(name, a):
    a = a or {}
    if name == "send_email":
        s = a.get("subject")
        return f"Send email to {a.get('to', '?')}" + (f" — subject: {s}" if s else "")
    if name == "outlook_forward_email":
        return f"Forward this email to {a.get('to', '?')}"
    if name == "outlook_reply_email":
        return "Reply-all to this email" if a.get("reply_all") else "Reply to this email"
    if name == "send_teams_message":
        return f"Send a Teams message to {a.get('to', '?')}: \"{(a.get('message') or '')[:60]}\""
    if name == "create_calendar_event":
        att = a.get("attendees")
        return (f"Create a meeting '{a.get('subject', '')}' with {att}" if att
                else f"Add '{a.get('subject', '')}' to your calendar")
    if name == "respond_to_meeting":
        return f"{a.get('response', 'respond to')} the meeting invite"
    if name == "outlook_move_emails":
        return f"Move {_count(a.get('ids'))} email(s) to '{a.get('to_folder', '')}'"
    if name == "outlook_delete_emails":
        return f"Delete {_count(a.get('ids'))} email(s) (to Deleted Items)"
    if name == "outlook_categorize":
        return f"Categorize {_count(a.get('ids'))} email(s) as '{a.get('category', '')}'"
    if name == "outlook_mark_read":
        return f"Mark {_count(a.get('ids'))} email(s) as {'read' if a.get('read', True) else 'unread'}"
    if name == "create_outlook_folders":
        return "Create folder(s): " + ", ".join(a.get("folders") or [])
    if name == "delete_outlook_folders":
        return "Delete folder(s): " + ", ".join(a.get("folders") or [])
    if name == "create_outlook_rule":
        return f"Create mail rule '{a.get('name', '')}'"
    if name == "set_outlook_signature":
        return "Set/update your email signature"
    if name == "outlook_contact":
        return f"Create contact '{a.get('name', '')}'" if (a.get("action") or "create") == "create" else "Find contacts"
    if name == "outlook_task":
        return f"Create task '{a.get('subject', '')}'" if (a.get("action") or "create") == "create" else "List tasks"
    if name == "schedule_task":
        return f"Create scheduled task '{a.get('name', '')}'"
    if name == "close_app":
        return f"Close {a.get('app', '')}"
    if name == "file_op":
        return f"File: {a.get('action', '')} {a.get('path', '')}".strip()
    return name


CHAT_SYSTEM_PROMPT = (
    "You are Maia, a friendly, sharp assistant that lives inside the Axon desktop app. "
    "Answer the user's question directly, clearly, and concisely, with light formatting when it "
    "helps. In THIS mode you are a conversational assistant: you do NOT control the computer or "
    "take actions on the PC. If the user actually wants something done on their machine (open an "
    "app, send an email, automate a task), tell them to switch to Agent mode. If the user attached "
    "a screenshot, use it to answer their question."
)


def chat(question, on_status=None, image_path=None, history=None):
    """'Ask Maia' mode: a direct conversational answer from the LLM — no tools, no desktop actions.
    Supports an attached screenshot (vision) and prior turns (history) for follow-up context.
    history is a list of {"role": "user"|"assistant", "content": "..."} from earlier in the chat."""
    if not os.getenv("OPENAI_API_KEY"):
        return "OPENAI_API_KEY is not set. Add it to assistant/.env"
    if on_status:
        on_status("Maia is thinking...")
    client = OpenAI()
    content = question
    if image_path and os.path.isfile(image_path):
        try:
            import base64
            import mimetypes
            with open(image_path, "rb") as _f:
                _b64 = base64.b64encode(_f.read()).decode("ascii")
            _mime = mimetypes.guess_type(image_path)[0] or "image/png"
            content = [
                {"type": "text", "text": (question or "Describe and answer about this screenshot.").strip()},
                {"type": "image_url", "image_url": {"url": f"data:{_mime};base64,{_b64}"}},
            ]
        except Exception:
            content = question
    messages = [{"role": "system", "content": CHAT_SYSTEM_PROMPT}]
    if history:
        messages.extend(history)            # prior turns, so follow-ups keep context
    messages.append({"role": "user", "content": content})
    try:
        resp = client.chat.completions.create(model=MODEL, messages=messages, temperature=0.4)
        return (resp.choices[0].message.content or "").strip() or "(no answer)"
    except Exception as e:
        return f"Error: {e}"


def run_task(question, on_status=None, should_cancel=None, on_approval=None, image_path=None):
    """Run a natural-language task. Calls on_status(str) with progress; returns final summary.
    If on_approval is given, it is called as on_approval(description)->bool before any action that
    sends/changes data, and the action runs only if it returns True (approval mode).
    If image_path is given (a screenshot the user attached in the composer), it is shown to the
    vision model alongside the question so the user can ask about what's on their screen."""
    def status(msg):
        if on_status:
            on_status(msg)

    def cancelled():
        return bool(should_cancel and should_cancel())

    def stop_result():
        status("[stopped] Stopped by user.")
        return "Stopped."

    if not os.getenv("OPENAI_API_KEY"):
        msg = "OPENAI_API_KEY is not set. Add it to assistant/.env"
        status(msg)
        return msg

    client = OpenAI()
    mcp = MCPClient()
    globals()["_APPROVAL_CB"] = on_approval  # tools that draft-then-send use this
    # If the user attached a screenshot, send it to the vision model with the question so they can
    # ask about what's on their screen (and the agent can still act on it if asked).
    user_content = question
    if image_path and os.path.isfile(image_path):
        try:
            import base64
            import mimetypes
            with open(image_path, "rb") as _f:
                _b64 = base64.b64encode(_f.read()).decode("ascii")
            _mime = mimetypes.guess_type(image_path)[0] or "image/png"
            user_content = [
                {"type": "text", "text": (question or "").strip()
                    + "\n\n[The user attached this screenshot. If they are asking about it, answer "
                      "directly from the image; otherwise use it as context for the task.]"},
                {"type": "image_url", "image_url": {"url": f"data:{_mime};base64,{_b64}"}},
            ]
        except Exception:
            user_content = question
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]
    status("Axon intelligence is thinking...")
    try:
        for _ in range(MAX_STEPS):
            if cancelled():
                return stop_result()
            resp = client.chat.completions.create(
                model=MODEL, messages=messages, tools=TOOLS, tool_choice="auto", temperature=0
            )
            if cancelled():
                return stop_result()
            msg = resp.choices[0].message

            assistant_msg = {"role": "assistant", "content": msg.content or ""}
            if msg.tool_calls:
                assistant_msg["tool_calls"] = [
                    {"id": tc.id, "type": "function",
                     "function": {"name": tc.function.name, "arguments": tc.function.arguments}}
                    for tc in msg.tool_calls
                ]
            messages.append(assistant_msg)

            if not msg.tool_calls:
                final = msg.content or "Done."
                status("[done] " + final)
                return final

            pending_images = []
            for tc in msg.tool_calls:
                if cancelled():
                    return stop_result()
                name = tc.function.name
                try:
                    args = json.loads(tc.function.arguments or "{}")
                except Exception:
                    args = {}
                short = ", ".join(f"{k}={v}" for k, v in args.items() if k != "app")
                status(f"-> {name}({short})" if short else f"-> {name}")

                # Approval gate: pause for the user's OK before sending/changing data.
                declined = False
                if on_approval and _needs_approval(name, args):
                    if not on_approval(_describe_action(name, args)):
                        declined = True
                if cancelled():
                    return stop_result()

                if declined:
                    desc = _describe_action(name, args)
                    status(f"[skipped] {desc}")
                    result = _result(
                        f"The user DECLINED this action, so it was NOT performed: {desc}. "
                        f"Do not retry it; continue with anything else they asked, or stop and report.", False)
                elif name == "update_todos":
                    todos = args.get("todos") or []
                    lines = []
                    for t in todos:
                        box = "[x]" if t.get("done") else "[ ]"
                        lines.append(f"  {box} {t.get('task', '')}")
                    ndone = sum(1 for t in todos if t.get("done"))
                    status("Checklist (%d/%d done):\n%s" % (ndone, len(todos), "\n".join(lines)))
                    result = _result("Checklist updated: %d of %d steps done." % (ndone, len(todos)))
                elif name == "browse":
                    result = _browse(args)
                elif name == "research_website":
                    result = _research_website(args)
                elif name == "read_file":
                    result = _read_file(args)
                elif name == "send_email":
                    result = _send_email(args)
                elif name == "create_outlook_folders":
                    result = _create_outlook_folders(args)
                elif name == "delete_outlook_folders":
                    result = _delete_outlook_folders(args)
                elif name == "list_outlook_folders":
                    result = _list_outlook_folders(args)
                elif name == "outlook_list_emails":
                    result = _outlook_list_emails(args)
                elif name == "outlook_move_emails":
                    result = _outlook_move_emails(args)
                elif name == "outlook_delete_emails":
                    result = _outlook_delete_emails(args)
                elif name == "outlook_forward_email":
                    result = _outlook_forward_email(args)
                elif name == "save_email":
                    result = _outlook_save_email(args)
                elif name == "get_open_email":
                    result = _get_open_email(args)
                elif name == "outlook_reply_email":
                    result = _outlook_reply_email(args)
                elif name == "outlook_mark_read":
                    result = _outlook_mark_read(args)
                elif name == "outlook_categorize":
                    result = _outlook_categorize(args)
                elif name == "create_outlook_rule":
                    result = _create_outlook_rule(args)
                elif name == "create_calendar_event":
                    result = _create_calendar_event(args)
                elif name == "set_outlook_signature":
                    result = _set_outlook_signature(args)
                elif name == "list_installed_apps":
                    result = _list_installed_apps(args)
                elif name == "send_teams_message":
                    result = _send_teams_message(args)
                elif name == "open_app":
                    result = _open_app(args)
                elif name == "close_app":
                    result = _close_app(mcp, args)
                elif name in ot.DISPATCH:
                    result = ot.DISPATCH[name](args)
                else:
                    result = mcp.call(name, args)

                if cancelled():
                    return stop_result()
                text, image = _extract(result)
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": text[:TOOL_TEXT_LIMIT]})
                if image:
                    pending_images.append((name, image))

            for name, image in pending_images:
                if cancelled():
                    return stop_result()
                messages.append({
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"Screenshot after {name}:"},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image}"}},
                    ],
                })

        status("Reached step limit without finishing.")
        return "Reached step limit without finishing."
    finally:
        globals()["_APPROVAL_CB"] = None
        mcp.close()


GUIDE_SYSTEM_PROMPT = """You are Axon intelligence, a friendly on-screen guide. The user is stuck and has
shared a screenshot of their whole screen plus a question. COACH them — do NOT perform any
actions yourself. Give SHORT, numbered, plain-English steps that reference what is actually
visible (e.g. "Top-right, click the blue 'Share' button", "In the left list, choose 'Inbox'").
Keep it to the next few concrete steps, not a huge wall.

Also pick the SINGLE most important spot for the user to look at / click next, and give its
location as fractions of the screenshot: x and y each from 0.0 to 1.0 (0,0 = top-left,
1,1 = bottom-right), plus a 1-3 word label.

Respond ONLY as JSON:
{"steps": ["1. ...", "2. ..."], "pointer": {"x": 0.0-1.0, "y": 0.0-1.0, "label": "..."}}
If there is no specific spot to point at, set "pointer" to null."""


GUIDE_LIVE_SYSTEM_PROMPT = """You are Axon intelligence, a friendly LIVE on-screen guide. You walk the user
through a task ONE STEP AT A TIME by looking at their screen. You NEVER act yourself — you tell them, in
plain words, the single next thing to do.

You are given the user's GOAL, the steps already guided, and a screenshot of their CURRENT screen.

CRITICAL: refer ONLY to controls that are ACTUALLY VISIBLE in the screenshot right now. Look carefully at
what is really on screen and describe THAT. NEVER invent or assume a control that isn't there — do not say
"three-dots menu", "hamburger menu", etc. unless you can actually SEE that exact icon. If you are not sure
where an action lives, point to the closest visible button that plausibly opens it and describe its real
appearance — do not make up a generic menu.

Describe the ONE next action so clearly the user can find it instantly with WORDS ALONE (there is no
pointer): copy the element's REAL visible text/label, say what its icon looks like (shape, colour), and say
WHERE it is — which corner/toolbar, and which real, visible items it sits between or next to. Everything
you say must come from THIS screenshot, never from a template.

If the exact action has no direct button, it usually opens from a nearby VISIBLE control — tell them to
click that real, visible control first, then guide the click inside the menu it opens on the NEXT step.
Refer to the application's own controls, not the browser tabs/address bar, unless the task is about the
browser. Do NOT copy any wording from these instructions into your answer.

Keep it to ONE short, specific step. Respond ONLY as JSON:
{
  "instruction": "one clear sentence about the REAL element on screen (its visible label/look + where it is)",
  "done": true|false
}
Set "done" to true only when the goal is fully accomplished on the current screen (with a brief confirming sentence)."""


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
    if not DOT_HWND:
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

        mon = u.MonitorFromWindow(ctypes.c_void_p(int(DOT_HWND)), 2)  # NEAREST
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


_VSCREEN_CACHE = {}


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


_ELEM_RE = re.compile(
    r"^\s*(\d+)\s+(.*?)\s+(?:Secondary Actions:.*?)?Frame:\s*\{x:\s*(-?\d+),\s*y:\s*(-?\d+),"
    r"\s*width:\s*(\d+),\s*height:\s*(\d+)\}", re.I)


def _parse_elements(text):
    """Parse get_app_state's tree into {index: (label, x, y, w, h)} (absolute screen px)."""
    elems = {}
    for ln in text.splitlines():
        m = _ELEM_RE.match(ln)
        if not m:
            continue
        idx = int(m.group(1))
        label = re.sub(r"\s+", " ", m.group(2)).strip()[:60]
        elems[idx] = (label, int(m.group(3)), int(m.group(4)), int(m.group(5)), int(m.group(6)))
    return elems


def _guide_decide(question, history, b64):
    client = OpenAI()
    done_txt = "; ".join(history) if history else "(none yet)"
    # The DECIDE pass must pick the RIGHT element (needs strong reasoning), so use the main model.
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": GUIDE_LIVE_SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "text", "text": f"GOAL: {question or 'Help me with what is on my screen.'}\n"
                                         f"Steps already guided: {done_txt}"},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    )
    return json.loads(resp.choices[0].message.content or "{}")


GUIDE_REFINE_PROMPT = """You are locating ONE UI element PRECISELY in a ZOOMED-IN screenshot that has a red
0-1000 coordinate grid (faint lines every 0.5, bold labelled lines every 10; 0,0 top-left, 1000,1000
bottom-right). You are told which element to find. Read its EXACT edges off the grid and return a box that
matches the element's real bounds — its true width AND height (a small icon = a small, short box). Do not
pad it. Respond ONLY as JSON: {"box": [x1, y1, x2, y2]}  (or {"box": null} if it is not visible)."""


def _guide_refine(b64, target):
    client = OpenAI()
    resp = client.chat.completions.create(
        model=GUIDE_MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": GUIDE_REFINE_PROMPT},
            {"role": "user", "content": [
                {"type": "text", "text": f"Find and tightly box this element: {target}"},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    )
    return json.loads(resp.choices[0].message.content or "{}")


def _refine_marker(img, data):
    """Coarse box -> zoom into that region with a fresh grid -> precise box -> monitor-fraction marker."""
    box1 = data.get("box")
    if not GUIDE_REFINE or not (isinstance(box1, (list, tuple)) and len(box1) == 4):
        return _marker_from(data)
    try:
        x1, x2 = sorted((float(box1[0]), float(box1[2])))
        y1, y2 = sorted((float(box1[1]), float(box1[3])))
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        hw = max((x2 - x1), 130)  # zoom region half-size in 0-1000 units (>=13% of monitor)
        hh = max((y2 - y1), 130)
        rl, rr = max(0.0, cx - hw), min(1000.0, cx + hw)
        rt, rb = max(0.0, cy - hh), min(1000.0, cy + hh)
        W, H = img.size
        crop = img.crop((int(rl / 1000 * W), int(rt / 1000 * H), int(rr / 1000 * W), int(rb / 1000 * H)))
        # Standard grid on the MAGNIFIED crop already resolves to ~0.5% of the full monitor per line
        # (the crop is ~25% of the screen), so it's precise AND readable.
        ref = _guide_refine(_grid_b64(crop), str(data.get("label") or "the target element"))
        b2 = ref.get("box")
        if isinstance(b2, (list, tuple)) and len(b2) == 4:
            bx1, bx2 = sorted((float(b2[0]), float(b2[2])))
            by1, by2 = sorted((float(b2[1]), float(b2[3])))
            mx1 = rl + bx1 / 1000 * (rr - rl)
            mx2 = rl + bx2 / 1000 * (rr - rl)
            my1 = rt + by1 / 1000 * (rb - rt)
            my2 = rt + by2 / 1000 * (rb - rt)
            pad = 3  # tight margin; let small buttons get small brackets
            return {"type": "box",
                    "fx": (mx1 - pad) / 1000.0, "fy": (my1 - pad) / 1000.0,
                    "fw": max(mx2 - mx1 + 2 * pad, 12) / 1000.0,
                    "fh": max(my2 - my1 + 2 * pad, 9) / 1000.0,  # allow short height for small buttons
                    "label": str(data.get("label") or "")}
    except Exception:
        pass
    return _marker_from(data)


def _marker_from(data):
    """Build a 'Click here' callout marker from the model's approximate grid point [x, y]."""
    pt = data.get("point")
    if isinstance(pt, (list, tuple)) and len(pt) == 2:
        try:
            return {"type": "clickhere",
                    "fx": min(max(float(pt[0]) / 1000.0, 0.0), 1.0),
                    "fy": min(max(float(pt[1]) / 1000.0, 0.0), 1.0),
                    "label": "Click here"}
        except Exception:
            pass
    return None


def guide_live(question, on_step=None, should_cancel=None, max_steps=25):
    """Live coaching with CLEAR TEXT instructions (no on-screen marker): describe the next thing to
    click using landmarks, wait for the user to act, then guide the next step.

    Calls on_step({"instruction", "marker": None, "done": bool}) for each step.
    """
    def cancelled():
        return bool(should_cancel and should_cancel())

    def emit(instruction, marker, done):
        if on_step:
            on_step({"instruction": instruction, "marker": marker, "done": done})

    if not os.getenv("OPENAI_API_KEY"):
        emit("OPENAI_API_KEY is not set. Add it to assistant/.env", None, True)
        return

    history = []
    for _ in range(max_steps):
        if cancelled():
            return
        try:
            img = _grab_dot_monitor_img()
        except Exception as e:
            emit(f"Could not capture the screen: {e}", None, True)
            return
        try:
            data = _guide_decide(question, history, _img_b64(img))
        except Exception as e:
            emit(f"Guidance failed: {e}", None, True)
            return
        if cancelled():
            return
        instruction = str(data.get("instruction") or "").strip()
        done = bool(data.get("done"))
        # Text-only guidance — no on-screen marker (the clear description is the guide).
        emit(instruction or "(thinking...)", None, done)
        if done:
            return
        if instruction:
            history.append(instruction)
        # Wait for the user to act, then for the screen to SETTLE, before re-guiding. This avoids
        # re-evaluating (and visibly moving the mark) on transient changes like hover tooltips,
        # chart animations, or a blinking cursor on dynamic pages.
        try:
            base = _screen_signature()
        except Exception:
            base = None
        waited = 0.0
        changed = False
        THRESH = 10  # ignore small/transient changes; only react to a real navigation/click
        while waited < 90 and not cancelled():
            time.sleep(0.4)
            waited += 0.4
            try:
                cur = _screen_signature()
            except Exception:
                break
            if not base:
                break
            if not changed:
                if _screens_differ(base, cur, thresh=THRESH):
                    changed = True  # the user did something; now wait for it to settle
                    base = cur
            else:
                if not _screens_differ(base, cur, thresh=THRESH):
                    break  # settled into the new state -> guide the next step
                base = cur
    if not cancelled():
        emit("That's as far as I can guide step-by-step — tell me if you're still stuck.", None, True)


def guide(question, on_status=None, should_cancel=None):
    """Coaching mode: screenshot the screen and return {steps_text, pointer:{x,y,label}|None}."""
    def status(m):
        if on_status:
            on_status(m)

    def cancelled():
        return bool(should_cancel and should_cancel())

    if not os.getenv("OPENAI_API_KEY"):
        return {"steps_text": "OPENAI_API_KEY is not set. Add it to assistant/.env", "pointer": None}
    status("Axon intelligence is looking at your screen...")
    if cancelled():
        return {"steps_text": "Stopped.", "pointer": None}
    try:
        b64 = _grab_screen_b64()
    except Exception as e:
        return {"steps_text": f"Could not capture the screen: {e}", "pointer": None}
    if cancelled():
        return {"steps_text": "Stopped.", "pointer": None}
    try:
        client = OpenAI()
        resp = client.chat.completions.create(
            model=MODEL,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": GUIDE_SYSTEM_PROMPT},
                {"role": "user", "content": [
                    {"type": "text", "text": question or "Help me with what is on my screen."},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ]},
            ],
        )
        if cancelled():
            return {"steps_text": "Stopped.", "pointer": None}
        data = json.loads(resp.choices[0].message.content or "{}")
    except Exception as e:
        return {"steps_text": f"Guidance failed: {e}", "pointer": None}
    steps = data.get("steps") or []
    steps_text = "\n".join(str(s) for s in steps) if isinstance(steps, list) else str(steps)
    pointer = data.get("pointer") if isinstance(data.get("pointer"), dict) else None
    status(steps_text or "(no steps returned)")
    return {"steps_text": steps_text, "pointer": pointer}


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # Windows console is cp1252 by default
    except Exception:
        pass
    q = " ".join(sys.argv[1:]) or input("Task: ")
    run_task(q, on_status=print)
