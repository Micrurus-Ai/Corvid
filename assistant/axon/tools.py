"""Tool registry: the JSON schemas advertised to the model (TOOLS) and the name->callable
DISPATCH map used by the agent loop. axon.office contributes its own schemas + dispatch."""
from axon import office as ot

from axon.browse import _browse, _research_website, setup_browser
from axon.screen import _take_screenshot
from axon.files import _read_file
from axon.messaging import _send_teams_message
from axon.mcp import _list_installed_apps, _open_app
from axon.outlook import (
    _send_email, _create_outlook_folders, _delete_outlook_folders, _list_outlook_folders,
    _outlook_list_emails, _outlook_move_emails, _outlook_delete_emails, _outlook_forward_email,
    _outlook_save_email, _get_open_email, _outlook_reply_email, _outlook_mark_read,
    _outlook_categorize, _create_outlook_rule, _create_calendar_event, _set_outlook_signature,
    learn_my_tone,
)
from axon.memory import remember, recall, forget
from axon.knowledge import set_project, ask_documents, index_documents
from axon.briefing import daily_briefing, inbox_triage
from axon.meeting import meeting_notes
from axon.triggers import add_email_trigger, list_email_triggers, remove_email_trigger, run_email_triggers
from axon.vision import ocr_image

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
            "name": "take_screenshot",
            "description": "Capture a screenshot of the screen and save it to a PNG file, returning the file PATH. Use this whenever the user asks to take a screenshot and then DO something with it (email it / send it as an attachment, save it, etc.). After calling this, pass the returned path to send_email's `attachments` to send it. By default it captures the monitor where the dot is; set scope='all' to capture every monitor.",
            "parameters": {
                "type": "object",
                "properties": {
                    "scope": {"type": "string", "description": "'dot' (default) = the screen the dot is on; 'all' = all monitors combined."},
                    "filename": {"type": "string", "description": "Optional PNG file name (default screenshot_<timestamp>.png), saved to the Downloads folder."},
                },
                "required": [],
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


TOOLS += ot.TOOLS


def _fn(name, description, properties, required):
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties, "required": required}}}


TOOLS += [
    _fn("remember",
        "Save a durable fact about the user to long-term memory (brand, key people, preferences, "
        "ongoing projects). Use whenever the user shares something worth remembering across sessions.",
        {"text": {"type": "string"},
         "category": {"type": "string", "description": "e.g. brand, people, preference, project"}},
        ["text"]),
    _fn("recall", "List what you remember about the user (optionally filtered by a query word).",
        {"query": {"type": "string"}}, []),
    _fn("forget", "Forget remembered facts whose text contains the query.",
        {"query": {"type": "string"}}, ["query"]),
    _fn("set_project", "Set the active project folder so 'ask my documents' searches it by default.",
        {"folder": {"type": "string"}}, ["folder"]),
    _fn("ask_documents",
        "Answer a question using the user's documents. Searches the project folder (or a given "
        "folder) with a semantic index (handles large folders, no file cap), reads "
        ".txt/.md/.csv/.docx/.pdf/.pptx/.xlsx/.html/.json, and answers with [filename] citations.",
        {"question": {"type": "string"},
         "folder": {"type": "string", "description": "Folder to search (optional if a project is set)"}},
        ["question"]),
    _fn("index_documents",
        "Build/refresh the search index for a folder so Q&A over a large folder (hundreds/thousands "
        "of files) is fast and complete. Run once up front; ask_documents keeps it updated after.",
        {"folder": {"type": "string", "description": "Folder to index (optional if a project is set)"}},
        []),
    _fn("daily_briefing",
        "Summarize the user's day from Outlook: today's calendar events, unread email (count + top "
        "senders), and open tasks. Use for 'what's my day / brief me / morning summary'.",
        {}, []),
    _fn("inbox_triage",
        "Triage the unread inbox: order emails by importance, each with a priority, one-line summary, "
        "and suggested action. Use for 'what needs my attention / triage my inbox'.",
        {}, []),
    _fn("setup_browser",
        "Open Axon's own browser so the user can sign in to the accounts they want Axon to use "
        "(Google, intranet, analytics). One-time — the logins are saved and reused for all future "
        "browsing. Use for 'sign in to my browser / set up browsing / add my Chrome account'.",
        {"url": {"type": "string", "description": "Optional first page to open (default Google sign-in)"}},
        []),
    _fn("learn_my_tone",
        "Analyse the user's recent Sent emails to learn their writing style (greeting, sign-off, "
        "formality, length, per-language NL/FR/EN notes) and save it, so future replies and drafts "
        "sound like them. Use for 'learn my tone / write like me / match my style'.",
        {}, []),
    _fn("meeting_notes",
        "Turn a meeting into notes: summary, decisions, action items, and a follow-up email draft. "
        "Accepts a recording file (audio=path to .mp3/.m4a/.mp4/.wav), a transcript file, or "
        "transcript text. Use for 'summarize this meeting / take minutes / meeting notes'.",
        {"audio": {"type": "string", "description": "Path to an audio/video recording"},
         "transcript": {"type": "string", "description": "Transcript text, or path to a .txt transcript"},
         "title": {"type": "string", "description": "Meeting title (optional)"}},
        []),
    _fn("add_email_trigger",
        "Add an automation rule: when an unread inbox email matches (from and/or subject_contains), "
        "do an action. action=move needs folder; action=categorize needs category; mark_read and "
        "draft_reply need neither (draft_reply writes a reply in your tone and leaves it open to approve "
        "— never auto-sends).",
        {"from": {"type": "string", "description": "sender name/email contains"},
         "subject_contains": {"type": "string"},
         "action": {"type": "string", "enum": ["move", "categorize", "mark_read", "draft_reply"]},
         "folder": {"type": "string"}, "category": {"type": "string"}},
        ["action"]),
    _fn("list_email_triggers", "List the configured email triggers (with their indexes).", {}, []),
    _fn("remove_email_trigger", "Remove an email trigger by index (from list_email_triggers).",
        {"index": {"type": "integer"}}, ["index"]),
    _fn("run_email_triggers",
        "Apply the configured email triggers to current unread inbox emails now. Schedule this "
        "(schedule_task) to make it run periodically for hands-off filing.",
        {}, []),
    _fn("ocr_image",
        "Extract the text (and tables, as markdown) from an image file — a screenshot, a photo of a "
        "document, etc. Pair with take_screenshot to read what's on screen.",
        {"path": {"type": "string"}, "instruction": {"type": "string", "description": "Optional: what to extract"}},
        ["path"]),
]

# name -> callable(args). update_todos and close_app are handled inline by run_task (they need the
# UI status callback / the live MCP client). Everything else routes through here.
DISPATCH = {
    "browse": _browse,
    "research_website": _research_website,
    "take_screenshot": _take_screenshot,
    "read_file": _read_file,
    "send_email": _send_email,
    "create_outlook_folders": _create_outlook_folders,
    "delete_outlook_folders": _delete_outlook_folders,
    "list_outlook_folders": _list_outlook_folders,
    "outlook_list_emails": _outlook_list_emails,
    "outlook_move_emails": _outlook_move_emails,
    "outlook_delete_emails": _outlook_delete_emails,
    "outlook_forward_email": _outlook_forward_email,
    "save_email": _outlook_save_email,
    "get_open_email": _get_open_email,
    "outlook_reply_email": _outlook_reply_email,
    "outlook_mark_read": _outlook_mark_read,
    "outlook_categorize": _outlook_categorize,
    "create_outlook_rule": _create_outlook_rule,
    "create_calendar_event": _create_calendar_event,
    "set_outlook_signature": _set_outlook_signature,
    "list_installed_apps": _list_installed_apps,
    "send_teams_message": _send_teams_message,
    "open_app": _open_app,
    "setup_browser": setup_browser,
}
DISPATCH.update(ot.DISPATCH)
DISPATCH.update({
    "remember": remember, "recall": recall, "forget": forget,
    "set_project": set_project, "ask_documents": ask_documents, "index_documents": index_documents,
    "daily_briefing": daily_briefing, "inbox_triage": inbox_triage,
    "learn_my_tone": learn_my_tone, "meeting_notes": meeting_notes,
    "add_email_trigger": add_email_trigger, "list_email_triggers": list_email_triggers,
    "remove_email_trigger": remove_email_trigger, "run_email_triggers": run_email_triggers,
    "ocr_image": ocr_image,
})
