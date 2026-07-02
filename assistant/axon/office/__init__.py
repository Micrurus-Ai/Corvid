"""Office/Windows COM automation toolkit: tool schemas (TOOLS) + dispatch (DISPATCH)."""
from axon.office.excel import excel
from axon.office.word import word
from axon.office.powerpoint import powerpoint
from axon.office.pdf import convert_to_pdf
from axon.office.misc import (
    file_op, outlook_contact, outlook_task, save_email_attachments,
    respond_to_meeting, schedule_task, system_query, speak,
)
# Structured read/edit engine (python-pptx / python-docx / openpyxl) — the Cursor-style
# read -> edit -> preview -> refine workflow on existing documents.
from axon.office.decks import ppt_read, ppt_edit
from axon.office.docs import word_read, word_edit
from axon.office.sheets import excel_read, excel_edit, excel_analyze
from axon.office.brand import set_brand, brand_deck, brand_doc
from axon.office.report import make_report
from axon.office.imaging import generate_image
from axon.office.pdf_tools import pdf_read, pdf_extract_tables, pdf_fill_form, pdf_form_fields


def _fn(name, description, properties, required):
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties, "required": required},
    }}


TOOLS = [
    _fn("excel",
        "Create or read an Excel workbook (no UI). action='write' builds a .xlsx from headers+rows with optional title, chart (bar/line/pie/column) and PDF export; action='read' returns a workbook's cell data. Great for professional data reports.",
        {"action": {"type": "string", "enum": ["write", "read"]},
         "path": {"type": "string", "description": "File path or name (defaults to Downloads, .xlsx)"},
         "title": {"type": "string"},
         "sheet_name": {"type": "string"},
         "headers": {"type": "array", "items": {"type": "string"}},
         "rows": {"type": "array", "items": {"type": "array", "items": {}}, "description": "Rows of cell values"},
         "chart": {"type": "string", "enum": ["bar", "line", "pie", "column"]},
         "pdf": {"type": "boolean", "description": "Also export a PDF"},
         "max_rows": {"type": "integer", "description": "read: max rows to return"}},
        ["action"]),
    _fn("word",
        "Create or read a Word document (no UI). action='write' builds a .docx from a title + paragraphs (each {text, style: normal/heading1/heading2/bullet/number}) with optional PDF export; action='read' extracts text from a .docx.",
        {"action": {"type": "string", "enum": ["write", "read"]},
         "path": {"type": "string", "description": "File path or name (defaults to Downloads, .docx)"},
         "title": {"type": "string"},
         "paragraphs": {"type": "array", "items": {"type": "object", "properties": {
             "text": {"type": "string"}, "style": {"type": "string"}}}},
         "body": {"type": "string", "description": "Alternative to paragraphs: plain text (one paragraph per line)"},
         "pdf": {"type": "boolean"}},
        ["action"]),
    _fn("powerpoint",
        "Create a PowerPoint deck (no UI) from slides (each {title, bullets:[...]}). Optional PDF export.",
        {"path": {"type": "string", "description": "File path or name (defaults to Downloads, .pptx)"},
         "slides": {"type": "array", "items": {"type": "object", "properties": {
             "title": {"type": "string"}, "bullets": {"type": "array", "items": {"type": "string"}}}}},
         "pdf": {"type": "boolean"}},
        ["slides"]),
    _fn("convert_to_pdf",
        "Convert an existing Office file (Word/Excel/PowerPoint/CSV/txt) to PDF in the same folder.",
        {"path": {"type": "string", "description": "Path to the file to convert"}},
        ["path"]),
    _fn("file_op",
        "Manage files and folders (no UI). actions: list, search (pattern + recursive), move, copy, rename, delete, mkdir, zip, unzip, exists, read_text, write_text, open.",
        {"action": {"type": "string", "enum": ["list", "search", "move", "copy", "rename", "delete", "mkdir", "zip", "unzip", "exists", "read_text", "write_text", "open"]},
         "path": {"type": "string", "description": "Target path (source for move/copy/zip)"},
         "dest": {"type": "string", "description": "Destination path / new name / zip output"},
         "pattern": {"type": "string", "description": "search: filename pattern e.g. *.pdf"},
         "recursive": {"type": "boolean", "description": "search: include subfolders"},
         "content": {"type": "string", "description": "write_text: the text to write"}},
        ["action"]),
    _fn("outlook_contact",
        "Create or find an Outlook contact. action='create' (name + optional email/phone/company/title); action='find' (query matches name/email/company).",
        {"action": {"type": "string", "enum": ["create", "find"]},
         "name": {"type": "string"}, "email": {"type": "string"}, "phone": {"type": "string"},
         "company": {"type": "string"}, "title": {"type": "string"},
         "query": {"type": "string", "description": "find: search text"}},
        ["action"]),
    _fn("outlook_task",
        "Create or list Outlook tasks. action='create' (subject + optional body/due/reminder, dates like '2026-07-01 09:00'); action='list' (open tasks).",
        {"action": {"type": "string", "enum": ["create", "list"]},
         "subject": {"type": "string"}, "body": {"type": "string"},
         "due": {"type": "string"}, "reminder": {"type": "string"}},
        ["action"]),
    _fn("save_email_attachments",
        "Save all attachments from an Outlook email (by its id from outlook_list_emails) to a folder (defaults to Documents).",
        {"id": {"type": "string"}, "folder": {"type": "string"}},
        ["id"]),
    _fn("respond_to_meeting",
        "Accept, tentatively accept, or decline a meeting invitation (by the email/meeting id).",
        {"id": {"type": "string"}, "response": {"type": "string", "enum": ["accept", "tentative", "decline"]},
         "send_response": {"type": "boolean", "description": "Send the response to the organizer (default true)"}},
        ["id", "response"]),
    _fn("schedule_task",
        "Create a Windows scheduled task to run a command on a schedule. frequency: once/daily/weekly; time like '08:00' or '2026-07-01 09:00'.",
        {"name": {"type": "string"}, "command": {"type": "string", "description": "Program/exe to run"},
         "arguments": {"type": "string"}, "time": {"type": "string"},
         "frequency": {"type": "string", "enum": ["once", "daily", "weekly"]},
         "day_of_week": {"type": "string", "description": "weekly: e.g. Monday"}},
        ["name", "command", "time"]),
    _fn("system_query",
        "Query the computer (no UI). what: 'processes' (top by memory), 'services', 'disks', or 'system' (OS/RAM/uptime). Optional filter.",
        {"what": {"type": "string", "enum": ["processes", "services", "disks", "system"]},
         "filter": {"type": "string"}},
        []),
    _fn("speak",
        "Speak text aloud through the computer's speakers (text-to-speech).",
        {"text": {"type": "string"}, "rate": {"type": "integer", "description": "-10 (slow) to 10 (fast)"}},
        ["text"]),
    # ---- Structured document editing (read -> edit -> preview, like editing code) ----
    _fn("ppt_read",
        "Outline a PowerPoint deck: each slide's index (1-based), title, bullets, shapes, and notes. "
        "Call this FIRST when editing an existing .pptx so you know what to change.",
        {"path": {"type": "string", "description": "Path to the .pptx"}}, ["path"]),
    _fn("ppt_edit",
        "Edit (or create) a PowerPoint deck with a list of ops, then open it as a live preview. Slides "
        "are 1-based. ops items: "
        "{op:'add_slide', layout:title|title_content|section|two_content|title_only|blank, title, bullets:[], notes, at}; "
        "{op:'set_title', slide, text}; {op:'set_bullets', slide, bullets:[]}; {op:'set_notes', slide, text}; "
        "{op:'add_image', slide, path, left, top, width, height (inches)}; {op:'add_table', slide, rows:[[...]], left, top, width, height}; "
        "{op:'add_chart', slide, chart_type:column|bar|line|pie|area, categories:[], series:{name:[vals]}, title, left, top, width, height}; "
        "{op:'set_bg', slide(or 'all'), color:'RRGGBB'}; {op:'move_slide', slide, to}; {op:'delete_slide', slide}.",
        {"path": {"type": "string", "description": "Path/name (defaults to Downloads, .pptx). Created if missing."},
         "ops": {"type": "array", "items": {"type": "object"}, "description": "List of edit operations (see above)."}},
        ["ops"]),
    _fn("word_read",
        "Outline a Word document: each non-empty paragraph as [index] (style) text, plus tables. "
        "Call this FIRST when editing an existing .docx (paragraph indexes come from here).",
        {"path": {"type": "string", "description": "Path to the .docx"}}, ["path"]),
    _fn("word_edit",
        "Edit (or create) a Word document with a list of ops, then open it as a live preview. Paragraph "
        "indexes come from word_read. ops items: "
        "{op:'add_heading', text, level (0=title,1,2,3)}; {op:'add_paragraph', text, style, bold, italic, size, color}; "
        "{op:'set_paragraph', index, text, bold, italic, size}; {op:'delete_paragraph', index}; "
        "{op:'replace_text', find, replace}; {op:'add_image', path, width (inches)}; "
        "{op:'add_table', rows:[[...]], style}; {op:'add_page_break'}.",
        {"path": {"type": "string", "description": "Path/name (defaults to Downloads, .docx). Created if missing."},
         "ops": {"type": "array", "items": {"type": "object"}, "description": "List of edit operations (see above)."}},
        ["ops"]),
    _fn("excel_read",
        "Outline a workbook: each sheet's size and a preview of the first rows. Call this FIRST when "
        "editing an existing .xlsx.",
        {"path": {"type": "string", "description": "Path to the .xlsx"}}, ["path"]),
    _fn("excel_edit",
        "Edit (or create) an Excel workbook with a list of ops, then open it as a live preview. ops items: "
        "{op:'set_cell', sheet, cell:'A1', value (a formula like '=SUM(B2:B4)' works)}; "
        "{op:'set_range', sheet, start:'A1', rows:[[...]]}; {op:'add_sheet', name}; {op:'delete_sheet', name}; "
        "{op:'format', sheet, range:'A1:B1', bold, italic, fill:'RRGGBB', font_color, number_format, align}; "
        "{op:'column_width', column:'A', width}; "
        "{op:'add_chart', sheet, chart_type:column|bar|line|pie, data:'B1:B4', categories:'A2:A4', title, anchor:'H2'}. "
        "Omit sheet to use the active sheet.",
        {"path": {"type": "string", "description": "Path/name (defaults to Downloads, .xlsx). Created if missing."},
         "ops": {"type": "array", "items": {"type": "object"}, "description": "List of edit operations (see above)."}},
        ["ops"]),
    # ---- Office power-ups ----
    _fn("excel_analyze",
        "Analyze a spreadsheet: compute count/sum/avg/min/max for each numeric column, add a Summary "
        "sheet with a chart, and report the stats. Use for 'analyze this spreadsheet'.",
        {"path": {"type": "string"}, "sheet": {"type": "string", "description": "Sheet name (default: active)"}},
        ["path"]),
    _fn("set_brand",
        "Save the user's brand so decks/docs can be styled consistently. Colors are RRGGBB.",
        {"primary": {"type": "string"}, "accent": {"type": "string"}, "dark": {"type": "string"},
         "light": {"type": "string"}, "text": {"type": "string"},
         "font": {"type": "string"}, "logo": {"type": "string", "description": "Path to a logo image"}},
        []),
    _fn("brand_deck",
        "Apply the saved brand to a whole PowerPoint deck (backgrounds, title colors/font, body text, "
        "logo top-right). Use for 'brand this deck' or 'make it prettier'.",
        {"path": {"type": "string", "description": "Path to the .pptx"}}, ["path"]),
    _fn("brand_doc",
        "Apply the saved brand to a Word document (brand-colored headings in the brand font, consistent body).",
        {"path": {"type": "string", "description": "Path to the .docx"}}, ["path"]),
    _fn("make_report",
        "Build a branded PDF report/proposal from structured content using CODE (accurate, real data "
        "charts, brand colors, optional logo — no image model, no Office needed). Use for 'make a PDF "
        "report/proposal/one-pager'. The logo is optional: only used if the brand has one or the user "
        "gives a logo path — ask the user if they want a logo when it matters.",
        {"title": {"type": "string"},
         "subtitle": {"type": "string"},
         "sections": {"type": "array", "description":
             "[{heading?, text? (string or [strings]), bullets?[], table?{headers[],rows[[]]}, "
             "chart?{type:'bar'|'line'|'pie', title?, categories[], series?{name:[vals]}, values?[]}}]",
             "items": {"type": "object"}},
         "logo": {"type": "string", "description": "Optional logo image path (else the saved brand logo)"},
         "filename": {"type": "string", "description": "Optional output file name"}},
        ["title"]),
    _fn("generate_image",
        "Generate an image from a text prompt and save it as a PNG (then use ppt_edit/word_edit add_image "
        "to place it). Good for slide graphics, illustrations, simple logos.",
        {"prompt": {"type": "string"}, "path": {"type": "string", "description": "Output PNG path (optional)"},
         "size": {"type": "string", "enum": ["1024x1024", "1536x1024", "1024x1536"]}},
        ["prompt"]),
    _fn("pdf_read",
        "Extract text from a PDF (optionally a page range like '1-5'). Read it, then YOU write the "
        "summary/answer. Use for 'summarize this PDF' or questions about a PDF.",
        {"path": {"type": "string"}, "pages": {"type": "string", "description": "e.g. '3' or '1-5'"}},
        ["path"]),
    _fn("pdf_extract_tables",
        "Extract all tables from a PDF into an Excel workbook (one sheet per table) and open it.",
        {"path": {"type": "string"}, "out": {"type": "string", "description": "Output .xlsx (optional)"}},
        ["path"]),
    _fn("pdf_form_fields",
        "List the fillable form field names in an AcroForm PDF (so you know what to fill).",
        {"path": {"type": "string"}}, ["path"]),
    _fn("pdf_fill_form",
        "Fill an AcroForm PDF's fields and save a copy. Call pdf_form_fields first to get field names.",
        {"path": {"type": "string"}, "fields": {"type": "object", "description": "{field_name: value}"},
         "out": {"type": "string"}},
        ["path", "fields"]),
]


DISPATCH = {
    "excel": excel,
    "word": word,
    "powerpoint": powerpoint,
    "convert_to_pdf": convert_to_pdf,
    "file_op": file_op,
    "outlook_contact": outlook_contact,
    "outlook_task": outlook_task,
    "save_email_attachments": save_email_attachments,
    "respond_to_meeting": respond_to_meeting,
    "schedule_task": schedule_task,
    "system_query": system_query,
    "speak": speak,
    "ppt_read": ppt_read,
    "ppt_edit": ppt_edit,
    "word_read": word_read,
    "word_edit": word_edit,
    "excel_read": excel_read,
    "excel_edit": excel_edit,
    "excel_analyze": excel_analyze,
    "set_brand": set_brand,
    "brand_deck": brand_deck,
    "brand_doc": brand_doc,
    "make_report": make_report,
    "generate_image": generate_image,
    "pdf_read": pdf_read,
    "pdf_extract_tables": pdf_extract_tables,
    "pdf_fill_form": pdf_fill_form,
    "pdf_form_fields": pdf_form_fields,
}
