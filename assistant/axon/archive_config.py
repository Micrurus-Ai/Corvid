"""Shared config for the Outlook add-in's 'suggest a download folder' feature. Written by the dot's
Settings panel, read by the add-in (C#). Lives in %APPDATA%\\AxonOutlook\\archive.json so both sides
see it.

Shape:
{
  "client_base":   "T:\\IF\\Sales\\AB\\SOP",       # required; sample code in it (e.g. AB) is the code slot
  "supplier_base": "T:\\IF\\Purchasing\\AB\\PO",   # optional; blank = suppliers use the same tree
  "country_codes": {"Belgium": "AB", "Germany": "AB", "France": "FR"},
  "save_mode":     "both",                         # email | attachments | both
  "default_subfolder": "Order\\MC"                 # optional; appended after the matched folder
}
"""
import os
import json


def _path():
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonOutlook")
    return os.path.join(base, "archive.json")


def load():
    try:
        with open(_path(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save(cfg):
    try:
        os.makedirs(os.path.dirname(_path()), exist_ok=True)
        with open(_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False
