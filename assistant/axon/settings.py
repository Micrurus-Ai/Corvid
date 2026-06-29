"""Persisted user settings (e.g. the inbox auto-filer on/off), stored next to the app."""
import os
import json

# axon/ lives inside the assistant dir, so the settings file sits one level up (next to the app).
_SETTINGS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "axon_settings.json"
)


def load_settings():
    try:
        with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass
