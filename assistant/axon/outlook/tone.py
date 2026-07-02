"""Learn the user's writing tone from their Sent items, save a reusable style guide, and expose it
so both the dot's drafting and the Outlook add-in's Reply sound like the user."""
import os

from openai import OpenAI

from axon.config import IS_WINDOWS, MODEL
from axon.util import _result
from axon.outlook._base import _run_outlook_ps


def _tone_path():
    """Shared with the Outlook add-in, which reads %APPDATA%\\AxonOutlook\\tone.txt at draft time."""
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonOutlook")
    return os.path.join(base, "tone.txt")


# Pull the body of the most recent Sent emails (newest first), skipping quoted history.
_SENT_PS = r'''
$ErrorActionPreference = "SilentlyContinue"
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$sent = $ns.GetDefaultFolder(5)   # olFolderSentMail
$items = $sent.Items
$items.Sort("[SentOn]", $true)
$n = 0
foreach ($m in $items) {
  if ($n -ge 40) { break }
  try {
    $b = [string]$m.Body
    # Cut off quoted history so we learn the user's OWN writing, not what they replied to.
    foreach ($mark in @("`r`nFrom:", "`r`nVan:", "-----Original", "________________________________")) {
      $i = $b.IndexOf($mark)
      if ($i -gt 0) { $b = $b.Substring(0, $i) }
    }
    $b = $b.Trim()
    if ($b.Length -lt 20) { continue }
    if ($b.Length -gt 1200) { $b = $b.Substring(0, 1200) }
    [Console]::Out.WriteLine("=====EMAIL=====")
    [Console]::Out.WriteLine($b)
    $n++
  } catch {}
}
'''


def learn_my_tone(args=None):
    """Read recent Sent items, derive a concise style guide, and save it for reuse."""
    if not IS_WINDOWS:
        return _result("Learning your tone needs the Windows Outlook app.", True)
    if not os.getenv("OPENAI_API_KEY"):
        return _result("No OpenAI key configured, so I can't analyse your writing.", True)
    out = _run_outlook_ps(_SENT_PS, {}, show=False)
    if not out or out.startswith("OL_ERROR") or "=====EMAIL=====" not in out:
        return _result("Couldn't read your Sent items: " + (out or "no output"), True)
    samples = out[:16000]
    prompt = (
        "Below are recent emails the user has SENT (quoted history removed). Write a concise STYLE "
        "GUIDE another writer could follow to sound exactly like this person when drafting their "
        "emails. Describe: typical greeting, sign-off, formality/warmth, sentence length, whether "
        "they use bullet points or emojis, and any signature phrases. If you see Dutch, French, or "
        "English, add a one-line note per language. Output plain-text bullet points, under 200 words. "
        "Do NOT summarise the email contents.\n\n" + samples
    )
    try:
        client = OpenAI()
        resp = client.chat.completions.create(
            model=MODEL, messages=[{"role": "user", "content": prompt}], temperature=0.2)
        guide = (resp.choices[0].message.content or "").strip()
    except Exception as e:
        return _result("Couldn't derive your tone: " + str(e), True)
    if not guide:
        return _result("Couldn't derive your tone (the model returned nothing).", True)
    try:
        p = _tone_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write(guide)
    except Exception as e:
        return _result("Derived your tone but couldn't save it: " + str(e), True)
    return _result(
        "Learned your writing tone from your Sent items. Axon will now match your style in replies "
        "and drafts (in the dot and the Outlook add-in).\n\n" + guide)


def my_tone():
    """Return the saved style guide (or '' if none yet) for injecting into draft prompts."""
    try:
        with open(_tone_path(), encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def save_tone(text):
    """Save the user's writing tone (set manually in Settings, or derived from Sent items)."""
    try:
        p = _tone_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write((text or "").strip())
        return True
    except Exception:
        return False
