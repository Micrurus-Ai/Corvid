"""Learn the user's writing tone from their Sent items, save a reusable style guide, and expose it
so both the dot's drafting and the Outlook add-in's Reply sound like the user."""
import os
import re

from openai import OpenAI

from axon.config import IS_WINDOWS, MODEL
from axon.util import _result
from axon.outlook._base import _run_outlook_ps


def _plain(text):
    """Strip Markdown markers (**bold**, *italic*, `code`, # headings) so the tone reads as clean
    plain text in the editable Settings box and when injected into draft prompts."""
    if not text:
        return text
    t = re.sub(r"\*+", "", text)                    # drop * and ** emphasis
    t = re.sub(r"`+", "", t)                         # drop code backticks
    t = re.sub(r"(?m)^\s{0,3}#{1,6}\s+", "", t)      # drop heading hashes
    return t.strip()


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
    # Cut off quoted history so we learn the user's OWN writing, not the emails they replied to.
    $cut = $b.Length
    # Outlook-style header blocks + forward markers that introduce quoted / forwarded content.
    foreach ($mark in @("`r`nFrom:", "`r`nVan:", "`r`nSent:", "`r`nVerzonden:", "-----Original", "-----Oorspronkelijk", "Forwarded message", "Begin forwarded message", "Doorgestuurd bericht", "________________________________")) {
      $i = $b.IndexOf($mark)
      if ($i -ge 0 -and $i -lt $cut) { $cut = $i }
    }
    # Reply intro lines ("On ... wrote:", Dutch "Op ... schreef:", French "Le ... a ecrit :") and > quotes.
    foreach ($pat in @('(?im)^\s*On .+ wrote:\s*$', '(?im)^\s*Op .+ (schreef|geschreven).*:\s*$', '(?im)^\s*Le .+ a .crit\s*:\s*$', '(?m)^>')) {
      $mm = [regex]::Match($b, $pat)
      if ($mm.Success -and $mm.Index -lt $cut) { $cut = $mm.Index }
    }
    $b = $b.Substring(0, $cut).Trim()
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
        return _result("Axon isn't set up with an API key yet, so I can't analyse your writing.", True)
    out = _run_outlook_ps(_SENT_PS, {}, show=False)
    if not out or out.startswith("OL_ERROR") or "=====EMAIL=====" not in out:
        return _result("Couldn't read your Sent items: " + (out or "no output"), True)
    samples = out[:16000]
    prompt = (
        "Below are recent emails the USER has SENT (their own outgoing messages, quoted history "
        "removed). Describe ONLY how the user themselves writes; ignore any quoted or forwarded text "
        "from other people that may remain. Write a concise STYLE GUIDE another writer could follow to "
        "sound exactly like this person when drafting their emails. Describe: typical greeting, "
        "sign-off, formality/warmth, sentence length, and whether they use bullet points or emojis. "
        "Note which languages they write in (Dutch, French, English) and the general tone per "
        "language. "
        "IMPORTANT: describe PATTERNS, not one-off examples. Do NOT quote specific phrases unless the "
        "SAME phrase clearly recurs across MULTIPLE emails. If you are not certain a phrase is the "
        "user's own recurring wording, leave it out entirely. Under 200 words. Do NOT summarise the "
        "email contents.\n\n"
        "Format: plain text with simple '- ' bullets. Do NOT use Markdown and do NOT wrap any words "
        "in asterisks (no * or **). Write labels as 'Greeting:' not '**Greeting**:'.\n\n" + samples
    )
    try:
        from axon.llm import text_llm
        client, model = text_llm()
        resp = client.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}], temperature=0.2)
        guide = _plain((resp.choices[0].message.content or "").strip())
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
            return _plain(f.read().strip())
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
