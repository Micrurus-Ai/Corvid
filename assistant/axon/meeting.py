"""Meeting -> notes: turn a meeting recording (or transcript) into a summary, decisions,
action items, and a ready-to-send follow-up email draft."""
import os

from openai import OpenAI

from axon.config import MODEL
from axon.util import _result
from axon.vision import transcribe_audio

# Whisper's hard upload limit is 25 MB; warn before a confusing API error.
_MAX_AUDIO_MB = 25
_AUDIO_EXTS = (".wav", ".mp3", ".m4a", ".mp4", ".webm", ".mpga", ".mpeg", ".ogg", ".flac")


def _read_transcript_file(path):
    try:
        with open(path, encoding="utf-8", errors="ignore") as f:
            return f.read().strip()
    except Exception:
        return ""


def meeting_notes(args):
    """Produce notes from a meeting. Accepts an audio/video recording (audio=path), a transcript
    file (transcript=path or .txt), or transcript text directly."""
    title = (args.get("title") or "Meeting").strip()
    transcript = (args.get("transcript") or "").strip()
    path = (args.get("audio") or args.get("path") or "").strip()

    text = ""
    if transcript and not transcript.lower().endswith((".txt", ".md")) and os.sep not in transcript and "/" not in transcript:
        text = transcript  # raw transcript text passed inline
    elif transcript:
        text = _read_transcript_file(transcript)
    elif path:
        if not os.path.exists(path):
            return _result(f"File not found: {path}", True)
        ext = os.path.splitext(path)[1].lower()
        if ext in (".txt", ".md"):
            text = _read_transcript_file(path)
        elif ext in _AUDIO_EXTS:
            try:
                mb = os.path.getsize(path) / (1024 * 1024)
            except Exception:
                mb = 0
            if mb > _MAX_AUDIO_MB:
                return _result(
                    f"That recording is {mb:.0f} MB — over the {_MAX_AUDIO_MB} MB transcription limit. "
                    "Export a smaller/compressed audio file (e.g. .m4a or .mp3), or paste the meeting "
                    "transcript instead.", True)
            text = transcribe_audio(path)
            if not text:
                return _result("Couldn't transcribe that recording (check the file and your audio key).", True)
        else:
            return _result(f"Unsupported file type '{ext}'. Use an audio/video recording or a .txt transcript.", True)
    else:
        return _result("Give me a meeting recording (audio=path), a transcript file, or transcript text.", True)

    if len(text) < 30:
        return _result("The transcript was too short to summarise.", True)
    if len(text) > 24000:
        text = text[:24000]

    prompt = (
        f"You are taking minutes for the meeting titled \"{title}\". From the transcript below, produce "
        "clear notes in this exact markdown layout. Write in the transcript's own language. Be concise "
        "and concrete; if a section has nothing, write '- None'.\n\n"
        "## Summary\n- (3-5 bullet points)\n\n"
        "## Decisions\n- (key decisions made)\n\n"
        "## Action items\n- Owner — task — due date if mentioned\n\n"
        "## Follow-up email (draft)\nA short, friendly recap email the organiser can send to attendees: "
        "greeting, 2-3 line summary, the action items, and a sign-off.\n\n"
        "Transcript:\n" + text
    )
    try:
        client = OpenAI()
        resp = client.chat.completions.create(
            model=MODEL, messages=[{"role": "user", "content": prompt}], temperature=0.3)
        notes = (resp.choices[0].message.content or "").strip()
    except Exception as e:
        return _result("Couldn't produce the notes: " + str(e), True)
    if not notes:
        return _result("Couldn't produce the notes (the model returned nothing).", True)
    return _result(notes + "\n\n_Tip: say \"send the follow-up as an email\" and I'll draft it in Outlook._")
