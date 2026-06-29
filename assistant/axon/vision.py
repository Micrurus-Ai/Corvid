"""Vision & voice helpers: OCR / extract text from an image (multimodal model), record from the
mic, and transcribe speech to text (for voice input)."""
import os
import base64
import tempfile

from openai import OpenAI

from axon.util import _result
from axon.config import MODEL


def _b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def ocr_image(args):
    """Extract all text (and tables, as markdown) from an image file. args: path, instruction?"""
    path = args.get("path")
    if not path or not os.path.exists(path):
        return _result(f"No such image: {path}", True)
    if not os.getenv("OPENAI_API_KEY"):
        return _result("No API key available for OCR.", True)
    instr = args.get("instruction") or (
        "Extract ALL text from this image exactly as written. Preserve any tables as markdown. "
        "Output only the extracted content, nothing else.")
    try:
        client = OpenAI()
        resp = client.chat.completions.create(
            model=MODEL, temperature=0,
            messages=[{"role": "user", "content": [
                {"type": "text", "text": instr},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{_b64(path)}"}}]}])
        return _result(resp.choices[0].message.content or "(no text found)")
    except Exception as e:
        return _result(f"OCR failed: {e}", True)


def transcribe_audio(path):
    """Transcribe an audio file to text. Returns '' on failure."""
    if not path or not os.path.exists(path) or not os.getenv("OPENAI_API_KEY"):
        return ""
    try:
        client = OpenAI()
        with open(path, "rb") as f:
            r = client.audio.transcriptions.create(
                model=os.getenv("ASSISTANT_STT_MODEL", "whisper-1"), file=f)
        return (getattr(r, "text", "") or "").strip()
    except Exception:
        return ""


def record_and_transcribe(seconds=6, samplerate=16000):
    """Record from the default microphone for a few seconds and return the transcribed text.
    Used by the composer's mic button. Returns '' if recording/transcription isn't available."""
    try:
        import sounddevice as sd
        import soundfile as sf
    except Exception:
        return ""
    try:
        audio = sd.rec(int(seconds * samplerate), samplerate=samplerate, channels=1)
        sd.wait()
        path = os.path.join(tempfile.gettempdir(), "axon_voice.wav")
        sf.write(path, audio, samplerate)
        return transcribe_audio(path)
    except Exception:
        return ""
