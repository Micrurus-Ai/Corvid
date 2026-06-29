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


class Recorder:
    """Continuous microphone recorder for the composer's mic button: start() begins capture,
    stop() writes a WAV and returns its path. Uses a raw int16 stream + the stdlib wave module
    (no numpy/soundfile needed), so it's light and reliable."""

    def __init__(self, samplerate=16000):
        self.samplerate = samplerate
        self._frames = []
        self._stream = None

    def start(self):
        try:
            import sounddevice as sd
        except Exception:
            return False
        self._frames = []
        try:
            self._stream = sd.RawInputStream(
                samplerate=self.samplerate, channels=1, dtype="int16",
                callback=lambda indata, n, t, s: self._frames.append(bytes(indata)))
            self._stream.start()
            return True
        except Exception:
            self._stream = None
            return False

    def stop(self):
        import wave
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass
            self._stream = None
        if not self._frames:
            return None
        path = os.path.join(tempfile.gettempdir(), "axon_voice.wav")
        try:
            wf = wave.open(path, "wb")
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(self.samplerate)
            wf.writeframes(b"".join(self._frames))
            wf.close()
        except Exception:
            return None
        self._frames = []
        return path


def play_audio(path):
    """Replay a recorded WAV (Windows, async)."""
    try:
        import winsound
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception:
        pass


def stop_audio():
    try:
        import winsound
        winsound.PlaySound(None, winsound.SND_PURGE)
    except Exception:
        pass
