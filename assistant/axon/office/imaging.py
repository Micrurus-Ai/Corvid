"""Generate images with OpenAI (for slides, documents, or standalone). Returns a saved PNG path
that the agent can then drop into a deck/doc via ppt_edit/word_edit add_image."""
import os
import base64
import urllib.request

from openai import OpenAI

from axon.util import _result
from axon.config import DOWNLOADS_DIR


def generate_image(args):
    """Create an image from a text prompt and save it as a PNG. args: prompt, path?, size?
    (1024x1024 | 1536x1024 | 1024x1536)."""
    prompt = args.get("prompt")
    if not prompt:
        return _result("Provide a prompt describing the image to generate.", True)
    out = args.get("path") or os.path.join(DOWNLOADS_DIR, "axon_image.png")
    try:
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    except Exception:
        pass
    if not os.getenv("OPENAI_API_KEY"):
        return _result("No API key available for image generation.", True)
    try:
        client = OpenAI()
        resp = client.images.generate(
            model=os.getenv("ASSISTANT_IMAGE_MODEL", "gpt-image-1"),
            prompt=prompt, size=args.get("size", "1024x1024"), n=1)
        d = resp.data[0]
        if getattr(d, "b64_json", None):
            with open(out, "wb") as f:
                f.write(base64.b64decode(d.b64_json))
        elif getattr(d, "url", None):
            urllib.request.urlretrieve(d.url, out)
        else:
            return _result("Image API returned no image data.", True)
        return _result(f"Image saved: {out}")
    except Exception as e:
        return _result(f"Image generation failed: {e}", True)
