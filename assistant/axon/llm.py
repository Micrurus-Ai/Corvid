"""Provider routing for the dot.

text_llm() returns the (client, model) for PURE-TEXT tasks — chat, tone learning, meeting notes,
document Q&A answers, inbox triage, draft replies. These go to Mistral when MISTRAL_API_KEY is set
(cheaper + EU data residency), otherwise OpenAI.

Resilience: when Mistral is the primary, OpenAI is wired in as an automatic BACKUP. If a Mistral
call errors, times out, or the service is down, the same request is transparently retried on OpenAI
(GPT) so the feature keeps working. Callers don't change — they still do
`client, model = text_llm(); client.chat.completions.create(model=model, ...)`.

Vision (agent/guide screen reading), voice transcription, image generation, and embeddings stay on
OpenAI and just use OpenAI() + MODEL directly — Mistral either can't do them or does them differently.
"""
import os
import sys

from openai import OpenAI

from axon.config import MODEL, MISTRAL_API_KEY, MISTRAL_BASE, TEXT_MODEL


class _FallbackCompletions:
    """.create() tries the primary (Mistral); on any failure it retries on the backup (OpenAI)."""

    def __init__(self, primary, backup, backup_model):
        self._primary = primary
        self._backup = backup
        self._backup_model = backup_model

    def create(self, **kwargs):
        try:
            return self._primary.chat.completions.create(**kwargs)
        except Exception as e:
            if self._backup is None:
                raise
            try:
                print(f"[axon.llm] Mistral unavailable ({e.__class__.__name__}: {e}); "
                      "falling back to OpenAI.", file=sys.stderr)
            except Exception:
                pass
            kwargs = dict(kwargs)
            kwargs["model"] = self._backup_model   # swap in the OpenAI model name
            return self._backup.chat.completions.create(**kwargs)


class _FallbackChat:
    def __init__(self, primary, backup, backup_model):
        self.completions = _FallbackCompletions(primary, backup, backup_model)


class FallbackClient:
    """Quacks like an OpenAI client for `.chat.completions.create`, but transparently retries the
    request on OpenAI whenever the primary (Mistral) errors or is unreachable."""

    def __init__(self, primary, backup, backup_model):
        self.chat = _FallbackChat(primary, backup, backup_model)


def text_llm():
    """(client, model) for a pure-text completion.

    - Mistral configured: Mistral is primary, OpenAI is the automatic backup (used only if Mistral
      fails/times out/is down). A short timeout + single retry means we fail over to OpenAI quickly
      instead of hanging when Mistral is unresponsive.
    - No Mistral key: plain OpenAI.
    """
    if MISTRAL_API_KEY:
        primary = OpenAI(api_key=MISTRAL_API_KEY, base_url=MISTRAL_BASE, timeout=30, max_retries=1)
        backup = OpenAI() if os.getenv("OPENAI_API_KEY") else None
        return FallbackClient(primary, backup, MODEL), TEXT_MODEL
    return OpenAI(), MODEL
