"""Minimal client for a local, OpenAI-compatible LLM (company Qwen).

Talks to an internal vLLM-style endpoint like:

    https://ray-gpu-prod.bt.local/llms/qwen_qwen3_14B_fp8/v1

using only the Python standard library (urllib + ssl), so it needs no extra
dependencies (openai / httpx). The internal server uses a self-signed cert, so
TLS verification is disabled by default -- this is why the whole AI feature is
OPT-IN and OFF by default: enabling it means the app makes a network call to
the configured internal host, which is a deliberate change from the app's
otherwise fully-offline behaviour.

Nothing here raises on import, and all calls fail soft (raise LLMError, which
callers turn into a friendly message) so the rest of the app is unaffected when
the LLM is disabled or unreachable.
"""

from __future__ import annotations

import json
import ssl
import urllib.request
import urllib.error
from dataclasses import dataclass
from typing import Optional

from tiq_assistant.core.exceptions import TIQAssistantError


class LLMError(TIQAssistantError):
    """Raised when the LLM is unreachable or returns an unusable response."""


@dataclass
class LLMConfig:
    """Connection settings for the local LLM (persisted in user settings)."""
    enabled: bool = False
    base_url: str = "https://ray-gpu-prod.bt.local/llms/qwen_qwen3_14B_fp8/v1"
    api_key: str = "dummy-token"
    model: str = ""            # empty => auto-detect via /models
    verify_ssl: bool = False   # internal self-signed cert
    timeout_seconds: int = 120  # Qwen on shared GPU can be slow; allow headroom
    disable_thinking: bool = True  # skip Qwen3 chain-of-thought for speed
    # Speech-to-text model: a size name ("base") that downloads from the Hub,
    # or a local folder path for offline use on locked-down machines.
    whisper_model: str = "base"
    # User-supplied domain terms/acronyms (comma-separated) to help speech
    # recognition and LLM name-correction (e.g. "RAG, EnGPT, Agentbot").
    custom_terms: str = ""


class LLMClient:
    """Tiny OpenAI-compatible chat client over urllib."""

    def __init__(self, config: LLMConfig):
        self.config = config
        self._resolved_model: Optional[str] = None  # cache to avoid re-hitting /models

    # ------------------------------------------------------------------ http

    def _ssl_context(self) -> Optional[ssl.SSLContext]:
        if self.config.base_url.lower().startswith("https") and not self.config.verify_ssl:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            return ctx
        return None

    def _post(self, path: str, payload: dict) -> dict:
        url = self.config.base_url.rstrip("/") + path
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=data,
            method="POST",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.config.api_key}",
            },
        )
        try:
            with urllib.request.urlopen(
                req, timeout=self.config.timeout_seconds, context=self._ssl_context()
            ) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8", "ignore")[:300]
            except Exception:
                pass
            raise LLMError(f"LLM HTTP {e.code}: {body or e.reason}")
        except urllib.error.URLError as e:
            raise LLMError(f"Could not reach the LLM at {url}: {e.reason}")
        except Exception as e:  # noqa: BLE001
            raise LLMError(f"LLM request failed: {e}")

    def _get(self, path: str) -> dict:
        url = self.config.base_url.rstrip("/") + path
        req = urllib.request.Request(url, method="GET",
                                     headers={"Authorization": f"Bearer {self.config.api_key}"})
        try:
            with urllib.request.urlopen(
                req, timeout=self.config.timeout_seconds, context=self._ssl_context()
            ) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001
            raise LLMError(f"LLM request failed: {e}")

    # --------------------------------------------------------------- public

    def resolve_model(self) -> str:
        """Return the configured model, or auto-detect the first available one.

        The auto-detected id is cached so we don't make an extra /models round
        trip before every completion.
        """
        if self.config.model:
            return self.config.model
        if self._resolved_model:
            return self._resolved_model
        data = self._get("/models")
        models = data.get("data") or []
        if not models:
            raise LLMError("LLM returned no available models.")
        self._resolved_model = models[0]["id"]
        return self._resolved_model

    def test_connection(self) -> str:
        """Ping the endpoint; return the resolved model id or raise LLMError."""
        model = self.resolve_model()
        # A tiny completion to prove the chat route works end-to-end.
        self.chat([{"role": "user", "content": "Reply with: ok"}], model=model, max_tokens=5)
        return model

    def chat(
        self,
        messages: list[dict],
        model: Optional[str] = None,
        temperature: float = 0.0,
        max_tokens: Optional[int] = None,
        json_mode: bool = False,
    ) -> str:
        """Run a chat completion and return the assistant message content."""
        # Copy messages so any in-place tweaks (e.g. /no_think) don't mutate the
        # caller's list, which is reused across retries.
        messages = [dict(m) for m in messages]
        payload: dict = {
            "model": model or self.resolve_model(),
            "messages": messages,
            "temperature": temperature,
        }
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens
        if json_mode:
            # vLLM/OpenAI-compatible servers accept this; harmless if ignored.
            payload["response_format"] = {"type": "json_object"}

        # Qwen3 models do slow chain-of-thought ("thinking") by default, which
        # can take tens of seconds of hidden generation and time the request
        # out. We don't need reasoning for structured entry extraction, so turn
        # it off. Different servers accept different switches, so we send all the
        # common ones (unknown fields are ignored by OpenAI-compatible servers).
        if self.config.disable_thinking:
            payload["chat_template_kwargs"] = {"enable_thinking": False}
            payload["enable_thinking"] = False
            extra_body = payload.setdefault("extra_body", {})
            extra_body["enable_thinking"] = False
            # Qwen also honours a "/no_think" directive in the message text; add
            # it to the last user message as a belt-and-suspenders switch that
            # works even when the API param is ignored.
            msgs = payload["messages"]
            for m in reversed(msgs):
                if m.get("role") == "user":
                    if "/no_think" not in (m.get("content") or ""):
                        m["content"] = "/no_think\n" + (m.get("content") or "")
                    break

        result = self._post("/chat/completions", payload)

        import logging as _logging
        _logging.getLogger(__name__).info("LLM raw response: %s", result)

        try:
            choice = result["choices"][0]
            message = choice.get("message", {})
        except (KeyError, IndexError, TypeError):
            raise LLMError(f"LLM response had no choices. Raw: {json.dumps(result)[:400]}")

        content = message.get("content") or ""
        finish_reason = choice.get("finish_reason")

        # Some Qwen/vLLM setups (thinking mode) put the answer in a separate
        # field instead of "content" -- observed field names include
        # "reasoning_content" and "reasoning". Fall back to whichever has text.
        if not content.strip():
            for alt in ("reasoning_content", "reasoning"):
                alt_val = message.get(alt)
                if alt_val and str(alt_val).strip():
                    content = str(alt_val)
                    break

        import re
        content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip()

        if not content:
            # Surface the diagnostic detail so the cause is visible.
            usage = result.get("usage") or {}
            detail = (
                f"finish_reason={finish_reason}, usage={usage}, "
                f"message_keys={list(message.keys())}"
            )
            if finish_reason == "length":
                raise LLMError(
                    "The model used all its output tokens on hidden reasoning and "
                    "returned no answer. Enable '/no_think' handling or use a "
                    f"smaller/faster model.\n\n[{detail}]"
                )
            raise LLMError(f"The model returned an empty response.\n\n[{detail}]")

        return content
