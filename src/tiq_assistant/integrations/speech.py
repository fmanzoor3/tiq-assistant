"""Local, offline speech-to-text via faster-whisper (optional dependency).

Voice input is a convenience layered on top of the always-available text box.
Both faster-whisper (transcription) and sounddevice (mic capture) are imported
lazily, so the app runs fine without them -- ``is_available()`` returns False
and the UI falls back to typing.

Nothing here talks to the network; the whisper model runs on-device.
"""

from __future__ import annotations

import logging
import tempfile
import wave
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Default to a small, fast model that runs on CPU. Can be overridden via config.
DEFAULT_MODEL_SIZE = "base"


def is_available() -> tuple[bool, str]:
    """Return (available, reason). available=True only if both deps import."""
    try:
        import sounddevice  # noqa: F401
    except Exception as e:  # noqa: BLE001
        return False, f"microphone capture unavailable (sounddevice): {e}"
    try:
        import faster_whisper  # noqa: F401
    except Exception as e:  # noqa: BLE001
        return False, f"speech-to-text unavailable (faster-whisper): {e}"
    return True, ""


class AudioRecorder:
    """Records mic audio to a temp WAV file until stopped."""

    def __init__(self, samplerate: int = 16000, channels: int = 1):
        self.samplerate = samplerate
        self.channels = channels
        self._frames: list = []
        self._stream = None

    def start(self) -> None:
        import sounddevice as sd
        import numpy as np  # sounddevice pulls numpy

        self._frames = []

        def _callback(indata, frames, time_info, status):  # noqa: ANN001
            if status:
                logger.debug("Audio status: %s", status)
            self._frames.append(indata.copy())

        self._stream = sd.InputStream(
            samplerate=self.samplerate,
            channels=self.channels,
            dtype="int16",
            callback=_callback,
        )
        self._stream.start()

    def stop(self) -> Optional[Path]:
        """Stop recording and return the path to a temp WAV, or None if empty."""
        if self._stream is None:
            return None
        self._stream.stop()
        self._stream.close()
        self._stream = None

        if not self._frames:
            return None

        import numpy as np

        audio = np.concatenate(self._frames, axis=0)
        tmp = Path(tempfile.gettempdir()) / "tiq_voice_entry.wav"
        with wave.open(str(tmp), "wb") as wf:
            wf.setnchannels(self.channels)
            wf.setsampwidth(2)  # int16
            wf.setframerate(self.samplerate)
            wf.writeframes(audio.tobytes())
        return tmp


_model_cache = {}


class ModelUnavailableError(Exception):
    """Raised when the whisper model can't be loaded (e.g. download blocked)."""


def _looks_like_path(model_ref: str) -> bool:
    """True if model_ref points at a local folder rather than a Hub name."""
    try:
        return Path(model_ref).exists()
    except Exception:
        return False


def transcribe(
    wav_path: Path,
    model_ref: str = DEFAULT_MODEL_SIZE,
    language: Optional[str] = None,
) -> str:
    """Transcribe a WAV file using faster-whisper (CPU int8).

    ``model_ref`` may be either a model size name ("base", "small", ...) which
    faster-whisper downloads from the Hugging Face Hub on first use, OR a path
    to a local model folder, which loads fully offline (no network). On locked-
    down machines where the Hub download is blocked (WinError 10013), point
    ``model_ref`` at a pre-downloaded model folder via Settings.
    """
    import os

    is_local = _looks_like_path(model_ref)

    if is_local:
        # Loading a local folder: force offline mode so no download is attempted.
        os.environ.setdefault("HF_HUB_OFFLINE", "1")
        os.environ.setdefault("TRANSFORMERS_OFFLINE", "1")
    else:
        # Downloading from the Hub on a corporate network: disable the
        # accelerated "xet" transfer path, which opens a socket that some
        # firewalls forbid (WinError 10013). This forces plain HTTPS downloads
        # that go through the normal proxy -- slower, but far more likely to work.
        os.environ.setdefault("HF_HUB_DISABLE_XET", "1")
        os.environ.setdefault("HF_XET_DISABLE", "1")

    from faster_whisper import WhisperModel

    if model_ref not in _model_cache:
        try:
            _model_cache[model_ref] = WhisperModel(
                model_ref, device="cpu", compute_type="int8"
            )
        except Exception as e:  # noqa: BLE001
            # Most commonly: Hub download blocked by corporate firewall.
            raise ModelUnavailableError(
                "Could not load the speech model. If this machine blocks "
                "internet downloads, download a faster-whisper model on another "
                "machine and set its folder path in Settings > AI Assistant "
                f"(Whisper model path).\n\nDetails: {e}"
            )
    model = _model_cache[model_ref]

    segments, _info = model.transcribe(str(wav_path), language=language)
    return " ".join(seg.text.strip() for seg in segments).strip()
