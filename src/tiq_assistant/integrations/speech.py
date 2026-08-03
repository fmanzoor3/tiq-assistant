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


def transcribe(wav_path: Path, model_size: str = DEFAULT_MODEL_SIZE, language: Optional[str] = None) -> str:
    """Transcribe a WAV file to text using faster-whisper (CPU int8)."""
    from faster_whisper import WhisperModel

    if model_size not in _model_cache:
        # int8 on CPU is a good speed/quality tradeoff for short clips.
        _model_cache[model_size] = WhisperModel(model_size, device="cpu", compute_type="int8")
    model = _model_cache[model_size]

    segments, _info = model.transcribe(str(wav_path), language=language)
    return " ".join(seg.text.strip() for seg in segments).strip()
