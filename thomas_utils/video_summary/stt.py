"""Speech-to-text via Whisper.

Prefers `faster-whisper` (CTranslate2, lighter and faster) and falls back
to OpenAI's reference `whisper` package when faster-whisper is unavailable.

Returns a list of timestamped TranscriptSegment objects.
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

from thomas_utils.video_summary.formatter import TranscriptSegment


class STTError(RuntimeError):
    """Raised when audio transcription fails or no backend is available."""


def transcribe(
    audio_path: Path,
    model_name: str = "base",
    language: Optional[str] = None,
    device: str = "auto",
) -> Tuple[List[TranscriptSegment], Optional[str]]:
    """Transcribe an audio file into timestamped segments.

    Returns:
        (segments, detected_language). detected_language may be None.
    """
    audio_path = Path(audio_path)
    if not audio_path.exists():
        raise STTError(f"Audio file not found: {audio_path}")

    try:
        return _transcribe_faster_whisper(audio_path, model_name, language, device)
    except ModuleNotFoundError:
        pass

    try:
        return _transcribe_openai_whisper(audio_path, model_name, language)
    except ModuleNotFoundError as e:
        raise STTError(
            "Whisper is not installed. Install one of:\n"
            "  pip install faster-whisper   (recommended)\n"
            "  pip install openai-whisper"
        ) from e


def _transcribe_faster_whisper(
    audio_path: Path,
    model_name: str,
    language: Optional[str],
    device: str,
) -> Tuple[List[TranscriptSegment], Optional[str]]:
    from faster_whisper import WhisperModel

    compute_type = "int8" if device in ("cpu", "auto") else "float16"
    model = WhisperModel(model_name, device=device, compute_type=compute_type)
    segments_iter, info = model.transcribe(
        str(audio_path),
        language=language,
        vad_filter=True,
        beam_size=1,
    )
    segments = [
        TranscriptSegment(start=float(s.start or 0.0), end=float(s.end or 0.0), text=(s.text or "").strip())
        for s in segments_iter
    ]
    detected = getattr(info, "language", None) or language
    return segments, detected


def _transcribe_openai_whisper(
    audio_path: Path,
    model_name: str,
    language: Optional[str],
) -> Tuple[List[TranscriptSegment], Optional[str]]:
    import whisper

    model = whisper.load_model(model_name)
    result = model.transcribe(str(audio_path), language=language, verbose=False)
    raw_segments = result.get("segments") or []
    segments = [
        TranscriptSegment(
            start=float(s.get("start") or 0.0),
            end=float(s.get("end") or 0.0),
            text=(s.get("text") or "").strip(),
        )
        for s in raw_segments
    ]
    return segments, result.get("language") or language
