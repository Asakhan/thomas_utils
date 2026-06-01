"""Speech-to-text via Whisper.

Prefers `faster-whisper` (CTranslate2, lighter and faster) and falls back
to OpenAI's reference `whisper` package when faster-whisper is unavailable.

Long audio files are split into chunks before transcription so that the
in-memory float32 audio buffer (and Whisper activations) stay bounded —
critical on small VMs where decoding a 1+ hour wav at once gets the
process OOM-killed.

Returns a list of timestamped TranscriptSegment objects.
"""

from __future__ import annotations

import gc
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Optional, Tuple

from thomas_utils.video_summary.formatter import TranscriptSegment


class STTError(RuntimeError):
    """Raised when audio transcription fails or no backend is available."""


# Default chunk length for long-audio splitting. ~10 minutes of 16kHz mono
# float32 audio is ~38 MB, which keeps faster-whisper's working set well
# under 1 GB even on small VMs.
_DEFAULT_CHUNK_SECONDS = 600


def transcribe(
    audio_path: Path,
    model_name: str = "base",
    language: Optional[str] = None,
    device: str = "auto",
    chunk_seconds: int = _DEFAULT_CHUNK_SECONDS,
) -> Tuple[List[TranscriptSegment], Optional[str]]:
    """Transcribe an audio file into timestamped segments.

    Returns:
        (segments, detected_language). detected_language may be None.
    """
    audio_path = Path(audio_path)
    if not audio_path.exists():
        raise STTError(f"Audio file not found: {audio_path}")

    try:
        return _transcribe_faster_whisper(
            audio_path, model_name, language, device, chunk_seconds
        )
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


def _probe_duration_seconds(audio_path: Path) -> Optional[float]:
    """Return audio duration via ffprobe, or None if unavailable."""
    if shutil.which("ffprobe") is None:
        return None
    try:
        proc = subprocess.run(
            [
                "ffprobe", "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                str(audio_path),
            ],
            capture_output=True, timeout=30,
        )
    except subprocess.SubprocessError:
        return None
    if proc.returncode != 0:
        return None
    try:
        return float(proc.stdout.decode("utf-8").strip())
    except ValueError:
        return None


def _split_audio_chunks(
    audio_path: Path, out_dir: Path, chunk_seconds: int
) -> List[Tuple[Path, float]]:
    """Use ffmpeg's segment muxer to slice wav into N-second chunks.

    Returns list of (chunk_path, start_offset_seconds), in order.
    """
    if shutil.which("ffmpeg") is None:
        raise STTError("ffmpeg is required to chunk long audio for transcription.")

    pattern = str(out_dir / "chunk_%05d.wav")
    cmd = [
        "ffmpeg", "-y", "-loglevel", "error",
        "-i", str(audio_path),
        "-f", "segment", "-segment_time", str(chunk_seconds),
        "-c", "copy", pattern,
    ]
    proc = subprocess.run(cmd, capture_output=True, timeout=1800)
    if proc.returncode != 0:
        raise STTError(
            "ffmpeg segment split failed:\n"
            + proc.stderr.decode("utf-8", "replace")[-1500:]
        )
    chunks = sorted(out_dir.glob("chunk_*.wav"))
    if not chunks:
        raise STTError("ffmpeg produced no audio chunks.")
    return [(p, i * float(chunk_seconds)) for i, p in enumerate(chunks)]


def _transcribe_faster_whisper(
    audio_path: Path,
    model_name: str,
    language: Optional[str],
    device: str,
    chunk_seconds: int,
) -> Tuple[List[TranscriptSegment], Optional[str]]:
    from faster_whisper import WhisperModel

    compute_type = "int8" if device in ("cpu", "auto") else "float16"
    model = WhisperModel(model_name, device=device, compute_type=compute_type)

    duration = _probe_duration_seconds(audio_path)
    needs_chunking = (
        chunk_seconds > 0 and duration is not None and duration > chunk_seconds * 1.1
    )

    if not needs_chunking:
        return _run_one(model, audio_path, language, time_offset=0.0)

    print(
        f"[stt] audio is {duration:.0f}s; splitting into "
        f"{chunk_seconds}s chunks to bound memory usage.",
        file=sys.stderr,
    )

    all_segments: List[TranscriptSegment] = []
    detected: Optional[str] = language
    with tempfile.TemporaryDirectory(prefix="thomas_stt_chunks_") as tmp:
        chunks = _split_audio_chunks(audio_path, Path(tmp), chunk_seconds)
        for idx, (chunk_path, offset) in enumerate(chunks, 1):
            print(
                f"[stt] transcribing chunk {idx}/{len(chunks)} "
                f"(offset={offset:.0f}s)",
                file=sys.stderr,
            )
            segs, lang = _run_one(model, chunk_path, language or detected, offset)
            all_segments.extend(segs)
            if lang and not detected:
                detected = lang
            # Free the decoded-audio buffer before the next chunk.
            gc.collect()

    return all_segments, detected


def _run_one(
    model,
    audio_path: Path,
    language: Optional[str],
    time_offset: float,
) -> Tuple[List[TranscriptSegment], Optional[str]]:
    segments_iter, info = model.transcribe(
        str(audio_path),
        language=language,
        vad_filter=True,
        beam_size=1,
    )
    out: List[TranscriptSegment] = []
    for s in segments_iter:
        text = (s.text or "").strip()
        if not text:
            continue
        out.append(TranscriptSegment(
            start=float(s.start or 0.0) + time_offset,
            end=float(s.end or 0.0) + time_offset,
            text=text,
        ))
    detected = getattr(info, "language", None) or language
    return out, detected


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
