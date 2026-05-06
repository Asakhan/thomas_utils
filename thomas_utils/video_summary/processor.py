"""Orchestrate video → markdown lecture-note conversion.

Pipeline:
  1. Extract audio (ffmpeg, mono 16 kHz wav).
  2. Transcribe via Whisper (faster-whisper preferred).
  3. Detect scene-change keyframes via OpenCV HSV histogram diff.
  4. For each scene, ask a multimodal LLM (Claude or GPT-4o) to summarize
     the transcript + the keyframe screenshot.
  5. Render lecture-note Markdown via formatter.

Long-running (long videos / API throttling) is handled with per-call
timeouts, retry/backoff, and a hard ceiling on the number of scenes.
"""

from __future__ import annotations

import base64
import json
import os
import random
import re
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from thomas_utils.video_summary.formatter import (
    LectureMetadata,
    Section,
    TranscriptSegment,
    format_timestamp,
    render_lecture_notes,
    transcript_excerpt_for_range,
)
from thomas_utils.video_summary.stt import STTError, transcribe


class VideoSummaryError(RuntimeError):
    """Top-level error type for the video_summary pipeline."""


# ---------------------------------------------------------------------------
# Audio extraction
# ---------------------------------------------------------------------------

def _ensure_ffmpeg() -> None:
    if shutil.which("ffmpeg") is None:
        raise VideoSummaryError(
            "ffmpeg binary not found in PATH. Install ffmpeg (e.g. apt install ffmpeg / "
            "brew install ffmpeg / winget install ffmpeg) and retry."
        )


def _extract_audio(video_path: Path, out_path: Path, timeout: int = 1800) -> Path:
    """Extract mono 16 kHz WAV using ffmpeg."""
    _ensure_ffmpeg()
    cmd = [
        "ffmpeg", "-y", "-i", str(video_path),
        "-vn", "-ac", "1", "-ar", "16000", "-f", "wav", str(out_path),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired as e:
        raise VideoSummaryError(f"ffmpeg audio extraction timed out after {timeout}s") from e
    if proc.returncode != 0:
        raise VideoSummaryError(
            "ffmpeg audio extraction failed:\n" + proc.stderr.decode("utf-8", "replace")[-2000:]
        )
    return out_path


# ---------------------------------------------------------------------------
# Scene-change detection (OpenCV HSV histogram diff)
# ---------------------------------------------------------------------------

@dataclass
class Keyframe:
    timestamp: float
    image_path: Path


def _detect_keyframes(
    video_path: Path,
    out_dir: Path,
    threshold: float = 0.55,
    sample_fps: float = 2.0,
    min_gap_seconds: float = 8.0,
    max_scenes: int = 40,
) -> Tuple[List[Keyframe], float]:
    """Detect scene-change keyframes by comparing HSV histograms of sampled frames.

    Returns (keyframes, total_duration_seconds). The list always contains the
    very first frame (scene #1 = intro).
    """
    try:
        import cv2  # type: ignore
        import numpy as np  # type: ignore
    except ModuleNotFoundError as e:
        raise VideoSummaryError(
            "OpenCV (cv2) is required for scene detection. "
            "Install with: pip install opencv-python numpy"
        ) from e

    cap = cv2.VideoCapture(str(video_path))
    if not cap.isOpened():
        raise VideoSummaryError(f"OpenCV could not open video: {video_path}")

    fps = cap.get(cv2.CAP_PROP_FPS) or 0.0
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    duration = (total_frames / fps) if fps > 0 else 0.0
    if fps <= 0:
        cap.release()
        raise VideoSummaryError("Could not determine FPS of input video.")

    step = max(1, int(round(fps / max(0.1, sample_fps))))

    out_dir.mkdir(parents=True, exist_ok=True)

    def _hist(frame) -> "np.ndarray":
        small = cv2.resize(frame, (320, 180))
        hsv = cv2.cvtColor(small, cv2.COLOR_BGR2HSV)
        h = cv2.calcHist([hsv], [0, 1], None, [50, 60], [0, 180, 0, 256])
        cv2.normalize(h, h, 0, 1, cv2.NORM_MINMAX)
        return h

    keyframes: List[Keyframe] = []
    prev_hist = None
    last_kept_t = -1e9
    frame_idx = 0
    saved = 0

    def _save(frame, t: float) -> Keyframe:
        nonlocal saved
        saved += 1
        path = out_dir / f"scene_{saved:03d}.jpg"
        cv2.imwrite(str(path), frame, [int(cv2.IMWRITE_JPEG_QUALITY), 85])
        return Keyframe(timestamp=t, image_path=path)

    while True:
        cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
        ok, frame = cap.read()
        if not ok or frame is None:
            break
        t = frame_idx / fps
        cur = _hist(frame)
        if prev_hist is None:
            keyframes.append(_save(frame, t))
            last_kept_t = t
        else:
            # correlation: 1.0 identical, lower means more change
            corr = cv2.compareHist(prev_hist, cur, cv2.HISTCMP_CORREL)
            change = 1.0 - float(corr)
            if change >= threshold and (t - last_kept_t) >= min_gap_seconds:
                keyframes.append(_save(frame, t))
                last_kept_t = t
                if len(keyframes) >= max_scenes:
                    break
        prev_hist = cur
        frame_idx += step

    cap.release()
    return keyframes, duration


# ---------------------------------------------------------------------------
# Section assembly: split transcript by keyframe boundaries
# ---------------------------------------------------------------------------

def _assemble_sections(
    keyframes: List[Keyframe],
    segments: List[TranscriptSegment],
    duration: float,
) -> List[Tuple[Keyframe, float, float, str]]:
    """Pair each keyframe with its (start, end) range and the transcript text inside it.

    Returns: list of (keyframe, start, end, transcript_text).
    """
    out: List[Tuple[Keyframe, float, float, str]] = []
    boundaries = [kf.timestamp for kf in keyframes] + [duration or (segments[-1].end if segments else 0.0)]
    for i, kf in enumerate(keyframes):
        start = kf.timestamp
        end = boundaries[i + 1]
        if end <= start:
            end = start + 1.0
        text = " ".join(
            (s.text or "").strip()
            for s in segments
            if s.end > start and s.start < end and (s.text or "").strip()
        )
        out.append((kf, start, end, text))
    return out


# ---------------------------------------------------------------------------
# LLM section summarization (Anthropic Claude / OpenAI GPT-4o)
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = (
    "당신은 강의 영상을 마크다운 강의노트로 정리하는 어시스턴트입니다. "
    "주어진 스크린샷과 해당 구간의 음성 스크립트를 바탕으로, 그 구간의 "
    "핵심 주제와 요약을 한국어로 작성하세요. "
    "출력은 반드시 다음 JSON 스키마만 사용합니다 (다른 텍스트 금지):\n"
    '{"title": "<짧은 섹션 제목>",'
    ' "bullets": ["핵심 포인트1", "핵심 포인트2", "..."],'
    ' "summary": "<2~4문장 요약>"}'
)


def _user_prompt_for_section(idx: int, start: float, end: float, transcript_text: str) -> str:
    excerpt = transcript_text.strip()
    if len(excerpt) > 6000:
        excerpt = excerpt[:6000] + " …"
    return (
        f"섹션 {idx} ({format_timestamp(start)} — {format_timestamp(end)})의 "
        f"스크린샷과 스크립트를 분석해 위 JSON 스키마로만 응답하세요.\n\n"
        f"스크립트:\n\"\"\"\n{excerpt or '(이 구간에는 음성이 거의 없음)'}\n\"\"\""
    )


def _parse_json_response(raw: str, fallback_title: str) -> Tuple[str, List[str], str]:
    """Be lenient about wrapping ```json fences."""
    s = raw.strip()
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s, flags=re.MULTILINE).strip()
    try:
        obj = json.loads(s)
    except Exception:
        return fallback_title, [], s[:500].strip()
    title = (obj.get("title") or fallback_title).strip()
    bullets = [str(b).strip() for b in (obj.get("bullets") or []) if str(b).strip()]
    summary = (obj.get("summary") or "").strip()
    return title, bullets, summary


def _summarize_section_anthropic(
    image_b64: str,
    user_prompt: str,
    model: str,
    timeout: int,
) -> str:
    try:
        import anthropic  # type: ignore
    except ModuleNotFoundError as e:
        raise VideoSummaryError(
            "anthropic SDK not installed. pip install anthropic"
        ) from e
    client = anthropic.Anthropic(timeout=timeout)
    msg = client.messages.create(
        model=model,
        max_tokens=1024,
        system=_SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}},
                {"type": "text", "text": user_prompt},
            ],
        }],
    )
    parts = []
    for block in msg.content or []:
        if getattr(block, "type", None) == "text":
            parts.append(block.text)
    return "".join(parts)


def _summarize_section_openai(
    image_b64: str,
    user_prompt: str,
    model: str,
    timeout: int,
) -> str:
    try:
        import openai  # type: ignore
    except ModuleNotFoundError as e:
        raise VideoSummaryError(
            "openai SDK not installed. pip install openai"
        ) from e
    client = openai.OpenAI(timeout=timeout)
    r = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "text", "text": user_prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}},
            ]},
        ],
        max_tokens=1024,
    )
    return (r.choices[0].message.content or "") if r.choices else ""


def _summarize_with_retry(
    provider: str,
    model: str,
    image_b64: str,
    user_prompt: str,
    timeout: int,
    retries: int = 4,
) -> str:
    last_err: Optional[Exception] = None
    for attempt in range(retries):
        try:
            if provider == "anthropic":
                return _summarize_section_anthropic(image_b64, user_prompt, model, timeout)
            if provider == "openai":
                return _summarize_section_openai(image_b64, user_prompt, model, timeout)
            raise VideoSummaryError(f"Unknown LLM provider: {provider}")
        except VideoSummaryError:
            raise
        except Exception as e:
            last_err = e
            msg = str(e).lower()
            transient = any(t in msg for t in ("rate", "timeout", "429", "503", "overloaded", "temporar"))
            if attempt == retries - 1 or not transient:
                break
            sleep = (2 ** attempt) + random.uniform(0, 0.5)
            time.sleep(sleep)
    raise VideoSummaryError(f"LLM call failed after {retries} attempts: {last_err}")


def _default_model(provider: str) -> str:
    return "claude-sonnet-4-6" if provider == "anthropic" else "gpt-4o"


# ---------------------------------------------------------------------------
# Public entrypoint
# ---------------------------------------------------------------------------

def convert_video(
    video_path: str | Path,
    output_path: str | Path,
    *,
    provider: str = "openai",
    model: Optional[str] = None,
    whisper_model: str = "base",
    language: Optional[str] = None,
    scene_threshold: float = 0.55,
    min_gap_seconds: float = 8.0,
    max_scenes: int = 40,
    api_timeout: int = 120,
    audio_timeout: int = 1800,
    screenshots_dir: Optional[str | Path] = None,
    title: Optional[str] = None,
) -> Path:
    """Convert a lecture video into a Markdown lecture-note document.

    Args:
        video_path: Input video file.
        output_path: Output Markdown file path.
        provider: "anthropic" (Claude) or "openai" (GPT-4o).
        model: Override default model name for the provider.
        whisper_model: Whisper model size (tiny / base / small / medium / large-v3).
        language: ISO language code for STT; None = auto-detect.
        scene_threshold: 0.0–1.0; higher = fewer scene cuts.
        min_gap_seconds: Minimum spacing between detected scenes.
        max_scenes: Hard cap to bound API cost on long videos.
        api_timeout: Per-LLM-call timeout (seconds).
        audio_timeout: ffmpeg audio extraction timeout (seconds).
        screenshots_dir: Where to write keyframe images. Default: <output>_assets/.
        title: Override the lecture title (default = video filename stem).

    Returns:
        Path of the written Markdown file.
    """
    video_path = Path(video_path)
    output_path = Path(output_path)
    if not video_path.exists():
        raise VideoSummaryError(f"Video not found: {video_path}")

    provider = provider.lower().strip()
    if provider not in ("anthropic", "openai"):
        raise VideoSummaryError(f"Unknown provider: {provider}. Use 'anthropic' or 'openai'.")
    model = model or _default_model(provider)

    try:
        from dotenv import load_dotenv  # type: ignore
        load_dotenv()
    except ModuleNotFoundError:
        pass

    required_env = "ANTHROPIC_API_KEY" if provider == "anthropic" else "OPENAI_API_KEY"
    if not os.environ.get(required_env):
        raise VideoSummaryError(
            f"{required_env} is not set. Put it in .env or export it before running."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    assets_dir = Path(screenshots_dir) if screenshots_dir else (
        output_path.parent / f"{output_path.stem}_assets"
    )
    assets_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        wav_path = Path(tmp) / "audio.wav"
        _extract_audio(video_path, wav_path, timeout=audio_timeout)

        try:
            segments, detected_lang = transcribe(
                wav_path, model_name=whisper_model, language=language
            )
        except STTError as e:
            raise VideoSummaryError(str(e)) from e

        keyframes, duration = _detect_keyframes(
            video_path,
            assets_dir,
            threshold=scene_threshold,
            min_gap_seconds=min_gap_seconds,
            max_scenes=max_scenes,
        )
        if not keyframes:
            raise VideoSummaryError("No keyframes were detected.")

        if duration <= 0 and segments:
            duration = segments[-1].end

        section_inputs = _assemble_sections(keyframes, segments, duration)

        sections: List[Section] = []
        for idx, (kf, start, end, transcript_text) in enumerate(section_inputs, start=1):
            image_b64 = base64.b64encode(kf.image_path.read_bytes()).decode("ascii")
            user_prompt = _user_prompt_for_section(idx, start, end, transcript_text)
            try:
                raw = _summarize_with_retry(
                    provider, model, image_b64, user_prompt, api_timeout
                )
            except VideoSummaryError as e:
                raw = json.dumps({
                    "title": f"섹션 {idx}",
                    "bullets": [],
                    "summary": f"(LLM 요약 실패: {e})",
                }, ensure_ascii=False)
            section_title, bullets, summary = _parse_json_response(raw, fallback_title=f"섹션 {idx}")
            sections.append(Section(
                index=idx,
                start=start,
                end=end,
                title=section_title,
                summary=summary,
                bullets=bullets,
                screenshot_path=kf.image_path,
                transcript_excerpt=transcript_excerpt_for_range(segments, start, end),
            ))

        metadata = LectureMetadata(
            title=title or video_path.stem,
            source_filename=video_path.name,
            duration_seconds=duration,
            section_count=len(sections),
            language=detected_lang,
            model=f"{provider}:{model}",
        )
        markdown = render_lecture_notes(metadata, sections, segments, output_path)
        output_path.write_text(markdown, encoding="utf-8")
        return output_path
