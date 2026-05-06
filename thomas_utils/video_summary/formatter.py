"""Markdown template for the lecture-note style output.

Produces a document with:
1. Header (title, metadata)
2. Table of contents (jump-links with timestamps)
3. Per-section blocks (timestamp range, screenshot, summary)
4. Full transcript with timestamps
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional


@dataclass
class TranscriptSegment:
    """One Whisper segment."""
    start: float
    end: float
    text: str


@dataclass
class Section:
    """One lecture section bound to a scene-change keyframe."""
    index: int
    start: float
    end: float
    title: str
    summary: str
    bullets: List[str] = field(default_factory=list)
    screenshot_path: Optional[Path] = None
    transcript_excerpt: str = ""


@dataclass
class LectureMetadata:
    """Top-level info about the source video."""
    title: str
    source_filename: str
    duration_seconds: float
    section_count: int
    language: Optional[str] = None
    model: Optional[str] = None
    generated_at: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M"))


def format_timestamp(seconds: float) -> str:
    """Format seconds as HH:MM:SS (or MM:SS if under one hour)."""
    if seconds is None or seconds < 0:
        seconds = 0.0
    total = int(round(seconds))
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


_SLUG_RE = re.compile(r"[^\w\-]+", re.UNICODE)


def slugify(text: str, fallback: str = "section") -> str:
    """Produce a stable Markdown anchor slug. Keeps Hangul and word chars."""
    text = unicodedata.normalize("NFKC", text or "").strip().lower()
    text = re.sub(r"\s+", "-", text)
    text = _SLUG_RE.sub("", text)
    return text or fallback


def _relative_image_path(screenshot: Optional[Path], output_path: Path) -> Optional[str]:
    """Express the screenshot path relative to the output markdown file when possible."""
    if screenshot is None:
        return None
    try:
        rel = Path(screenshot).resolve().relative_to(output_path.parent.resolve())
        return rel.as_posix()
    except ValueError:
        return Path(screenshot).as_posix()


def render_lecture_notes(
    metadata: LectureMetadata,
    sections: Iterable[Section],
    transcript: Iterable[TranscriptSegment],
    output_path: Path,
) -> str:
    """Render the full lecture-note Markdown document."""
    sections = list(sections)
    transcript = list(transcript)

    lines: List[str] = []
    lines.append(f"# {metadata.title}")
    lines.append("")
    meta_bits = [
        f"생성일: {metadata.generated_at}",
        f"소스: `{metadata.source_filename}`",
        f"길이: {format_timestamp(metadata.duration_seconds)}",
        f"섹션 수: {metadata.section_count}",
    ]
    if metadata.language:
        meta_bits.append(f"언어: {metadata.language}")
    if metadata.model:
        meta_bits.append(f"모델: {metadata.model}")
    lines.append("> " + "  |  ".join(meta_bits))
    lines.append("")

    lines.append("## 목차")
    lines.append("")
    for sec in sections:
        anchor = f"section-{sec.index}-{slugify(sec.title)}"
        lines.append(f"- [{format_timestamp(sec.start)} — {sec.title}](#{anchor})")
    lines.append("")
    lines.append("---")
    lines.append("")

    for sec in sections:
        anchor = f"section-{sec.index}-{slugify(sec.title)}"
        lines.append(f'<a id="{anchor}"></a>')
        lines.append(
            f"## {sec.index}. {sec.title} "
            f"({format_timestamp(sec.start)} — {format_timestamp(sec.end)})"
        )
        lines.append("")
        rel = _relative_image_path(sec.screenshot_path, output_path)
        if rel:
            lines.append(f"![섹션 {sec.index} 스크린샷]({rel})")
            lines.append("")
        if sec.bullets:
            lines.append("**핵심 포인트**")
            lines.append("")
            for b in sec.bullets:
                lines.append(f"- {b}")
            lines.append("")
        if sec.summary:
            lines.append("**요약**")
            lines.append("")
            lines.append(sec.summary.strip())
            lines.append("")
        if sec.transcript_excerpt:
            lines.append("<details><summary>해당 구간 스크립트</summary>")
            lines.append("")
            lines.append(sec.transcript_excerpt.strip())
            lines.append("")
            lines.append("</details>")
            lines.append("")
        lines.append("---")
        lines.append("")

    lines.append("## 전체 스크립트")
    lines.append("")
    if not transcript:
        lines.append("_(스크립트가 비어 있습니다.)_")
    else:
        for seg in transcript:
            text = (seg.text or "").strip()
            if not text:
                continue
            lines.append(f"- `[{format_timestamp(seg.start)}]` {text}")
    lines.append("")

    out = "\n".join(lines)
    out = re.sub(r"\n{3,}", "\n\n", out).rstrip() + "\n"
    return out


def transcript_excerpt_for_range(
    segments: Iterable[TranscriptSegment],
    start: float,
    end: float,
    max_chars: int = 1200,
) -> str:
    """Stitch transcript segments overlapping [start, end) into one block, capped to max_chars."""
    chunks: List[str] = []
    total = 0
    for seg in segments:
        if seg.end <= start or seg.start >= end:
            continue
        text = (seg.text or "").strip()
        if not text:
            continue
        line = f"[{format_timestamp(seg.start)}] {text}"
        if total + len(line) > max_chars:
            chunks.append("…")
            break
        chunks.append(line)
        total += len(line)
    return "\n".join(chunks)
