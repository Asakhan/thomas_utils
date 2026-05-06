"""Tests for video_summary (formatter + CLI parsing).

Heavyweight pieces (Whisper, OpenCV, ffmpeg, real LLM calls) require external
binaries / models / API keys, so they are exercised through the formatter
plus argparse-only paths here.
"""

from pathlib import Path

import pytest


def test_format_timestamp_under_hour():
    from thomas_utils.video_summary.formatter import format_timestamp

    assert format_timestamp(0) == "00:00"
    assert format_timestamp(65) == "01:05"
    assert format_timestamp(599) == "09:59"


def test_format_timestamp_over_hour():
    from thomas_utils.video_summary.formatter import format_timestamp

    assert format_timestamp(3600) == "01:00:00"
    assert format_timestamp(3725) == "01:02:05"


def test_slugify_keeps_hangul_and_words():
    from thomas_utils.video_summary.formatter import slugify

    assert slugify("Intro to ML") == "intro-to-ml"
    # 한글은 유지되어야 함
    assert "강의" in slugify("강의 소개")
    assert slugify("") == "section"
    assert slugify("!!!") == "section"


def test_render_lecture_notes_structure(tmp_path: Path):
    from thomas_utils.video_summary.formatter import (
        LectureMetadata,
        Section,
        TranscriptSegment,
        render_lecture_notes,
    )

    out_path = tmp_path / "out.md"
    screenshot = tmp_path / "out_assets" / "scene_001.jpg"
    screenshot.parent.mkdir(parents=True, exist_ok=True)
    screenshot.write_bytes(b"fake-png")

    meta = LectureMetadata(
        title="테스트 강의",
        source_filename="lecture.mp4",
        duration_seconds=125.0,
        section_count=1,
        language="ko",
        model="anthropic:claude-sonnet-4-6",
    )
    sections = [Section(
        index=1,
        start=0.0,
        end=125.0,
        title="강의 소개",
        summary="이 섹션은 강의 전반을 소개합니다.",
        bullets=["배경", "목표"],
        screenshot_path=screenshot,
        transcript_excerpt="[00:00] 안녕하세요",
    )]
    transcript = [TranscriptSegment(start=0.0, end=5.0, text="안녕하세요")]

    md = render_lecture_notes(meta, sections, transcript, out_path)

    assert md.startswith("# 테스트 강의")
    assert "## 목차" in md
    assert "[00:00 — 강의 소개]" in md
    assert "강의 소개 (00:00 — 02:05)" in md
    assert "out_assets/scene_001.jpg" in md  # relative path resolution
    assert "## 전체 스크립트" in md
    assert "`[00:00]` 안녕하세요" in md


def test_transcript_excerpt_truncates_to_max_chars():
    from thomas_utils.video_summary.formatter import (
        TranscriptSegment,
        transcript_excerpt_for_range,
    )

    segs = [TranscriptSegment(start=i, end=i + 1, text="x" * 200) for i in range(50)]
    out = transcript_excerpt_for_range(segs, start=0, end=50, max_chars=300)
    assert out.endswith("…")
    assert len(out) <= 400  # rough cap including timestamp prefixes


def test_cli_parser_defaults():
    from thomas_utils.video_summary.cli import build_parser

    p = build_parser()
    args = p.parse_args(["--input", "video.mp4"])
    assert args.input == "video.mp4"
    assert args.provider == "anthropic"
    assert args.whisper_model == "base"
    assert args.scene_threshold == pytest.approx(0.55)
    assert args.max_scenes == 40


def test_cli_parser_overrides():
    from thomas_utils.video_summary.cli import build_parser

    p = build_parser()
    args = p.parse_args([
        "-i", "v.mp4", "-o", "out.md",
        "--provider", "openai", "--model", "gpt-4o",
        "--whisper-model", "small", "--language", "ko",
        "--scene-threshold", "0.7", "--max-scenes", "10",
    ])
    assert args.provider == "openai"
    assert args.model == "gpt-4o"
    assert args.whisper_model == "small"
    assert args.language == "ko"
    assert args.scene_threshold == pytest.approx(0.7)
    assert args.max_scenes == 10


def test_cli_main_returns_error_for_missing_video(tmp_path: Path, capsys: pytest.CaptureFixture):
    from thomas_utils.video_summary.cli import main

    code = main(["--input", str(tmp_path / "nope.mp4")])
    assert code == 1
    err = capsys.readouterr().err
    assert "not found" in err.lower()


def test_main_cli_video2md_subcommand_registered():
    """thomas-utils CLI should expose `video2md` alongside pdf2md / pptx2md."""
    from thomas_utils import cli as main_cli

    # Ensure the helper is exported and callable
    assert callable(getattr(main_cli, "_video2md"))
