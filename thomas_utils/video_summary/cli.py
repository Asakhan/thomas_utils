"""CLI for `python -m thomas_utils.video_summary`."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from thomas_utils.video_summary.processor import VideoSummaryError, convert_video


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python -m thomas_utils.video_summary",
        description="Convert a lecture video into a Markdown lecture-note (audio + visual + LLM).",
    )
    p.add_argument("--input", "-i", required=True, metavar="VIDEO", help="Input video path")
    p.add_argument(
        "--output", "-o",
        metavar="OUT.md",
        help="Output Markdown path (default: output/<INPUT_STEM>.md)",
    )
    p.add_argument(
        "--provider",
        choices=("anthropic", "openai"),
        default="anthropic",
        help="Multimodal LLM provider (default: anthropic)",
    )
    p.add_argument("--model", help="Override LLM model name")
    p.add_argument(
        "--whisper-model",
        default="base",
        help="Whisper model: tiny / base / small / medium / large-v3 (default: base)",
    )
    p.add_argument("--language", help="STT language code (e.g. ko, en). Default: auto-detect")
    p.add_argument(
        "--scene-threshold",
        type=float, default=0.55,
        help="0.0–1.0; higher = fewer scene cuts (default: 0.55)",
    )
    p.add_argument(
        "--min-gap-seconds", type=float, default=8.0,
        help="Minimum gap between detected scenes (default: 8.0)",
    )
    p.add_argument(
        "--max-scenes", type=int, default=40,
        help="Hard cap on detected scenes to bound API cost (default: 40)",
    )
    p.add_argument("--api-timeout", type=int, default=120, help="Per-LLM-call timeout (s)")
    p.add_argument("--audio-timeout", type=int, default=1800, help="ffmpeg audio extract timeout (s)")
    p.add_argument("--screenshots-dir", help="Override directory for keyframe images")
    p.add_argument("--title", help="Override the lecture title in the output")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    video = Path(args.input)
    if not video.exists():
        print(f"Error: video not found: {video}", file=sys.stderr)
        return 1

    out_path = Path(args.output) if args.output else Path("output") / (video.stem + ".md")

    try:
        written = convert_video(
            video_path=video,
            output_path=out_path,
            provider=args.provider,
            model=args.model,
            whisper_model=args.whisper_model,
            language=args.language,
            scene_threshold=args.scene_threshold,
            min_gap_seconds=args.min_gap_seconds,
            max_scenes=args.max_scenes,
            api_timeout=args.api_timeout,
            audio_timeout=args.audio_timeout,
            screenshots_dir=args.screenshots_dir,
            title=args.title,
        )
    except VideoSummaryError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return 130

    print(f"Wrote {written}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
