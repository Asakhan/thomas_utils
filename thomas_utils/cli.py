"""CLI for thomas_utils: PDF and PowerPoint -> Markdown."""

import argparse
import sys
from pathlib import Path


def _parse_pages(s: str) -> list[int]:
    """Parse --pages '0,1,2' or '0-2' into 0-based indices."""
    out: list[int] = []
    for part in s.replace(" ", "").split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return sorted(set(out))


def _pdf2md(args: argparse.Namespace) -> int:
    from thomas_utils.converters import convert

    pdf = Path(args.input)
    if not pdf.exists():
        print(f"Error: file not found: {pdf}", file=sys.stderr)
        return 1
    if not pdf.suffix.lower() == ".pdf":
        print(f"Error: expected .pdf file, got: {pdf}", file=sys.stderr)
        return 1

    out_path = Path(args.output) if args.output else Path("output") / (pdf.stem + ".md")
    pages = _parse_pages(args.pages) if args.pages else None

    try:
        md = convert(str(pdf), pages=pages, engine=args.engine)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ImportError as e:
        if "marker" in str(e).lower() or "marker" in str(args.engine).lower():
            print("Error: marker engine requires 'pip install thomas-utils[marker]'", file=sys.stderr)
        else:
            print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(md, encoding="utf-8")
    print(f"Wrote {out_path}")
    return 0


def _pptx2md(args: argparse.Namespace) -> int:
    from thomas_utils.converters import convert_pptx

    pptx = Path(args.input)
    if not pptx.exists():
        print(f"Error: file not found: {pptx}", file=sys.stderr)
        return 1
    if not pptx.suffix.lower() == ".pptx":
        print(f"Error: expected .pptx file, got: {pptx}", file=sys.stderr)
        return 1

    # 마크다운은 항상 output/ 폴더에 저장
    out_path = Path("output") / (Path(args.output).name if args.output else (pptx.stem + ".md"))

    try:
        md = convert_pptx(
            str(pptx),
            use_llm=getattr(args, "pptx_use_llm", False),
            engine=getattr(args, "pptx_engine", "python-pptx"),
            use_llm_multimodal=getattr(args, "pptx_use_llm_multimodal", False),
        )
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(md, encoding="utf-8")
    print(f"Wrote {out_path}")
    return 0


def _video2md(args: argparse.Namespace) -> int:
    from thomas_utils.video_summary.processor import VideoSummaryError, convert_video

    video = Path(args.input)
    if not video.exists():
        print(f"Error: file not found: {video}", file=sys.stderr)
        return 1

    out_path = Path(args.output) if args.output else Path("output") / (video.stem + ".md")
    try:
        convert_video(
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
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    print(f"Wrote {out_path}")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="thomas-utils",
        description="PDF / PowerPoint / Video to Markdown.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    pdf2md_p = subparsers.add_parser("pdf2md", help="Convert PDF to Markdown")
    pdf2md_p.add_argument("input", metavar="INPUT.pdf", help="Input PDF path")
    pdf2md_p.add_argument("-o", "--output", metavar="OUTPUT.md", help="Output Markdown path (default: output/INPUT.md)")
    pdf2md_p.add_argument(
        "--pages",
        metavar="LIST",
        help="0-based page indices, e.g. 0,1,2 or 0-5 (default: all)",
    )
    pdf2md_p.add_argument(
        "--engine",
        choices=("pymupdf", "marker"),
        default="pymupdf",
        help="Conversion engine (default: pymupdf)",
    )
    pdf2md_p.set_defaults(_run=_pdf2md)

    pptx2md_p = subparsers.add_parser("pptx2md", help="Convert PowerPoint to Markdown")
    pptx2md_p.add_argument("input", metavar="INPUT.pptx", help="Input PPTX path")
    pptx2md_p.add_argument("-o", "--output", metavar="OUTPUT.md", help="Output Markdown path (default: output/INPUT.md)")
    pptx2md_p.add_argument(
        "--slides",
        metavar="LIST",
        help="0-based slide indices (currently ignored, all slides are converted)",
    )
    pptx2md_p.add_argument(
        "--pptx-use-llm",
        action="store_true",
        help="Use LLM to polish extracted markdown (requires pptx-llm extra)",
    )
    pptx2md_p.add_argument(
        "--engine",
        choices=("python-pptx", "unstructured"),
        default="python-pptx",
        dest="pptx_engine",
        help="PPTX conversion engine (default: python-pptx)",
    )
    pptx2md_p.add_argument(
        "--pptx-use-llm-multimodal",
        action="store_true",
        help="Render each slide to image and convert via vision LLM (GPT-4o); needs pywin32 (Windows) or LibreOffice + pymupdf",
    )
    pptx2md_p.set_defaults(_run=_pptx2md)

    video2md_p = subparsers.add_parser("video2md", help="Convert lecture video to Markdown notes")
    video2md_p.add_argument("input", metavar="INPUT.mp4", help="Input video path")
    video2md_p.add_argument("-o", "--output", metavar="OUTPUT.md",
                            help="Output Markdown path (default: output/INPUT.md)")
    video2md_p.add_argument("--provider", choices=("anthropic", "openai"), default="openai",
                            help="Multimodal LLM provider (default: openai)")
    video2md_p.add_argument("--model", help="Override LLM model name")
    video2md_p.add_argument("--whisper-model", default="base",
                            help="Whisper model size (default: base)")
    video2md_p.add_argument("--language", help="STT language code (default: auto)")
    video2md_p.add_argument("--scene-threshold", type=float, default=0.55,
                            help="0.0–1.0; higher = fewer cuts (default: 0.55)")
    video2md_p.add_argument("--min-gap-seconds", type=float, default=8.0,
                            help="Minimum gap between scenes (default: 8.0)")
    video2md_p.add_argument("--max-scenes", type=int, default=40,
                            help="Hard cap on detected scenes (default: 40)")
    video2md_p.add_argument("--api-timeout", type=int, default=120,
                            help="Per-LLM-call timeout in seconds (default: 120)")
    video2md_p.add_argument("--audio-timeout", type=int, default=1800,
                            help="ffmpeg audio extract timeout in seconds (default: 1800)")
    video2md_p.add_argument("--screenshots-dir", help="Override directory for keyframe images")
    video2md_p.add_argument("--title", help="Override the lecture title in the output")
    video2md_p.set_defaults(_run=_video2md)

    args = parser.parse_args()
    run = getattr(args, "_run", None)
    if run is None:
        parser.print_help()
        sys.exit(0)
    sys.exit(run(args))
