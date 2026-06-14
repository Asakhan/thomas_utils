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


def _humanize_all(args: argparse.Namespace) -> int:
    from thomas_utils.humanize_kr.batch import humanize_batch
    from thomas_utils.humanize_kr.processor import HumanizeError

    source_dir = Path(args.source)
    output_dir = Path(args.output)

    if not source_dir.exists():
        print(f"Error: source 폴더가 존재하지 않습니다: {source_dir}", file=sys.stderr)
        print(f"  → 먼저 `mkdir {source_dir}` 후 .md/.txt 파일을 넣어 주세요.", file=sys.stderr)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "humanize_log.json"

    def _progress(idx: int, total: int, path: Path) -> None:
        print(f"[{idx}/{total}] 윤문 중: {path.relative_to(source_dir)}", flush=True)

    try:
        results = humanize_batch(
            source_dir=source_dir,
            output_dir=output_dir,
            provider=args.provider,
            model=args.model,
            api_timeout=args.api_timeout,
            suffix=args.suffix,
            halt_change_rate=args.halt_change_rate,
            warn_change_rate=args.warn_change_rate,
            log_path=log_path,
            on_progress=_progress,
        )
    except HumanizeError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    if not results:
        print(f"처리할 .md/.txt/.markdown/.mdx 파일이 {source_dir} 에 없습니다.")
        return 1

    ok = sum(1 for r in results if r.success)
    halted = sum(1 for r in results if r.halted)
    failed = sum(1 for r in results if not r.success)
    print()
    print(f"=== 일괄 윤문 완료 ===")
    print(f"총 {len(results)}건 / 성공 {ok} / 강제 중단 {halted} / 실패 {failed}")
    for r in results:
        status = "OK" if r.success else "FAIL"
        if r.halted:
            status = "HALT"
        print(f"  [{status}] {r.input_path.name} "
              f"등급 {r.grade_before}→{r.grade_after} "
              f"(개선 {r.improvement_percent:.1f}%, 변경율 {r.change_rate_percent:.1f}%)"
              + (f" — {r.error}" if r.error else ""))
    print(f"\n로그: {log_path}")
    return 0 if failed == 0 else 1


def _humanize_text(args: argparse.Namespace) -> int:
    from thomas_utils.humanize_kr.processor import HumanizeError, humanize_text

    # 입력 텍스트: 인자(TEXT) 우선, 없으면 표준입력(파이프/리다이렉트)에서 읽음
    if args.text:
        text = args.text
    elif not sys.stdin.isatty():
        text = sys.stdin.read()
    else:
        print("Error: 윤문할 텍스트를 인자로 주거나 표준입력으로 전달하세요.", file=sys.stderr)
        print('  예: thomas-utils humanize "윤문할 한국어 문장"', file=sys.stderr)
        print('      echo "..." | thomas-utils humanize', file=sys.stderr)
        return 1

    if not text.strip():
        print("Error: 입력 텍스트가 비어 있습니다.", file=sys.stderr)
        return 1

    try:
        r = humanize_text(
            text,
            provider=args.provider,
            model=args.model,
            api_timeout=args.api_timeout,
            halt_change_rate=args.halt_change_rate,
            warn_change_rate=args.warn_change_rate,
        )
    except HumanizeError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    # --quiet: 윤문 결과 본문만 출력 (파이프라인 친화적)
    if args.quiet:
        sys.stdout.write(r.rewritten_text)
        if not r.rewritten_text.endswith("\n"):
            sys.stdout.write("\n")
        return 0

    print(r.rewritten_text)
    print("\n" + "─" * 40, file=sys.stderr)
    print(
        f"등급 {r.before.grade}→{r.after.grade} "
        f"(개선 {r.improvement_percent:.1f}%, 변경율 {r.change_rate_percent:.1f}%)",
        file=sys.stderr,
    )
    if r.halted:
        print(f"중단: {r.halt_reason} → 원본을 그대로 출력했습니다.", file=sys.stderr)
    for w in r.warnings:
        print(f"경고: {w}", file=sys.stderr)
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="thomas-utils",
        description="PDF / PowerPoint / Video to Markdown, plus Korean humanization.",
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

    humanize_p = subparsers.add_parser(
        "humanize-all",
        help="source/ 폴더의 모든 한국어 문서에서 AI 말투를 제거하고 output/ 에 저장",
    )
    humanize_p.add_argument("--source", default="source",
                            help="입력 폴더 (기본: source)")
    humanize_p.add_argument("--output", default="output",
                            help="출력 폴더 (기본: output)")
    humanize_p.add_argument("--provider", choices=("openai", "anthropic"),
                            default="openai",
                            help="윤문 LLM provider (기본: openai)")
    humanize_p.add_argument("--model",
                            help="모델 강제 지정 (기본: openai=gpt-4o, anthropic=claude-sonnet-4-6)")
    humanize_p.add_argument("--api-timeout", type=int, default=180,
                            help="LLM 호출당 타임아웃(초) (기본: 180)")
    humanize_p.add_argument("--suffix", default=".humanized",
                            help="출력 파일 stem 접미사 (기본: .humanized)")
    humanize_p.add_argument("--halt-change-rate", type=float, default=50.0,
                            help="이 변경율(%%)을 넘으면 원본 유지 (기본: 50.0)")
    humanize_p.add_argument("--warn-change-rate", type=float, default=30.0,
                            help="이 변경율(%%)을 넘으면 경고 (기본: 30.0)")
    humanize_p.set_defaults(_run=_humanize_all)

    humanize_text_p = subparsers.add_parser(
        "humanize",
        help="텍스트를 직접 입력받아 즉시 윤문 결과를 출력 (파일 불필요)",
    )
    humanize_text_p.add_argument(
        "text", nargs="?", metavar="TEXT",
        help="윤문할 한국어 텍스트. 생략하면 표준입력(stdin)에서 읽음",
    )
    humanize_text_p.add_argument("--provider", choices=("openai", "anthropic"),
                                 default="openai",
                                 help="윤문 LLM provider (기본: openai)")
    humanize_text_p.add_argument("--model",
                                 help="모델 강제 지정 (기본: openai=gpt-4o, anthropic=claude-sonnet-4-6)")
    humanize_text_p.add_argument("--api-timeout", type=int, default=180,
                                 help="LLM 호출 타임아웃(초) (기본: 180)")
    humanize_text_p.add_argument("--halt-change-rate", type=float, default=50.0,
                                 help="이 변경율(%%)을 넘으면 원본 유지 (기본: 50.0)")
    humanize_text_p.add_argument("--warn-change-rate", type=float, default=30.0,
                                 help="이 변경율(%%)을 넘으면 경고 (기본: 30.0)")
    humanize_text_p.add_argument("-q", "--quiet", action="store_true",
                                 help="윤문 결과 본문만 출력(등급/경고 등 메타 정보 숨김)")
    humanize_text_p.set_defaults(_run=_humanize_text)

    args = parser.parse_args()
    run = getattr(args, "_run", None)
    if run is None:
        parser.print_help()
        sys.exit(0)
    sys.exit(run(args))
