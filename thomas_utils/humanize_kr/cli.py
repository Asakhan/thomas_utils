"""humanize_kr 단독 CLI — `python -m thomas_utils.humanize_kr ...`.

상위 통합 CLI(`thomas-utils humanize-all`)는 `thomas_utils/cli.py` 에 있다.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from thomas_utils.humanize_kr.batch import humanize_batch
from thomas_utils.humanize_kr.processor import HumanizeError, humanize_file


def _progress(idx: int, total: int, path: Path) -> None:
    print(f"[{idx}/{total}] 윤문 중: {path}", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="thomas-utils-humanize",
        description="한국어 문서의 AI 말투(번역투·기계적 병렬·과편집)를 제거합니다.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # batch (= humanize-all)
    p_batch = sub.add_parser("all", help="source/ 폴더의 모든 문서를 일괄 윤문")
    p_batch.add_argument("--source", default="source", help="입력 폴더 (기본: source)")
    p_batch.add_argument("--output", default="output", help="출력 폴더 (기본: output)")
    p_batch.add_argument("--provider", choices=("openai", "anthropic"), default="openai")
    p_batch.add_argument("--model", help="모델 강제 지정")
    p_batch.add_argument("--api-timeout", type=int, default=180)
    p_batch.add_argument("--suffix", default=".humanized", help="출력 파일 접미사 (기본: .humanized)")
    p_batch.add_argument("--halt-change-rate", type=float, default=50.0)
    p_batch.add_argument("--warn-change-rate", type=float, default=30.0)
    p_batch.set_defaults(_run=_run_all)

    # text (직접 텍스트 입력 → 즉시 윤문 결과 출력)
    p_text = sub.add_parser("text", help="텍스트를 직접 입력받아 즉시 윤문 (파일 불필요)")
    p_text.add_argument("text", nargs="?", help="윤문할 텍스트. 생략 시 표준입력에서 읽음")
    p_text.add_argument("--provider", choices=("openai", "anthropic"), default="openai")
    p_text.add_argument("--model")
    p_text.add_argument("--api-timeout", type=int, default=180)
    p_text.add_argument("--halt-change-rate", type=float, default=50.0)
    p_text.add_argument("--warn-change-rate", type=float, default=30.0)
    p_text.add_argument("-q", "--quiet", action="store_true",
                        help="윤문 결과 본문만 출력")
    p_text.set_defaults(_run=_run_text)

    # 단일 파일
    p_one = sub.add_parser("one", help="단일 파일을 윤문")
    p_one.add_argument("input")
    p_one.add_argument("-o", "--output", required=True)
    p_one.add_argument("--report")
    p_one.add_argument("--provider", choices=("openai", "anthropic"), default="openai")
    p_one.add_argument("--model")
    p_one.add_argument("--api-timeout", type=int, default=180)
    p_one.add_argument("--halt-change-rate", type=float, default=50.0)
    p_one.add_argument("--warn-change-rate", type=float, default=30.0)
    p_one.set_defaults(_run=_run_one)

    args = parser.parse_args()
    sys.exit(args._run(args))


def _run_all(args: argparse.Namespace) -> int:
    src = Path(args.source)
    out = Path(args.output)
    log = out / "humanize_log.json"
    try:
        results = humanize_batch(
            source_dir=src,
            output_dir=out,
            provider=args.provider,
            model=args.model,
            api_timeout=args.api_timeout,
            suffix=args.suffix,
            halt_change_rate=args.halt_change_rate,
            warn_change_rate=args.warn_change_rate,
            log_path=log,
            on_progress=_progress,
        )
    except HumanizeError as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1

    ok = sum(1 for r in results if r.success)
    halted = sum(1 for r in results if r.halted)
    print(f"\n=== 완료: {ok}/{len(results)} 성공 (강제 중단 {halted}건) ===")
    print(f"로그: {log}")
    return 0 if results else 1


def _run_text(args: argparse.Namespace) -> int:
    from thomas_utils.humanize_kr.processor import humanize_text

    if args.text:
        text = args.text
    elif not sys.stdin.isatty():
        text = sys.stdin.read()
    else:
        print("오류: 텍스트를 인자로 주거나 표준입력으로 전달하세요.", file=sys.stderr)
        return 1
    if not text.strip():
        print("오류: 입력 텍스트가 비어 있습니다.", file=sys.stderr)
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
        print(f"오류: {e}", file=sys.stderr)
        return 1

    if args.quiet:
        sys.stdout.write(r.rewritten_text)
        if not r.rewritten_text.endswith("\n"):
            sys.stdout.write("\n")
        return 0

    print(r.rewritten_text)
    print(f"\n[등급 {r.before.grade}→{r.after.grade} "
          f"개선 {r.improvement_percent:.1f}%, 변경율 {r.change_rate_percent:.1f}%]",
          file=sys.stderr)
    if r.halted:
        print(f"중단: {r.halt_reason} → 원본을 그대로 출력했습니다.", file=sys.stderr)
    for w in r.warnings:
        print(f"경고: {w}", file=sys.stderr)
    return 0


def _run_one(args: argparse.Namespace) -> int:
    try:
        r = humanize_file(
            args.input,
            args.output,
            provider=args.provider,
            model=args.model,
            api_timeout=args.api_timeout,
            report_path=args.report,
            halt_change_rate=args.halt_change_rate,
            warn_change_rate=args.warn_change_rate,
        )
    except HumanizeError as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1
    print(f"입력 등급 {r.before.grade} → 출력 등급 {r.after.grade} "
          f"(개선 {r.improvement_percent:.1f}%, 변경율 {r.change_rate_percent:.1f}%)")
    if r.halted:
        print(f"중단: {r.halt_reason} → 원본을 그대로 저장했습니다.")
    for w in r.warnings:
        print(f"경고: {w}")
    print(f"출력: {args.output}")
    if args.report:
        print(f"리포트: {args.report}")
    return 0
