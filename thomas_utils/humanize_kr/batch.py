"""`source/` 폴더 일괄 윤문 처리.

지원 확장자: .md, .txt, .markdown, .mdx
출력 파일명: `<stem>.humanized.md` 와 `<stem>.report.json`
요약 로그: `output/humanize_log.json`
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional

from thomas_utils.humanize_kr.processor import (
    HumanizeError,
    HumanizeResult,
    humanize_file,
)


SUPPORTED_SUFFIXES = {".md", ".markdown", ".mdx", ".txt"}


@dataclass
class BatchItemResult:
    input_path: Path
    output_path: Optional[Path]
    report_path: Optional[Path]
    success: bool
    error: str = ""
    grade_before: str = ""
    grade_after: str = ""
    improvement_percent: float = 0.0
    change_rate_percent: float = 0.0
    halted: bool = False

    def to_dict(self) -> dict:
        return {
            "input": str(self.input_path),
            "output": str(self.output_path) if self.output_path else None,
            "report": str(self.report_path) if self.report_path else None,
            "success": self.success,
            "error": self.error,
            "grade_before": self.grade_before,
            "grade_after": self.grade_after,
            "improvement_percent": round(self.improvement_percent, 2),
            "change_rate_percent": round(self.change_rate_percent, 2),
            "halted": self.halted,
        }


def _iter_source_files(source_dir: Path) -> List[Path]:
    if not source_dir.exists():
        return []
    files: List[Path] = []
    for p in sorted(source_dir.rglob("*")):
        if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES:
            files.append(p)
    return files


def humanize_batch(
    source_dir: Path,
    output_dir: Path,
    *,
    provider: str = "openai",
    model: Optional[str] = None,
    api_timeout: int = 180,
    suffix: str = ".humanized",
    halt_change_rate: float = 50.0,
    warn_change_rate: float = 30.0,
    log_path: Optional[Path] = None,
    on_progress=None,
) -> List[BatchItemResult]:
    source_dir = Path(source_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    files = _iter_source_files(source_dir)
    results: List[BatchItemResult] = []

    for idx, in_path in enumerate(files, start=1):
        rel = in_path.relative_to(source_dir)
        stem_with_suffix = rel.stem + suffix
        out_path = output_dir / rel.with_name(stem_with_suffix + rel.suffix).name
        # nested 디렉토리는 구조 보존
        out_path = output_dir / rel.with_name(stem_with_suffix + rel.suffix)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        report_path = out_path.with_suffix(".report.json")

        if on_progress:
            on_progress(idx, len(files), in_path)

        try:
            r: HumanizeResult = humanize_file(
                in_path,
                out_path,
                provider=provider,
                model=model,
                api_timeout=api_timeout,
                report_path=report_path,
                halt_change_rate=halt_change_rate,
                warn_change_rate=warn_change_rate,
            )
            results.append(BatchItemResult(
                input_path=in_path,
                output_path=out_path,
                report_path=report_path,
                success=True,
                grade_before=r.before.grade,
                grade_after=r.after.grade,
                improvement_percent=r.improvement_percent,
                change_rate_percent=r.change_rate_percent,
                halted=r.halted,
            ))
        except HumanizeError as e:
            results.append(BatchItemResult(
                input_path=in_path,
                output_path=None,
                report_path=None,
                success=False,
                error=str(e),
            ))

    if log_path is not None:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_path.write_text(
            json.dumps(
                {
                    "generated_at": datetime.now(timezone.utc).isoformat(),
                    "source_dir": str(source_dir),
                    "output_dir": str(output_dir),
                    "provider": provider,
                    "model": model,
                    "items": [r.to_dict() for r in results],
                    "summary": {
                        "total": len(results),
                        "success": sum(1 for r in results if r.success),
                        "halted": sum(1 for r in results if r.halted),
                        "failed": sum(1 for r in results if not r.success),
                    },
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    return results
