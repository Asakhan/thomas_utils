"""AI 티 탐지기.

`thomas_utils.humanize_kr.patterns.REGEX_PATTERNS` 를 사용해 텍스트를 스캔하고
패턴별 발견 횟수와 위치를 보고한다. 또 문장 길이 분산·종결어미 반복 같은
LLM-only 카테고리(E-1, E-2)는 통계 기반으로 보조 측정한다.

탐지 결과는 `DetectionReport` 데이터클래스로 직렬화 가능한 형태(`to_dict`)로 반환한다.
"""

from __future__ import annotations

import re
import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from thomas_utils.humanize_kr.patterns import (
    CATEGORY_NAMES,
    Detection,
    REGEX_PATTERNS,
)


_SENTENCE_SPLIT = re.compile(r"(?<=[.!?。…])\s+|(?<=[다요죠])\s+(?=[가-힣A-Z])")
_ENDING_TAIL = re.compile(r"(이다|입니다|한다|합니다|것이다|된다|됩니다)[.\s]*$")


@dataclass
class DetectionReport:
    detections: List[Detection]
    severity_counts: Dict[str, int]   # {"S1": 3, "S2": 7, "S3": 0}
    category_counts: Dict[str, int]   # {"A": 5, "C": 2, ...}
    pattern_counts: Dict[str, int]    # {"A-1": 2, ...}
    sentence_count: int
    sentence_length_stdev: float
    ending_repetition_top: List[Tuple[str, int]]
    grade: str                        # "A" | "B" | "C" | "D"
    grade_reason: str

    def to_dict(self) -> Dict:
        return {
            "grade": self.grade,
            "grade_reason": self.grade_reason,
            "severity_counts": self.severity_counts,
            "category_counts": {
                f"{k} ({CATEGORY_NAMES.get(k, k)})": v
                for k, v in self.category_counts.items()
            },
            "pattern_counts": self.pattern_counts,
            "sentence_count": self.sentence_count,
            "sentence_length_stdev": round(self.sentence_length_stdev, 2),
            "ending_repetition_top": self.ending_repetition_top,
            "detections": [
                {
                    "code": d.code,
                    "name": d.name,
                    "severity": d.severity,
                    "category": d.category,
                    "span": list(d.span),
                    "snippet": d.snippet,
                    "advice": d.advice,
                }
                for d in self.detections
            ],
        }


def _split_sentences(text: str) -> List[str]:
    raw = _SENTENCE_SPLIT.split(text)
    return [s.strip() for s in raw if s and len(s.strip()) >= 2]


def _ending_repetitions(sentences: List[str]) -> List[Tuple[str, int]]:
    counter: Counter[str] = Counter()
    for s in sentences:
        m = _ENDING_TAIL.search(s)
        if m:
            counter[m.group(1)] += 1
    return counter.most_common(5)


def detect(text: str) -> DetectionReport:
    """텍스트에 대한 AI 티 패턴을 탐지하고 등급을 매긴다."""
    detections: List[Detection] = []
    pattern_counts: Counter[str] = Counter()
    severity_counts: Counter[str] = Counter()
    category_counts: Counter[str] = Counter()

    for p in REGEX_PATTERNS:
        if p.regex is None:
            continue
        for m in p.regex.finditer(text):
            snippet = text[max(0, m.start() - 20):min(len(text), m.end() + 20)]
            detections.append(Detection(
                code=p.code, name=p.name, severity=p.severity, category=p.category,
                span=(m.start(), m.end()), snippet=snippet, advice=p.advice,
            ))
            pattern_counts[p.code] += 1
            severity_counts[p.severity] += 1
            category_counts[p.category] += 1

    sentences = _split_sentences(text)
    if len(sentences) >= 2:
        lengths = [len(s) for s in sentences]
        stdev = statistics.pstdev(lengths)
    else:
        stdev = 0.0

    endings = _ending_repetitions(sentences)
    if sentences and endings and endings[0][1] >= max(4, len(sentences) // 3):
        # E-2 보조 탐지
        detections.append(Detection(
            code="E-2", name="동일 종결어미 반복", severity="S2", category="E",
            span=(0, 0), snippet=f"종결어미 '{endings[0][0]}' {endings[0][1]}회",
            advice="과거형/명사형/연결어미로 변주.",
        ))
        pattern_counts["E-2"] += 1
        severity_counts["S2"] += 1
        category_counts["E"] += 1

    if len(sentences) >= 5 and stdev < 8.0:
        # E-1 보조 탐지
        detections.append(Detection(
            code="E-1", name="문장 길이 저분산", severity="S2", category="E",
            span=(0, 0), snippet=f"문장 길이 표준편차 {stdev:.2f}",
            advice="짧은(10~15자)·긴(80자+) 문장을 의도적으로 섞는다.",
        ))
        pattern_counts["E-1"] += 1
        severity_counts["S2"] += 1
        category_counts["E"] += 1

    grade, grade_reason = _grade(severity_counts)

    return DetectionReport(
        detections=detections,
        severity_counts=dict(severity_counts),
        category_counts=dict(category_counts),
        pattern_counts=dict(pattern_counts),
        sentence_count=len(sentences),
        sentence_length_stdev=stdev,
        ending_repetition_top=endings,
        grade=grade,
        grade_reason=grade_reason,
    )


def _grade(severity_counts: Counter[str]) -> Tuple[str, str]:
    s1 = severity_counts.get("S1", 0)
    s2 = severity_counts.get("S2", 0)
    if s1 == 0 and s2 <= 2:
        return "A", f"S1=0, S2={s2}"
    if s1 == 0 and s2 <= 4:
        return "B", f"S1=0, S2={s2}"
    if s1 <= 2:
        return "C", f"S1={s1}, S2={s2} — 1차 윤문 후 재검토 필요"
    return "D", f"S1={s1}, S2={s2} — 사람의 검토 권장"


def improvement_pct(before: DetectionReport, after: DetectionReport) -> float:
    """심각도 가중치를 적용한 개선율(%) — S1 가중치 3, S2 가중치 1."""
    def score(r: DetectionReport) -> int:
        return r.severity_counts.get("S1", 0) * 3 + r.severity_counts.get("S2", 0)
    b, a = score(before), score(after)
    if b == 0:
        return 0.0
    return max(0.0, (b - a) / b * 100.0)
