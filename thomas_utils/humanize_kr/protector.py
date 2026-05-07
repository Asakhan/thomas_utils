"""보호 규칙 (Do-NOT 리스트).

다음에 해당하는 부분은 윤문 단계에서 절대 수정되지 않아야 한다:
- 수치 / 단위 / 날짜 / 시각
- 직접 인용구 (큰따옴표 / 작은따옴표 / 백틱)
- 코드 블록 (```~```) 과 인라인 코드 (`...`)
- URL / 이메일
- 마크다운 이미지 / 링크 텍스트의 URL 부분
- 영문 고유명사 후보 (대문자 시작 영단어 연쇄)
- 인용/각주 표기 [1], [Smith 2020] 등

`extract_protected_spans` 는 (start, end, kind, text) 튜플 리스트를 반환한다.
`verify_preserved` 는 보호 스팬의 텍스트가 출력에 모두 보존되었는지 검증한다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Tuple


@dataclass(frozen=True)
class ProtectedSpan:
    start: int
    end: int
    kind: str   # "number" | "date" | "quote" | "code" | "url" | "email" | "proper_noun" | "citation"
    text: str


# 정규식 모음 (우선순위 — 코드/URL 등이 먼저 매칭되어야 따옴표 규칙과 충돌하지 않음)
_PATTERNS: List[Tuple[str, re.Pattern[str]]] = [
    # 코드 블록은 가장 먼저 (다른 규칙이 내용을 잘못 잡아내지 않게)
    ("code", re.compile(r"```[\s\S]*?```", re.MULTILINE)),
    ("code", re.compile(r"`[^`\n]+`")),
    # URL / 이메일
    ("url", re.compile(r"https?://[^\s)\]]+")),
    ("email", re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")),
    # 마크다운 이미지 / 링크의 URL 부분
    ("url", re.compile(r"!\[[^\]]*\]\([^)]+\)")),
    # 인용 / 각주
    ("citation", re.compile(r"\[(?:\d+|[A-Z][A-Za-z]+\s*\d{4})\]")),
    # 직접 인용구 (한국어 큰따옴표, ASCII 따옴표, 단·복수, 길이 제한 없음)
    ("quote", re.compile(r"\"[^\"\n]{1,200}\"")),
    ("quote", re.compile(r"“[^”\n]{1,200}”")),
    ("quote", re.compile(r"'[^'\n]{1,200}'")),
    ("quote", re.compile(r"‘[^’\n]{1,200}’")),
    # 날짜 / 시각 — 2026-05-07, 2026/5/7, 2026년 5월 7일, 14:30
    ("date", re.compile(r"\b\d{4}[-./]\d{1,2}[-./]\d{1,2}\b")),
    ("date", re.compile(r"\d{4}\s*년\s*\d{1,2}\s*월(?:\s*\d{1,2}\s*일)?")),
    ("date", re.compile(r"\b\d{1,2}\s*월\s*\d{1,2}\s*일\b")),
    ("date", re.compile(r"\b\d{1,2}:\d{2}(?::\d{2})?\b")),
    # 수치 + 선택적 단위 — 12.5%, 3,400원, 1억 2천, 5kg, 100MHz
    ("number", re.compile(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?(?:\s*(?:%|개|원|달러|kg|km|m|cm|mm|ms|s|MB|GB|TB|MHz|GHz|°C|°F))?")),
    ("number", re.compile(r"\b\d+(?:\.\d+)?(?:\s*(?:%|개|원|달러|kg|km|m|cm|mm|ms|s|MB|GB|TB|MHz|GHz|°C|°F))")),
    ("number", re.compile(r"\b\d+(?:\.\d+)?\b")),
    # 영문 고유명사 후보 (2~5 단어 대문자 시작 시퀀스)
    ("proper_noun", re.compile(r"\b[A-Z][A-Za-z0-9]+(?:\s+[A-Z][A-Za-z0-9]+){0,4}\b")),
]


def extract_protected_spans(text: str) -> List[ProtectedSpan]:
    """겹치지 않는 최장 보호 스팬 목록을 반환한다."""
    spans: List[ProtectedSpan] = []
    occupied = [False] * (len(text) + 1)

    for kind, regex in _PATTERNS:
        for m in regex.finditer(text):
            s, e = m.start(), m.end()
            if any(occupied[s:e]):
                continue
            spans.append(ProtectedSpan(s, e, kind, text[s:e]))
            for i in range(s, e):
                occupied[i] = True

    spans.sort(key=lambda x: x.start)
    return spans


def collect_protected_strings(text: str) -> List[str]:
    """검증용 — 중복 제거된 보호 텍스트 리스트(긴 것 우선)."""
    seen: dict[str, None] = {}
    for sp in extract_protected_spans(text):
        seen.setdefault(sp.text, None)
    return sorted(seen.keys(), key=len, reverse=True)


def verify_preserved(original: str, rewritten: str) -> Tuple[bool, List[str]]:
    """보호 스팬이 윤문본에 모두 그대로 들어 있는지 확인.

    수치·고유명사·인용구는 글자 그대로 보존되어야 한다. 누락된 항목은 리스트로 반환.
    매우 짧은 단편(1~2자 숫자, 일반어와 충돌 위험)은 건너뛴다.
    """
    missing: List[str] = []
    for s in collect_protected_strings(original):
        # 너무 짧은 일반 숫자(예: 단일 자리)와 단일 영문 첫문자는 검증에서 제외
        if len(s) < 2:
            continue
        if s not in rewritten:
            missing.append(s)
    return (len(missing) == 0, missing)
