"""윤문 오케스트레이션 — 탐지 → 보호스팬 추출 → LLM 윤문 → 검증 → 리포트.

워크플로:
  1. `detect()` 로 입력 본문에서 AI 티 스팬 식별
  2. `extract_protected_spans()` 로 수정 금지 영역 식별
  3. LLM 에 시스템 프롬프트(do-not 리스트, 4대 원칙) + 사용자 프롬프트(원문, 탐지 리포트, 보호 어구) 전달
  4. 응답을 받아 보호 스팬이 모두 보존됐는지 검증
  5. 출력본에 대해 다시 detect() 를 돌려 개선율 계산
  6. 변경율(문자 단위) 50% 초과면 강제 중단(원문 유지), 30% 초과면 경고
  7. 결과를 `HumanizeResult` 로 반환

LLM provider 는 OpenAI(gpt-4o, 기본) 와 Anthropic(claude-sonnet-4-6) 둘 다 지원.
"""

from __future__ import annotations

import json
import os
import random
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from thomas_utils.humanize_kr.detector import (
    DetectionReport,
    detect,
    improvement_pct,
)
from thomas_utils.humanize_kr.protector import (
    collect_protected_strings,
    extract_protected_spans,
    verify_preserved,
)


class HumanizeError(RuntimeError):
    """윤문 파이프라인의 최상위 오류."""


# ---------------------------------------------------------------------------
# 결과 데이터클래스
# ---------------------------------------------------------------------------

@dataclass
class HumanizeResult:
    input_path: Optional[Path]
    output_path: Optional[Path]
    original_text: str
    rewritten_text: str
    before: DetectionReport
    after: DetectionReport
    improvement_percent: float
    change_rate_percent: float
    protected_missing: List[str]
    halted: bool
    halt_reason: str = ""
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict:
        return {
            "input_path": str(self.input_path) if self.input_path else None,
            "output_path": str(self.output_path) if self.output_path else None,
            "halted": self.halted,
            "halt_reason": self.halt_reason,
            "warnings": self.warnings,
            "improvement_percent": round(self.improvement_percent, 2),
            "change_rate_percent": round(self.change_rate_percent, 2),
            "protected_missing": self.protected_missing,
            "before": self.before.to_dict(),
            "after": self.after.to_dict(),
        }


# ---------------------------------------------------------------------------
# 시스템 프롬프트 — 4대 원칙 + Do-NOT 리스트
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
당신은 한국어 글의 'AI 티(AI-tell)'를 제거하는 윤문 전문가입니다.

[4대 원칙]
1. 의미 불변(Meaning Invariant): 사실·주장·수치·고유명사·직접 인용은 100% 보존합니다. 단 한 글자도 바꾸지 마세요.
2. 증거 기반 편집(Evidence-Based): 탐지된 패턴 스팬에만 외과적으로 손대고, 그 외 문장은 그대로 둡니다.
3. 장르 보존(Genre Preservation): 칼럼은 칼럼답게, 보고서는 보고서답게. 스타일을 바꾸지 않습니다.
4. 과편집 방지(Anti-Over-Editing): 변경율 30% 경고, 50% 초과 금지.

[Do-NOT 리스트 — 절대 수정 금지]
- 모든 수치(정수·소수·퍼센트·금액·단위)와 날짜·시각
- 따옴표 안의 직접 인용구(", ", ', ', “, ”, ‘, ’ 모두)
- 코드 블록(```...```)과 인라인 코드(`...`)
- URL · 이메일 주소 · 마크다운 이미지/링크의 URL
- 영문 고유명사(제품명·인명·기관명 등 대문자 시작 영단어)
- 인용·각주 표기 [1], [Smith 2020]

[제거 대상 AI 패턴 핵심]
A. 번역투: '~에 대해', '~를 통해', '~에 있어서', '~라는 점에서', '가지고 있다', 이중피동 '~되어진다', '~에 의해'(피동), '만들어지다/이루어지다', '그/그녀' 직역
B. 영문 인용 과다: '한글(English)' 반복 병기, 영어 단어 그대로 사용
C. 구조 패턴: 기계적 '첫째/둘째/셋째', 이모지 남발, '이 섹션에서는~다룬다' 안내문, '먼저·반면·결국' 3단 공식, 콜론 부제 헤딩, 연결어미 직후 쉼표('~지만,' '~고,'), 문서당 쉼표 50%+
D. AI 시그니처: '결론적으로', '시사하는 바가 크다', '주목할 만하다', '매우 중요하다', '압도적/폭발적/파격적/혁신적'
E. 리듬 균일: 모든 문장 30~50자, 동일 종결어미('~이다.' 4회+), 균일한 문단 길이, 경어체 혼재
F. 수식 과다: '매우/정말/진짜로', 동의어 중복('중요하고 핵심적인'), '-적/-성/-화/-tion' 남용
G. 헤지 과다: '~할 수 있을 것으로 보인다', '~가능성이 있을 수 있다', '양쪽 모두/균형/신중하게'
H. 접속사 군집: 문단 첫머리 '또한/따라서/즉/나아가', '하지만+그러나' 혼용, '이는/이 점에서' 메타 진입어
I. 형식 명사 과다: '~한 것이다', '주목할 점은', '~할 필요가 있다', '~이/가 필요하다'
J. 시각 장식 과다: 산문에 **볼드** 남발, 따옴표 강조 5회+, 엠대시(—) 남발

[작업 지침]
- 탐지된 스팬만 외과적으로 수정합니다. 단정·연결어 다양화·종결어미 변주·문장 길이 변화로 자연스럽게 만듭니다.
- 마크다운 구조(헤딩 단계, 리스트, 표, 이미지 참조, 코드 블록, 링크)는 그대로 유지합니다.
- 보호 항목(수치, 고유명사, 직접 인용, URL, 코드)은 스펠링 한 글자도 바꾸지 마세요.
- 추가 정보·새 사실·새 평가어를 만들어 넣지 마세요. 빠진 정보를 채우지도 마세요.
- 단순 요약·문장 삭제로 의미가 흐려지지 않게 합니다. 의심스러우면 원문 그대로 둡니다.

[출력 형식]
다음 JSON 만 반환하세요(코드펜스/추가 텍스트 금지):
{"rewritten": "<윤문된 전체 본문(마크다운)>"}
"""


# ---------------------------------------------------------------------------
# LLM 호출
# ---------------------------------------------------------------------------

def _user_prompt(text: str, report: DetectionReport, protected: List[str]) -> str:
    # 프롬프트 비대화 방지 — 너무 많은 보호 어구는 상위 50개만 (긴 것 우선)
    top_protected = protected[:50]
    detection_summary = []
    for d in report.detections[:60]:
        detection_summary.append(
            f"- [{d.code} {d.severity}] {d.name}: …{d.snippet}…"
        )
    detection_block = "\n".join(detection_summary) if detection_summary else "(자동 탐지된 항목 없음 — 그래도 4대 원칙대로 가벼운 윤문)"

    return (
        "다음 한국어 본문을 4대 원칙·Do-NOT 리스트에 맞춰 윤문하세요.\n\n"
        f"[자동 탐지 리포트 — 상위 {len(detection_summary)}건]\n"
        f"{detection_block}\n\n"
        "[보호 어구(절대 글자 변경 금지) — 상위]\n"
        + ("\n".join(f"- {s}" for s in top_protected) if top_protected else "(없음)")
        + "\n\n[원문 시작]\n" + text + "\n[원문 끝]\n\n"
        '응답: {"rewritten": "..."} JSON 만.'
    )


def _call_openai(system: str, user: str, model: str, timeout: int) -> str:
    try:
        import openai  # type: ignore
    except ModuleNotFoundError as e:
        raise HumanizeError(
            "openai SDK 가 설치되어 있지 않습니다. `pip install openai` 또는 "
            "`pip install \"thomas-utils[humanize]\"` 로 설치하세요."
        ) from e
    client = openai.OpenAI(timeout=timeout)
    r = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        response_format={"type": "json_object"},
        temperature=0.3,
        max_tokens=16000,
    )
    return (r.choices[0].message.content or "") if r.choices else ""


def _call_anthropic(system: str, user: str, model: str, timeout: int) -> str:
    try:
        import anthropic  # type: ignore
    except ModuleNotFoundError as e:
        raise HumanizeError(
            "anthropic SDK 가 설치되어 있지 않습니다. `pip install anthropic` 또는 "
            "`pip install \"thomas-utils[humanize]\"` 로 설치하세요."
        ) from e
    client = anthropic.Anthropic(timeout=timeout)
    msg = client.messages.create(
        model=model,
        max_tokens=8192,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    parts: List[str] = []
    for block in msg.content or []:
        if getattr(block, "type", None) == "text":
            parts.append(block.text)
    return "".join(parts)


def _call_with_retry(
    provider: str, model: str, system: str, user: str, timeout: int, retries: int = 4
) -> str:
    last_err: Optional[Exception] = None
    for attempt in range(retries):
        try:
            if provider == "openai":
                return _call_openai(system, user, model, timeout)
            if provider == "anthropic":
                return _call_anthropic(system, user, model, timeout)
            raise HumanizeError(f"Unknown provider: {provider}")
        except HumanizeError:
            raise
        except Exception as e:
            last_err = e
            msg = str(e).lower()
            transient = any(t in msg for t in ("rate", "timeout", "429", "503", "overloaded", "temporar"))
            if attempt == retries - 1 or not transient:
                break
            time.sleep((2 ** attempt) + random.uniform(0, 0.5))
    raise HumanizeError(f"LLM 호출 실패({retries}회 시도): {last_err}")


def _parse_json_response(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s, flags=re.MULTILINE).strip()
    try:
        obj = json.loads(s)
    except Exception as e:
        raise HumanizeError(f"LLM 응답을 JSON 으로 파싱하지 못했습니다: {e}\n원본: {raw[:500]}")
    rewritten = obj.get("rewritten")
    if not isinstance(rewritten, str) or not rewritten.strip():
        raise HumanizeError("LLM 응답의 'rewritten' 필드가 비어 있거나 문자열이 아닙니다.")
    return rewritten


# ---------------------------------------------------------------------------
# 청크 분할 — 큰 문서를 헤딩/문단 경계로 쪼개 LLM 출력 토큰 한도를 회피
# ---------------------------------------------------------------------------

_HEADING_LINE_RE = re.compile(r"^#{1,6} ")


def _heading_offsets(text: str) -> List[int]:
    """코드펜스 밖에 있는 헤딩 라인의 시작 오프셋(문자 단위)을 반환."""
    out: List[int] = []
    in_fence = False
    pos = 0
    for line in text.splitlines(keepends=True):
        stripped = line.lstrip()
        if stripped.startswith("```"):
            in_fence = not in_fence
        elif not in_fence and _HEADING_LINE_RE.match(line):
            out.append(pos)
        pos += len(line)
    return out


def _split_at_headings(text: str) -> List[str]:
    """헤딩 경계로 본문을 섹션 리스트로 자른다(코드펜스 보존)."""
    bs = _heading_offsets(text)
    if not bs or bs[0] != 0:
        bs = [0] + bs
    bs = sorted(set(bs))
    sections: List[str] = []
    for i, start in enumerate(bs):
        end = bs[i + 1] if i + 1 < len(bs) else len(text)
        if start < end:
            sections.append(text[start:end])
    return sections


def _split_paragraphs_safe(text: str, hard_max: int) -> List[str]:
    """코드펜스를 깨지 않으며 빈 줄(`\\n\\n`) 단위로 자르는 폴백."""
    paragraphs: List[str] = []
    in_fence = False
    buf: List[str] = []
    for line in text.splitlines(keepends=True):
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
        if line.strip() == "" and not in_fence and buf:
            paragraphs.append("".join(buf))
            buf = [line]
        else:
            buf.append(line)
    if buf:
        paragraphs.append("".join(buf))

    chunks: List[str] = []
    current = ""
    for p in paragraphs:
        if len(p) > hard_max:
            if current:
                chunks.append(current)
                current = ""
            chunks.append(p)
            continue
        if not current or len(current) + len(p) <= hard_max:
            current += p
        else:
            chunks.append(current)
            current = p
    if current:
        chunks.append(current)
    return chunks


def _chunk_markdown(text: str, target: int = 6000, hard_max: int = 8000) -> List[str]:
    """마크다운을 헤딩 → 문단 순으로 자른다.

    각 청크는 hard_max 자 이내가 되도록 하고, target 자 근처까지 헤딩 섹션을 그리디 결합한다.
    """
    if len(text) <= target:
        return [text]

    sections = _split_at_headings(text)

    expanded: List[str] = []
    for sec in sections:
        if len(sec) <= hard_max:
            expanded.append(sec)
        else:
            expanded.extend(_split_paragraphs_safe(sec, hard_max))

    chunks: List[str] = []
    current = ""
    for sec in expanded:
        if not current:
            current = sec
        elif len(current) + len(sec) <= target:
            current += sec
        else:
            chunks.append(current)
            current = sec
    if current:
        chunks.append(current)
    return chunks


# ---------------------------------------------------------------------------
# 변경율 (간이 — 문자 단위 토큰 차이)
# ---------------------------------------------------------------------------

def _change_rate_pct(before: str, after: str) -> float:
    """문자 다중집합 거리 / 입력 길이. 0 ~ 100 (%)."""
    if not before:
        return 0.0
    from collections import Counter
    cb = Counter(before)
    ca = Counter(after)
    diff = sum((cb - ca).values()) + sum((ca - cb).values())
    return min(100.0, diff / max(1, len(before)) * 100.0)


# ---------------------------------------------------------------------------
# 공개 API
# ---------------------------------------------------------------------------

_DEFAULT_MODELS = {"openai": "gpt-4o", "anthropic": "claude-sonnet-4-6"}


def _ensure_env(provider: str) -> None:
    try:
        from dotenv import load_dotenv  # type: ignore
        load_dotenv()
    except ModuleNotFoundError:
        pass
    key = "OPENAI_API_KEY" if provider == "openai" else "ANTHROPIC_API_KEY"
    if not os.environ.get(key):
        raise HumanizeError(
            f"{key} 가 설정되어 있지 않습니다. .env 에 추가하거나 환경변수로 export 하세요."
        )


def humanize_text(
    text: str,
    *,
    provider: str = "openai",
    model: Optional[str] = None,
    api_timeout: int = 180,
    halt_change_rate: float = 50.0,
    warn_change_rate: float = 30.0,
) -> HumanizeResult:
    """한 문서 본문 문자열을 윤문한다.

    입력이 LLM 출력 토큰 한도를 넘길 만큼 크면 헤딩 경계로 청크 분할 후
    각 청크를 따로 윤문해 재결합한다.
    """
    provider = provider.lower().strip()
    if provider not in ("openai", "anthropic"):
        raise HumanizeError(f"Unknown provider: {provider}")
    model = model or _DEFAULT_MODELS[provider]
    _ensure_env(provider)

    before = detect(text)
    protected = collect_protected_strings(text)

    chunks = _chunk_markdown(text)
    if len(chunks) == 1:
        user = _user_prompt(text, before, protected)
        raw = _call_with_retry(provider, model, _SYSTEM_PROMPT, user, api_timeout)
        rewritten = _parse_json_response(raw)
    else:
        rewritten_parts: List[str] = []
        for i, chunk in enumerate(chunks, 1):
            chunk_report = detect(chunk)
            chunk_protected = collect_protected_strings(chunk)
            user = _user_prompt(chunk, chunk_report, chunk_protected)
            raw = _call_with_retry(provider, model, _SYSTEM_PROMPT, user, api_timeout)
            piece = _parse_json_response(raw)
            if not piece.endswith("\n"):
                piece += "\n"
            rewritten_parts.append(piece)
        rewritten = "".join(rewritten_parts)

    ok, missing = verify_preserved(text, rewritten)
    change_rate = _change_rate_pct(text, rewritten)

    warnings: List[str] = []
    halted = False
    halt_reason = ""

    if not ok:
        # 보호 어구가 깨졌으면 강제 롤백 — 안전 우선
        warnings.append(
            f"보호 어구 {len(missing)}건이 윤문본에 누락되어 원본을 유지합니다."
        )
        rewritten = text
        halted = True
        halt_reason = f"보호 어구 누락 ({len(missing)}건)"
    elif change_rate >= halt_change_rate:
        warnings.append(
            f"변경율 {change_rate:.1f}% (≥{halt_change_rate}%) 한도를 초과해 원본을 유지합니다."
        )
        rewritten = text
        halted = True
        halt_reason = f"변경율 초과 ({change_rate:.1f}%)"
    elif change_rate >= warn_change_rate:
        warnings.append(
            f"변경율 {change_rate:.1f}% (≥{warn_change_rate}%) — 결과를 사람이 검토하세요."
        )

    after = detect(rewritten)
    return HumanizeResult(
        input_path=None,
        output_path=None,
        original_text=text,
        rewritten_text=rewritten,
        before=before,
        after=after,
        improvement_percent=improvement_pct(before, after),
        change_rate_percent=change_rate,
        protected_missing=missing,
        halted=halted,
        halt_reason=halt_reason,
        warnings=warnings,
    )


def humanize_file(
    input_path: str | Path,
    output_path: str | Path,
    *,
    provider: str = "openai",
    model: Optional[str] = None,
    api_timeout: int = 180,
    report_path: Optional[str | Path] = None,
    halt_change_rate: float = 50.0,
    warn_change_rate: float = 30.0,
) -> HumanizeResult:
    """단일 파일을 읽어 윤문하고 결과를 출력 경로에 저장한다.

    `report_path` 가 주어지면 탐지·검증 리포트를 JSON 으로 함께 저장한다.
    """
    in_p = Path(input_path)
    out_p = Path(output_path)
    if not in_p.exists():
        raise HumanizeError(f"입력 파일을 찾을 수 없습니다: {in_p}")
    text = in_p.read_text(encoding="utf-8")

    result = humanize_text(
        text,
        provider=provider,
        model=model,
        api_timeout=api_timeout,
        halt_change_rate=halt_change_rate,
        warn_change_rate=warn_change_rate,
    )
    result.input_path = in_p
    result.output_path = out_p

    out_p.parent.mkdir(parents=True, exist_ok=True)
    out_p.write_text(result.rewritten_text, encoding="utf-8")

    if report_path is not None:
        rp = Path(report_path)
        rp.parent.mkdir(parents=True, exist_ok=True)
        rp.write_text(
            json.dumps(result.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    return result
