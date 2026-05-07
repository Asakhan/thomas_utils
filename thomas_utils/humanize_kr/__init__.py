"""Humanize-KR: 한글 문서에서 AI 말투를 제거하는 윤문 도구.

레퍼런스: https://github.com/epoko77-ai/im-not-ai (MIT)
- AI 패턴 분류 체계(40+ 패턴, 10개 카테고리, S1/S2/S3 심각도)
- 보호 규칙(수치·고유명사·인용구·코드블록 등 무수정)
- 윤문 후 내용 보존 검증과 변경율 한도(30% 경고 / 50% 강제 중단)
"""

from thomas_utils.humanize_kr.processor import (
    HumanizeError,
    HumanizeResult,
    humanize_file,
    humanize_text,
)

__all__ = [
    "humanize_file",
    "humanize_text",
    "HumanizeError",
    "HumanizeResult",
]
