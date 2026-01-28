# thomas_utils

PDF를 내용 손실을 최소화하면서 빠르게 Markdown으로 변환하는 도구입니다.  
속도 우선(**PyMuPDF4LLM**)과 품질 우선(**marker-pdf**) 엔진을 모두 지원합니다.

## 가장 빠르게 쓰기

```bash
python -m pip install thomas-utils
thomas-utils pdf2md INPUT.pdf
# 또는
python -m thomas_utils pdf2md INPUT.pdf
```

`pip`이 인식되지 않으면 `python -m pip`(또는 `py -m pip`)을 사용하세요.  
경로를 지정하지 않으면 `output/INPUT.md`로 UTF-8로 저장됩니다.

## 설치

- 기본(빠른 변환만):  
  `python -m pip install thomas-utils`
- marker 엔진까지 사용:  
  `python -m pip install "thomas-utils[marker]"`

## CLI 사용법

```bash
thomas-utils pdf2md INPUT.pdf [-o OUTPUT.md] [--pages 0,1,2] [--engine pymupdf|marker]
```

| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `INPUT.pdf` | 변환할 PDF 경로 | (필수) |
| `-o`, `--output` | 출력 Markdown 경로 | `output/INPUT.md` |
| `--pages` | 변환할 페이지 (0-based, 쉼표·범위). 예: `0,1,2` 또는 `0-5` | 전체 |
| `--engine` | `pymupdf`(속도) 또는 `marker`(품질) | `pymupdf` |

예:

```bash
thomas-utils pdf2md report.pdf -o docs/report.md
thomas-utils pdf2md report.pdf --pages 0-2 --engine pymupdf
thomas-utils pdf2md report.pdf --engine marker
```

## 내용 손실 없이 쓰기

- **지원**: 제목, 표, 리스트, 볼드/이탤릭, 이미지 참조 등.
- **제한**:
  - 복잡한 수식·다단·레이아웃은 `--engine marker`를 쓰는 편이 더 나을 수 있습니다.
  - marker 엔진은 `--pages`를 지원하지 않으며, 항상 전체 문서를 변환합니다.

### 엔진별 특성

| 구분 | PyMuPDF4LLM (`pymupdf`) | marker-pdf (`marker`) |
|------|-------------------------|------------------------|
| **속도** | 매우 빠름 (GPU 불필요) | 상대적으로 느림 (PyTorch, GPU 권장) |
| **내용 보존** | 제목/표/리스트/볼드/이탤릭 등 기본 구조 | 테이블·수식(LaTeX)·코드블록·다단·각주·헤더/푸터 제거까지 처리 |
| **의존성** | `pymupdf4llm`만 사용 | Python 3.10+, PyTorch, `marker-pdf` |

## Python API

```python
from thomas_utils.converters import convert

md = convert("document.pdf", pages=[0, 1], engine="pymupdf")
# 또는 고품질 모드:
# md = convert("document.pdf", engine="marker")
```

- `convert(pdf_path, pages=None, engine="pymupdf")`  
  - `pdf_path`: PDF 파일 경로 (`str` 또는 `pathlib.Path`)
  - `pages`: 변환할 0-based 페이지 인덱스 리스트. `None`이면 전체.
  - `engine`: `"pymupdf"` 또는 `"marker"`
- 반환값: UTF-8 Markdown 문자열.

## 테스트

```bash
python -m pip install -e ".[test]"
pytest tests -v
```

## 라이선스

MIT License. see [LICENSE](LICENSE).
