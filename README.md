# thomas_utils

PDF와 PowerPoint를 내용 손실을 최소화하면서 Markdown으로 변환하는 도구입니다.

- **PDF**: 속도 우선(**PyMuPDF4LLM**) 또는 품질 우선(**marker-pdf**) 엔진 선택 가능.
- **PowerPoint**: **python-pptx**로 구조화 마크다운(Type, Layout, Title, Subtitle, Content) 추출. 표·리스트·코드블록·시각적 순서 지원. 선택적으로 **Unstructured** 엔진, **LLM 보정**, **멀티모달(슬라이드 이미지 → GPT-4o 비전)** 지원.

## 가장 빠르게 쓰기

**로컬에서 개발/실행할 때** (프로젝트 폴더에서):

```bash
python -m pip install -e .
python -m thomas_utils pdf2md INPUT.pdf
python -m thomas_utils pptx2md INPUT.pptx
```

**PyPI에서 설치한 경우**:

```bash
python -m pip install thomas-utils
thomas-utils pdf2md INPUT.pdf
thomas-utils pptx2md INPUT.pptx
```

`thomas-utils`가 인식되지 않으면 **반드시** `python -m thomas_utils ...` 를 사용하세요. (모듈 이름은 **밑줄** `thomas_utils` 이며, 하이픈 `thomas-utils` 가 아님.)  
`pip`이 인식되지 않으면 `python -m pip`(또는 `py -m pip`)을 사용하세요.  
경로를 지정하지 않으면 `output/INPUT.md`로 UTF-8로 저장됩니다.

**참고**: 프로젝트에 `.venv`가 이미 있으면 해당 가상환경을 활성화한 뒤 `python -m pip install -e .` 로 설치하고 `python -m thomas_utils ...` 로 실행하면 됩니다. `venv` 생성 시 권한 오류가 나면 기존 `.venv`를 사용하세요.

## 가상환경 및 테스트 (권장)

프로그램은 **가상환경**에서 실행하는 것을 전제로 합니다. `requirements.txt`로 의존성을 설치한 뒤, 1·2단계(PPT 구조화 변환)는 **pytest**로 검증할 수 있습니다.

```bash
# 가상환경 생성 및 활성화 (Windows)
python -m venv .venv
.venv\Scripts\activate

# 의존성 설치 (requirements.txt)
pip install -r requirements.txt
pip install -e .

# 1·2단계 동작 테스트
pytest tests/test_pptx.py -v
```

**3단계(LLM 보정)** 를 사용하려면 프로젝트 루트에 `.env` 파일을 두고 `OPENAI_API_KEY`를 설정하세요.  
`python-dotenv`가 `.env`를 읽어 OpenAI API 호출 시 해당 키를 사용합니다.

## 설치

- **기본**: `python -m pip install thomas-utils`
- **PDF marker 엔진**: `python -m pip install "thomas-utils[marker]"`
- **PPT LLM 보정**: `python -m pip install "thomas-utils[pptx-llm]"`
- **PPT 멀티모달(비전)**: `python -m pip install "thomas-utils[pptx-multimodal]"` (Windows: pywin32 + PowerPoint, 그 외: LibreOffice + pymupdf)
- **PPT Unstructured 엔진**: `python -m pip install "thomas-utils[unstructured]"`
- **PPT 수식(OMML→LaTeX)**: `python -m pip install "thomas-utils[pptx-math]"`

## CLI 사용법

### PDF 변환

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

### PowerPoint 변환

```bash
thomas-utils pptx2md INPUT.pptx [-o OUTPUT.md] [--slides LIST]
```

| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `INPUT.pptx` | 변환할 PPTX 경로 | (필수) |
| `-o`, `--output` | 출력 Markdown 경로 | `output/INPUT.md` |
| `--slides` | 변환할 슬라이드 (현재는 무시, 전체 슬라이드 변환) | 전체 |
| `--pptx-use-llm` | LLM으로 추출 마크다운 문장 다듬기 | 꺼짐 |
| `--engine` | `python-pptx` 또는 `unstructured` | `python-pptx` |
| `--pptx-use-llm-multimodal` | 슬라이드를 이미지로 렌더 후 GPT-4o 비전으로 마크다운 변환 | 꺼짐 |

예:

```bash
thomas-utils pptx2md presentation.pptx -o docs/presentation.md
thomas-utils pptx2md presentation.pptx
thomas-utils pptx2md presentation.pptx --pptx-use-llm
thomas-utils pptx2md presentation.pptx --pptx-use-llm-multimodal -o result.md
thomas-utils pptx2md presentation.pptx --engine unstructured
```

**참고**: PowerPoint 변환 시 마크다운만 생성되며, 이미지(PNG)는 추출하지 않습니다. 출력 파일은 항상 `output/` 폴더에 저장됩니다.

**출력 형식**: 각 슬라이드는 `## Slide N`, **Type** (Title Slide / Content Slide / Section Divider), **Layout**, **Title**, **Subtitle**, `### Content`(표·리스트·코드블록) 구조로 출력됩니다.

**멀티모달 LLM** (`--pptx-use-llm-multimodal`): 각 슬라이드를 이미지로 만든 뒤 GPT-4o 비전 API로 마크다운을 생성합니다.  
- **Windows**: Microsoft PowerPoint 설치 + `pip install pywin32` (또는 `pip install "thomas-utils[pptx-multimodal]"`). PowerPoint 창이 잠깐 보일 수 있습니다. LibreOffice 불필요.  
- **그 외**: LibreOffice(`soffice`)가 PATH에 있고 `pip install pymupdf` 필요.  
- `.env`에 `OPENAI_API_KEY` 설정 필요.

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

### PDF 변환

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

### PowerPoint 변환

```python
from thomas_utils.converters import convert_pptx

md = convert_pptx("presentation.pptx")
# LLM 보정: convert_pptx("presentation.pptx", use_llm=True)
# 멀티모달(비전): convert_pptx("presentation.pptx", use_llm_multimodal=True)
# Unstructured 엔진: convert_pptx("presentation.pptx", engine="unstructured")
```

- `convert_pptx(pptx_path, slides=None, use_llm=False, engine="python-pptx", use_llm_multimodal=False)`  
  - `pptx_path`: PPTX 파일 경로 (`str` 또는 `pathlib.Path`)
  - `slides`: 현재는 무시됨 (전체 슬라이드 변환)
  - `use_llm`: True면 추출 마크다운을 LLM으로 보정 (`.env`의 `OPENAI_API_KEY` 필요)
  - `engine`: `"python-pptx"` 또는 `"unstructured"`
  - `use_llm_multimodal`: True면 슬라이드를 이미지로 렌더 후 GPT-4o 비전으로 변환 (Windows: PowerPoint + pywin32, 그 외: LibreOffice + pymupdf)
- 반환값: UTF-8 Markdown 문자열 (구조: ## Slide N, **Type**, **Layout**, **Title**, **Subtitle**, ### Content).
- 이미지는 마크다운에 포함하지 않습니다.

## 테스트

```bash
python -m pip install -e ".[test]"
pytest tests -v
```

## 라이선스

MIT License. see [LICENSE](LICENSE).
