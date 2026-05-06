# thomas_utils

PDF, PowerPoint, 그리고 강의 영상을 내용 손실을 최소화하면서 Markdown으로 변환하는 도구입니다.

- **PDF**: 속도 우선(**PyMuPDF4LLM**) 또는 품질 우선(**marker-pdf**) 엔진 선택 가능.
- **PowerPoint**: **python-pptx**로 구조화 마크다운(Type, Layout, Title, Subtitle, Content) 추출. 표·리스트·코드블록·시각적 순서 지원. 선택적으로 **Unstructured** 엔진, **LLM 보정**, **멀티모달(슬라이드 이미지 → GPT-4o 비전)** 지원.
- **Video (강의)**: **Whisper**(STT) + **OpenCV**(장면 전환 감지) + **멀티모달 LLM**(Claude / GPT-4o)으로 강의 영상을 목차·섹션 요약·스크린샷·전체 스크립트가 포함된 마크다운 노트로 변환.

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
- **Video → Markdown**: `python -m pip install "thomas-utils[video-summary]"` (시스템에 `ffmpeg` 바이너리 필요)

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

### Video 변환 (강의 영상 → 마크다운 노트)

```bash
# 메인 CLI를 통해
thomas-utils video2md INPUT.mp4 [-o OUTPUT.md] [--provider anthropic|openai] [--whisper-model base] ...

# 또는 전용 모듈로 직접
python -m thomas_utils.video_summary --input INPUT.mp4 --output OUTPUT.md
```

| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `--input`, `-i` | 변환할 영상 경로 | (필수) |
| `--output`, `-o` | 출력 Markdown 경로 | `output/INPUT.md` |
| `--provider` | `anthropic`(Claude) 또는 `openai`(GPT-4o) | `anthropic` |
| `--model` | 모델 이름 강제 지정 | provider 기본값 |
| `--whisper-model` | `tiny` / `base` / `small` / `medium` / `large-v3` | `base` |
| `--language` | STT 언어 코드 (예: `ko`, `en`) | 자동 감지 |
| `--scene-threshold` | 장면 전환 민감도(0–1, 클수록 적게 감지) | `0.55` |
| `--min-gap-seconds` | 인접 장면 사이 최소 간격(초) | `8.0` |
| `--max-scenes` | 최대 섹션 수(긴 영상의 API 비용 상한) | `40` |
| `--api-timeout` | LLM 한 번 호출당 타임아웃(초) | `120` |
| `--audio-timeout` | ffmpeg 음성 추출 타임아웃(초) | `1800` |
| `--screenshots-dir` | 키프레임 이미지 저장 디렉터리 | `<OUTPUT>_assets/` |
| `--title` | 출력 강의 제목 강제 지정 | 영상 파일명 |

#### 동영상 변환 실행 전 셋팅 체크리스트

`video2md`를 실행하기 전에 아래 항목을 순서대로 점검하세요.

**1. 시스템 의존성 (필수)**

- `ffmpeg` 바이너리가 PATH에 설치되어 있어야 합니다.
  - Linux: `apt install ffmpeg`
  - macOS: `brew install ffmpeg`
  - Windows: 공식 빌드 설치 후 PATH 등록
- 확인: `ffmpeg -version`

**2. Python 패키지 설치 (필수)**

```bash
pip install "thomas-utils[video-summary]"
# 로컬 개발 시
pip install -e ".[video-summary]"
```
→ `opencv-python`, `faster-whisper`, `anthropic`/`openai` 등이 함께 설치됩니다.

**3. API 키 — `.env` 파일 (필수, provider에 따라 택1)**

프로젝트 루트의 `.env`에 사용할 provider에 맞는 키를 설정합니다.

```env
ANTHROPIC_API_KEY=sk-ant-...   # --provider anthropic (기본값) 사용 시
OPENAI_API_KEY=sk-...          # --provider openai 사용 시
```

**4. 실행 시 결정해야 할 옵션**

| 항목 | 결정 포인트 |
|------|------------|
| `--provider` | `anthropic`(Claude, 기본) / `openai`(GPT-4o) — 위 API 키와 일치해야 함 |
| `--whisper-model` | `tiny`/`base`/`small`/`medium`/`large-v3` — 한국어 강의는 `small` 이상 권장 |
| `--language` | 강의 언어 미리 지정(예: `ko`)하면 STT 정확도↑ |
| `--scene-threshold` | 기본 `0.55`. 슬라이드가 자주 바뀌면 ↑, 적게 잡히면 ↓ |
| `--min-gap-seconds` | 기본 `8.0`. 너무 잘게 잘리면 ↑ |
| `--max-scenes` | 기본 `40`. **API 비용 상한** — 긴 영상에서 반드시 점검 |
| `--api-timeout` / `--audio-timeout` | 긴 영상이면 충분히 키워두기 |
| `--screenshots-dir` | 기본 `<OUTPUT>_assets/` — 별도 보관 위치 원할 때만 |
| `--title` | 파일명이 아닌 별도 제목으로 출력하려면 지정 |

**5. 디스크/리소스 점검**

- 출력은 `output/` 폴더로 저장되며 **키프레임 이미지**가 `<OUTPUT>_assets/`에 함께 쌓입니다 — 충분한 디스크 여유 확인.
- Whisper `medium`/`large-v3`는 메모리·시간 비용이 크므로 GPU 또는 충분한 RAM 권장.

**빠른 사전 점검 명령**

```bash
ffmpeg -version
python -c "import faster_whisper, cv2, anthropic; print('ok')"
grep -E 'ANTHROPIC_API_KEY|OPENAI_API_KEY' .env
```

**출력 구조**: 강의 제목·메타정보 → **목차**(타임스탬프 점프) → 섹션별(스크린샷 + 핵심 포인트 + 요약 + 해당 구간 스크립트) → **전체 스크립트**(타임스탬프 포함). 키프레임 이미지는 `<OUTPUT>_assets/` 폴더에 저장되고 마크다운에서 상대 경로로 참조됩니다.

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

### Video 변환

```python
from thomas_utils.video_summary import convert_video

out_path = convert_video(
    video_path="lecture.mp4",
    output_path="output/lecture.md",
    provider="anthropic",          # 또는 "openai"
    whisper_model="base",
    scene_threshold=0.55,
    max_scenes=40,
)
```

- `convert_video(video_path, output_path, *, provider="anthropic", model=None, whisper_model="base", language=None, scene_threshold=0.55, min_gap_seconds=8.0, max_scenes=40, api_timeout=120, audio_timeout=1800, screenshots_dir=None, title=None) -> Path`
- 반환값: 실제로 작성된 Markdown 파일의 경로.
- 키프레임 이미지는 `<OUTPUT>_assets/` 또는 `screenshots_dir`에 저장됩니다.

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
