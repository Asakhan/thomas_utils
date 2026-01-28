"""PowerPoint -> Markdown via Unstructured (optional engine)."""

from pathlib import Path
from typing import List, Union


def convert_unstructured(pptx_path: Union[str, Path]) -> str:
    """Convert PPTX to Markdown using Unstructured. Output follows ## Slide N, ### Content template."""
    try:
        from unstructured.partition.pptx import partition_pptx
    except ImportError as e:
        raise ImportError(
            "Unstructured is not installed. Install with: pip install 'thomas-utils[unstructured]'"
        ) from e
    path = Path(pptx_path)
    if not path.exists():
        raise FileNotFoundError(f"PPTX not found: {path}")
    if path.suffix.lower() != ".pptx":
        raise ValueError(f"Expected .pptx file, got: {path}")

    elements = partition_pptx(str(path))
    # Group by page_number if present; otherwise treat as single slide
    slides_content: List[List[str]] = []
    current: List[str] = []
    current_page: int = 0

    for el in elements:
        page = getattr(getattr(el, "metadata", None), "page_number", None) if el else None
        text = (el.text or "").strip() if hasattr(el, "text") else ""
        if not text:
            continue
        if page is not None and page != current_page and current:
            slides_content.append(current)
            current = []
            current_page = page
        current.append(text)
    if current:
        slides_content.append(current)

    if not slides_content:
        slides_content = [[""]]

    md_parts: List[str] = []
    for i, parts in enumerate(slides_content):
        block_lines = [
            f"## Slide {i + 1}",
            "**Type**: Content Slide",
            "",
            "### Content",
            "",
            "\n\n".join(parts).strip() or "",
        ]
        md_parts.append("\n".join(block_lines))
        if i < len(slides_content) - 1:
            md_parts.append("\n---\n\n")
    return "\n".join(md_parts)
