"""PyMuPDF4LLM-backed PDF -> Markdown conversion."""

from pathlib import Path
from typing import List, Optional, Union

import pymupdf4llm


def convert(
    pdf_path: Union[str, Path],
    pages: Optional[List[int]] = None,
) -> str:
    """Convert PDF to Markdown using PyMuPDF4LLM.

    Args:
        pdf_path: Path to the PDF file.
        pages: Optional 0-based page indices to convert. None means all pages.

    Returns:
        UTF-8 Markdown string.
    """
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {path}")
    md = pymupdf4llm.to_markdown(str(path), pages=pages)
    return md if isinstance(md, str) else md.decode("utf-8")
