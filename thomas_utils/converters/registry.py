"""Engine registry and unified convert() API."""

from pathlib import Path
from typing import List, Optional, Union

_ENGINES = ("pymupdf", "marker")


def get_engine(name: str) -> str:
    """Return engine name if supported, else raise ValueError."""
    n = name.lower().strip()
    if n not in _ENGINES:
        raise ValueError(f"Unknown engine: {name}. Choose from {_ENGINES}.")
    return n


def convert(
    pdf_path: Union[str, Path],
    pages: Optional[List[int]] = None,
    engine: str = "pymupdf",
) -> str:
    """Convert PDF to Markdown.

    Args:
        pdf_path: Path to the PDF file.
        pages: Optional 0-based page indices. None = all pages.
               For engine "marker", pages may be ignored (full doc converted).
        engine: "pymupdf" (fast, default) or "marker" (high-fidelity).

    Returns:
        UTF-8 Markdown string.
    """
    eng = get_engine(engine)
    if eng == "pymupdf":
        from thomas_utils.converters.pymupdf_impl import convert as _convert

        return _convert(pdf_path, pages=pages)
    if eng == "marker":
        from thomas_utils.converters.marker_impl import convert as _convert

        return _convert(pdf_path, pages=pages)
    raise ValueError(f"Unknown engine: {engine}")
