"""Tests for PDF -> Markdown conversion."""

from pathlib import Path

import pytest


def _make_sample_pdf(path: Path) -> None:
    """Create a minimal one-page PDF with known text using pymupdf."""
    import pymupdf

    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Test Heading\n\nBody text for thomas_utils.")
    doc.save(str(path))
    doc.close()


def test_convert_pymupdf_non_empty(tmp_path: Path) -> None:
    """convert() returns non-empty markdown for a minimal PDF."""
    from thomas_utils.converters import convert

    pdf_path = tmp_path / "sample.pdf"
    _make_sample_pdf(pdf_path)
    result = convert(str(pdf_path), engine="pymupdf")
    assert result
    assert "Test" in result or "Heading" in result or "Body" in result or "thomas_utils" in result


def test_convert_pymupdf_pages_subset(tmp_path: Path) -> None:
    """convert() with pages=[0] returns content from first page only."""
    from thomas_utils.converters import convert

    pdf_path = tmp_path / "sample.pdf"
    _make_sample_pdf(pdf_path)
    result = convert(str(pdf_path), pages=[0], engine="pymupdf")
    assert result
    assert "Test" in result or "Heading" in result or "Body" in result


def test_convert_missing_file() -> None:
    """convert() raises FileNotFoundError for missing path."""
    from thomas_utils.converters import convert

    with pytest.raises(FileNotFoundError, match="not found"):
        convert("/nonexistent/sample.pdf", engine="pymupdf")


def test_get_engine() -> None:
    """get_engine() accepts pymupdf and marker, rejects others."""
    from thomas_utils.converters import get_engine

    assert get_engine("pymupdf") == "pymupdf"
    assert get_engine("marker") == "marker"
    with pytest.raises(ValueError, match="Unknown engine"):
        get_engine("invalid")
