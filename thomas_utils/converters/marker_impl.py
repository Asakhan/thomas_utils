"""marker-pdf-backed PDF -> Markdown conversion.

Requires: pip install thomas-utils[marker]
Pages are ignored for this engine; the full document is converted.
"""

from pathlib import Path
from typing import List, Optional, Union


def convert(
    pdf_path: Union[str, Path],
    pages: Optional[List[int]] = None,
) -> str:
    """Convert PDF to Markdown using marker-pdf.

    Args:
        pdf_path: Path to the PDF file.
        pages: Ignored for marker engine (full document is always converted).

    Returns:
        UTF-8 Markdown string.
    """
    # pages intentionally unused: marker API does not expose page range in PdfConverter
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {path}")

    from marker.converters.pdf import PdfConverter
    from marker.models import create_model_dict
    from marker.output import text_from_rendered

    converter = PdfConverter(artifact_dict=create_model_dict())
    rendered = converter(str(path))
    text, _, _ = text_from_rendered(rendered)
    return text if isinstance(text, str) else str(text)
