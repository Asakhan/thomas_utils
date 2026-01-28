"""CLI for thomas_utils: PDF -> Markdown."""

import argparse
import sys
from pathlib import Path


def _parse_pages(s: str) -> list[int]:
    """Parse --pages '0,1,2' or '0-2' into 0-based indices."""
    out: list[int] = []
    for part in s.replace(" ", "").split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return sorted(set(out))


def _pdf2md(args: argparse.Namespace) -> int:
    from thomas_utils.converters import convert

    pdf = Path(args.input)
    if not pdf.exists():
        print(f"Error: file not found: {pdf}", file=sys.stderr)
        return 1
    if not pdf.suffix.lower() == ".pdf":
        print(f"Error: expected .pdf file, got: {pdf}", file=sys.stderr)
        return 1

    out_path = Path(args.output) if args.output else Path("output") / (pdf.stem + ".md")
    pages = _parse_pages(args.pages) if args.pages else None

    try:
        md = convert(str(pdf), pages=pages, engine=args.engine)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ImportError as e:
        if "marker" in str(e).lower() or "marker" in str(args.engine).lower():
            print("Error: marker engine requires 'pip install thomas-utils[marker]'", file=sys.stderr)
        else:
            print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(md, encoding="utf-8")
    print(f"Wrote {out_path}")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="thomas-utils",
        description="PDF to Markdown — fast (pymupdf) or high-fidelity (marker).",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    pdf2md_p = subparsers.add_parser("pdf2md", help="Convert PDF to Markdown")
    pdf2md_p.add_argument("input", metavar="INPUT.pdf", help="Input PDF path")
    pdf2md_p.add_argument("-o", "--output", metavar="OUTPUT.md", help="Output Markdown path (default: output/INPUT.md)")
    pdf2md_p.add_argument(
        "--pages",
        metavar="LIST",
        help="0-based page indices, e.g. 0,1,2 or 0-5 (default: all)",
    )
    pdf2md_p.add_argument(
        "--engine",
        choices=("pymupdf", "marker"),
        default="pymupdf",
        help="Conversion engine (default: pymupdf)",
    )
    pdf2md_p.set_defaults(_run=_pdf2md)

    args = parser.parse_args()
    run = getattr(args, "_run", None)
    if run is None:
        parser.print_help()
        sys.exit(0)
    sys.exit(run(args))
