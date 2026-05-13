"""Extract page 2 of a real Catalyst PDF into a tiny standalone PDF fixture.

Run from the repo root:
    python tests/fixtures/extract_page2.py "data/March PDF/02 March 2026.pdf"
"""
from pathlib import Path
import shutil
import sys

OUT = Path(__file__).parent / "holdings_page2_sample.pdf"


def main(src_pdf: str) -> None:
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        shutil.copy(src_pdf, OUT)
        print(f"wrote {OUT} (full PDF, install pypdf for a single-page extract)")
        return

    reader = PdfReader(src_pdf)
    writer = PdfWriter()
    writer.add_page(reader.pages[1])  # page 2, 0-indexed
    with open(OUT, "wb") as fh:
        writer.write(fh)
    print(f"wrote {OUT} (page 2 only)")


if __name__ == "__main__":
    main(sys.argv[1])
