import pandas as pd
import pytest

from attribution.adapters import pdf_weights


@pytest.fixture
def page2_pdf(fixture_dir):
    p = fixture_dir / "holdings_page2_sample.pdf"
    if not p.exists():
        pytest.skip("page2 PDF fixture not present; run extract_page2.py")
    return p


def test_returns_long_frame_with_expected_columns(page2_pdf):
    df = pdf_weights.read_page2(page2_pdf)

    assert list(df.columns) == ["ticker", "benchmark_weight"]
    assert df["ticker"].dtype == object
    assert df["benchmark_weight"].dtype == "float64"


def test_tickers_are_3_letter_jse_codes(page2_pdf):
    df = pdf_weights.read_page2(page2_pdf)
    assert df["ticker"].str.match(r"^[A-Z]{3,4}$").all()


def test_weight_column_sums_close_to_one(page2_pdf):
    df = pdf_weights.read_page2(page2_pdf)
    s = df["benchmark_weight"].sum()
    assert 0.995 <= s <= 1.005, f"weight sum was {s}"


def test_contains_known_j803_constituents(page2_pdf):
    df = pdf_weights.read_page2(page2_pdf)
    tickers = set(df["ticker"])
    assert {"ATT", "GRT", "NRP", "RDF"}.issubset(tickers)


def test_rejection_path_returns_none(tmp_path):
    pytest.importorskip("reportlab")
    from reportlab.pdfgen.canvas import Canvas

    fake = tmp_path / "garbage.pdf"
    c = Canvas(str(fake))
    c.drawString(72, 720, "This is not a Catalyst holdings sheet")
    c.showPage()
    c.drawString(72, 720, "Neither is this")
    c.showPage()
    c.save()

    assert pdf_weights.read_page2(fake) is None
