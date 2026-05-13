from pathlib import Path

import pandas as pd
import pytest

from attribution.weights import loader as wl


@pytest.fixture
def fake_loader(tmp_path, monkeypatch):
    """Build a loader run against a temp PDF dir + CSV dir with controlled adapter behaviour."""
    pdf_dir = tmp_path / "pdfs"
    csv_dir = tmp_path / "csvs"
    pdf_dir.mkdir()
    csv_dir.mkdir()

    # Create empty placeholder PDFs and CSVs so file-existence checks pass.
    (pdf_dir / "02 March 2026.pdf").touch()  # date 1: PDF present, CSV present → PDF wins
    (csv_dir / "J803_20260302.csv").touch()

    (csv_dir / "J803_20260303.csv").touch()  # date 2: PDF missing, CSV present → CSV used

    # date 3: neither present → no_source

    def fake_pdf_read(path):
        if "02 March" in str(path):
            return pd.DataFrame({"ticker": ["AAA", "BBB"], "benchmark_weight": [0.6, 0.4]})
        return None

    def fake_csv_read(path):
        return {
            "closing": pd.DataFrame({"ticker": ["AAA", "BBB"], "weight": [0.5, 0.5]}),
            "opening": pd.DataFrame({"ticker": ["AAA", "BBB"], "weight": [0.5, 0.5]}),
        }

    monkeypatch.setattr(wl, "_read_pdf", fake_pdf_read)
    monkeypatch.setattr(wl, "_read_csv", fake_csv_read)

    return pdf_dir, csv_dir


def test_pdf_wins_when_present(fake_loader):
    pdf_dir, csv_dir = fake_loader
    frame, dq = wl.load(
        dates=[pd.Timestamp("2026-03-02")],
        pdf_dir=pdf_dir,
        csv_dir=csv_dir,
        benchmark="J803",
    )
    row = frame[frame["ticker"] == "AAA"].iloc[0]
    assert row["weight_source"] == "pdf"
    assert row["benchmark_weight"] == 0.6


def test_csv_fallback_when_pdf_missing(fake_loader):
    pdf_dir, csv_dir = fake_loader
    frame, dq = wl.load(
        dates=[pd.Timestamp("2026-03-03")],
        pdf_dir=pdf_dir,
        csv_dir=csv_dir,
        benchmark="J803",
    )
    assert (frame["weight_source"] == "csv_close").all()
    assert frame["benchmark_weight"].sum() == pytest.approx(1.0)


def test_no_source_records_data_quality(fake_loader):
    pdf_dir, csv_dir = fake_loader
    frame, dq = wl.load(
        dates=[pd.Timestamp("2026-03-04")],
        pdf_dir=pdf_dir,
        csv_dir=csv_dir,
        benchmark="J803",
    )
    assert frame.empty
    assert any(r["category"] == "no_source" and r["date"] == pd.Timestamp("2026-03-04") for r in dq)


def test_pdf_rejection_falls_back_to_csv(tmp_path, monkeypatch):
    pdf_dir = tmp_path / "pdfs"
    csv_dir = tmp_path / "csvs"
    pdf_dir.mkdir()
    csv_dir.mkdir()
    (pdf_dir / "02 March 2026.pdf").touch()
    (csv_dir / "J803_20260302.csv").touch()

    monkeypatch.setattr(wl, "_read_pdf", lambda path: None)  # rejection
    monkeypatch.setattr(
        wl, "_read_csv",
        lambda path: {
            "closing": pd.DataFrame({"ticker": ["AAA"], "weight": [1.0]}),
            "opening": pd.DataFrame({"ticker": ["AAA"], "weight": [1.0]}),
        },
    )

    frame, dq = wl.load(
        dates=[pd.Timestamp("2026-03-02")],
        pdf_dir=pdf_dir,
        csv_dir=csv_dir,
        benchmark="J803",
    )
    assert (frame["weight_source"] == "csv_close").all()
    assert any(r["category"] == "pdf_rejected" for r in dq)
