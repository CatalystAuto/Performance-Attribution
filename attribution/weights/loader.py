"""Source-routed weights loader.

For each requested date:
  1. Attempt PDF extraction → if valid, use it with weight_source="pdf".
  2. Otherwise attempt closing-weight CSV → use with weight_source="csv_close".
  3. Otherwise record `no_source` in data_quality and skip the date.

A single date never mixes sources.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from attribution.adapters import csv_weights as _csv_mod
from attribution.adapters import pdf_weights as _pdf_mod

# Indirection so tests can monkeypatch without touching the real adapters.
def _read_pdf(path: Path) -> pd.DataFrame | None:
    return _pdf_mod.read_page2(path)


def _read_csv(path: Path) -> dict[str, pd.DataFrame]:
    return _csv_mod.read(path)


def _pdf_path_for(pdf_dir: Path, date: pd.Timestamp) -> Path:
    return pdf_dir / f"{date.strftime('%d %B %Y')}.pdf"


def _csv_path_for(csv_dir: Path, benchmark: str, date: pd.Timestamp) -> Path:
    return csv_dir / f"{benchmark}_{date.strftime('%Y%m%d')}.csv"


def load(
    dates: list[pd.Timestamp],
    pdf_dir: Path,
    csv_dir: Path,
    benchmark: str,
) -> tuple[pd.DataFrame, list[dict]]:
    frames: list[pd.DataFrame] = []
    dq: list[dict] = []

    for d in dates:
        pdf_path = _pdf_path_for(pdf_dir, d)
        csv_path = _csv_path_for(csv_dir, benchmark, d)

        used: pd.DataFrame | None = None
        source: str | None = None

        if pdf_path.exists():
            try:
                parsed = _read_pdf(pdf_path)
            except Exception as e:
                dq.append({
                    "date": d, "severity": "warning",
                    "category": "pdf_error", "ticker": None,
                    "note": f"PDF parse error: {e}",
                })
                parsed = None
            if parsed is not None and not parsed.empty:
                used = parsed
                source = "pdf"
            else:
                dq.append({
                    "date": d, "severity": "info",
                    "category": "pdf_rejected", "ticker": None,
                    "note": "PDF adapter rejected page 2; falling back to CSV.",
                })

        if used is None and csv_path.exists():
            try:
                blocks = _read_csv(csv_path)
            except Exception as e:
                dq.append({
                    "date": d, "severity": "warning",
                    "category": "csv_error", "ticker": None,
                    "note": f"CSV parse error: {e}",
                })
                blocks = None
            if blocks is not None and "closing" in blocks:
                used = blocks["closing"].rename(columns={"weight": "benchmark_weight"})
                source = "csv_close"

        if used is None:
            dq.append({
                "date": d, "severity": "warning",
                "category": "no_source", "ticker": None,
                "note": "Neither PDF nor CSV available for this date.",
            })
            continue

        used = used.copy()
        used["date"] = d
        used["weight_source"] = source
        frames.append(used[["date", "ticker", "benchmark_weight", "weight_source"]])

    if not frames:
        return pd.DataFrame(
            columns=["date", "ticker", "benchmark_weight", "weight_source"]
        ), dq
    return pd.concat(frames, ignore_index=True), dq
