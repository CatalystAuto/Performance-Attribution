"""Write the run output: 5-sheet workbook + parquet sidecar of contributions."""
from __future__ import annotations

from pathlib import Path

import pandas as pd


_SHEET_ORDER = [
    "summary",
    "daily_benchmark_return",
    "contributions",
    "data_quality",
    "run_metadata",
]


def _summary_to_df(summary: dict) -> pd.DataFrame:
    return pd.DataFrame([{"field": k, "value": v} for k, v in summary.items()])


def _metadata_to_df(metadata: dict) -> pd.DataFrame:
    return pd.DataFrame([{"field": k, "value": v} for k, v in metadata.items()])


def write(
    out_xlsx: str | Path,
    *,
    contributions: pd.DataFrame,
    daily: pd.DataFrame,
    data_quality: pd.DataFrame,
    summary: dict,
    metadata: dict,
) -> None:
    out_xlsx = Path(out_xlsx)
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    sheets = {
        "summary": _summary_to_df(summary),
        "daily_benchmark_return": daily,
        "contributions": contributions,
        "data_quality": data_quality,
        "run_metadata": _metadata_to_df(metadata),
    }

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xl:
        for name in _SHEET_ORDER:
            sheets[name].to_excel(xl, sheet_name=name, index=False)

    sidecar = out_xlsx.with_name(out_xlsx.stem + "_contributions.parquet")
    contributions.to_parquet(sidecar, index=False)
