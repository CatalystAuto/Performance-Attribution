from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from attribution.report import excel as ex


@pytest.fixture
def report_inputs():
    contributions = pd.DataFrame({
        "date": [pd.Timestamp("2026-03-02")] * 2,
        "ticker": ["AAA", "BBB"],
        "benchmark_weight": [0.5, 0.5],
        "daily_return": [0.01, -0.01],
        "contribution": [0.005, -0.005],
        "weight_source": ["pdf", "pdf"],
    })
    daily = pd.DataFrame({
        "date": [pd.Timestamp("2026-03-02")],
        "benchmark_return_computed": [0.0],
        "benchmark_return_published": [0.0],
        "diff_decimal": [0.0],
        "diff_bp": [0.0],
        "exceeds_tolerance": [False],
        "weight_source": ["pdf"],
    })
    data_quality = pd.DataFrame(
        [{"date": pd.Timestamp("2026-03-02"), "severity": "info", "category": "x", "ticker": None, "note": "n"}]
    )
    summary = {
        "month": "2026-03",
        "benchmark": "J803",
        "n_days": 1,
        "cumulative_return_computed": 0.0,
        "cumulative_return_published": 0.0,
        "max_abs_diff_bp": 0.0,
        "n_breaches": 0,
        "n_dq_issues": 1,
    }
    metadata = {
        "pdf_dir": "data/March PDF",
        "csv_dir": "data/Benchmark Weights/2026-03/J803",
        "returns": "data/March Stock Return.xlsx",
        "tolerance": 1e-7,
        "git_sha": "abc1234",
        "timestamp": "2026-05-13T09:00:00",
        "exit_code": 0,
    }
    return contributions, daily, data_quality, summary, metadata


def test_workbook_has_all_sheets(report_inputs, tmp_path):
    contributions, daily, dq, summary, metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        contributions=contributions,
        daily=daily,
        data_quality=dq,
        summary=summary,
        metadata=metadata,
    )

    assert out_xlsx.exists()
    wb = openpyxl.load_workbook(out_xlsx)
    assert set(wb.sheetnames) == {"summary", "daily_benchmark_return", "contributions", "data_quality", "run_metadata"}


def test_parquet_sidecar_written(report_inputs, tmp_path):
    contributions, daily, dq, summary, metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        contributions=contributions,
        daily=daily,
        data_quality=dq,
        summary=summary,
        metadata=metadata,
    )

    sidecar = tmp_path / "benchmark_attribution_2026-03_contributions.parquet"
    assert sidecar.exists()
    round_trip = pd.read_parquet(sidecar)
    assert len(round_trip) == len(contributions)
