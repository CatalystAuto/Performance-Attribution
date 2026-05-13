from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from attribution.report import excel as ex


@pytest.fixture
def report_inputs():
    """Two-ticker, two-date scenario for testing the 3-sheet wide format."""
    dates = [pd.Timestamp("2026-03-02"), pd.Timestamp("2026-03-03")]

    weights = pd.DataFrame({
        "date": dates * 2,
        "ticker": ["AAA", "AAA", "BBB", "BBB"],
        "benchmark_weight": [0.5, 0.5, 0.5, 0.5],
        "weight_source": ["pdf"] * 4,
    })

    returns = pd.DataFrame({
        "date": dates * 2,
        "ticker": ["AAA", "AAA", "BBB", "BBB"],
        "daily_return": [0.01, 0.02, -0.01, -0.02],
    })

    contributions = pd.DataFrame({
        "date": dates * 2,
        "ticker": ["AAA", "AAA", "BBB", "BBB"],
        "benchmark_weight": [0.5, 0.5, 0.5, 0.5],
        "daily_return": [0.01, 0.02, -0.01, -0.02],
        "contribution": [0.005, 0.01, -0.005, -0.01],
        "weight_source": ["pdf"] * 4,
    })

    daily = pd.DataFrame({
        "date": dates,
        "benchmark_return_computed": [0.0, 0.0],
        "benchmark_return_published": [0.001, 0.002],
        "diff_decimal": [0.0, 0.0],
        "diff_bp": [0.0, 0.0],
        "exceeds_tolerance": [False, False],
        "weight_source": ["pdf", "pdf"],
    })

    ignored_indices = pd.DataFrame({
        "date": dates,
        "ticker": ["JSAPY Index", "JSAPY Index"],
        "value": [0.0015, 0.0025],
    })

    ticker_metadata = pd.DataFrame({
        "ticker": ["AAA", "BBB"],
        "name": ["ALPHA LTD", "BRAVO LTD"],
    })

    return weights, returns, contributions, daily, ignored_indices, ticker_metadata


def test_workbook_has_three_sheets_with_expected_names(report_inputs, tmp_path):
    weights, returns, contributions, daily, ignored_indices, ticker_metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        weights=weights,
        returns=returns,
        contributions=contributions,
        daily=daily,
        ignored_indices=ignored_indices,
        ticker_metadata=ticker_metadata,
    )

    assert out_xlsx.exists()
    wb = openpyxl.load_workbook(out_xlsx)
    assert wb.sheetnames == ["Weights", "Returns", "Calculation"]


def test_calculation_sheet_has_instruction_blank_and_summary_rows(report_inputs, tmp_path):
    weights, returns, contributions, daily, ignored_indices, ticker_metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        weights=weights,
        returns=returns,
        contributions=contributions,
        daily=daily,
        ignored_indices=ignored_indices,
        ticker_metadata=ticker_metadata,
    )

    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["Calculation"]

    # Row 1, col A is the instructional string
    assert ws.cell(1, 1).value is not None
    assert ws.cell(1, 1).value.startswith("Each cell below")
    # Row 1 cols B onwards should be empty
    assert ws.cell(1, 2).value is None

    # Row 2 is the header
    assert ws.cell(2, 1).value == "Ticker"
    assert ws.cell(2, 2).value == "Name"
    assert ws.cell(2, 3).value == "2026-03-02"

    # Bottom of sheet: PORT, IDX, DIFF rows (last 3 rows)
    last_three = [
        ws.cell(ws.max_row - 2, 1).value,
        ws.cell(ws.max_row - 1, 1).value,
        ws.cell(ws.max_row, 1).value,
    ]
    assert last_three == ["PORT", "IDX", "DIFF"]

    # Blank row just before PORT
    blank_row_idx = ws.max_row - 3
    assert ws.cell(blank_row_idx, 1).value is None


def test_weights_sheet_shape(report_inputs, tmp_path):
    weights, returns, contributions, daily, ignored_indices, ticker_metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        weights=weights,
        returns=returns,
        contributions=contributions,
        daily=daily,
        ignored_indices=ignored_indices,
        ticker_metadata=ticker_metadata,
    )

    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["Weights"]
    # Header row + 2 tickers = 3 rows, Ticker+Name+2 dates = 4 cols
    assert ws.max_row == 3
    assert ws.max_column == 4
    assert ws.cell(1, 1).value == "Ticker"
    assert ws.cell(1, 3).value == "2026-03-02"
    assert ws.cell(2, 1).value == "AAA SJ"


def test_returns_sheet_has_j803tr_and_jsapy_rows(report_inputs, tmp_path):
    weights, returns, contributions, daily, ignored_indices, ticker_metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        weights=weights,
        returns=returns,
        contributions=contributions,
        daily=daily,
        ignored_indices=ignored_indices,
        ticker_metadata=ticker_metadata,
    )

    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["Returns"]
    # Header + 2 tickers + J803TR + JSAPY = 5 rows
    assert ws.max_row == 5
    assert ws.cell(4, 1).value == "J803TR"
    assert ws.cell(5, 1).value == "JSAPY"


def test_parquet_sidecar_written(report_inputs, tmp_path):
    weights, returns, contributions, daily, ignored_indices, ticker_metadata = report_inputs
    out_xlsx = tmp_path / "benchmark_attribution_2026-03.xlsx"

    ex.write(
        out_xlsx,
        weights=weights,
        returns=returns,
        contributions=contributions,
        daily=daily,
        ignored_indices=ignored_indices,
        ticker_metadata=ticker_metadata,
    )

    sidecar = tmp_path / "benchmark_attribution_2026-03_contributions.parquet"
    assert sidecar.exists()
    round_trip = pd.read_parquet(sidecar)
    assert len(round_trip) == len(contributions)
