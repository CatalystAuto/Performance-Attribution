import pandas as pd
import pytest

from attribution.adapters import xlsx_returns


def test_returns_frame_is_long_with_expected_dtypes(fixture_dir):
    returns, published, dq = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert list(returns.columns) == ["date", "ticker", "daily_return"]
    assert returns["date"].dtype == "datetime64[ns]"
    assert returns["ticker"].dtype == object
    assert returns["daily_return"].dtype == "float64"


def test_returns_frame_contains_three_stocks_three_unique_dates(fixture_dir):
    returns, _, _ = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert set(returns["ticker"].unique()) == {"AAA", "BBB", "CCC"}
    # 4 date columns in the source, but one is a duplicate → 3 unique dates.
    assert returns["date"].nunique() == 3


def test_returns_frame_keeps_first_occurrence_of_duplicate_date(fixture_dir):
    returns, _, _ = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    # The duplicate 20260303 column has value 0.99; the first 20260303 column
    # has value 0.0. We must keep the first.
    row = returns[(returns["ticker"] == "AAA") & (returns["date"] == pd.Timestamp("2026-03-03"))]
    assert row["daily_return"].iloc[0] == 0.0


def test_published_series_holds_j803tr_row(fixture_dir):
    _, published, _ = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert list(published.columns) == ["date", "benchmark_return_published"]
    row = published[published["date"] == pd.Timestamp("2026-03-02")]
    # Value 0.001 is the J803TR row's 20260302 cell in the fixture; the
    # JSAPY row's value at the same date (0.002) must NOT appear here.
    assert row["benchmark_return_published"].iloc[0] == pytest.approx(0.001)


def test_jsapy_row_is_logged_as_ignored(fixture_dir):
    _, _, dq = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    ignored = [r for r in dq if r["category"] == "ignored_index_row"]
    assert len(ignored) == 1
    assert ignored[0]["ticker"] == "JSAPY Index"


def test_duplicate_date_column_is_recorded_in_data_quality(fixture_dir):
    _, _, dq = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    dup_rows = [r for r in dq if r["category"] == "duplicate_return_column"]
    assert len(dup_rows) == 1
    assert dup_rows[0]["date"] == pd.Timestamp("2026-03-03")


def test_base_date_returns_are_zero(fixture_dir):
    returns, _, _ = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    base_rows = returns[returns["date"] == pd.Timestamp("2026-02-28")]
    assert (base_rows["daily_return"] == 0).all()
