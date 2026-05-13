import pandas as pd
import pytest

from attribution.adapters import xlsx_returns


def test_returns_frame_is_long_with_expected_dtypes(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert list(result.returns.columns) == ["date", "ticker", "daily_return"]
    assert result.returns["date"].dtype == "datetime64[ns]"
    assert result.returns["ticker"].dtype == object
    assert result.returns["daily_return"].dtype == "float64"


def test_returns_frame_contains_three_stocks_three_unique_dates(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert set(result.returns["ticker"].unique()) == {"AAA", "BBB", "CCC"}
    # 4 date columns in the source, but one is a duplicate → 3 unique dates.
    assert result.returns["date"].nunique() == 3


def test_returns_frame_keeps_first_occurrence_of_duplicate_date(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    # The duplicate 20260303 column has value 0.99; the first 20260303 column
    # has value 0.0. We must keep the first.
    row = result.returns[
        (result.returns["ticker"] == "AAA")
        & (result.returns["date"] == pd.Timestamp("2026-03-03"))
    ]
    assert row["daily_return"].iloc[0] == 0.0


def test_published_series_holds_j803tr_row(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert list(result.published.columns) == ["date", "benchmark_return_published"]
    row = result.published[result.published["date"] == pd.Timestamp("2026-03-02")]
    # Value 0.001 is the J803TR row's 20260302 cell in the fixture; the
    # JSAPY row's value at the same date (0.002) must NOT appear here.
    assert row["benchmark_return_published"].iloc[0] == pytest.approx(0.001)


def test_jsapy_row_is_logged_as_ignored(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    ignored = [r for r in result.data_quality if r["category"] == "ignored_index_row"]
    assert len(ignored) == 1
    assert ignored[0]["ticker"] == "JSAPY Index"


def test_duplicate_date_column_is_recorded_in_data_quality(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    dup_rows = [r for r in result.data_quality if r["category"] == "duplicate_return_column"]
    assert len(dup_rows) == 1
    assert dup_rows[0]["date"] == pd.Timestamp("2026-03-03")


def test_base_date_returns_are_zero(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    base_rows = result.returns[result.returns["date"] == pd.Timestamp("2026-02-28")]
    assert (base_rows["daily_return"] == 0).all()


def test_ignored_indices_contains_jsapy_values(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert list(result.ignored_indices.columns) == ["date", "ticker", "value"]
    jsapy_rows = result.ignored_indices[result.ignored_indices["ticker"] == "JSAPY Index"]
    assert len(jsapy_rows) > 0
    # JSAPY value at 2026-03-02 should be 0.002 (from fixture)
    row = jsapy_rows[jsapy_rows["date"] == pd.Timestamp("2026-03-02")]
    assert row["value"].iloc[0] == pytest.approx(0.002)


def test_ticker_metadata_has_correct_columns_and_names(fixture_dir):
    result = xlsx_returns.read(fixture_dir / "tiny_returns.xlsx")

    assert list(result.ticker_metadata.columns) == ["ticker", "name"]
    assert set(result.ticker_metadata["ticker"]) == {"AAA", "BBB", "CCC"}
    # Name should match source Name column
    row = result.ticker_metadata[result.ticker_metadata["ticker"] == "AAA"]
    assert row["name"].iloc[0] == "ALPHA LTD"
