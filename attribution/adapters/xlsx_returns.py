"""Adapter: wide-format returns xlsx → canonical long-format frames.

Returns a tuple of:
  - returns_frame:        (date, ticker, daily_return) long-format
  - benchmark_published:  (date, benchmark_return_published)
  - data_quality_rows:    list of dicts describing any issues encountered
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

PUBLISHED_TICKER_LABEL = "J803TR Index"
IGNORED_INDEX_LABELS: frozenset[str] = frozenset({"JSAPY Index"})


def _parse_date_header(cell) -> tuple[pd.Timestamp | None, bool]:
    """Return (Timestamp, is_pandas_duplicate) for a column header cell.

    Accepts:
      - Integer cells (e.g. 20260302) → (Timestamp, False)
      - String cells that are plain YYYYMMDD (e.g. "20260302") → (Timestamp, False)
      - String cells in pandas auto-dedup format "YYYYMMDD.N" (e.g. "20260303.1")
        → (Timestamp for YYYYMMDD, True)
      - Anything else → (None, False)
    """
    if cell is None:
        return None, False
    s = str(cell).strip()
    # Check for pandas duplicate-column suffix: "YYYYMMDD.N"
    is_dedup = False
    if len(s) > 8 and '.' in s:
        base, _, suffix = s.rpartition('.')
        if len(base) == 8 and base.isdigit() and suffix.isdigit():
            # Validate that base is actually a parseable YYYYMMDD date
            try:
                pd.Timestamp(base)
            except (ValueError, TypeError):
                return None, False
            s = base
            is_dedup = True
    if len(s) != 8 or not s.isdigit():
        return None, False
    try:
        return pd.Timestamp(s), is_dedup
    except (ValueError, TypeError):
        return None, False


def read(path: str | Path) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    raw = pd.read_excel(path, sheet_name="Return", header=0)

    id_cols = ["Name", "BB Ticker", "JSE code"]
    missing = [c for c in id_cols if c not in raw.columns]
    if missing:
        raise ValueError(f"Returns sheet missing identifier columns: {missing}")

    data_quality: list[dict] = []

    # Pair each non-id column with the date it parses to (if any).
    # Pandas automatically renames duplicate column headers as "YYYYMMDD.1",
    # "YYYYMMDD.2", etc. – we detect those via _parse_date_header's is_dedup flag.
    date_cols: list[tuple[str, pd.Timestamp]] = []
    seen_dates: dict[pd.Timestamp, str] = {}
    for col in raw.columns:
        if col in id_cols:
            continue
        ts, is_dedup = _parse_date_header(col)
        if ts is None:
            continue
        if ts in seen_dates or is_dedup:
            data_quality.append({
                "date": ts,
                "severity": "warning",
                "category": "duplicate_return_column",
                "ticker": None,
                "note": f"Duplicate column header {col!r}; keeping first occurrence ({seen_dates.get(ts, str(col))!r}).",
            })
            continue
        seen_dates[ts] = str(col)
        date_cols.append((col, ts))

    if not date_cols:
        raise ValueError("Returns sheet has no recognizable YYYYMMDD date columns.")

    # Route non-stock rows by BB Ticker. The published target is one
    # specific row; any other index-style row is logged and discarded.
    is_published = raw["BB Ticker"] == PUBLISHED_TICKER_LABEL
    is_ignored = raw["BB Ticker"].isin(IGNORED_INDEX_LABELS)
    published_raw = raw.loc[is_published]
    ignored_raw = raw.loc[is_ignored]
    stocks_raw = raw.loc[~(is_published | is_ignored)]

    if published_raw.empty:
        raise ValueError(f"Returns sheet has no row with BB Ticker == {PUBLISHED_TICKER_LABEL!r}.")
    if len(published_raw) > 1:
        raise ValueError(f"Returns sheet has multiple {PUBLISHED_TICKER_LABEL!r} rows.")

    for label in ignored_raw["BB Ticker"].tolist():
        data_quality.append({
            "date": None,
            "severity": "info",
            "category": "ignored_index_row",
            "ticker": label,
            "note": f"Row with BB Ticker == {label!r} present in returns xlsx; ignored.",
        })

    src_cols = [c for c, _ in date_cols]
    col_to_date = dict(date_cols)

    # Stock returns: melt wide → long, rename JSE code → ticker.
    melted = stocks_raw[["JSE code", *src_cols]].melt(
        id_vars=["JSE code"], value_vars=src_cols,
        var_name="src_col", value_name="daily_return",
    )
    melted["date"] = melted["src_col"].map(col_to_date)
    returns_frame = (
        melted.rename(columns={"JSE code": "ticker"})
        .loc[:, ["date", "ticker", "daily_return"]]
        .reset_index(drop=True)
    )
    returns_frame["ticker"] = returns_frame["ticker"].astype(object)
    returns_frame["daily_return"] = returns_frame["daily_return"].astype("float64")
    returns_frame["date"] = returns_frame["date"].astype("datetime64[ns]")

    # Published series.
    published_melted = published_raw[src_cols].melt(
        var_name="src_col", value_name="benchmark_return_published",
    )
    published_melted["date"] = published_melted["src_col"].map(col_to_date)
    benchmark_published = (
        published_melted[["date", "benchmark_return_published"]]
        .astype({"benchmark_return_published": "float64"})
        .reset_index(drop=True)
    )
    benchmark_published["date"] = benchmark_published["date"].astype("datetime64[ns]")

    return returns_frame, benchmark_published, data_quality
