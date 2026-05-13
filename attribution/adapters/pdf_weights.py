"""Adapter: Catalyst portfolio PDF → benchmark-weight long frame.

Strategy: open page 2, pull all words, then bucket each word into a
column by x-coordinate. Required output columns: ticker, benchmark_weight.

The PDF page is a wide landscape table (~4953 × 3499 pts) with wrapped
multi-line headers, so structural table-extraction (extract_tables) is
unreliable — the Name/Ticker columns come back entirely as None in data
rows, and the headers merge into one giant concatenated string. We rely
on calibrated x-coordinate bands instead.

CALIBRATION NOTES (02 March 2026.pdf, page 2):
  Page size : 4953.18 × 3499.76 pts
  Headers at top ≈ 367.8:
    'Name'       x0 =  105.9  →  name column starts here
    'Ticker'     x0 =  234.4  →  ticker code starts here
    'Benchmark'  x0 =  282.4  \
    'Weight'     x0 =  293.0  /  → "Actual Benchmark Weight" column
    'Permitte'   x0 =  457.5  →  right boundary of benchmark column

  Data rows (first constituent, NEPI ROCK / NRP, top ≈ 565.5):
    Ticker word  'NRP'   x0 ≈ 221
    Suffix       'SJ'    x0 ≈ 243
    Benchmark %  '14.10%' x0 ≈ 305

  Data rows vary slightly (x0 = 305–311), so we use a comfortable band.

Chosen bands:
  _TICKER_BAND    = (213, 248)   # 3-4 char uppercase JSE code
  _BENCHMARK_BAND = (288, 345)   # Actual Benchmark Weight (% or decimal)

Row bucketing uses 5-pt tolerance (rows are ≈14 pt apart).

If the page does not look like a Catalyst holdings page (no recognizable
columns, no parseable weights, weight sum outside ±0.005 of 1.0), the
function returns None so the loader can fall back.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd
import pdfplumber

from attribution import config

log = logging.getLogger(__name__)

# Calibrated x-coordinate bands (left, right) on page 2 of Catalyst PDFs.
# See module docstring for derivation.
_TICKER_BAND: tuple[float, float] = (213.0, 248.0)
_BENCHMARK_BAND: tuple[float, float] = (288.0, 345.0)

_HEADER_KEYWORDS = ("Benchmark", "Ticker", "Name")

# JSE ticker: exactly 3–4 uppercase letters (e.g. NRP, L2D, NEPI)
_TICKER_RE = re.compile(r"^[A-Z]{3,4}$")

# Bloomberg full identifier suffix appended to every JSE ticker code.
_TICKER_SUFFIX = " SJ Equity"


def _clean_ticker(raw: str) -> str | None:
    """Strip the Bloomberg ` SJ Equity` suffix and validate as a JSE code.

    Returns the cleaned 3-4 letter uppercase code, or None if the input
    doesn't match the JSE-code pattern after cleaning.
    """
    if not raw:
        return None
    candidate = raw.replace(_TICKER_SUFFIX, "").strip().upper()
    return candidate if _TICKER_RE.fullmatch(candidate) else None


def _is_valid_ticker(text: str) -> bool:
    return bool(_TICKER_RE.fullmatch(text))


def read_page2(path: str | Path) -> pd.DataFrame | None:
    """Extract ticker + benchmark weight from page 2 of a Catalyst PDF.

    Returns
    -------
    pd.DataFrame with columns ['ticker', 'benchmark_weight'] (float, decimal)
    or None if the page doesn't look like a Catalyst holdings sheet.
    """
    with pdfplumber.open(str(path)) as pdf:
        n_pages = len(pdf.pages)
        if n_pages == 0:
            log.info("PDF %s has no pages; rejecting.", path)
            return None
        # Use page index 1 (i.e. physical page 2) for the full Catalyst PDF.
        # Fall back to index 0 when given a single-page extract (e.g. test fixtures).
        page_index = 1 if n_pages >= 2 else 0
        pg = pdf.pages[page_index]
        words = pg.extract_words(x_tolerance=3, y_tolerance=3)

    if not words:
        log.info("Page 2 of %s yielded no words; rejecting.", path)
        return None

    text_blob = " ".join(w["text"] for w in words)
    if not all(kw in text_blob for kw in _HEADER_KEYWORDS):
        log.info("Page 2 of %s lacks expected headers; rejecting.", path)
        return None

    # Identify the approximate top-coordinate of the header row by finding
    # where 'Ticker' appears in the ticker-band x range.
    header_top = None
    for w in words:
        if w["text"] == "Ticker" and _TICKER_BAND[0] <= w["x0"] < _TICKER_BAND[1] + 20:
            header_top = w["top"]
            break

    if header_top is None:
        # Fallback: find 'Ticker' anywhere on the page
        for w in words:
            if w["text"] == "Ticker":
                header_top = w["top"]
                break

    if header_top is None:
        log.info("Page 2 of %s: could not locate Ticker header; rejecting.", path)
        return None

    # Only process data rows (below the header)
    data_words = [w for w in words if w["top"] > header_top + 5]

    # Group words into rows using a 5-pt bucket on the `top` coordinate.
    # (Row pitch is ~14 pts; 5-pt bucket keeps multi-glyph words together.)
    rows: dict[int, list[dict]] = {}
    for w in data_words:
        key = round(w["top"] / 5) * 5
        rows.setdefault(key, []).append(w)

    extracted: list[tuple[str, float]] = []

    for _key, row_words in sorted(rows.items()):
        # ── ticker column ────────────────────────────────────────────────
        ticker_words = [
            w for w in row_words
            if _TICKER_BAND[0] <= w["x0"] < _TICKER_BAND[1]
            and _is_valid_ticker(w["text"])
            and w["text"] != "SJ"   # "SJ" is the exchange suffix, not a ticker
        ]
        if not ticker_words:
            continue
        ticker = ticker_words[0]["text"]

        # ── benchmark weight column ──────────────────────────────────────
        bench_words = [
            w for w in row_words
            if _BENCHMARK_BAND[0] <= w["x0"] < _BENCHMARK_BAND[1]
        ]
        if not bench_words:
            continue

        raw_weight = "".join(w["text"] for w in bench_words).strip()
        # Detect whether the value is given as a percentage (e.g. "14.10%")
        is_percent = "%" in raw_weight
        weight_text = raw_weight.replace("%", "").replace(",", ".").strip()
        try:
            weight = float(weight_text)
        except ValueError:
            continue

        # Normalise to decimal fraction:
        # - If the raw text had a "%" suffix, always divide by 100.
        # - If no "%" suffix, treat as decimal already (shouldn't happen in
        #   this PDF, but guard against it with the > 1.5 heuristic).
        if is_percent:
            weight /= 100.0
        elif weight > 1.5:
            # Defensive fallback for any PDF that reports % values without a literal `%` suffix.
            weight /= 100.0

        extracted.append((ticker, weight))

    if not extracted:
        log.info("Page 2 of %s yielded no rows after parsing; rejecting.", path)
        return None

    df = pd.DataFrame(extracted, columns=["ticker", "benchmark_weight"])
    # Deduplicate: some tickers appear twice (e.g. TEXTON has two rows)
    df = df.groupby("ticker", as_index=False)["benchmark_weight"].sum()

    # Reject zero-weight-only frames (no real J803 constituents)
    nonzero = df[df["benchmark_weight"] > 0]
    if nonzero.empty:
        log.info("Page 2 of %s: all benchmark weights are zero; rejecting.", path)
        return None

    # Keep only J803 constituents (non-zero weight) to avoid polluting the
    # attribution calculation with zeros for non-constituents.
    df = nonzero.reset_index(drop=True).copy()

    s = df["benchmark_weight"].sum()
    if not (config.WEIGHT_SUM_LOWER <= s <= config.WEIGHT_SUM_UPPER):
        log.info(
            "Page 2 of %s weight sum %s outside [%s, %s]; rejecting.",
            path, s, config.WEIGHT_SUM_LOWER, config.WEIGHT_SUM_UPPER,
        )
        return None

    # Ensure classic object dtype for the ticker column regardless of pandas version.
    df["ticker"] = df["ticker"].astype(object)
    df["benchmark_weight"] = df["benchmark_weight"].astype("float64")

    return df
