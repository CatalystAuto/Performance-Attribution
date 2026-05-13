# Benchmark Attribution Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Python package + CLI that, for a given month, computes the daily benchmark return and per-stock contributions for the J803 South African Listed Property index, and validates the result against the published `J803TR Index` (total-return) series carried in the returns workbook.

**Architecture:** Per-source adapters convert messy inputs (page-2 PDFs, opening/closing-weight CSVs, wide-format returns xlsx) into canonical long-format DataFrames. A source-routed loader prefers PDFs and falls back to closing-weight CSVs. Compute, validate, report, and CLI layers consume only the canonical shapes.

**Tech Stack:** Python 3.11+ (interpreter at hand: 3.14.5), pandas, openpyxl, pdfplumber 0.11.9, pyarrow, pytest + pytest-cov.

**Spec:** [`docs/superpowers/specs/2026-05-12-benchmark-attribution-pipeline-design.md`](../specs/2026-05-12-benchmark-attribution-pipeline-design.md)

---

## File Structure

```
attribution/
├── __init__.py
├── __main__.py                 # entry point: `python -m attribution …`
├── config.py                   # constants (tolerance, weight-sum bounds, paths)
├── adapters/
│   ├── __init__.py
│   ├── pdf_weights.py          # pdfplumber → (date, ticker, benchmark_weight)
│   ├── csv_weights.py          # J803/J253 CSV → closing + opening weights
│   └── xlsx_returns.py         # wide-returns xlsx → long returns + J803TR published series
├── weights/
│   ├── __init__.py
│   └── loader.py               # source-routed weights loader, per-date precedence
├── compute/
│   ├── __init__.py
│   └── benchmark.py            # join + contribution + daily benchmark return
├── validate/
│   ├── __init__.py
│   └── published.py            # computed vs J803TR published row, tolerance flagging
├── report/
│   ├── __init__.py
│   └── excel.py                # 5-sheet workbook writer + parquet sidecar
└── cli.py                      # argparse, orchestration, exit codes

tests/
├── __init__.py
├── conftest.py                 # shared fixtures
├── fixtures/
│   ├── README.md
│   ├── make_fixtures.py        # regenerable build script for synthetic fixtures
│   ├── tiny_returns.xlsx       # built by make_fixtures.py
│   ├── tiny_weights_both_blocks.csv
│   ├── tiny_weights_closing_only.csv
│   ├── tiny_weights_bad_sum.csv
│   └── holdings_page2_sample.pdf   # extracted real page 2 (one-time copy)
├── adapters/
│   ├── __init__.py
│   ├── test_xlsx_returns.py
│   ├── test_csv_weights.py
│   └── test_pdf_weights.py
├── weights/
│   ├── __init__.py
│   └── test_loader.py
├── compute/
│   ├── __init__.py
│   └── test_benchmark.py
├── validate/
│   ├── __init__.py
│   └── test_published.py
├── report/
│   ├── __init__.py
│   └── test_excel.py
└── test_cli_smoke.py

pyproject.toml
.python-version                 # 3.14
README.md                       # short usage note
```

**Responsibility split:** Each adapter knows one source format and nothing else. The loader knows source precedence and nothing about source formats. Compute, validate, and report know only the canonical schemas. The CLI is a thin orchestrator. This is the seam where future Brinson + portfolio work will splice in.

---

## Task 1: Project scaffolding

**Files:**
- Create: `pyproject.toml`
- Create: `.python-version`
- Create: `attribution/__init__.py`
- Create: `attribution/adapters/__init__.py`
- Create: `attribution/weights/__init__.py`
- Create: `attribution/compute/__init__.py`
- Create: `attribution/validate/__init__.py`
- Create: `attribution/report/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `tests/test_smoke.py` (a trivial test so pytest has something to discover)

- [ ] **Step 1: Write `pyproject.toml`**

```toml
[project]
name = "attribution"
version = "0.1.0"
description = "Catalyst benchmark attribution pipeline"
requires-python = ">=3.11"
dependencies = [
    "pandas>=2.0",
    "openpyxl>=3.1",
    "pdfplumber>=0.11",
    "pyarrow>=14.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "pytest-cov>=4.0",
]

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[tool.setuptools.packages.find]
include = ["attribution*"]

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-ra"
```

- [ ] **Step 2: Write `.python-version`**

```
3.14
```

- [ ] **Step 3: Write all empty `__init__.py` files**

Each of: `attribution/__init__.py`, `attribution/adapters/__init__.py`, `attribution/weights/__init__.py`, `attribution/compute/__init__.py`, `attribution/validate/__init__.py`, `attribution/report/__init__.py`, `tests/__init__.py` is a zero-byte file.

- [ ] **Step 4: Write `tests/conftest.py`**

```python
from pathlib import Path
import pytest

FIXTURE_DIR = Path(__file__).parent / "fixtures"

@pytest.fixture
def fixture_dir() -> Path:
    return FIXTURE_DIR
```

- [ ] **Step 5: Write a discovery smoke test `tests/test_smoke.py`**

```python
def test_pytest_runs():
    assert True
```

- [ ] **Step 6: Install and run**

Run: `C:/Users/jonathan/AppData/Local/Python/pythoncore-3.14-64/python.exe -m pip install -e ".[dev]"`
Then: `C:/Users/jonathan/AppData/Local/Python/pythoncore-3.14-64/python.exe -m pytest -v`
Expected: `1 passed`.

- [ ] **Step 7: Commit**

```bash
git add pyproject.toml .python-version attribution tests
git commit -m "Scaffold attribution package and pytest setup"
```

---

## Task 2: `attribution.config`

**Files:**
- Create: `attribution/config.py`
- Create: `tests/test_config.py`

- [ ] **Step 1: Write the failing test `tests/test_config.py`**

```python
from pathlib import Path

from attribution import config


def test_tolerance_is_1e_7():
    assert config.TOLERANCE_DECIMAL == 1e-7


def test_weight_sum_bounds():
    assert config.WEIGHT_SUM_LOWER == 0.995
    assert config.WEIGHT_SUM_UPPER == 1.005


def test_default_benchmark_is_j803():
    assert config.DEFAULT_BENCHMARK == "J803"


def test_default_paths_are_paths():
    assert isinstance(config.DEFAULT_PDF_DIR, Path)
    assert isinstance(config.DEFAULT_CSV_ROOT, Path)
    assert isinstance(config.DEFAULT_RETURNS, Path)
    assert isinstance(config.DEFAULT_OUT_DIR, Path)


def test_csv_dir_for_run():
    p = config.csv_dir_for("2026-03", "J803")
    assert p == config.DEFAULT_CSV_ROOT / "2026-03" / "J803"
```

- [ ] **Step 2: Run the test and confirm it fails**

Run: `python -m pytest tests/test_config.py -v`
Expected: `ImportError` or `AttributeError` (no `config` module yet).

- [ ] **Step 3: Implement `attribution/config.py`**

```python
"""Project-wide constants.

Tolerances and default paths are centralized here so the rest of the
package never hard-codes them.
"""
from pathlib import Path

TOLERANCE_DECIMAL: float = 1e-7
WEIGHT_SUM_LOWER: float = 0.995
WEIGHT_SUM_UPPER: float = 1.005

DEFAULT_BENCHMARK: str = "J803"
DEFAULT_PDF_DIR: Path = Path("data/March PDF")
DEFAULT_CSV_ROOT: Path = Path("data/Benchmark Weights")
DEFAULT_RETURNS: Path = Path("data/March Stock Return.xlsx")
DEFAULT_OUT_DIR: Path = Path("output")


def csv_dir_for(month: str, benchmark: str) -> Path:
    """Resolve `<DEFAULT_CSV_ROOT>/<month>/<benchmark>`.

    `month` is in `YYYY-MM` form, e.g. `"2026-03"`.
    """
    return DEFAULT_CSV_ROOT / month / benchmark
```

- [ ] **Step 4: Run and confirm pass**

Run: `python -m pytest tests/test_config.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add attribution/config.py tests/test_config.py
git commit -m "Add config module with tolerance and path defaults"
```

---

## Task 3: Build fixture-builder script and synthetic fixtures

**Files:**
- Create: `tests/fixtures/make_fixtures.py`
- Create: `tests/fixtures/README.md`
- Create (by running script): `tests/fixtures/tiny_returns.xlsx`
- Create (by hand): `tests/fixtures/tiny_weights_both_blocks.csv`
- Create (by hand): `tests/fixtures/tiny_weights_closing_only.csv`
- Create (by hand): `tests/fixtures/tiny_weights_bad_sum.csv`

These fixtures back the adapter tests in tasks 4, 5, 6.

- [ ] **Step 1: Write `tests/fixtures/make_fixtures.py`**

```python
"""Regenerable builder for the synthetic test fixtures.

Run from the repo root:
    python tests/fixtures/make_fixtures.py

Re-creates `tiny_returns.xlsx` from scratch. CSV fixtures are
hand-maintained alongside this script (small enough to read at a glance).
"""
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent


def build_tiny_returns() -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Return"

    # Header: 3 ID columns + 4 date columns (the last two date columns are
    # 20260303, appearing twice on purpose to exercise the dedupe rule).
    ws.append([
        "Name", "BB Ticker", "JSE code",
        20260228, 20260302, 20260303, 20260303,
    ])

    # Three stocks. 20260228 is the base date (all zeros).
    ws.append(["ALPHA LTD",   "AAA SJ Equity", "AAA", 0.0,  0.01, 0.00,  0.99])
    ws.append(["BRAVO LTD",   "BBB SJ Equity", "BBB", 0.0, -0.02, 0.00,  0.99])
    ws.append(["CHARLIE LTD", "CCC SJ Equity", "CCC", 0.0,  0.05, 0.00,  0.99])

    # Published benchmark row — this is the validation target. The Name
    # column is intentionally blank to match the real file's shape.
    ws.append([None, "J803TR Index", None, 0.0, 0.001, 0.000, 0.999])

    # A second index row that the pipeline must ignore (and log to
    # data_quality as `ignored_index_row`).
    ws.append([None, "JSAPY Index", None, 0.0, 0.002, 0.000, 0.999])

    out = HERE / "tiny_returns.xlsx"
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    build_tiny_returns()
```

- [ ] **Step 2: Run the builder**

Run: `python tests/fixtures/make_fixtures.py`
Expected: `wrote tests/fixtures/tiny_returns.xlsx`.

- [ ] **Step 3: Write `tests/fixtures/tiny_weights_both_blocks.csv`**

```
J803,Closing Weights
AAA,0.5
BBB,0.3
CCC,0.2

J803,Opening Weights
AAA,0.45
BBB,0.30
CCC,0.25
```

- [ ] **Step 4: Write `tests/fixtures/tiny_weights_closing_only.csv`**

```
J803,Closing Weights
AAA,0.5
BBB,0.3
CCC,0.2
```

- [ ] **Step 5: Write `tests/fixtures/tiny_weights_bad_sum.csv`**

```
J803,Closing Weights
AAA,0.5
BBB,0.3
CCC,0.05

J803,Opening Weights
AAA,0.45
BBB,0.30
CCC,0.25
```

(Closing block sums to 0.85, well outside the ±0.005 bound — exercises the rejection path.)

- [ ] **Step 6: Write `tests/fixtures/README.md`**

```markdown
# Test fixtures

Synthetic mini-files used by the adapter tests.

- `tiny_returns.xlsx` — regenerable via `python tests/fixtures/make_fixtures.py`.
- `tiny_weights_*.csv` — hand-maintained CSVs covering the three CSV scenarios.
- `holdings_page2_sample.pdf` — a one-page extract of page 2 of one real Catalyst PDF, committed once.
  Regenerate with: `python tests/fixtures/extract_page2.py "data/March PDF/02 March 2026.pdf"`.
```

- [ ] **Step 7: Commit**

```bash
git add tests/fixtures/
git commit -m "Add synthetic fixtures + builder script for adapter tests"
```

---

## Task 4: `attribution.adapters.xlsx_returns`

**Files:**
- Create: `attribution/adapters/xlsx_returns.py`
- Create: `tests/adapters/__init__.py`
- Create: `tests/adapters/test_xlsx_returns.py`

- [ ] **Step 1: Write the failing test `tests/adapters/test_xlsx_returns.py`**

```python
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
```

- [ ] **Step 2: Add `tests/adapters/__init__.py` (empty file)**

- [ ] **Step 3: Run the test to confirm it fails**

Run: `python -m pytest tests/adapters/test_xlsx_returns.py -v`
Expected: ImportError on `attribution.adapters.xlsx_returns`.

- [ ] **Step 4: Implement `attribution/adapters/xlsx_returns.py`**

```python
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


def _parse_date_header(cell) -> pd.Timestamp | None:
    """Return a Timestamp if `cell` parses as YYYYMMDD, else None.

    Accepts integer cells (e.g. 20260302) and string cells (e.g. "20260302").
    """
    if cell is None:
        return None
    s = str(cell).strip()
    if len(s) != 8 or not s.isdigit():
        return None
    try:
        return pd.Timestamp(s)
    except (ValueError, TypeError):
        return None


def read(path: str | Path) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    raw = pd.read_excel(path, sheet_name="Return", header=0)

    id_cols = ["Name", "BB Ticker", "JSE code"]
    missing = [c for c in id_cols if c not in raw.columns]
    if missing:
        raise ValueError(f"Returns sheet missing identifier columns: {missing}")

    data_quality: list[dict] = []

    # Pair each non-id column with the date it parses to (if any).
    date_cols: list[tuple[str, pd.Timestamp]] = []
    seen_dates: dict[pd.Timestamp, str] = {}
    for col in raw.columns:
        if col in id_cols:
            continue
        ts = _parse_date_header(col)
        if ts is None:
            continue
        if ts in seen_dates:
            data_quality.append({
                "date": ts,
                "severity": "warning",
                "category": "duplicate_return_column",
                "ticker": None,
                "note": f"Duplicate column header {col!r}; keeping first occurrence ({seen_dates[ts]!r}).",
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
        .astype({"ticker": str})
        .reset_index(drop=True)
    )
    returns_frame["daily_return"] = returns_frame["daily_return"].astype("float64")

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

    return returns_frame, benchmark_published, data_quality
```

- [ ] **Step 5: Run the test and confirm it passes**

Run: `python -m pytest tests/adapters/test_xlsx_returns.py -v`
Expected: 6 passed.

- [ ] **Step 6: Commit**

```bash
git add attribution/adapters/xlsx_returns.py tests/adapters/
git commit -m "Add xlsx_returns adapter with dedupe + J803TR split + JSAPY ignore"
```

---

## Task 5: `attribution.adapters.csv_weights`

**Files:**
- Create: `attribution/adapters/csv_weights.py`
- Create: `tests/adapters/test_csv_weights.py`

- [ ] **Step 1: Write the failing test `tests/adapters/test_csv_weights.py`**

```python
import pytest

from attribution.adapters import csv_weights


def test_both_blocks_returns_closing_and_opening(fixture_dir):
    result = csv_weights.read(fixture_dir / "tiny_weights_both_blocks.csv")

    assert set(result) == {"closing", "opening"}
    closing = result["closing"]
    assert list(closing.columns) == ["ticker", "weight"]
    assert closing.set_index("ticker")["weight"].to_dict() == {
        "AAA": 0.5, "BBB": 0.3, "CCC": 0.2,
    }
    assert result["opening"].set_index("ticker")["weight"].to_dict() == {
        "AAA": 0.45, "BBB": 0.30, "CCC": 0.25,
    }


def test_closing_only_is_rejected(fixture_dir):
    with pytest.raises(ValueError, match="Opening Weights"):
        csv_weights.read(fixture_dir / "tiny_weights_closing_only.csv")


def test_bad_sum_is_rejected(fixture_dir):
    with pytest.raises(ValueError, match="weight sum"):
        csv_weights.read(fixture_dir / "tiny_weights_bad_sum.csv")
```

- [ ] **Step 2: Run the test and confirm it fails**

Run: `python -m pytest tests/adapters/test_csv_weights.py -v`
Expected: ImportError on `attribution.adapters.csv_weights`.

- [ ] **Step 3: Implement `attribution/adapters/csv_weights.py`**

```python
"""Adapter: J803/J253 daily weights CSV → closing + opening DataFrames.

Each CSV has two blocks separated by a blank row:

    <index>,Closing Weights
    AAA,0.5
    BBB,0.3
    ...
                                  ← blank row
    <index>,Opening Weights
    AAA,0.45
    BBB,0.30
    ...
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from attribution import config

_BLOCK_LABELS = {"Closing Weights": "closing", "Opening Weights": "opening"}


def _validate_block(label: str, df: pd.DataFrame) -> None:
    weight_sum = df["weight"].sum()
    if not (config.WEIGHT_SUM_LOWER <= weight_sum <= config.WEIGHT_SUM_UPPER):
        raise ValueError(
            f"{label.title()} block weight sum {weight_sum:.6f} outside "
            f"[{config.WEIGHT_SUM_LOWER}, {config.WEIGHT_SUM_UPPER}]"
        )


def read(path: str | Path) -> dict[str, pd.DataFrame]:
    """Return `{"closing": df, "opening": df}` with columns `ticker, weight`.

    Raises ValueError if either block is missing or its weight column fails
    the sum-to-1 sanity check.
    """
    text = Path(path).read_text(encoding="utf-8")
    lines = [ln.rstrip("\r\n") for ln in text.splitlines()]

    blocks: dict[str, list[tuple[str, float]]] = {}
    current_label: str | None = None
    for raw in lines:
        if raw.strip() == "":
            current_label = None
            continue

        parts = [p.strip() for p in raw.split(",")]
        if len(parts) != 2:
            continue

        head, body = parts
        if body in _BLOCK_LABELS:
            current_label = _BLOCK_LABELS[body]
            blocks.setdefault(current_label, [])
            continue

        if current_label is None:
            continue

        try:
            weight = float(body)
        except ValueError as e:
            raise ValueError(f"Non-numeric weight {body!r} in {current_label} block") from e
        blocks[current_label].append((head, weight))

    for label in ("closing", "opening"):
        if label not in blocks or not blocks[label]:
            raise ValueError(f"{label.title()} Weights block missing or empty")

    result: dict[str, pd.DataFrame] = {}
    for label, rows in blocks.items():
        df = pd.DataFrame(rows, columns=["ticker", "weight"]).astype(
            {"ticker": str, "weight": "float64"}
        )
        _validate_block(label, df)
        result[label] = df.reset_index(drop=True)

    return result
```

- [ ] **Step 4: Run the test and confirm it passes**

Run: `python -m pytest tests/adapters/test_csv_weights.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add attribution/adapters/csv_weights.py tests/adapters/test_csv_weights.py
git commit -m "Add csv_weights adapter with two-block parsing + sum check"
```

---

## Task 6: `attribution.adapters.pdf_weights`

**Files:**
- Create: `tests/fixtures/extract_page2.py`
- Create (by running script): `tests/fixtures/holdings_page2_sample.pdf`
- Create: `attribution/adapters/pdf_weights.py`
- Create: `tests/adapters/test_pdf_weights.py`

This task has a calibration step. The real PDF page 2 is a wide landscape table; the column boundaries must be tuned against pdfplumber's character positions before the adapter is robust. Plan: spike with `extract_words`/`chars`, identify x-coordinates of column edges, encode them as a module-level constant, then test.

- [ ] **Step 1: Write `tests/fixtures/extract_page2.py`**

```python
"""Extract page 2 of a real Catalyst PDF into a tiny standalone PDF fixture.

Run from the repo root:
    python tests/fixtures/extract_page2.py "data/March PDF/02 March 2026.pdf"
"""
from pathlib import Path
import shutil
import sys
import tempfile

import pdfplumber

OUT = Path(__file__).parent / "holdings_page2_sample.pdf"


def main(src_pdf: str) -> None:
    # pdfplumber can't write PDFs; use pypdf if available, else copy whole
    # PDF as the fixture (still small enough to commit).
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        # Fallback: copy the entire PDF.
        shutil.copy(src_pdf, OUT)
        print(f"wrote {OUT} (full PDF, install pypdf for a single-page extract)")
        return

    reader = PdfReader(src_pdf)
    writer = PdfWriter()
    writer.add_page(reader.pages[1])  # page 2, 0-indexed
    with open(OUT, "wb") as fh:
        writer.write(fh)
    print(f"wrote {OUT} (page 2 only)")


if __name__ == "__main__":
    main(sys.argv[1])
```

- [ ] **Step 2: Install pypdf and run the extractor**

Run: `python -m pip install pypdf`
Run: `python tests/fixtures/extract_page2.py "data/March PDF/02 March 2026.pdf"`
Expected: `wrote tests/fixtures/holdings_page2_sample.pdf (page 2 only)`.

- [ ] **Step 3: Calibration spike — print page 2 layout**

Write and run a throwaway script (or do this interactively in a REPL):

```python
import pdfplumber
import json

with pdfplumber.open("tests/fixtures/holdings_page2_sample.pdf") as pdf:
    pg = pdf.pages[0]
    words = pg.extract_words(x_tolerance=3, y_tolerance=3)
    # Look for the header words: Name, Ticker, Benchmark
    for w in words[:50]:
        print(round(w["x0"], 1), round(w["x1"], 1), round(w["top"], 1), repr(w["text"]))

    # Try extract_tables with default settings:
    tables = pg.extract_tables()
    print(f"{len(tables)} tables found")
    if tables:
        for i, t in enumerate(tables):
            print(f"table {i}: {len(t)} rows x {len(t[0]) if t else 0} cols")
            for row in t[:3]:
                print("  ", row)
```

Goal: identify the **x-coordinate** of the `Benchmark` column header (left edge) and the start/end x of the `Name` and `Ticker` columns. Record these as `_COLUMN_BOUNDS` constants in the implementation below. Also confirm whether `extract_tables()` resolves the table without explicit settings; if not, provide an `explicit_vertical_lines` setting.

This step is exploratory — the engineer prints, examines, and writes down the numbers. No assertion.

- [ ] **Step 4: Write the failing test `tests/adapters/test_pdf_weights.py`**

```python
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
    # All tickers should be 3-4 letters, no 'SJ Equity' suffix.
    assert df["ticker"].str.match(r"^[A-Z]{3,4}$").all()


def test_weight_column_sums_close_to_one(page2_pdf):
    df = pdf_weights.read_page2(page2_pdf)
    s = df["benchmark_weight"].sum()
    assert 0.995 <= s <= 1.005, f"weight sum was {s}"


def test_contains_known_j803_constituents(page2_pdf):
    df = pdf_weights.read_page2(page2_pdf)
    tickers = set(df["ticker"])
    # These are J803 constituents observed in the CSV for 2026-03-02.
    assert {"ATT", "GRT", "NRP", "RDF"}.issubset(tickers)


def test_rejection_path_returns_none(tmp_path):
    """A non-Catalyst PDF must be rejected (return None) so the loader can fall back."""
    # Build a tiny garbage PDF with reportlab if available, else skip.
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
```

- [ ] **Step 5: Run the test and confirm it fails**

Run: `python -m pytest tests/adapters/test_pdf_weights.py -v`
Expected: ImportError on `attribution.adapters.pdf_weights`.

- [ ] **Step 6: Implement `attribution/adapters/pdf_weights.py`**

The implementation depends on the calibration from Step 3. The pattern below uses `extract_words` and a column-by-x-band approach — it's more robust than `extract_tables` for the wide, irregular header. Replace the placeholder x-bounds with the numbers identified in Step 3.

```python
"""Adapter: Catalyst portfolio PDF → benchmark-weight long frame.

Strategy: open page 2, pull all words, then bucket each word into a
column by x-coordinate. Required output columns: ticker, benchmark_weight.

The PDF page is a wide landscape table with wrapped multi-line headers,
so structural table-extraction (`extract_tables`) is unreliable. We rely
on three calibrated x-coordinate bands instead: NAME, TICKER, BENCHMARK.

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
# Numbers below are placeholders — UPDATE FROM CALIBRATION SPIKE.
_NAME_BAND: tuple[float, float] = (0.0, 600.0)
_TICKER_BAND: tuple[float, float] = (600.0, 900.0)
_BENCHMARK_BAND: tuple[float, float] = (900.0, 1100.0)

_HEADER_KEYWORDS = ("Benchmark", "Ticker", "Name")
_TICKER_SUFFIX = " SJ Equity"


def _clean_ticker(raw: str) -> str | None:
    """Map Bloomberg-style ticker to 3-letter JSE code, or return None."""
    if not raw:
        return None
    candidate = raw.replace(_TICKER_SUFFIX, "").strip().upper()
    return candidate if re.fullmatch(r"[A-Z]{3,4}", candidate) else None


def read_page2(path: str | Path) -> pd.DataFrame | None:
    with pdfplumber.open(str(path)) as pdf:
        if len(pdf.pages) < 2:
            log.info("PDF %s has fewer than 2 pages; rejecting.", path)
            return None
        pg = pdf.pages[1]
        words = pg.extract_words(x_tolerance=3, y_tolerance=3)

    text_blob = " ".join(w["text"] for w in words)
    if not all(kw in text_blob for kw in _HEADER_KEYWORDS):
        log.info("Page 2 of %s lacks expected headers; rejecting.", path)
        return None

    # Bucket words by row using `top` coordinate.
    rows: dict[int, list[dict]] = {}
    for w in words:
        key = round(w["top"] / 10) * 10  # 10-pt row bucket
        rows.setdefault(key, []).append(w)

    extracted: list[tuple[str, float]] = []
    for _key, row_words in rows.items():
        ticker_words = [w for w in row_words if _TICKER_BAND[0] <= w["x0"] < _TICKER_BAND[1]]
        bench_words = [w for w in row_words if _BENCHMARK_BAND[0] <= w["x0"] < _BENCHMARK_BAND[1]]
        if not ticker_words or not bench_words:
            continue

        ticker = _clean_ticker(" ".join(w["text"] for w in ticker_words))
        if ticker is None:
            continue

        weight_text = "".join(w["text"] for w in bench_words).replace("%", "").strip()
        try:
            weight = float(weight_text)
        except ValueError:
            continue
        # Page 2 reports weights as percentages or decimals depending on
        # the formatting. Normalize: anything > 1.5 is treated as a %.
        if weight > 1.5:
            weight /= 100.0

        extracted.append((ticker, weight))

    if not extracted:
        log.info("Page 2 of %s yielded no rows after parsing; rejecting.", path)
        return None

    df = pd.DataFrame(extracted, columns=["ticker", "benchmark_weight"])
    # Deduplicate (a multi-line ticker may have produced two row buckets).
    df = df.groupby("ticker", as_index=False)["benchmark_weight"].sum()

    s = df["benchmark_weight"].sum()
    if not (config.WEIGHT_SUM_LOWER <= s <= config.WEIGHT_SUM_UPPER):
        log.info("Page 2 of %s weight sum %s outside bounds; rejecting.", path, s)
        return None

    return df.reset_index(drop=True)
```

- [ ] **Step 7: Run the test and confirm it passes**

Run: `python -m pytest tests/adapters/test_pdf_weights.py -v`
Expected: 5 passed (or 4 passed + 1 skipped if reportlab is unavailable).

If `test_weight_column_sums_close_to_one` or `test_contains_known_j803_constituents` fails, return to Step 3 and refine the calibration constants. The most likely cause is that the x-bands chosen straddle the Benchmark column boundary, so weights get merged with the adjacent `Permitted` column.

- [ ] **Step 8: Commit**

```bash
git add attribution/adapters/pdf_weights.py tests/adapters/test_pdf_weights.py tests/fixtures/
git commit -m "Add pdf_weights adapter with calibrated column bands"
```

---

## Task 7: `attribution.weights.loader`

**Files:**
- Create: `attribution/weights/loader.py`
- Create: `tests/weights/__init__.py`
- Create: `tests/weights/test_loader.py`

- [ ] **Step 1: Write the failing test `tests/weights/test_loader.py`**

```python
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

    # Monkeypatch the adapters: pretend 20260302's PDF parses cleanly,
    # 20260303 has no PDF (we won't even try), and the CSVs always parse
    # cleanly with two stocks summing to 1.
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
```

- [ ] **Step 2: Add `tests/weights/__init__.py` (empty)**

- [ ] **Step 3: Run the test and confirm it fails**

Run: `python -m pytest tests/weights/test_loader.py -v`
Expected: ImportError on `attribution.weights.loader`.

- [ ] **Step 4: Implement `attribution/weights/loader.py`**

```python
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
    return pdf_dir / date.strftime("%d %B %Y") + ".pdf" if False else pdf_dir / f"{date.strftime('%d %B %Y')}.pdf"


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
```

- [ ] **Step 5: Run the test and confirm it passes**

Run: `python -m pytest tests/weights/test_loader.py -v`
Expected: 4 passed.

- [ ] **Step 6: Commit**

```bash
git add attribution/weights/loader.py tests/weights/
git commit -m "Add weights loader with PDF-first source routing"
```

---

## Task 8: `attribution.compute.benchmark`

**Files:**
- Create: `attribution/compute/benchmark.py`
- Create: `tests/compute/__init__.py`
- Create: `tests/compute/test_benchmark.py`

- [ ] **Step 1: Write the failing test `tests/compute/test_benchmark.py`**

```python
import pandas as pd
import pytest

from attribution.compute import benchmark as bm


@pytest.fixture
def two_day_inputs():
    weights = pd.DataFrame({
        "date": [pd.Timestamp("2026-03-02")] * 3 + [pd.Timestamp("2026-03-03")] * 3,
        "ticker": ["AAA", "BBB", "CCC"] * 2,
        "benchmark_weight": [0.5, 0.3, 0.2, 0.6, 0.3, 0.1],
        "weight_source": ["pdf"] * 6,
    })
    returns = pd.DataFrame({
        "date": [pd.Timestamp("2026-03-02")] * 3 + [pd.Timestamp("2026-03-03")] * 3,
        "ticker": ["AAA", "BBB", "CCC"] * 2,
        "daily_return": [0.01, -0.02, 0.05, 0.0, 0.02, -0.01],
    })
    return weights, returns


def test_contribution_is_weight_times_return(two_day_inputs):
    weights, returns = two_day_inputs
    contributions, daily = bm.compute(weights, returns)

    row = contributions[
        (contributions["date"] == pd.Timestamp("2026-03-02"))
        & (contributions["ticker"] == "AAA")
    ].iloc[0]
    assert row["contribution"] == pytest.approx(0.005)


def test_daily_return_sums_contributions(two_day_inputs):
    weights, returns = two_day_inputs
    _, daily = bm.compute(weights, returns)

    day1 = daily[daily["date"] == pd.Timestamp("2026-03-02")].iloc[0]
    assert day1["benchmark_return_computed"] == pytest.approx(0.5*0.01 + 0.3*-0.02 + 0.2*0.05)

    day2 = daily[daily["date"] == pd.Timestamp("2026-03-03")].iloc[0]
    assert day2["benchmark_return_computed"] == pytest.approx(0.6*0.0 + 0.3*0.02 + 0.1*-0.01)


def test_missing_return_propagates_nan(two_day_inputs):
    weights, returns = two_day_inputs
    # Drop CCC's return on day 1.
    returns = returns.drop(returns.index[2])
    contributions, daily = bm.compute(weights, returns)

    day1 = daily[daily["date"] == pd.Timestamp("2026-03-02")].iloc[0]
    assert pd.isna(day1["benchmark_return_computed"]), "missing return must not silently become zero"


def test_weight_source_propagates(two_day_inputs):
    weights, returns = two_day_inputs
    contributions, daily = bm.compute(weights, returns)
    assert (contributions["weight_source"] == "pdf").all()
    assert (daily["weight_source"] == "pdf").all()
```

- [ ] **Step 2: Add `tests/compute/__init__.py` (empty)**

- [ ] **Step 3: Run the test and confirm it fails**

Run: `python -m pytest tests/compute/test_benchmark.py -v`
Expected: ImportError on `attribution.compute.benchmark`.

- [ ] **Step 4: Implement `attribution/compute/benchmark.py`**

```python
"""Compute benchmark contributions and daily benchmark return.

The math:
    contribution[t,i] = benchmark_weight[t,i] * daily_return[t,i]
    benchmark_return[t] = sum over i of contribution[t,i]

Left-join semantics: a ticker present in `weights` but missing from
`returns` produces NaN. The daily benchmark return for any date with a
NaN contribution is itself NaN — we never impute zero.
"""
from __future__ import annotations

import pandas as pd


def compute(
    weights: pd.DataFrame,
    returns: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    contributions = weights.merge(returns, on=["date", "ticker"], how="left")
    contributions["contribution"] = (
        contributions["benchmark_weight"] * contributions["daily_return"]
    )

    contributions = contributions[
        ["date", "ticker", "benchmark_weight", "daily_return", "contribution", "weight_source"]
    ]

    # min_count=1 ensures groupby.sum() with all-NaN returns NaN, not 0.
    daily = (
        contributions.groupby("date", as_index=False)
        .agg(
            benchmark_return_computed=("contribution", lambda s: s.sum(min_count=len(s))),
            weight_source=("weight_source", "first"),
        )
    )
    return contributions, daily
```

Note: the lambda `lambda s: s.sum(min_count=len(s))` forces NaN-on-any-NaN behaviour for each date — exactly the spec rule.

- [ ] **Step 5: Run the test and confirm it passes**

Run: `python -m pytest tests/compute/test_benchmark.py -v`
Expected: 4 passed.

- [ ] **Step 6: Commit**

```bash
git add attribution/compute/benchmark.py tests/compute/
git commit -m "Add compute.benchmark: contributions + daily benchmark return"
```

---

## Task 9: `attribution.validate.published`

**Files:**
- Create: `attribution/validate/published.py`
- Create: `tests/validate/__init__.py`
- Create: `tests/validate/test_published.py`

- [ ] **Step 1: Write the failing test `tests/validate/test_published.py`**

```python
import pandas as pd

from attribution.validate import published


def _computed(values):
    return pd.DataFrame({
        "date": pd.to_datetime(["2026-03-02", "2026-03-03", "2026-03-04"]),
        "benchmark_return_computed": values,
        "weight_source": ["pdf"] * 3,
    })


def _published(values):
    return pd.DataFrame({
        "date": pd.to_datetime(["2026-03-02", "2026-03-03", "2026-03-04"]),
        "benchmark_return_published": values,
    })


def test_exact_match_does_not_exceed():
    out = published.validate(_computed([0.01, 0.02, 0.03]), _published([0.01, 0.02, 0.03]))
    assert out["exceeds_tolerance"].sum() == 0


def test_diff_just_below_tolerance_does_not_exceed():
    out = published.validate(
        _computed([0.01 + 5e-8, 0.02, 0.03]),
        _published([0.01, 0.02, 0.03]),
    )
    assert bool(out.iloc[0]["exceeds_tolerance"]) is False


def test_diff_at_tolerance_does_not_exceed():
    # Tolerance is `|diff| > 1e-7`, so exactly at 1e-7 must NOT exceed.
    out = published.validate(
        _computed([0.01 + 1e-7, 0.02, 0.03]),
        _published([0.01, 0.02, 0.03]),
    )
    assert bool(out.iloc[0]["exceeds_tolerance"]) is False


def test_diff_just_above_tolerance_exceeds():
    out = published.validate(
        _computed([0.01 + 2e-7, 0.02, 0.03]),
        _published([0.01, 0.02, 0.03]),
    )
    assert bool(out.iloc[0]["exceeds_tolerance"]) is True


def test_diff_bp_is_decimal_times_10000():
    out = published.validate(_computed([0.0101, 0.02, 0.03]), _published([0.01, 0.02, 0.03]))
    assert out.iloc[0]["diff_bp"] == 1.0  # 0.0001 → 1 bp
```

- [ ] **Step 2: Add `tests/validate/__init__.py` (empty)**

- [ ] **Step 3: Run the test and confirm it fails**

Run: `python -m pytest tests/validate/test_published.py -v`
Expected: ImportError on `attribution.validate.published`.

- [ ] **Step 4: Implement `attribution/validate/published.py`**

```python
"""Validate computed daily benchmark return against the published index series.

In v1 the published target is the `J803TR Index` row of the returns
xlsx; this module is generic over which series it receives and does not
encode the index code.
"""
from __future__ import annotations

import pandas as pd

from attribution import config


def validate(
    computed: pd.DataFrame,
    published: pd.DataFrame,
) -> pd.DataFrame:
    merged = computed.merge(published, on="date", how="outer")
    merged["diff_decimal"] = (
        merged["benchmark_return_computed"] - merged["benchmark_return_published"]
    )
    merged["diff_bp"] = merged["diff_decimal"] * 10_000.0
    merged["exceeds_tolerance"] = (
        merged["diff_decimal"].abs() > config.TOLERANCE_DECIMAL
    )
    return merged[[
        "date",
        "benchmark_return_computed",
        "benchmark_return_published",
        "diff_decimal",
        "diff_bp",
        "exceeds_tolerance",
        "weight_source",
    ]]
```

- [ ] **Step 5: Run the test and confirm it passes**

Run: `python -m pytest tests/validate/test_published.py -v`
Expected: 5 passed.

- [ ] **Step 6: Commit**

```bash
git add attribution/validate/published.py tests/validate/
git commit -m "Add validate.published comparing computed vs published index"
```

---

## Task 10: `attribution.report.excel`

**Files:**
- Create: `attribution/report/excel.py`
- Create: `tests/report/__init__.py`
- Create: `tests/report/test_excel.py`

- [ ] **Step 1: Write the failing test `tests/report/test_excel.py`**

```python
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
```

- [ ] **Step 2: Add `tests/report/__init__.py` (empty)**

- [ ] **Step 3: Run the test and confirm it fails**

Run: `python -m pytest tests/report/test_excel.py -v`
Expected: ImportError on `attribution.report.excel`.

- [ ] **Step 4: Implement `attribution/report/excel.py`**

```python
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
```

- [ ] **Step 5: Run the test and confirm it passes**

Run: `python -m pytest tests/report/test_excel.py -v`
Expected: 2 passed.

- [ ] **Step 6: Commit**

```bash
git add attribution/report/excel.py tests/report/
git commit -m "Add report.excel: 5-sheet workbook + parquet sidecar"
```

---

## Task 11: CLI + `__main__`

**Files:**
- Create: `attribution/cli.py`
- Create: `attribution/__main__.py`
- Create: `tests/test_cli_smoke.py`

- [ ] **Step 1: Write the failing test `tests/test_cli_smoke.py`**

```python
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

from attribution.adapters import xlsx_returns
from attribution.compute import benchmark as bm


def test_end_to_end_on_synthetic_month(tmp_path, fixture_dir, monkeypatch):
    """Run the CLI end-to-end against tiny synthetic inputs.

    Layout:
      tmp/
        returns.xlsx                  ← copy of tiny_returns.xlsx
        pdfs/                          ← empty (forces CSV fallback)
        csvs/2026-03/J803/J803_20260302.csv, J803_20260303.csv
        out/
    """
    import shutil

    pdf_dir = tmp_path / "pdfs"
    pdf_dir.mkdir()
    csv_root = tmp_path / "csvs"
    csv_dir = csv_root / "2026-03" / "J803"
    csv_dir.mkdir(parents=True)
    out_dir = tmp_path / "out"

    # Returns file: copy tiny fixture.
    returns_path = tmp_path / "returns.xlsx"
    shutil.copy(fixture_dir / "tiny_returns.xlsx", returns_path)

    # CSVs for the two non-base dates in tiny_returns.xlsx (20260302, 20260303).
    csv_body = (
        "J803,Closing Weights\n"
        "AAA,0.4\n"
        "BBB,0.3\n"
        "CCC,0.3\n"
        "\n"
        "J803,Opening Weights\n"
        "AAA,0.4\n"
        "BBB,0.3\n"
        "CCC,0.3\n"
    )
    (csv_dir / "J803_20260302.csv").write_text(csv_body)
    (csv_dir / "J803_20260303.csv").write_text(csv_body)

    result = subprocess.run(
        [
            sys.executable, "-m", "attribution", "run",
            "--month", "2026-03",
            "--benchmark", "J803",
            "--pdf-dir", str(pdf_dir),
            "--csv-dir", str(csv_dir),
            "--returns", str(returns_path),
            "--out", str(out_dir),
        ],
        capture_output=True, text=True,
    )

    # Exit code may be 2 (validation breaches likely with random tiny weights),
    # but never 1.
    assert result.returncode in (0, 2), result.stderr

    out_xlsx = out_dir / "benchmark_attribution_2026-03.xlsx"
    assert out_xlsx.exists(), f"missing output: stderr={result.stderr}"
    wb = openpyxl.load_workbook(out_xlsx)
    assert "daily_benchmark_return" in wb.sheetnames
    assert "contributions" in wb.sheetnames
```

- [ ] **Step 2: Run the test and confirm it fails**

Run: `python -m pytest tests/test_cli_smoke.py -v`
Expected: `No module named attribution.__main__` or argparse failure.

- [ ] **Step 3: Implement `attribution/__main__.py`**

```python
from attribution.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Implement `attribution/cli.py`**

```python
"""CLI entry point: `python -m attribution run …`.

Exit codes:
  0 — success, all days within tolerance.
  1 — hard error (e.g. missing returns xlsx).
  2 — completed with tolerance breaches and/or data-quality warnings.
"""
from __future__ import annotations

import argparse
import logging
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

from attribution import config
from attribution.adapters import xlsx_returns
from attribution.compute import benchmark as bm
from attribution.report import excel as report_excel
from attribution.validate import published as published_v
from attribution.weights import loader as weights_loader

log = logging.getLogger("attribution")


def _git_sha() -> str:
    try:
        out = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True, text=True, check=False,
        )
        return out.stdout.strip() or "unknown"
    except FileNotFoundError:
        return "unknown"


def _trading_dates_for_month(returns_frame: pd.DataFrame, month: str) -> list[pd.Timestamp]:
    """Return distinct, sorted dates in `returns_frame` that fall in `month`.

    `month` is in YYYY-MM form. Excludes the base date (assumed to be the
    last day of the prior month).
    """
    year, mo = (int(x) for x in month.split("-"))
    dates = returns_frame.loc[
        (returns_frame["date"].dt.year == year)
        & (returns_frame["date"].dt.month == mo),
        "date",
    ].drop_duplicates().sort_values().tolist()
    return dates


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="attribution")
    sub = parser.add_subparsers(dest="command", required=True)

    run = sub.add_parser("run", help="Run the benchmark attribution pipeline")
    run.add_argument("--month", required=True, help="Month in YYYY-MM form, e.g. 2026-03")
    run.add_argument("--benchmark", default=config.DEFAULT_BENCHMARK)
    run.add_argument("--pdf-dir", type=Path, default=config.DEFAULT_PDF_DIR)
    run.add_argument("--csv-dir", type=Path, default=None,
                     help="Defaults to <DEFAULT_CSV_ROOT>/<month>/<benchmark>")
    run.add_argument("--returns", type=Path, default=config.DEFAULT_RETURNS)
    run.add_argument("--out", type=Path, default=config.DEFAULT_OUT_DIR)
    run.add_argument("--tolerance", type=float, default=config.TOLERANCE_DECIMAL)
    run.add_argument("--log-level", default="INFO")

    args = parser.parse_args(argv)
    logging.basicConfig(level=args.log_level, format="%(levelname)s %(name)s: %(message)s")

    if args.tolerance != config.TOLERANCE_DECIMAL:
        config.TOLERANCE_DECIMAL = args.tolerance

    csv_dir = args.csv_dir or config.csv_dir_for(args.month, args.benchmark)

    log.info("Reading returns from %s", args.returns)
    try:
        returns, published, returns_dq = xlsx_returns.read(args.returns)
    except Exception as e:
        log.error("Failed to read returns xlsx: %s", e)
        return 1

    dates = _trading_dates_for_month(returns, args.month)
    if not dates:
        log.error("No returns columns found for month %s", args.month)
        return 1

    log.info("Loading weights for %d trading dates", len(dates))
    weights, weights_dq = weights_loader.load(
        dates=dates, pdf_dir=args.pdf_dir, csv_dir=csv_dir, benchmark=args.benchmark,
    )

    if weights.empty:
        log.error("Loader produced zero weight rows; nothing to compute.")
        return 1

    contributions, daily_computed = bm.compute(weights, returns)
    daily_full = published_v.validate(daily_computed, published)

    # Per-row data quality: missing returns.
    missing = contributions[contributions["daily_return"].isna()]
    missing_dq = [
        {"date": r["date"], "severity": "warning", "category": "missing_return",
         "ticker": r["ticker"], "note": "Weight present but no return in xlsx."}
        for r in missing.to_dict(orient="records")
    ]

    dq_rows = returns_dq + weights_dq + missing_dq
    dq_frame = pd.DataFrame(dq_rows) if dq_rows else pd.DataFrame(
        columns=["date", "severity", "category", "ticker", "note"]
    )

    n_breaches = int(daily_full["exceeds_tolerance"].fillna(False).sum())
    cum_computed = float((1 + daily_full["benchmark_return_computed"].fillna(0)).prod() - 1)
    cum_published = float((1 + daily_full["benchmark_return_published"].fillna(0)).prod() - 1)

    summary = {
        "month": args.month,
        "benchmark": args.benchmark,
        "n_days": len(dates),
        "cumulative_return_computed": cum_computed,
        "cumulative_return_published": cum_published,
        "max_abs_diff_bp": float(daily_full["diff_bp"].abs().max()) if len(daily_full) else 0.0,
        "n_breaches": n_breaches,
        "n_dq_issues": len(dq_rows),
    }

    exit_code = 2 if (n_breaches > 0 or dq_rows) else 0

    metadata = {
        "pdf_dir": str(args.pdf_dir),
        "csv_dir": str(csv_dir),
        "returns": str(args.returns),
        "tolerance": args.tolerance,
        "git_sha": _git_sha(),
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "exit_code": exit_code,
    }

    out_xlsx = args.out / f"benchmark_attribution_{args.month}.xlsx"
    report_excel.write(
        out_xlsx,
        contributions=contributions,
        daily=daily_full,
        data_quality=dq_frame,
        summary=summary,
        metadata=metadata,
    )
    log.info("Wrote %s (exit %d)", out_xlsx, exit_code)
    return exit_code
```

- [ ] **Step 5: Run the test and confirm it passes**

Run: `python -m pytest tests/test_cli_smoke.py -v`
Expected: 1 passed.

- [ ] **Step 6: Run the full test suite**

Run: `python -m pytest -v`
Expected: All tests pass (one possibly skipped if reportlab missing).

- [ ] **Step 7: Commit**

```bash
git add attribution/cli.py attribution/__main__.py tests/test_cli_smoke.py
git commit -m "Add CLI entry point and end-to-end smoke test"
```

---

## Task 12: Real-data validation run

**Files:**
- Modify: `README.md` (create with usage notes if it doesn't yet exist)

- [ ] **Step 1: Run the pipeline against the real March 2026 inputs**

Run:
```
python -m attribution run --month 2026-03 --benchmark J803 --log-level INFO
```

Expected: an `output/benchmark_attribution_2026-03.xlsx` file appears. Exit code is 0 or 2 (never 1).

- [ ] **Step 2: Inspect the output**

Open `output/benchmark_attribution_2026-03.xlsx`:
- `summary`: `n_days == 22`, `cumulative_return_computed` close to `cumulative_return_published`.
- `daily_benchmark_return`: 22 rows, `weight_source` mostly `"pdf"` for 02–27 March and `"csv_close"` for 30–31 March.
- `validation` (within `daily_benchmark_return`): rows where `exceeds_tolerance == True` correspond to days with PDF-rounded weights. Note these in the README troubleshooting section.
- `contributions`: 22 × 21 ≈ 462 rows (one per ticker per day).
- `data_quality`: any missing-return rows or PDF rejections.

- [ ] **Step 3: Write `README.md` with one short usage block**

```markdown
# Attribution Pipeline

Catalyst benchmark attribution pipeline.

## Quick start

    pip install -e ".[dev]"
    python -m attribution run --month 2026-03

The output workbook lands in `output/benchmark_attribution_2026-03.xlsx`.

Exit codes:
- `0` — success, every day within tolerance.
- `1` — hard error (missing inputs, malformed returns xlsx).
- `2` — completed with one or more tolerance breaches or data-quality warnings.

Spec: `docs/superpowers/specs/2026-05-12-benchmark-attribution-pipeline-design.md`.
```

- [ ] **Step 4: Commit**

```bash
git add README.md
git commit -m "Add README with quick-start usage"
```

---

## Self-Review

**Spec coverage:** Every spec section has a task implementing it: §2 inputs → tasks 4–7; §3 architecture → task 1 layout; §4 data contracts → adapter + compute + validate tasks; §5 adapters → tasks 4–6; §6 loader → task 7; §7 compute → task 8; §8 validation → task 9; §9 error handling → cli.py + loader behaviours; §10 CLI → task 11; §11 output → task 10; §12 testing → covered per task; §13 config → task 2; §14 deps → pyproject in task 1.

**Placeholder scan:** No TBDs in steps. One *deliberate* calibration step in Task 6 (Step 3) — that's an exploratory action with concrete output (x-coord constants to plug into Step 6); not a placeholder, but the engineer must do it.

**Type consistency:** `weights_frame` columns `(date, ticker, benchmark_weight, weight_source)` match across loader output, compute input, and contributions/daily output. `returns_frame` columns `(date, ticker, daily_return)` match across xlsx_returns output and compute input. The `weight_source` propagation is tested in Task 8 Step 1.
