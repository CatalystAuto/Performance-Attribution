# Benchmark Attribution Pipeline — Design

- **Date:** 2026-05-12
- **Status:** Draft (awaiting user review)
- **Scope:** Benchmark side only. Portfolio holdings and Brinson decomposition are explicitly out of scope and follow in later sub-projects.

## 1. Goal

Build a Python package + CLI that, for a given month, computes:

1. The daily benchmark return as `Σ(weight × daily_return)` over the benchmark's constituents.
2. The per-stock contribution `weight × daily_return` (the row-level table later Brinson stages will consume).

Then validate the computed daily benchmark return against the published `JSAPY Index` return contained in the returns workbook, flagging any day where `|computed − published| > 1e-7` (0.001 bp).

The first month of operation is **March 2026** against benchmark **J803**.

## 2. Data sources

### 2.1 Returns — `data/March Stock Return.xlsx`

- Single sheet, `Return`.
- Wide layout: rows = stocks, columns = dates. First three columns are identifiers (`Name`, `BB Ticker`, `JSE code`); remaining columns are daily returns headed `YYYYMMDD`.
- Returns are simple daily returns in decimal form (e.g., `-0.0188` for −1.88%).
- The first date column (`20260228`) is a base date and all return values are `0`.
- The last row is the published benchmark return labelled `JSAPY Index` in the `BB Ticker` column.
- Known data-quality issues:
  - `20260331` appears as a duplicate column header. Resolution rule: keep the first occurrence, record the duplicate in `data_quality`.
  - The xlsx universe (25 stocks) is a superset of the J803 universe (21 stocks in current CSVs). The extras are non-benchmark names held by portfolios and are ignored on the benchmark side.

### 2.2 Primary weights — daily portfolio holdings PDFs

- Folder: `data/March PDF/` containing files named `DD March 2026.pdf`.
- Page 2 of each PDF is the holdings sheet. The page is landscape, very wide (4953 × 3499 pt), with fully selectable text (24,208 characters in the inspected sample; zero rasterized images).
- Required columns: `Name`, `Ticker`, `Benchmark` (= benchmark weight in decimal). Other columns on page 2 (`Active Weighting`, `Over/Under`, prices, units, etc.) are ignored in v1 but will be relevant when the portfolio side lands.
- `Ticker` on page 2 is the Bloomberg ticker (`XXX SJ Equity` style). The adapter strips the ` SJ Equity` suffix to derive the 3-letter JSE code used as the canonical join key.
- Coverage in the current folder: 02 March 2026 through 27 March 2026 (20 trading days). 30 + 31 March 2026 are not present in PDF form — those days use the CSV fallback.

### 2.3 Backup weights — `data/Benchmark Weights/<YYYY-MM>/<index>/<index>_<YYYYMMDD>.csv`

- For March 2026 + J803: `data/Benchmark Weights/2026-03/J803/J803_YYYYMMDD.csv`, one CSV per trading day. 22 files present, covering every trading day of the month including 30 + 31 March.
- Each CSV contains two blocks separated by a blank row:
  - Block 1: header row `<index_code>,Closing Weights` then `JSE_code,weight` rows.
  - Block 2: header row `<index_code>,Opening Weights` then `JSE_code,weight` rows.
- The pipeline uses **closing** weights as the fallback source. Opening weights are read by the adapter and exposed for future use but unused in v1.
- The CSV's JSE code (e.g., `ATT`) matches the `JSE code` column in the returns xlsx — no translation required.

### 2.4 Index choice

J803 is the benchmark for March 2026. `--benchmark J803` is the default in v1. The CLI accepts `J253` for symmetry, but no J253 acceptance work is in scope for v1.

## 3. Architecture

```
attribution/
├── adapters/
│   ├── pdf_weights.py
│   ├── csv_weights.py
│   └── xlsx_returns.py
├── weights/
│   └── loader.py
├── compute/
│   └── benchmark.py
├── validate/
│   └── jsapy.py
├── report/
│   └── excel.py
├── cli.py
└── config.py
tests/
docs/superpowers/specs/
output/                 # gitignored
```

Each adapter converts one source format into a canonical pandas DataFrame. Compute / validate / report consume only those canonical shapes; they never see raw source formats. Adding a new data source = adding one new adapter file.

## 4. Canonical data contracts

### 4.1 `returns_frame`
| column         | dtype          | notes                                 |
| -------------- | -------------- | ------------------------------------- |
| `date`         | datetime64[ns] | normalized to date, no tz             |
| `ticker`       | str            | 3-letter JSE code                     |
| `daily_return` | float64        | simple return in decimal              |

### 4.2 `benchmark_published`
| column                       | dtype          | notes                              |
| ---------------------------- | -------------- | ---------------------------------- |
| `date`                       | datetime64[ns] |                                    |
| `benchmark_return_published` | float64        | from the `JSAPY Index` row of xlsx |

### 4.3 `weights_frame`
| column             | dtype          | notes                          |
| ------------------ | -------------- | ------------------------------ |
| `date`             | datetime64[ns] |                                |
| `ticker`           | str            | 3-letter JSE code              |
| `benchmark_weight` | float64        | decimal, sums to ~1.0 per date |
| `weight_source`    | str            | `"pdf"` or `"csv_close"`       |

### 4.4 `contributions_frame` (produced by `compute.benchmark`)
| column             | dtype          | notes                                       |
| ------------------ | -------------- | ------------------------------------------- |
| `date`             | datetime64[ns] |                                             |
| `ticker`           | str            |                                             |
| `benchmark_weight` | float64        |                                             |
| `daily_return`     | float64        | may be `NaN` if return missing for ticker   |
| `contribution`     | float64        | `benchmark_weight * daily_return`           |
| `weight_source`    | str            | propagated from `weights_frame`             |

### 4.5 `daily_benchmark_frame`
| column                       | dtype          | notes                                 |
| ---------------------------- | -------------- | ------------------------------------- |
| `date`                       | datetime64[ns] |                                       |
| `benchmark_return_computed`  | float64        | `Σ contribution` over the date        |
| `benchmark_return_published` | float64        | outer-joined from `benchmark_published` |
| `diff_decimal`               | float64        | computed − published                  |
| `diff_bp`                    | float64        | `diff_decimal * 10_000`               |
| `exceeds_tolerance`          | bool           | `\|diff_decimal\| > 1e-7`             |
| `weight_source`              | str            | source label for the date (see §6.3)  |

## 5. Adapter responsibilities

### 5.1 `xlsx_returns`

- Read the single `Return` sheet.
- Identify the date columns: header values that parse as 8-digit dates in `YYYYMMDD` form, accepting either integer cells (e.g., `20260302`) or string cells (e.g., `"20260302"`). Convert to `datetime64`.
- Resolve duplicate date columns: keep the first occurrence; record each duplicate as a `data_quality` row with `category="duplicate_return_column"`.
- Identify the index row by `BB Ticker == "JSAPY Index"`; split it out into `benchmark_published`.
- Melt the remaining stock rows from wide to long, dropping the base date (28 Feb) returns where requested by the loader, and rename `JSE code` → `ticker`.
- Return `(returns_frame, benchmark_published, data_quality_rows)`.

### 5.2 `csv_weights`

- Read a single CSV. Split on the blank row into two blocks.
- Parse each block: first row is the header (`<index>,Closing Weights` or `<index>,Opening Weights`); subsequent rows are `(ticker, weight)`.
- Return both blocks as separate DataFrames keyed by source label.
- Reject and raise if either block is missing, malformed, or the weight column does not sum to within ±0.005 of 1.0.

### 5.3 `pdf_weights`

- Open the PDF with pdfplumber; read page 2 (`pdf.pages[1]`).
- Extract the table using `page.extract_tables(...)` with column-boundary tuning calibrated against the inspected sample. Header rows wrap across multiple lines; the adapter needs to consume the wrapped header and identify the columns `Name`, `Ticker`, `Benchmark` by best-match.
- Strip ` SJ Equity` from `Ticker` to yield the 3-letter JSE code.
- Reject the page (return empty) if any of the following hold:
  - The `Benchmark` column cannot be located.
  - Any row has a non-numeric weight after coercion.
  - The weight column does not sum to within ±0.005 of 1.0.
- Calibration of column boundaries is a one-time step done during implementation; the calibrated boundaries live in `pdf_weights.py` as a module-level constant.

## 6. Loader (`weights/loader.py`)

### 6.1 Inputs
- A list of trading dates to load. The CLI builds this list from the date columns present in the returns xlsx for the requested `--month`, excluding the base date (28 Feb for the March run) unless explicitly opted in. This makes the returns xlsx the calendar of record; PDFs and CSVs are sources, not calendars.
- Paths to the PDF folder and the CSV folder.
- Benchmark code (e.g., `"J803"`).

### 6.2 Per-date precedence
1. Attempt PDF extraction. If the adapter returns a non-empty, validated frame, use it with `weight_source = "pdf"`.
2. Otherwise attempt CSV-closing extraction. If the adapter returns a validated frame, use it with `weight_source = "csv_close"`.
3. Otherwise record the date as `status="no_source"` in `data_quality` and skip it.

### 6.3 Mixing rule

A single date never mixes sources. `weight_source` is the same for every row of a given date, and is propagated unchanged into the per-date `daily_benchmark_frame` row (`"pdf"` or `"csv_close"`).

### 6.4 Output

A `weights_frame` covering the dates for which a source was found, plus a list of `data_quality` rows describing rejections and misses.

## 7. Compute (`compute/benchmark.py`)

```python
contributions = weights.merge(returns, on=["date", "ticker"], how="left")
contributions["contribution"] = (
    contributions["benchmark_weight"] * contributions["daily_return"]
)
daily = contributions.groupby("date", as_index=False)["contribution"].sum()
```

Rules:
- **Left join, not inner.** A ticker present in weights but missing from returns produces `daily_return = NaN`, `contribution = NaN`, and a `NaN` benchmark return for that date. Missing returns are recorded in `data_quality` with `category="missing_return"`. We do not impute zero — that would bias the benchmark return downward and mask the data issue.
- **No weight renormalization.** Weights are used as reported. If the source sums to 0.9998 instead of 1.0000, the computed return reflects that. Renormalization would mask weight-extraction bugs and obscure the comparison to the published JSAPY return.
- **Base date.** The 28 Feb base date is included in the run only if explicitly requested via the trading-date list; by default the CLI runs March-only dates. When included, all return values are 0 and contributions are 0.

Cumulative return reported in the `summary` sheet is **compound**: `Π(1 + r_t) − 1`.

## 8. Validation (`validate/jsapy.py`)

```python
val = (
    daily_computed.merge(published, on="date", how="outer")
    .assign(diff_decimal=lambda d: d.benchmark_return_computed - d.benchmark_return_published)
    .assign(diff_bp=lambda d: d.diff_decimal * 10_000)
    .assign(exceeds_tolerance=lambda d: d.diff_decimal.abs() > config.TOLERANCE_DECIMAL)
)
```

- `TOLERANCE_DECIMAL = 1e-7` (0.001 bp). Held in `config.py`; overridable via `--tolerance`.
- Note: 1e-7 is tighter than the precision implied by PDF weight rounding (typically 4 decimal places). Frequent breaches are expected on PDF-sourced days; this is intentional — the tolerance is a visibility tool, not a hard correctness contract for v1.

## 9. Error handling

| Category               | Example                                                | Behaviour                                                                   |
| ---------------------- | ------------------------------------------------------ | --------------------------------------------------------------------------- |
| **Hard error**         | Returns xlsx missing or unparseable                    | Raise, exit `1`, do not write partial output                                |
| **Adapter rejection**  | PDF page-2 weight sum outside ±0.005 of 1.0            | Adapter returns empty; loader falls back to CSV; logged in `data_quality`   |
| **Loader miss**        | No PDF *and* no CSV for a date                         | Date skipped; logged with `status="no_source"`; pipeline continues          |
| **Validation breach**  | `\|diff\|` > 1e-7 on one or more days                  | Run completes, writes report, exits `2`                                     |

Exit codes:
- `0` — success, every day within tolerance.
- `1` — hard error.
- `2` — completed with one or more tolerance breaches *or* data-quality warnings.

Logging via the Python `logging` module. CLI flag `--log-level` (default `INFO`).

## 10. CLI

Entry point: `python -m attribution run`.

```
python -m attribution run \
    --month       2026-03 \
    --benchmark   J803 \
    --pdf-dir     "data/March PDF" \
    --csv-dir     "data/Benchmark Weights/2026-03/J803" \
    --returns     "data/March Stock Return.xlsx" \
    --out         "output/" \
    --tolerance   1e-7 \
    --log-level   INFO
```

All paths have defaults in `config.py` that resolve to the project-relative locations above so the bare command `python -m attribution run --month 2026-03` works.

Only the `run` sub-command exists in v1. Anything else (`refresh-cache`, `validate`, etc.) is YAGNI.

## 11. Output

One Excel workbook per run: `output/benchmark_attribution_<YYYY-MM>.xlsx`.

Sheets:

| Sheet                    | Contents                                                                                              |
| ------------------------ | ----------------------------------------------------------------------------------------------------- |
| `summary`                | Month, benchmark, # days, cumulative return computed vs published, max abs diff (bp), # breaches, # data-quality issues |
| `daily_benchmark_return` | `date, return_computed, return_published, diff_bp, exceeds_tolerance, weight_source`                  |
| `contributions`          | `date, ticker, benchmark_weight, daily_return, contribution, weight_source`                           |
| `data_quality`           | `date, severity, category, ticker, note`                                                              |
| `run_metadata`           | Resolved input paths, git SHA, tolerance, timestamp, exit code                                        |

A sidecar `output/benchmark_attribution_<YYYY-MM>_contributions.parquet` is also written for cheap programmatic re-use by future Brinson stages.

## 12. Testing

- **Adapter unit tests** with fixture files in `tests/fixtures/`:
  - `pdf_weights`: a snapshot of page 2 of one real PDF, plus synthetic minimal PDFs covering one good case, one malformed-but-recoverable case, and one rejection case.
  - `csv_weights`: hand-written tiny CSVs covering closing-only, closing+opening, and malformed-block cases.
  - `xlsx_returns`: a tiny xlsx with 3 stocks × 3 dates + a JSAPY row, including a duplicate column to exercise the dedupe rule.
- **Loader tests** for the source-routing precedence (PDF wins; PDF rejected → CSV; both missing → skip).
- **Compute tests** with hand-calculated golden values for a 3×3 mini scenario.
- **Validation tests** for tolerance-edge cases (exactly at, just over, just under 1e-7).
- **One end-to-end smoke test** that runs the CLI on a 3-day fixture month and asserts the output workbook has the expected sheets with non-empty content.

Tooling: `pytest`, `pytest-cov`. No mocking of file I/O — tests use real fixture files.

Coverage targets:
- `compute/`, `validate/`, `weights/loader.py`: > 90%.
- Adapters: characterization tests rather than line-coverage targets — the variance comes from the source format, not the adapter logic.

## 13. Configuration (`config.py`)

A single module of constants:

```python
TOLERANCE_DECIMAL = 1e-7
WEIGHT_SUM_LOWER  = 0.995
WEIGHT_SUM_UPPER  = 1.005
DEFAULT_BENCHMARK = "J803"
DEFAULT_PDF_DIR   = Path("data/March PDF")
DEFAULT_CSV_ROOT  = Path("data/Benchmark Weights")
DEFAULT_RETURNS   = Path("data/March Stock Return.xlsx")
DEFAULT_OUT_DIR   = Path("output")
```

The CSV directory for a given run is resolved as `DEFAULT_CSV_ROOT / month / benchmark`.

## 14. Dependencies

- `pandas` (canonical frames)
- `openpyxl` (xlsx read + Excel report write)
- `pdfplumber` (already installed at 0.11.9)
- `pyarrow` (Parquet sidecar)
- `pytest`, `pytest-cov` (dev only)

A `requirements.txt` (or `pyproject.toml`) pins these. Python 3.14 is the current interpreter; the package targets >= 3.11.

## 15. Out of scope (explicit)

- Portfolio-side weights, transactions, cash, and corporate-action handling.
- Brinson allocation / selection / interaction decomposition.
- Multi-period linking algorithms (Carino, geometric, etc.).
- Universe reconciliation across benchmark and portfolio.
- Any UI beyond the CLI; no web report, no dashboard.
- Backfill of pre-March 2026 months (the architecture supports it; v1 only ships March 2026).
- J253 acceptance testing (the CLI accepts `--benchmark J253` but it is not validated in v1).

## 16. Open questions deferred to implementation

- Exact PDF column-boundary heuristics for page 2. To be calibrated against the 20 real PDFs during implementation; if `pdfplumber.extract_tables` proves unreliable for that page, fall back to character-level positional parsing of the same page. Either way, the adapter contract does not change.
- Whether a future `--use-opening-weights` flag is worth surfacing. Not in v1.
