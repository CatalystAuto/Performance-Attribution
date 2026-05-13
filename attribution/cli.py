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
    """Distinct, sorted dates in `returns_frame` falling in `month` (YYYY-MM)."""
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
