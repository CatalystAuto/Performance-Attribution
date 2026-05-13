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
