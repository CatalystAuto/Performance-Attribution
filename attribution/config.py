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
