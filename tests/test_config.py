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
