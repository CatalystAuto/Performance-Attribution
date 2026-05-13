import pandas as pd
import pytest

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
    assert out.iloc[0]["diff_bp"] == pytest.approx(1.0)  # 0.0001 → 1 bp
