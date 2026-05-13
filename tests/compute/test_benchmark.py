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
