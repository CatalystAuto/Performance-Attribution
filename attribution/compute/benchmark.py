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

    # min_count=len(s) ensures groupby.sum() with any-NaN returns NaN, not 0.
    daily = (
        contributions.groupby("date", as_index=False)
        .agg(
            benchmark_return_computed=("contribution", lambda s: s.sum(min_count=len(s))),
            weight_source=("weight_source", "first"),
        )
    )
    return contributions, daily
