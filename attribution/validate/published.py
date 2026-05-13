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
