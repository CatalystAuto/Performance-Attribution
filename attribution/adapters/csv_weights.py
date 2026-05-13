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
