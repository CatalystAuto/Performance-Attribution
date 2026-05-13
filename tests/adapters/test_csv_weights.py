import pytest

from attribution.adapters import csv_weights


def test_both_blocks_returns_closing_and_opening(fixture_dir):
    result = csv_weights.read(fixture_dir / "tiny_weights_both_blocks.csv")

    assert set(result) == {"closing", "opening"}
    closing = result["closing"]
    assert list(closing.columns) == ["ticker", "weight"]
    assert closing.set_index("ticker")["weight"].to_dict() == {
        "AAA": 0.5, "BBB": 0.3, "CCC": 0.2,
    }
    assert result["opening"].set_index("ticker")["weight"].to_dict() == {
        "AAA": 0.45, "BBB": 0.30, "CCC": 0.25,
    }


def test_closing_only_is_rejected(fixture_dir):
    with pytest.raises(ValueError, match="Opening Weights"):
        csv_weights.read(fixture_dir / "tiny_weights_closing_only.csv")


def test_bad_sum_is_rejected(fixture_dir):
    with pytest.raises(ValueError, match="weight sum"):
        csv_weights.read(fixture_dir / "tiny_weights_bad_sum.csv")
