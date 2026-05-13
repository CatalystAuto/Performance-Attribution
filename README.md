# Attribution Pipeline

Catalyst benchmark attribution pipeline.

## Quick start

    pip install -e ".[dev]"
    python -m attribution run --month 2026-03

The output workbook lands in `output/benchmark_attribution_2026-03.xlsx`.

Exit codes:
- `0` — success, every day within tolerance.
- `1` — hard error (missing inputs, malformed returns xlsx).
- `2` — completed with one or more tolerance breaches or data-quality warnings.

Spec: `docs/superpowers/specs/2026-05-12-benchmark-attribution-pipeline-design.md`.
