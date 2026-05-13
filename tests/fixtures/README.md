# Test fixtures

Synthetic mini-files used by the adapter tests.

- `tiny_returns.xlsx` — regenerable via `python tests/fixtures/make_fixtures.py`.
- `tiny_weights_*.csv` — hand-maintained CSVs covering the three CSV scenarios.
- `holdings_page2_sample.pdf` — a one-page extract of page 2 of one real Catalyst PDF, committed once.
  Regenerate with: `python tests/fixtures/extract_page2.py "data/March PDF/02 March 2026.pdf"`.
