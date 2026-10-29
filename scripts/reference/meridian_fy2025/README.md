# Reference build: Meridian Technologies Group, FY2025 (sample data)

These scripts built the deliverables in the first full test cycle on the synthetic sample dataset.
They're kept here as a **worked example of the structure**, not as a pipeline to re-run:

| Script | Phase | Role |
|---|---|---|
| `build_model.py` | 1-3 | Parses inputs, consolidates, pickles a model (`_model.pkl`) |
| `build_workbook.py` | 6 | Formula-driven history workbook from the model |
| `build_memo.py` | 6 | Executive memo (.docx); every figure pulled from the model |
| `build_deck.py` | 6 | Board deck (.pptx) with native charts |
| `build_methodology.py` | 6 | Methodology & data-lineage document |
| `sanity_check.py` | 8 | Quick printout of model totals |

**Do not reuse these for another cycle as-is.** They hardcode FY2025 periods, the US01/UK01/IN01
entities, USD, the eight separate per-category input files the first run used, `data/inputs` /
`data/outputs`, and, in the memo and deck, narrative facts specific to that dataset (the Vantage churn,
the September revolver draw, and so on).

For a new cycle, Claude writes fresh scripts to `<outputs>/_work/<Period>/` that follow this same
split (model, then workbook, then memo and deck reading from the model). Those scripts use
`scripts/brlib.py` for paths and safe writes, and take every fact from that cycle's data.
