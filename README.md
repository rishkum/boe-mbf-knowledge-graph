# MBF knowledge graph: portable demo

A self-contained copy of the Bank of England MBF research knowledge graph, for demos and sharing.
No API keys, no network calls, and no full publication text (each row links to the source instead).

| File | What it is |
|---|---|
| `report/BoE_MBF_research_insights.docx` | The write-up: worked examples of what the knowledge graph can do for a central bank economist. |
| `exports/BoE_MBF_KG.xlsx` | The whole graph as an Excel workbook. Start with the README sheet. Works offline on any PC with Excel. |
| `exports/csv/*.csv` | The same tables as CSV files (UTF-8, open directly in Excel). |
| `streamlit_app.py` | Interactive explorer: causal map, grade evidence, audit foresight, trace scenarios, briefing packs, gaps. |
| `assets/vis-network.min.js` | Network-drawing library (vis-network 9.1.9, Apache-2.0/MIT), bundled so the causal map works offline. |
| `kg.sqlite` | The slim database the app reads (1.8 MB). |
| `build_demo_data.py` | Rebuilds everything above from the full knowledge-graph project (not included here). |

## Demo with Excel only (no installs)

Copy `exports/BoE_MBF_KG.xlsx` to the work PC. Useful views:

- **Claims** sheet: filter `driver` and `outcome` (e.g. *Margin and collateral calls* → *Fire sales*) to see every
  claim with its quote, evidence type and link.
- **Causal links** sheet: sort by `publications` to see the best-evidenced links.
- **Coverage matrix** sheet: zeros are research gaps.
- **Opposing claims** sheet: where findings differ and why.

## Run the app locally

Needs Python 3.11+.

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

It opens at http://localhost:8501.

## Deploy to Posit Connect Cloud (free plan)

1. At [connect.posit.cloud](https://connect.posit.cloud), sign in with GitHub and choose **Publish → Streamlit**.
2. Pick this repository, branch `main`, and primary file `streamlit_app.py`.
3. Connect Cloud installs `requirements.txt` and gives you a URL. Pushing to `main` redeploys.

The free plan has 20 active hours a month. Check whether your plan lets you restrict who can view the app before
publishing; this bundle contains only metadata, short quotes and model-written summaries.

## Rebuild after the graph changes

From the full project: `.venv/bin/python demo/build_demo_data.py`, then commit and push here.

## Known limitations

- Direction, hedging, evidence type and concept matching were assigned by a language model and have not been
  reviewed by experts.
- The US Treasuries concept also picks up generic interest-rate statements.
- About 14% of quotes are the closest source passage rather than the model's exact quote.
