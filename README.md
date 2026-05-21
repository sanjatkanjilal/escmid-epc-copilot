# ESCMID EPC Copilot

An end-to-end pipeline for analysing the ESCMID Global Congress scientific programme and supporting the Education Programme Committee (EPC) in reviewing session proposals for the 2027 Stockholm congress.

---

## Overview

The project has two main components:

1. **Programme Analyser** (`escmid_analyser.py`) — extracts, classifies, and visualises sessions from ESCMID Global programme PDFs (2021–2026)
2. **Proposal Reviewer** (`proposal_reviewer.py`) — parses submitted session proposals, tags them, scores them against the historical programme, and provides an interactive review dashboard

---

## Repository Structure

```
escmid-epc-copilot/
├── escmid_analyser.py          # Main programme analysis pipeline
├── proposal_reviewer.py        # Proposal review pipeline
├── build_review_html.py        # HTML dashboard generator (importable module)
├── parse_proposals_vscode.py   # HTML parser for ESCMID proposal pages
├── requirements.txt
├── data/
│   ├── programmes/             # Source PDFs (2021–2026, not committed)
│   ├── proposals/              # Proposal HTML files (view-source saves, not committed)
│   ├── tagging_cache.json      # API response cache (ESCMID categories + 94 tags)
│   └── output/
│       ├── sessions_raw.json       # Extracted + classified sessions
│       ├── ESCMID_Dashboard.html   # Interactive programme analysis dashboard
│       ├── ESCMID_Programmes.xlsx  # Programme data workbook
│       ├── proposals_tagged.json   # Parsed + scored proposals
│       ├── Proposal_Review.html    # Interactive proposal review dashboard
│       └── Proposal_Review.xlsx    # Proposal workbook with scoring columns
└── README.md
```

---

## Programme Analyser

### What it does

- Extracts sessions from programme PDFs using **pdfplumber** (2022–2026) with word-level bounding box column separation, and pdftotext for the 2021 online format
- Classifies each session into the **12 official ESCMID categories** (1–12) and their subcategories via the Anthropic API (cached)
- Applies a **94-tag clinical/methodological taxonomy** to each session
- Extracts individual **talk titles and speaker names** within each session in the same pdfplumber pass (no separate step needed)
- Generates an **interactive HTML dashboard** and an **Excel workbook**

### 94-Tag Taxonomy

Organised into 12 groups: Methods · Study Design · ClinMicro · Infectious Diseases · Treatments · Syndromes · Special Hosts · AMR Pathogens · Microbiome · Public Health · Professional · Region

### Commands

```bash
# Re-extract from PDFs using existing category + tag cache (recommended)
python escmid_analyser.py --rebuild

# Re-extract and re-apply 94 tags via API (~$3–5, ~90 min, cached after first run)
python escmid_analyser.py --skip-tagging --add-tags

# Full pipeline from scratch (expensive — only needed once)
python escmid_analyser.py --api-extract --add-tags

# Add talks to existing sessions_raw.json without re-running anything
python escmid_analyser.py --talks-only
```

### Dashboard Tabs

| Tab | Contents |
|-----|----------|
| Overview | Stacked bar chart of sessions by ESCMID category per year + session type mix |
| Trends | All-category trend lines + per-category subcategory breakdowns |
| Heatmap | Two side-by-side heatmaps: ESCMID subcategories × year (left), 94 tags × year (right) |
| Network | D3 force-directed graph — 1,028 nodes (one per curated session), edges weighted by shared tags |
| Explore | Searchable/filterable session table with dropdowns for year, category, subcategory, tag. Click any row for full details including talks |
| Gaps | Under-represented ESCMID subcategories and tags — potential gaps for 2027 proposals |
| People | Index of chairs and speakers across all years, searchable by name, filterable by year/role/category |

---

## Proposal Reviewer

### What it does

- Parses ESCMID proposal detail pages (saved as view-source HTML files) using BeautifulSoup
- Extracts: session title, category, subcategory, type, proposing entities, chairs, reserve chairs, champion, topic titles, speakers, and motivation/description
- Applies the same **94-tag taxonomy** to each proposal via API or keyword fallback
- Scores each proposal against the historical programme:
  - **Novelty** (0–1): how different is this from any previous ESCMID session? (tag Jaccard similarity)
  - **Trend** (-1 to +1): is this topic growing or shrinking in the programme (2022→2026)?
- Computes **heuristic AI ratings** (C1–C8) from available data; C9–C11 marked N/A
- Optionally generates **full Claude evaluations** on all 11 EPC criteria via `--ai-review`
- Generates an interactive HTML review dashboard and Excel scoring workbook

### Setup

1. For each assigned proposal, open it in Chrome, View Source (`Cmd+U`), save as `.html` into `data/proposals/`
2. Files with `_list_` or `_keynote_interview_` in the name are index pages — the parser skips them automatically
3. Individual proposals have `_proposal_view_id_` in the filename

### Commands

```bash
# Parse proposals, score, compute heuristic AI ratings, generate outputs (free)
python proposal_reviewer.py --skip-tagging

# Also apply 94 tags via API (~$0.50 for ~200 proposals)
python proposal_reviewer.py

# Full Claude evaluation on all 11 EPC criteria (~$1–2 for ~200 proposals)
python proposal_reviewer.py --skip-tagging --ai-review
```

### Serving the dashboard

The review dashboard uses `localStorage` to save your ratings — open via HTTP, not `file://`:

```bash
cd data/output
python3 -m http.server 8080
# Open: http://localhost:8080/Proposal_Review.html
```

### Dashboard Tabs

| Tab | Contents |
|-----|----------|
| Proposals | Full searchable table — click any row to open the Review Form |
| Review Form | Split panel: proposal details (left) + rating form (right). C1–C11 criteria with hover descriptions and AI dot indicators. Overall star rating. Notes. ← → keyboard navigation. Download CSV at any time. |
| Scores | Side-by-side novelty and trend bar charts — click any bar to open that proposal in the Review Form |
| Criteria | Reference descriptions for all 11 EPC scoring criteria |

### EPC Scoring Criteria

| | Criterion |
|--|-----------|
| C1 | Hot / timely / controversial |
| C2 | Not duplicated from recent meetings |
| C3 | Cross-disciplinary / wide appeal |
| C4 | Basic + translational + clinical integration |
| C5 | Appropriate session format |
| C6 | Relevant collaborators involved |
| C7 | Adheres to session format rules |
| C8 | Proper description provided |
| C9 | Gender & geographic balance |
| C10 | Best speakers / not self-serving |
| C11 | Engages young investigators |

### Excel Output

- **Proposals sheet**: one row per proposal with all parsed fields, tags, novelty/trend scores, most similar historical session, and 11 yellow scoring columns (C1–C11 + Overall + Notes)
- **Scoring Guide sheet**: full descriptions of all 11 criteria

---

## Installation

```bash
conda create -n escmid-epc-copilot python=3.12
conda activate escmid-epc-copilot
pip install -r requirements.txt
```

**requirements.txt** includes: `anthropic`, `pdfplumber`, `openpyxl`, `beautifulsoup4`, `python-dotenv`, `tqdm`

Set your API key in a `.env` file:
```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## Data Notes

- **PDF extraction**: pdfplumber uses word-level bounding boxes to separate the two-column programme layout. Talks are extracted in the same pass — no separate step needed for 2022–2026.
- **2021 format**: The 2021 ECCMID (online) programme uses a different session code format (S##) and falls back to pdftotext extraction.
- **Caching**: All API responses (ESCMID category classification, 94-tag assignments, AI proposal ratings) are cached in `data/tagging_cache.json`. Re-running any step is free if the cache is intact.
- **Category rebuild**: If `sessions_raw.json` gets overwritten with bad data, run `python escmid_analyser.py --rebuild` — this re-extracts sessions with pdfplumber but restores categories and tags from the cache.

---

## Future Plans

- **Speaker bios tab**: On-demand web search lookup for chairs and speakers in the Proposal Reviewer, with AI-summarised professional backgrounds. Planned as a button in the Review Form rather than a pre-fetched batch to control cost.
- **Improved PDF parsing**: Replace the current pdfplumber column-separation approach with a structured text input pipeline — copying clean session blocks directly from the ESCMID programme viewer (which produces well-structured plain text) into a dedicated parser. This would eliminate remaining title truncation and chair-bleed artefacts.
