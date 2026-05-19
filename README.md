# ESCMID EPC Copilot

Extracts all sessions from ESCMID annual programme PDFs, classifies them
using the **official ESCMID taxonomy** (12 categories, ~70 subcategories)
via the Anthropic API, and produces:

1. **`ESCMID_Programmes.xlsx`** — 13 sheets: overview matrix + one per ESCMID category
2. **`ESCMID_Dashboard.html`** — self-contained interactive dashboard (no server needed)

---

## Repository structure

```
escmid_analyser.py       ← main script
requirements.txt         ← Python dependencies
.env.example             ← API key template (copy to .env)
data/
  programmes/            ← place PDF files here (gitignored)
  output/                ← generated files land here (gitignored)
  tagging_cache.json     ← API responses cached here (gitignored by default)
```

---

## Setup

### 1. Clone and install

```bash
git clone https://github.com/YOUR_USERNAME/escmid-epc-copilot.git
cd escmid-epc-copilot
pip install -r requirements.txt
```

### 2. Install pdftotext (from poppler)

```bash
# macOS
brew install poppler

# Ubuntu / Debian
sudo apt install poppler-utils

# Windows: https://github.com/oschwartz10612/poppler-windows/releases
```

### 3. Set your Anthropic API key

```bash
cp .env.example .env
# Open .env and paste your key — get one at console.anthropic.com
```

### 4. Add PDF files

Place programme PDFs in `data/programmes/`. Expected filenames are set in
`PROGRAMME_FILES` near the top of `escmid_analyser.py` — edit to match
whatever filenames you have.

### 5. Run

```bash
# Full run — API extraction + classification (recommended, ~$2-4 for all years)
python escmid_analyser.py --api-extract

# Classification only (faster if PDFs extract cleanly)
python escmid_analyser.py

# Keyword fallback — free, less accurate
python escmid_analyser.py --skip-tagging

# Specific years only
python escmid_analyser.py --years 2025,2026

# Dashboard only (re-use existing tagging cache)
python escmid_analyser.py --skip-excel

# All options
python escmid_analyser.py --help
```

---

## Cost and caching

With `--api-extract` using `claude-sonnet-4-20250514`:

| Sessions | Approximate cost | Approximate time |
|---|---|---|
| ~1,625 (all 6 years) | $2–4 | 15–20 min |
| ~350 (one year) | $0.50 | 3–5 min |

Results are cached in `data/tagging_cache.json`. **Re-runs are free** —
the script checks the cache before making any API call.

### Sharing the cache with collaborators

The cache is gitignored by default but contains no secrets and is safe to share.

```bash
# Option A: commit it once for the team
git add -f data/tagging_cache.json
git commit -m "add tagging cache"

# Option B: share the file directly (email, Slack, shared drive)
# Collaborators place it at data/tagging_cache.json before running
```

---

## ESCMID Official Taxonomy

Source: [escmid.org](https://www.escmid.org/congress-events/escmid-global/proposal-entry/escmid-global-categories-and-subcategories/)

| # | Category | Subcategories |
|---|---|---|
| 1 | Viral infection & disease | 10 (incl. COVID-19, HIV, hepatitis) |
| 2 | Bacterial infection & disease | 9 |
| 3 | Bacterial susceptibility & resistance | 8 |
| 4 | Diagnostic microbiology | 11 (incl. AI & digital health: **4j**) |
| 5 | New antibacterial agents, PK/PD & stewardship | 6 |
| 6 | Fungal infection & disease | 5 |
| 7 | Parasitic diseases, travel medicine & migrant health | 6 |
| 8 | Healthcare-associated infections & IPC | 9 |
| 9 | Fundamental microbiology, pathogenesis & immunity | 6 |
| 10 | Immune compromise & transplant ID | 6 |
| 11 | Public health & vaccines | 8 |
| 12 | Professional & educational affairs | 5 |

---

## Improving the script

| What to change | Where in the script |
|---|---|
| Add a new congress year | `PROGRAMME_FILES` dict (~line 95) |
| Tune keyword classification rules | `KEYWORD_RULES` list |
| Change which session types are extracted | `INCLUDE_TYPES_MODERN` / `INCLUDE_TYPES_2021` |
| Edit the API classification prompt | `build_tagging_prompt()` |
| Edit the combined extraction + classification prompt | `build_extract_classify_prompt()` |
| Change dashboard styling | `generate_dashboard_html()` |
| Change Excel formatting | `write_overview_sheet()` / `write_category_sheet()` |

---

## Contributing

Pull requests welcome. Suggested areas:

- Improved PDF extraction for 2021 (online ECCMID format)
- Better handling of sessions that only appear in overview grids
- Additional congress years
- Dashboard improvements (subcategory drill-down, new chart types)
- Tagging quality evaluation against hand-labelled ground truth

---

## Notes on PDFs

Programme PDFs are not included — they are subject to ESCMID copyright.
Request them from ESCMID or use your own copies. Filenames are configurable
in `PROGRAMME_FILES`.

---

## Licence

MIT — see `LICENSE`.
