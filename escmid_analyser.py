#!/usr/bin/env python3
"""
ESCMID EPC Copilot
==================
Extracts all sessions from ESCMID annual programme PDFs,
classifies them using the official ESCMID taxonomy via the
Anthropic API, and produces:

  1. An Excel workbook  — one sheet per ESCMID category (12) + overview
  2. An interactive HTML dashboard — trends, heatmap, network, explore, gaps

Quick start
-----------
1. Install dependencies:
       pip install -r requirements.txt
       # macOS: brew install poppler   (for pdftotext)
       # Ubuntu: sudo apt install poppler-utils
       # Windows: https://github.com/oschwartz10612/poppler-windows

2. Set your Anthropic API key — either in a .env file:
       echo 'ANTHROPIC_API_KEY=sk-ant-...' > .env
   or in your shell:
       export ANTHROPIC_API_KEY="sk-ant-..."

3. Place PDF files in  data/programmes/  (configure PROGRAMME_FILES below)

4. Run:
       python escmid_analyser.py
       python escmid_analyser.py --years 2025,2026          # specific years
       python escmid_analyser.py --skip-tagging              # keyword fallback only
       python escmid_analyser.py --skip-excel                # dashboard only

Outputs land in  data/output/
  ESCMID_Programmes.xlsx
  ESCMID_Dashboard.html
  sessions_raw.json         (intermediate — safe to delete)
  tagging_cache.json        (API responses cached here — keep to avoid re-billing)

Improving the script
--------------------
- Add / remove PROGRAMME_FILES entries to cover more years
- Tune KEYWORD_RULES for better fallback accuracy
- Modify INCLUDE_TYPES to control which session formats are extracted
- The API tagging prompt is in build_tagging_prompt() — edit it to improve accuracy
- The HTML dashboard template is in generate_dashboard_html()

"""

# ── Standard library ──────────────────────────────────────────────────────────
import os, re, json, time, subprocess, sys, argparse
from pathlib import Path
from collections import defaultdict, Counter

# ── Load .env file if present (pip install python-dotenv) ─────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # .env loading is optional; key can also be set in the shell directly

# ── Optional third-party ──────────────────────────────────────────────────────
try:
    import anthropic as _anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from tqdm import tqdm as _tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    _tqdm = lambda x, **_: x


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — edit this section for your setup
# ══════════════════════════════════════════════════════════════════════════════

BASE_DIR      = Path(__file__).parent
PROGRAMMES_DIR = BASE_DIR / "data" / "programmes"
OUTPUT_DIR    = BASE_DIR / "data" / "output"
CACHE_FILE    = BASE_DIR / "data" / "tagging_cache.json"

# Map year → PDF filename (relative to PROGRAMMES_DIR)
PROGRAMME_FILES = {
    "2021": "FinalProgramme_2021.pdf",
    "2022": "FinalProgramme_2022_Lisbon.pdf",
    "2023": "FinalProgramme_2023_Copenhagen.pdf",
    "2024": "FinalProgramme_2024_Barcelona.pdf",
    "2025": "FinalProgramme_2025_Vienna.pdf",
    "2026": "FinalProgramme_2026_Munich.pdf",
}

# Session type codes to include (prefix of session codes like OS001, SY010…)
# 2021 uses S## format; 2022+ uses OS/SY/EW/ME/KN/LB/EF/PM/CS/IS
INCLUDE_TYPES_MODERN = {"OS","SY","EW","ME","KN","LB","EF","PM","CS","IS"}
INCLUDE_TYPES_2021   = {"S"}    # 2021 uses S## codes

# Session code prefix → canonical session type
# These prefixes are consistent across all ECCMID/ESCMID years.
# Used as a reliable fallback when the type cannot be inferred from context text.
PREFIX_TO_TYPE = {
    "OS": "Oral Session",    "SY": "Symposium",
    "EW": "Educational",     "ME": "Meet-the-Expert",
    "KN": "Keynote",         "LB": "Late-Breaking",
    "EF": "ePoster Flash",   "PM": "Pipeline",
    "CS": "Case Session",    "IS": "IS / Integrated",
}

# Session types shown in the clustering network (keep manageable)
NETWORK_TYPES = {"OS","SY","KN","LB","PM","ME","S"}

# Anthropic API settings
ANTHROPIC_MODEL   = "claude-sonnet-4-20250514"
API_DELAY         = 0.4   # seconds between calls
MAX_RETRIES       = 3
SAVE_CACHE_EVERY  = 25    # write cache every N sessions


# ══════════════════════════════════════════════════════════════════════════════
# ESCMID OFFICIAL TAXONOMY
# Source: https://www.escmid.org/congress-events/escmid-global/proposal-entry/
#         escmid-global-categories-and-subcategories/
# ══════════════════════════════════════════════════════════════════════════════

CATEGORIES = {
    1:  {"name":"Viral infection & disease",
         "short":"Viral disease","color":"#4fc3f7","subcats":{
         "1a":"HIV/AIDS","1b":"Viral hepatitis","1c":"Influenza & respiratory viruses",
         "1d":"Herpesviruses","1e":"Emerging/re-emerging viral diseases","1f":"Diagnostic virology (other)",
         "1g":"Viral epidemiology – general","1h":"Antiviral drugs (other)",
         "1i":"Fundamental & applied virology","1j":"COVID-19"}},
    2:  {"name":"Bacterial infection & disease",
         "short":"Bacterial disease","color":"#fbbf24","subcats":{
         "2a":"TB & other mycobacterial infections","2b":"Severe sepsis, bacteraemia & endocarditis",
         "2c":"Community-acquired respiratory infections","2d":"Community-acquired abdominal/GI infections",
         "2e":"Community-acquired UTI & genital tract","2f":"Community-acquired SSTI, bone & joint",
         "2g":"Community-acquired CNS & invasive infections","2h":"Zoonotic bacterial infections",
         "2i":"Other intracellular or rare bacteria"}},
    3:  {"name":"Bacterial susceptibility & resistance",
         "short":"AMR","color":"#f87171","subcats":{
         "3a":"Resistance surveillance: community","3b":"Resistance surveillance: healthcare",
         "3c":"Susceptibility testing methods","3d":"Resistance mechanisms",
         "3e":"Resistance detection/prediction","3f":"Clinical outcome of resistant infections",
         "3g":"Spread of resistance (ecology, One Health)","3h":"Policy aspects of AMR"}},
    4:  {"name":"Diagnostic microbiology",
         "short":"Diagnostics","color":"#a78bfa","subcats":{
         "4a":"Diagnostic bacteriology","4b":"Laboratory management",
         "4c":"MALDI-TOF & proteomic methods","4d":"Molecular diagnostics (incl POCT)",
         "4e":"Strain typing & surveillance","4f":"Whole genome sequencing",
         "4g":"Microbiome studies","4h":"Clinical metagenomics",
         "4i":"Bioinformatics tools & pipelines","4j":"Artificial intelligence & digital health",
         "4k":"Other novel diagnostic technologies"}},
    5:  {"name":"New antibacterial agents, PK/PD & stewardship",
         "short":"Antibacterials & AMS","color":"#34d399","subcats":{
         "5a":"Drug discovery & new compounds","5b":"Pharmacokinetics/pharmacodynamics",
         "5c":"New/repurposed agents: clinical studies","5d":"Antimicrobial stewardship & prescribing",
         "5e":"Safety, hypersensitivity & adverse effects","5f":"Pharmacoepidemiology/pharmacoeconomics"}},
    6:  {"name":"Fungal infection & disease",
         "short":"Mycology","color":"#fb923c","subcats":{
         "6a":"Fundamental mycology","6b":"Fungal disease epidemiology",
         "6c":"Diagnostic mycology","6d":"Antifungal susceptibility & resistance",
         "6e":"Antifungal drugs & treatment"}},
    7:  {"name":"Parasitic diseases, travel medicine & migrant health",
         "short":"Parasitology & Travel","color":"#6ee7b7","subcats":{
         "7a":"Fundamental parasitology","7b":"Parasitic disease epidemiology",
         "7c":"Diagnostic parasitology","7d":"Antiparasitic drugs & treatment",
         "7e":"Antiparasitic susceptibility & resistance","7f":"Travel medicine & migrant health"}},
    8:  {"name":"Healthcare-associated infections & IPC",
         "short":"HAI & IPC","color":"#38bdf8","subcats":{
         "8a":"Intravascular catheter-related infections","8b":"Foreign-body & implant infections",
         "8c":"Surgical site infections","8d":"Healthcare-associated pneumonia (VAP)",
         "8e":"Hospital epidemiology, transmission & surveillance",
         "8f":"Other HAIs (CDI, outbreaks)","8g":"Infection control interventions & trials",
         "8h":"Disinfection & sterilisation","8i":"Healthcare workers & IPC"}},
    9:  {"name":"Fundamental microbiology, pathogenesis & immunity",
         "short":"Basic science","color":"#e879f9","subcats":{
         "9a":"Microbial pathogenesis & virulence","9b":"Host-pathogen interaction",
         "9c":"Pre-clinical biofilm studies","9d":"Experimental & cellular microbiology",
         "9e":"Fundamental science using Omics","9f":"Immune response to infection"}},
    10: {"name":"Immune compromise & transplant ID",
         "short":"Immunocompromised","color":"#f472b6","subcats":{
         "10a":"Host genetics & primary immunodeficiency","10b":"Solid organ transplantation",
         "10c":"Haematopoietic stem cell transplantation","10d":"Cell-based therapies",
         "10e":"Cancer treatment & neutropaenia","10f":"Other immunosuppression"}},
    11: {"name":"Public health & vaccines",
         "short":"Public health","color":"#86efac","subcats":{
         "11a":"General vaccinology","11b":"Antiviral vaccines","11c":"Antibacterial vaccines",
         "11d":"Other preventive modalities","11e":"Food, water & environmental health",
         "11f":"Veterinary microbiology & One Health","11g":"Global health & health security",
         "11h":"Infections in low-resource settings"}},
    12: {"name":"Professional & educational affairs",
         "short":"Professional","color":"#94a3b8","subcats":{
         "12a":"Professional affairs & career development","12b":"Publishing, ethics & academic affairs",
         "12c":"Medical education for CM/ID","12d":"Diversity & equality",
         "12e":"Advocacy & role of patients"}},
}

# Flat lookup: subcat code → (cat_num, cat_name, subcat_name)
SUBCAT_LOOKUP = {
    sc: (cn, cat["name"], sn)
    for cn, cat in CATEGORIES.items()
    for sc, sn in cat["subcats"].items()
}

# Prompt string: all subcategories for the API
ALL_SUBCATS_STR = "\n".join(
    f"  {sc}: {sn}  [{CATEGORIES[cn]['name']}]"
    for cn, cat in CATEGORIES.items()
    for sc, sn in cat["subcats"].items()
)


# ══════════════════════════════════════════════════════════════════════════════
# PDF EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def pdf_to_text(path: Path) -> str:
    """Extract raw text from PDF using pdftotext (poppler)."""
    result = subprocess.run(
        ["pdftotext", "-layout", str(path), "-"],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftotext failed: {result.stderr[:200]}")
    return result.stdout


# Regex for 2022–2026 session codes (OS001, SY010, EW005, …)
MODERN_CODE_RE = re.compile(
    r'\b((?:OS|SY|EW|ME|LB|EF|KN|PM|CS|IS)\d{2,4})\b'
)
# Regex for 2021 session codes (S26, S222, …)
Y2021_CODE_RE = re.compile(r'\b(S\d{2,3})\b')

SESSION_TYPE_MAP = [
    ("Oral Session",      ["oral session", "oral case session", "mini oral flash"]),
    ("Symposium",         ["symposium"]),
    ("Educational",       ["educational", "workshop"]),
    ("Meet-the-Expert",   ["meet-the-expert", "meet the expert"]),
    ("Keynote",           ["keynote lecture", "keynote"]),
    ("Late-Breaking",     ["late-breaking", "late breaking"]),
    ("ePoster Flash",     ["eposter flash", "poster flash", "eposter review"]),
    ("Pipeline",          ["pipeline"]),
    ("Case Session",      ["case session", "oral case"]),
    ("IS / Integrated",   ["integrated symposium", "integrated workshop"]),
]

# Compile a single regex that matches across line breaks within a context string
SESSION_TYPE_RE = re.compile(
    r"\b(\d+[\s.,/-]*(?:5|hour|hour)s?[\s-]*)?"  # optional "N-hour"
    r"(oral\s+(?:case\s+)?session|mini\s+oral\s+flash|"
    r"symposium|educational|workshop|meet[-\s]the[-\s]expert|"
    r"keynote|late[-\s]breaking|eposter\s+(?:flash|review)|"
    r"poster\s+flash|pipeline|case\s+session|integrated\s+symposium)",
    re.IGNORECASE | re.DOTALL
)


def classify_type(context: str) -> str:
    """Classify session type from context text."""
    # First try the compiled regex (handles split lines)
    m = SESSION_TYPE_RE.search(context)
    if m:
        raw = m.group(0).lower()
        if "oral" in raw:          return "Oral Session"
        if "symposium" in raw:     return "Symposium"
        if "educational" in raw or "workshop" in raw: return "Educational"
        if "meet" in raw:          return "Meet-the-Expert"
        if "keynote" in raw:       return "Keynote"
        if "late" in raw:          return "Late-Breaking"
        if "eposter" in raw or "poster flash" in raw: return "ePoster Flash"
        if "pipeline" in raw:      return "Pipeline"
        if "case" in raw:          return "Case Session"
        if "integrated" in raw:    return "IS / Integrated"
    # Fallback: keyword scan
    lc = context.lower()
    for label, keywords in SESSION_TYPE_MAP:
        if any(k in lc for k in keywords):
            return label
    return "Other"


def stitch_fragments(frags: list) -> str:
    """
    Join multi-line title fragments from a narrow overview-table column.
    Handles hyphenated line-breaks (fragment ends with '-') and plain wraps.
    Stops at boilerplate lines (Chair, Hall, code, time, etc.).
    """
    STOP = re.compile(
        r'^(Chairs?|Hall|Arena|Co-organised|\d{2}:\d{2}|'
        r'(?:OS|SY|EW|ME|LB|EF|KN|PM|CS|IS)\d{2,4})',
        re.I
    )
    result = ""
    for frag in frags:
        frag = frag.strip()
        if not frag or STOP.match(frag):
            if result:
                break
            continue
        if result.endswith("-"):
            result = result[:-1] + frag   # de-hyphenate
        elif result:
            result += " " + frag
        else:
            result = frag
        if len(result) > 120 and not frag.endswith("-"):
            break
    return result.strip()


def is_faculty_listing_line(line: str) -> bool:
    """
    Return True if the line contains 2+ 'PersonName  CODE' pairs — the hallmark
    of a faculty index page.

    Catches 2025 format: 'Ajjampur Sitara Sr  EW005  Cameron Alexandra  PM6 ...'
    Does NOT fire on single-pair lines like 'Muge Cevik  ME113  13:30 - 14:30'
    (chair listed next to session code — that is a legitimate session entry).
    """
    pairs = re.findall(
        r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z\'\-]+){0,3}\s+[A-Z]{2,}\d+',
        line
    )
    return len(pairs) >= 2


def is_faculty_index(context_lines: list) -> bool:
    """
    Detect speaker/faculty index pages. Two formats observed:
      2022/2026: 'Surname, Firstname, City, Country  OS042'  (comma-separated)
      2025:      'Firstname Lastname  EW005'                  (no commas, name before code)
    """
    ctx = " ".join(context_lines[:5])
    # Format 1 — comma-separated with city/country
    if re.search(
        r'[A-Z][a-z]+,\s+[A-Z][a-z][a-z]+[^,]*,\s+[A-Z][a-z]+[^,]*,\s+[A-Z]',
        ctx
    ):
        return True
    # Format 2 — 'Firstname Lastname  CODE' on 2+ lines in context
    name_code = re.compile(
        r'^[A-Z][a-z]+(?:\s+[A-Z][a-z\'\-]+){1,4}\s{2,}[A-Z]{2,}\d+'
    )
    if sum(1 for l in context_lines[:6] if name_code.match(l.strip())) >= 2:
        return True
    return False


def extract_title_from_context(context_lines: list) -> tuple:
    """
    Extract (title, session_type) from column-aware context lines.

    Improvements:
    - Stitches vertically-wrapped title fragments (handles hyphenated line-breaks)
    - Detects and skips faculty-index contexts
    - Strips abstract codes (e.g. 'S0367 09:00') leaked into titles
    - Validates minimum title quality
    """
    if is_faculty_index(context_lines):
        return "", "Other"

    session_type = "Other"
    title_frags  = []
    found_type   = False

    for i, line in enumerate(context_lines):
        lc = line.lower().strip()
        if not lc:
            continue

        # Detect session-type line
        if not found_type and any(k in lc for k in [
                "oral session", "symposium", "educational", "workshop",
                "meet-the-expert", "keynote", "late-breaking", "poster flash",
                "pipeline", "case session", "oral case", "integrated"]):
            session_type = classify_type(line)
            found_type   = True
            title_frags  = []
            for k in range(i + 1, len(context_lines)):
                frag = context_lines[k].strip()
                if not frag:
                    continue
                if re.match(r'^(Chairs?|Hall|Arena|Co-organised)', frag, re.I):
                    break
                if MODERN_CODE_RE.match(frag) or Y2021_CODE_RE.match(frag):
                    break
                if re.match(r'^\d{2}:\d{2}', frag):
                    break
                title_frags.append(frag)
                if len(" ".join(title_frags)) > 130 and not frag.endswith("-"):
                    break
            break

    # Fallback: no type line found — collect from first real non-boilerplate line
    if not title_frags:
        for line in context_lines[1:]:
            c = line.strip()
            if (len(c) > 8
                    and not MODERN_CODE_RE.match(c)
                    and not Y2021_CODE_RE.match(c)
                    and not re.match(r'^\d{2}:\d{2}', c)
                    and not re.match(r'^(Chairs?|Hall|Arena)', c, re.I)):
                title_frags.append(c)

    title = stitch_fragments(title_frags)

    # Strip abstract codes leaked into the title (e.g. "Title S0367 09:00 ...")
    title = re.sub(r'\s+[A-Z]\d{4}\s+\d{2}:\d{2}.*$', '', title)
    # Strip page-footer text (e.g. "ESCMID Global | Munich, Germany 2026 29")
    title = re.sub(r'\s+ESCMID\s+Global\s*\|.*$', '', title, flags=re.I)
    title = re.sub(r'\s+please\s+convert\s+to.*$', '', title, flags=re.I)
    # Strip right-column bleed (2+ space gap = column boundary in overview tables)
    title = re.split(r'\s{2,}', title)[0].strip()
    # Final whitespace normalisation
    title = re.sub(r'\s+', ' ', title).strip()

    return title, session_type

def column_ctx(lines: list, line_idx: int, code_col: int, n: int = 14) -> list[str]:
    """
    Extract context lines from the same column as the session code.

    Uses the code's own position as the column anchor rather than a computed
    midpoint — this correctly handles both narrow overview-table columns (where
    two codes sit 15-20 chars apart) and wide full-page columns.

    For right-column codes: extract from code_col leftward, so we don't bleed
    into adjacent left-column text on the same line.
    For left-column codes: extract up to code_col end of neighbouring content.
    """
    code_line = lines[line_idx]
    line_len  = len(code_line)

    # Determine column: right if the code sits in the right half of this line
    mid   = max(line_len // 2, 20)
    right = code_col > mid

    # Use the code's own position as column boundary (with a small buffer)
    col_start = max(0, code_col - 2)    # right col: start just before the code
    col_end   = code_col                 # left col: stop just before code's column

    # Typical overview-table column width is ~22 chars; cap right extractions
    # so we don't bleed into the immediately adjacent column.
    COL_WIDTH = 26
    ctx = []
    for j in range(line_idx, min(line_idx + n, len(lines))):
        l = lines[j]
        if right:
            part = l[col_start: col_start + COL_WIDTH].strip()
        else:
            # Gap-split: take content up to the first 3+ space run.
            # More reliable than a fixed midpoint for left-column titles that
            # extend past the computed midpoint (e.g. long session titles).
            gap_parts = re.split(r'\s{3,}', l)
            part = gap_parts[0].strip() if gap_parts else l[:mid].strip()
        if part:
            ctx.append(part)
    return ctx

def extract_sessions_modern(text: str, year: str) -> list[dict]:
    """
    Extract sessions from 2022-2026 PDFs (OS/SY/... codes).

    Two-pass approach to avoid overview-table contamination:
      Pass 1 — single-code lines only (the detailed programme listing).
               These give clean one-session-per-line context.
      Pass 2 — multi-code lines for any codes not found in Pass 1.
               These are overview/schedule grids; column-aware extraction
               is used to isolate the correct column.
    """
    lines = text.split('\n')
    sessions = []
    seen = set()

    def process_line(i: int, line: str):
        for m in MODERN_CODE_RE.finditer(line):
            code = m.group(1)
            if code[:2] not in INCLUDE_TYPES_MODERN or code in seen:
                continue
            # Skip lines with 2+ Name+Code pairs — faculty index pages
            if is_faculty_listing_line(line):
                continue
            ctx = column_ctx(lines, i, m.start())
            title, stype = extract_title_from_context(ctx)
            if not title or len(title) < 8:
                return
            # Reject titles that still look like person names after extraction
            if re.match(r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3}$', title):
                return
            seen.add(code)
            sessions.append({"code": code, "year": year,
                              "type": stype, "title": title,
                              "_raw_ctx": "\n".join(ctx[:5])})

    # Pass 1: single-code lines (higher quality)
    for i, line in enumerate(lines):
        if len(MODERN_CODE_RE.findall(line)) == 1:
            process_line(i, line)

    # Pass 2: multi-code lines for anything missed
    for i, line in enumerate(lines):
        if len(MODERN_CODE_RE.findall(line)) > 1:
            process_line(i, line)

    return sessions


def extract_sessions_2021(text: str) -> list[dict]:
    """
    Extract sessions from 2021 PDF.
    2021 (online) uses S## codes and a different two-column layout.
    Session structure: S##\\t\\tHH:MM - HH:MM\\tCATEGORY
                           \\t\\tDURATION-hour Session Type
                       Title of session
    """
    lines = text.split('\n')
    sessions = []
    seen = set()

    for i, line in enumerate(lines):
        for m in Y2021_CODE_RE.finditer(line):
            code = m.group(1)
            if code in seen:
                continue
            if not re.search(r'\d{2}:\d{2}', line):
                continue

            # 2021 is two-column. The code appears at a horizontal offset.
            # Extract only the right-hand portion of each context line,
            # starting from roughly where the code sits, to avoid picking up
            # left-column authors/abstracts as the session title.
            col_start = max(0, m.start() - 5)
            ctx = []
            for j in range(i, min(i + 8, len(lines))):
                right = lines[j][col_start:].strip()
                if right:
                    ctx.append(right)

            title, stype = extract_title_from_context(ctx)
            if not title or len(title) < 8:
                continue
            # Reject author lines (e.g. "A. Smith*") or stray times
            if re.match(r'^[A-Z]\. [A-Z]', title):
                continue
            if re.match(r'^\d{2}:\d{2}', title):
                continue
            seen.add(code)
            sessions.append({"code": code, "year": "2021",
                              "type": stype, "title": title})
    return sessions


def extract_all(pdf_map: dict) -> list[dict]:
    """Extract sessions from all PDFs. pdf_map: {year: Path}"""
    all_sessions = []
    for year, path in sorted(pdf_map.items()):
        if not path.exists():
            print(f"  [SKIP] {year}: file not found ({path})")
            continue
        print(f"  Extracting {year} …", end=" ", flush=True)
        try:
            text = pdf_to_text(path)
            if year == "2021":
                sessions = extract_sessions_2021(text)
            else:
                sessions = extract_sessions_modern(text, year)
            print(f"{len(sessions)} sessions")
            all_sessions.extend(sessions)
        except Exception as exc:
            print(f"ERROR — {exc}")
    return all_sessions


# ══════════════════════════════════════════════════════════════════════════════
# KEYWORD FALLBACK TAGGER
# (used when --skip-tagging or ANTHROPIC_API_KEY not set)
# ══════════════════════════════════════════════════════════════════════════════

# Each rule: (primary_cat, primary_subcat, [keywords…])
# Rules are evaluated in order; first match wins.
KEYWORD_RULES = [
    # ── Category 4j: AI / digital health (check first — cross-cutting) ──────
    (4, "4j", ["artificial intelligence","machine learning","deep learning",
               "neural network","large language model","chatgpt","llm",
               "natural language processing","nlp","digital health"," ai-",
               "ai tool","ai model","generative ai","computer vision"]),
    # ── Category 1: Viral ────────────────────────────────────────────────────
    (1, "1j", ["covid","sars-cov","covid-19","coronavirus"]),
    (1, "1a", ["hiv","aids","antiretroviral","art therapy","pre-exposure prophylaxis"]),
    (1, "1b", ["hepatitis b","hepatitis c","hepatitis e","hbv","hcv","hev","viral hepatitis"]),
    (1, "1c", ["influenza","rsv","respiratory syncytial","respiratory virus","flu vaccine",
               "metapneumovirus","parainfluenza","rhinovirus"]),
    (1, "1d", ["herpes","cmv","cytomegalovirus","ebv","epstein-barr","vzv","hsv",
               "varicella zoster","kaposi"]),
    (1, "1e", ["mpox","monkeypox","ebola","marburg","dengue","zika","west nile",
               "emerging virus","hantavirus","rift valley","nipah"]),
    (1, "1i", ["phage","bacteriophage","phage therapy"]),
    # ── Category 2: Bacterial disease ────────────────────────────────────────
    (2, "2a", ["tuberculosis","mycobacteri","mycobacteria"," tb ","tb treatment",
               "latent tb","non-tuberculous mycobacteria","ntm"]),
    (2, "2b", ["sepsis","bacteraemia","bacteremia","bloodstream infection",
               "endocarditis","infective endocarditis","bacteraemic"]),
    (2, "2c", ["community-acquired pneumonia","cap ","lower respiratory","bronchitis",
               "community respiratory"]),
    (2, "2d", ["intraabdominal","peritonitis","abdominal sepsis","gastrointestinal infection",
               "cdiff","c. difficile","clostridioides","clostridium difficile"]),
    (2, "2e", ["urinary tract infection","uti ","cystitis","pyelonephritis","urinary infection"]),
    (2, "2f", ["skin and soft tissue","cellulitis","necrotising fasciitis","osteomyelitis",
               "septic arthritis","bone infection","joint infection"]),
    (2, "2g", ["meningitis","encephalitis","brain abscess","cns infection","neurological infection"]),
    (2, "2h", ["zoonotic","lyme disease","brucella","leptospira","rickettsial"]),
    # ── Category 3: AMR ───────────────────────────────────────────────────────
    (3, "3a", ["resistance surveillance","amr surveillance","epidemiology of resistance",
               "prevalence of resistance"]),
    (3, "3b", ["healthcare-associated resistance","hospital-acquired resistance","nosocomial amr"]),
    (3, "3c", ["susceptibility testing","mic determination","breakpoint","eucast","clsi",
               "disk diffusion","e-test","broth microdilution"]),
    (3, "3d", ["resistance mechanism","beta-lactamase","carbapenemase","kpc","ndm",
               "oxa-","mcr-","mobile genetic","plasmid transfer"]),
    (3, "3e", ["resistance prediction","resistance detection","rapid amr","resistome"]),
    (3, "3f", ["outcome of resistant","clinical outcome of resistant","mdr infection outcome"]),
    (3, "3g", ["spread of resistance","one health amr","amr ecology","amr reservoir"]),
    (3, "3h", ["amr policy","antimicrobial policy","amr economic","amr governance"]),
    # ── Category 4: Diagnostics ───────────────────────────────────────────────
    (4, "4a", ["blood culture","diagnostic bacteriology","culture-based diagnosis"]),
    (4, "4b", ["laboratory management","lab automation","laboratory quality","laboratory data"]),
    (4, "4c", ["maldi-tof","mass spectrometry","proteomics"]),
    (4, "4d", ["molecular diagnostic","pcr","point-of-care","poct","syndromic testing",
               "multiplex pcr","rapid molecular"]),
    (4, "4e", ["strain typing","genomic epidemiology","molecular typing","wgmlst","cgmlst"]),
    (4, "4f", ["whole genome sequencing","wgs","next-generation sequencing","ngs",
               "nanopore","sequencing-based"]),
    (4, "4g", ["microbiome","gut microbiota","microbiota","dysbiosis","16s rrna"]),
    (4, "4h", ["metagenomics","metagenomic","clinical metagenomics","shotgun sequencing"]),
    (4, "4i", ["bioinformatics","pipeline","software tool","computational"]),
    (4, "4k", ["novel diagnostic","lateral flow","immunochromatography","biosensor"]),
    # ── Category 5: Antibacterials & AMS ──────────────────────────────────────
    (5, "5a", ["drug discovery","new antibiotic","new compound","novel antibiotic",
               "drug design","antibiotic pipeline","antimicrobial candidate"]),
    (5, "5b", ["pharmacokinetics","pharmacodynamics","pk/pd","therapeutic drug monitoring",
               "population pk","drug dosing","monte carlo simulation"]),
    (5, "5c", ["ceftazidime","cefiderocol","imipenem-cilastatin","meropenem-vaborbactam",
               "ceftolozane","cefepime","aztreonam","new beta-lactam","repurposed antibiotic",
               "clinical trial antibiotic"]),
    (5, "5d", ["antimicrobial stewardship","antibiotic stewardship","prescribing",
               "antibiotic use","rational antibiotic","de-escalation","iv-to-oral"]),
    (5, "5e", ["adverse effect","drug safety","hypersensitivity","allergy antibiotic",
               "nephrotoxicity","hepatotoxicity"]),
    (5, "5f", ["pharmacoepidemiology","cost-effectiveness","pharmacoeconomics"]),
    # ── Category 6: Mycology ──────────────────────────────────────────────────
    (6, "6a", ["aspergillus","candida","cryptococcus","mucor","mould","fungal pathogen",
               "mycobiome","fungal virulence"]),
    (6, "6b", ["invasive fungal infection","invasive aspergillosis","candidaemia",
               "fungal epidemiology","ifd "]),
    (6, "6c", ["diagnostic mycology","fungal diagnosis","galactomannan","beta-glucan",
               "fungal pcr"]),
    (6, "6d", ["antifungal resistance","antifungal susceptibility","azole resistance",
               "echinocandin resistance"]),
    (6, "6e", ["antifungal treatment","antifungal therapy","voriconazole","isavuconazole",
               "amphotericin","caspofungin","anidulafungin","micafungin"]),
    # ── Category 7: Parasitology & Travel ─────────────────────────────────────
    (7, "7a", ["parasite","plasmodium","leishmania","trypanosoma","helminth","schistosoma"]),
    (7, "7b", ["malaria epidemiology","parasitic disease epidemiology","parasite burden"]),
    (7, "7c", ["malaria diagnosis","parasitic diagnosis","microscopy parasite","rapid diagnostic test"]),
    (7, "7d", ["antimalarial","antiparasitic","malaria treatment","parasitic treatment"]),
    (7, "7f", ["travel medicine","traveller","returning traveller","tropical disease",
               "migrant health","refugee health"]),
    # ── Category 8: HAI & IPC ─────────────────────────────────────────────────
    (8, "8a", ["catheter","cvc","central line","clabsi","intravascular catheter"]),
    (8, "8b", ["prosthetic joint infection","pji","implant infection","foreign body infection"]),
    (8, "8c", ["surgical site infection","ssi ","perioperative infection"]),
    (8, "8d", ["ventilator-associated","vap ","hospital-acquired pneumonia","hap "]),
    (8, "8e", ["nosocomial","hospital epidemiology","hospital transmission",
               "infection surveillance","hospital outbreak","decolonisation"]),
    (8, "8f", ["healthcare-associated infection","hai ","clostridioides difficile","cdiff",
               "c. diff","cdiff infection"]),
    (8, "8g", ["infection control intervention","ipc intervention","hand hygiene",
               "contact precautions","isolation"]),
    (8, "8h", ["disinfection","sterilisation","decontamination","medical device"]),
    (8, "8i", ["healthcare worker","hcw infection","occupational infection","staff vaccination"]),
    # ── Category 9: Basic science ─────────────────────────────────────────────
    (9, "9a", ["pathogenesis","virulence factor","toxin","pathogenicity"]),
    (9, "9b", ["host-pathogen","biofilm","animal model","in vitro model"]),
    (9, "9d", ["cellular microbiology","experimental microbiology","in vitro study"]),
    (9, "9e", ["omics","transcriptomics","proteomics","metabolomics","genomics"]),
    (9, "9f", ["immune response","innate immunity","adaptive immunity","t-cell","antibody response"]),
    # ── Category 10: Immunocompromised ────────────────────────────────────────
    (10, "10e", ["neutropaenia","febrile neutropaenia","haematological malignancy",
                 "cancer patient infection","oncology infection"]),
    (10, "10b", ["solid organ transplant","liver transplant","kidney transplant",
                 "lung transplant","organ transplant"]),
    (10, "10c", ["stem cell transplant","hsct","allogeneic transplant","autologous transplant"]),
    (10, "10d", ["car-t","cell therapy","adoptive therapy"]),
    (10, "10f", ["biological therapy","immunosuppression","corticosteroid",
                 "rheumatological","tnf inhibitor","biologic agent"]),
    # ── Category 11: Public health & vaccines ─────────────────────────────────
    (11, "11a", ["vaccine development","vaccine safety","vaccine efficacy","vaccinology",
                 "vaccine immunogenicity"]),
    (11, "11c", ["pneumococcal vaccine","meningococcal vaccine","antibacterial vaccine"]),
    (11, "11e", ["food safety","water safety","environmental health","vector epidemiology"]),
    (11, "11f", ["one health","veterinary","zoonosis","animal reservoir","animal antimicrobial"]),
    (11, "11g", ["pandemic preparedness","health security","global health","biosafety",
                 "biosecurity","climate change infection"]),
    (11, "11h", ["low-resource","lmic","sub-saharan","africa","developing countr",
                 "resource-limited","health equity","vulnerable population"]),
    # ── Category 12: Professional ─────────────────────────────────────────────
    (12, "12a", ["career","professional development","fellowship","training programme"]),
    (12, "12b", ["ethics","publishing","authorship","peer review","academic"]),
    (12, "12c", ["medical education","teaching","curriculum","competency","osce"]),
    (12, "12d", ["diversity","equality","equity","gender","inclusion"]),
    (12, "12e", ["patient","advocacy","patient perspective","patient involvement"]),
]


def keyword_tag(session: dict) -> dict:
    """Keyword-based fallback classification. Returns classification dict."""
    title = session["title"].lower()
    for cat, subcat, keywords in KEYWORD_RULES:
        if any(k in title for k in keywords):
            return {"primary_cat": cat, "primary_subcat": subcat,
                    "secondary_cat": None, "secondary_subcat": None, "confidence": "low"}
    return {"primary_cat": 4, "primary_subcat": "", "secondary_cat": None,
            "secondary_subcat": None, "confidence": "low"}


# ══════════════════════════════════════════════════════════════════════════════
# ANTHROPIC API TAGGER
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = (
    "You are an expert in infectious diseases and clinical microbiology, "
    "classifying ESCMID conference sessions. Respond ONLY with valid JSON "
    "with no extra text, preamble, or markdown fences."
)


def build_tagging_prompt(session: dict) -> str:
    return f"""Classify this ESCMID Global conference session into the official taxonomy.

Session code : {session['code']}
Session type : {session['type']}
Session title: {session['title']}

Official ESCMID subcategories:
{ALL_SUBCATS_STR}

Return JSON only:
{{
  "primary_cat"    : <integer 1-12>,
  "primary_subcat" : "<e.g. 4j>",
  "secondary_cat"  : <integer 1-12 or null>,
  "secondary_subcat": "<e.g. 5d or null>",
  "confidence"     : "high" | "medium" | "low"
}}

Guidelines:
- primary_cat/subcat = single best fit
- secondary_cat/subcat = second category if clearly relevant, else null
- For sessions spanning multiple domains (e.g. AI applied to AMS) use both fields"""


def load_cache() -> dict:
    return json.loads(CACHE_FILE.read_text()) if CACHE_FILE.exists() else {}


def save_cache(cache: dict):
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(json.dumps(cache, indent=2))


EXTRACT_CLASSIFY_PROMPT_SYS = (
    "You are an expert in infectious diseases and clinical microbiology. "
    "You are extracting and classifying sessions from a scientific conference programme PDF. "
    "The raw text may have artefacts from a two-column PDF layout — ignore stray text that "
    "belongs to an adjacent column. Respond ONLY with valid JSON, no extra text."
)

def build_extract_classify_prompt(raw_ctx: str) -> str:
    """
    Combined prompt: clean up raw PDF context AND classify in one API call.
    Used with --api-extract flag.
    """
    return f"""Below is raw text extracted from a two-column conference programme PDF.
It may contain artefacts: truncated titles, adjacent-column text bleed-through,
wrapped lines, author names, chair names, or partial content.

<raw_context>
{raw_ctx}
</raw_context>

1. Extract the SESSION TITLE — the descriptive name of the session (not the abstract
   titles of individual talks within it). Clean up any column artefacts.

2. Extract the SESSION TYPE from: Oral Session / Symposium / Educational /
   Meet-the-Expert / Keynote / Late-Breaking / ePoster Flash / Pipeline / Case Session / Other

3. Classify into the best-fit ESCMID official subcategory:
{ALL_SUBCATS_STR}

Return JSON only:
{{
  "title"          : "<clean session title>",
  "type"           : "<session type>",
  "primary_cat"    : <integer 1-12>,
  "primary_subcat" : "<e.g. 4j>",
  "secondary_cat"  : <integer 1-12 or null>,
  "secondary_subcat": "<e.g. 5d or null>",
  "confidence"     : "high" | "medium" | "low"
}}"""


def api_extract_and_classify(session: dict, client, cache: dict) -> dict:
    """
    Uses the API to BOTH clean the session title/type AND classify it.
    More accurate than separate extraction + classification for noisy PDFs.
    Enabled with --api-extract flag.
    """
    raw_ctx = session.get("_raw_ctx", session.get("title", ""))
    cache_key = f"extract|{session['year']}|{session['code']}|{raw_ctx[:80]}"

    if cache_key in cache:
        return cache[cache_key]

    prompt = build_extract_classify_prompt(raw_ctx)

    for attempt in range(MAX_RETRIES):
        try:
            resp = client.messages.create(
                model=ANTHROPIC_MODEL,
                max_tokens=400,
                system=EXTRACT_CLASSIFY_PROMPT_SYS,
                messages=[{"role": "user", "content": prompt}]
            )
            raw = resp.content[0].text.strip()
            raw = re.sub(r'^```json?\s*|\s*```$', '', raw)
            result = json.loads(raw)

            result["title"]           = str(result.get("title") or session["title"])
            result["type"]            = str(result.get("type", session.get("type","Other")))
            result["primary_cat"]     = int(result.get("primary_cat", 4))
            result["primary_subcat"]  = str(result.get("primary_subcat") or "")
            result["secondary_cat"]   = result.get("secondary_cat")
            result["secondary_subcat"]= result.get("secondary_subcat")
            result["confidence"]      = result.get("confidence", "medium")

            cache[cache_key] = result
            return result
        except (json.JSONDecodeError, KeyError, ValueError):
            if attempt == MAX_RETRIES - 1:
                break
            time.sleep(1.0)
        except Exception as exc:
            if attempt == MAX_RETRIES - 1:
                print(f"\n  API error: {exc}")
                break
            time.sleep(API_DELAY * (attempt + 2))

    return {"title": session["title"], "type": session.get("type","Other"),
            "primary_cat": 4, "primary_subcat": "", "secondary_cat": None,
            "secondary_subcat": None, "confidence": "low"}


def api_tag_session(session: dict, client, cache: dict) -> dict:
    """Call Anthropic API to classify a session; use cache if available."""
    key = f"{session['year']}|{session['code']}|{session['title'][:60]}"
    if key in cache:
        return cache[key]

    for attempt in range(MAX_RETRIES):
        try:
            resp = client.messages.create(
                model=ANTHROPIC_MODEL,
                max_tokens=256,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": build_tagging_prompt(session)}]
            )
            raw = resp.content[0].text.strip()
            raw = re.sub(r'^```json?\s*|\s*```$', '', raw)
            result = json.loads(raw)
            result["primary_cat"]     = int(result.get("primary_cat", 4))
            result["primary_subcat"]  = str(result.get("primary_subcat") or "")
            result["secondary_cat"]   = result.get("secondary_cat")
            result["secondary_subcat"]= result.get("secondary_subcat")
            result["confidence"]      = result.get("confidence", "medium")
            cache[key] = result
            return result
        except (json.JSONDecodeError, KeyError, ValueError):
            if attempt == MAX_RETRIES - 1:
                break
            time.sleep(1.0)
        except Exception as exc:
            if attempt == MAX_RETRIES - 1:
                print(f"\n  API error for {session['code']}: {exc}")
                break
            time.sleep(API_DELAY * (attempt + 2))

    fallback = {"primary_cat": 4, "primary_subcat": "", "secondary_cat": None,
                "secondary_subcat": None, "confidence": "low"}
    cache[key] = fallback
    return fallback


def enrich_session(session: dict, tags: dict) -> dict:
    """Merge tagging results into session dict."""
    cat_num  = tags["primary_cat"]
    subcat   = tags.get("primary_subcat", "")
    cat      = CATEGORIES.get(cat_num, CATEGORIES[4])

    session["escmid_cat"]       = cat_num
    session["escmid_cat_name"]  = cat["name"]
    session["escmid_cat_short"] = cat["short"]
    session["escmid_color"]     = cat["color"]
    session["escmid_subcat"]    = subcat
    session["escmid_subcat_name"] = cat["subcats"].get(subcat, "")
    session["escmid_cat2"]      = tags.get("secondary_cat")
    session["escmid_subcat2"]   = tags.get("secondary_subcat") or ""
    session["confidence"]       = tags.get("confidence", "medium")

    # Prefix is AUTHORITATIVE for known prefixes — always overrides context
    # text extraction, which can be wrong due to two-column layout contamination.
    prefix = session.get("code", "")[:2]
    if prefix in PREFIX_TO_TYPE:
        session["type"] = PREFIX_TO_TYPE[prefix]

    return session


def tag_all(sessions: list, use_api: bool, api_extract: bool = False) -> list:
    """
    Tag all sessions.

    use_api=True      — call Anthropic API for classification only
                        (uses pre-extracted title/type from PDF parsing)
    api_extract=True  — call Anthropic API for BOTH extraction clean-up
                        AND classification in one call. Recommended for
                        best accuracy; costs the same as use_api=True.
    """
    cache = load_cache()
    client = None

    if use_api or api_extract:
        if not HAS_ANTHROPIC:
            print("  anthropic not installed — using keyword fallback")
            use_api = api_extract = False
        else:
            api_key = os.environ.get("ANTHROPIC_API_KEY", "")
            if not api_key:
                print("  ANTHROPIC_API_KEY not set — using keyword fallback")
                use_api = api_extract = False
            else:
                client = _anthropic.Anthropic(api_key=api_key)
                if api_extract:
                    print("  Mode: API extraction + classification (--api-extract)")
                else:
                    print("  Mode: API classification only (keyword extraction)")

    iterator = _tqdm(sessions, desc="  Tagging") if HAS_TQDM else sessions
    tagged = []

    for i, session in enumerate(iterator):
        if api_extract and client:
            result = api_extract_and_classify(session, client, cache)
            # Update session title/type with cleaned versions
            session["title"] = result.pop("title", session["title"])
            session["type"]  = result.pop("type",  session["type"])
            tags = result
        elif use_api and client:
            tags = api_tag_session(session, client, cache)
        else:
            tags = keyword_tag(session)

        if (use_api or api_extract) and i % SAVE_CACHE_EVERY == 0:
            save_cache(cache)
            time.sleep(API_DELAY)

        tagged.append(enrich_session(session, tags))

    save_cache(cache)
    return tagged


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def _border():
    thin = Side(style="thin", color="DDDDDD")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def _hdr_cell(ws, row, col, value, fill_hex="1F4E79"):
    c = ws.cell(row=row, column=col, value=value)
    c.fill    = PatternFill("solid", start_color=fill_hex)
    c.font    = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border  = _border()
    return c

def _data_cell(ws, row, col, value, fill_hex="FFFFFF"):
    c = ws.cell(row=row, column=col, value=value)
    c.fill    = PatternFill("solid", start_color=fill_hex)
    c.font    = Font(name="Arial", size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.border  = _border()
    return c


def write_overview_sheet(ws, sessions: list, years: list):
    """Summary matrix: ESCMID category × year."""
    _hdr_cell(ws, 1, 1, "ESCMID Category")
    for j, yr in enumerate(years, 2):
        _hdr_cell(ws, 1, j, yr)
    _hdr_cell(ws, 1, len(years)+2, "TOTAL")
    ws.row_dimensions[1].height = 30

    counts = defaultdict(lambda: defaultdict(int))
    for s in sessions:
        counts[s["escmid_cat"]][s["year"]] += 1

    for row_i, (cat_num, cat) in enumerate(CATEGORIES.items(), 2):
        h = cat["color"].lstrip("#") + "44"  # translucent tint
        _data_cell(ws, row_i, 1, f"{cat_num}. {cat['name']}", h)
        ws.cell(row_i, 1).font = Font(bold=True, name="Arial", size=10)
        total = 0
        for j, yr in enumerate(years, 2):
            v = counts[cat_num].get(yr, 0)
            total += v
            _data_cell(ws, row_i, j, v or "", h)
            ws.cell(row_i, j).alignment = Alignment(horizontal="center")
        _data_cell(ws, row_i, len(years)+2, total, h)
        ws.cell(row_i, len(years)+2).font = Font(bold=True, name="Arial", size=10)

    ws.column_dimensions["A"].width = 46
    for j in range(2, len(years)+3):
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 10
    ws.freeze_panes = "B2"


def write_category_sheet(ws, sessions: list, cat_num: int, years: list):
    """One sheet per ESCMID category."""
    cat = CATEGORIES[cat_num]
    col_hex = cat["color"].lstrip("#")

    headers = ["Year","Code","Type","Session Title",
               "Primary Subcategory","Subcategory Name",
               "Secondary Cat","Confidence"]
    for j, h in enumerate(headers, 1):
        _hdr_cell(ws, 1, j, h)
    ws.row_dimensions[1].height = 30

    # Sort by year descending
    rows = sorted(sessions, key=lambda x: x["year"], reverse=True)

    for i, s in enumerate(rows, 2):
        fill = col_hex + ("33" if i % 2 == 0 else "11")
        cat2 = CATEGORIES.get(s.get("escmid_cat2"))
        cat2_str = (f"{s['escmid_cat2']} – {cat2['short']}"
                    if cat2 else "")
        vals = [s["year"], s["code"], s["type"], s["title"],
                s.get("escmid_subcat",""), s.get("escmid_subcat_name",""),
                cat2_str, s.get("confidence","")]
        for j, v in enumerate(vals, 1):
            _data_cell(ws, i, j, v, fill)

    widths = [8, 10, 18, 72, 10, 44, 22, 10]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1,j).column_letter].width = w
    ws.freeze_panes = "A2"


def generate_excel(sessions: list, path: Path):
    if not HAS_OPENPYXL:
        print("  openpyxl not installed — skipping Excel output")
        return

    years = sorted(set(s["year"] for s in sessions))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"
    write_overview_sheet(ws, sessions, years)

    for cat_num, cat in CATEGORIES.items():
        cat_sessions = [s for s in sessions if s.get("escmid_cat") == cat_num]
        sheet_name = f"{cat_num}. {cat['short']}"[:31]
        write_category_sheet(wb.create_sheet(sheet_name), cat_sessions, cat_num, years)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"  Excel  → {path}")


# ══════════════════════════════════════════════════════════════════════════════
# HTML DASHBOARD GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def build_dashboard_data(sessions: list) -> dict:
    years = sorted(set(s["year"] for s in sessions))

    cat_year  = defaultdict(lambda: defaultdict(int))
    cat_type  = defaultdict(lambda: defaultdict(int))
    subcat_year = defaultdict(lambda: defaultdict(int))

    for s in sessions:
        cn = s.get("escmid_cat", 4)
        cat_year[cn][s["year"]] += 1
        cat_type[cn][s["type"]] += 1
        sc = s.get("escmid_subcat", "")
        if sc:
            subcat_year[sc][s["year"]] += 1

    # Build row list for explore + network
    rows = []
    for s in sessions:
        rows.append({
            "id"        : len(rows),
            "year"      : s["year"],
            "code"      : s["code"],
            "type"      : s["type"],
            "title"     : s["title"],
            "cat"       : s.get("escmid_cat", 4),
            "cat_short" : s.get("escmid_cat_short", ""),
            "cat_name"  : s.get("escmid_cat_name", ""),
            "subcat"    : s.get("escmid_subcat", ""),
            "subcat_name": s.get("escmid_subcat_name", ""),
            "cat2"      : s.get("escmid_cat2"),
            "color"     : s.get("escmid_color", "#617d9b"),
            "confidence": s.get("confidence", "medium"),
        })

    # Network: only major types to keep manageable (~500–1000 nodes)
    net_rows = [r for r in rows if r["type"] in NETWORK_TYPES]
    # Cap at 800 nodes for performance: sample proportionally across categories
    if len(net_rows) > 800:
        from random import sample, seed
        seed(42)
        by_cat = defaultdict(list)
        for r in net_rows:
            by_cat[r["cat"]].append(r)
        # Keep proportional representation
        target = 800
        sampled = []
        for cat_list in by_cat.values():
            k = max(1, round(len(cat_list) * target / len(net_rows)))
            sampled.extend(sample(cat_list, min(k, len(cat_list))))
        net_rows = sorted(sampled, key=lambda x: (x["year"], x["cat"]))

    # Re-index network rows
    for i, r in enumerate(net_rows):
        r["net_id"] = i
    net_id_map = {r["id"]: r["net_id"] for r in net_rows}

    # Build edges: same category + year → connected; cap at 2000
    edges = []
    seen  = set()
    for i, r1 in enumerate(net_rows):
        for j, r2 in enumerate(net_rows):
            if i >= j:
                continue
            same_cat  = r1["cat"] == r2["cat"]
            same_year = r1["year"] == r2["year"]
            cross_cat = r1.get("cat2") == r2["cat"] if r1.get("cat2") else False

            weight = 0
            if same_cat and same_year:
                weight = 2
            elif same_cat:
                weight = 1
            elif cross_cat and same_year:
                weight = 1

            if weight > 0:
                key = (min(i,j), max(i,j))
                if key not in seen and len(edges) < 2000:
                    seen.add(key)
                    edges.append({"s": i, "t": j, "w": weight})

    return {
        "rows"       : rows,
        "net_rows"   : net_rows,
        "net_edges"  : edges,
        "years"      : years,
        "year_total" : dict(Counter(s["year"] for s in sessions)),
        "cat_year"   : {str(k): dict(v) for k, v in cat_year.items()},
        "cat_type"   : {str(k): dict(v) for k, v in cat_type.items()},
        "subcat_year": {k: dict(v) for k, v in subcat_year.items()},
        "categories" : {
            str(k): {"name": v["name"], "short": v["short"],
                     "color": v["color"], "subcats": v["subcats"]}
            for k, v in CATEGORIES.items()
        },
    }


def generate_dashboard_html(data_json: str) -> str:
    """Return the complete HTML string for the dashboard."""

    YEAR_COLORS = {
        "2021": "#94a3b8", "2022": "#7dd3fc", "2023": "#6ee7b7",
        "2024": "#fbbf24", "2025": "#4fc3f7", "2026": "#818cf8",
    }
    year_colors_js = json.dumps(YEAR_COLORS)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ESCMID Global — Programme Analysis</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.8.5/d3.min.js"></script>
<style>
:root{{--bg:#080e1a;--surf:#0f1d2e;--surf2:#162336;--bdr:#1e3450;--text:#c8d8ec;--dim:#617d9b;--acc:#4fc3f7}}
*{{box-sizing:border-box;margin:0;padding:0}}
html{{scroll-behavior:smooth}}
body{{background:var(--bg);color:var(--text);font-family:'Sora',sans-serif;font-size:13.5px;line-height:1.6;height:100vh;display:flex;flex-direction:column;overflow:hidden}}
header{{padding:14px 36px;border-bottom:1px solid var(--bdr);background:linear-gradient(135deg,#080e1a,#0a1830,#0c1e35);flex-shrink:0;display:flex;align-items:center;gap:24px;flex-wrap:wrap}}
.logo{{font-size:18px;font-weight:700;color:#e8f4ff;letter-spacing:-.02em}}.logo span{{color:var(--acc)}}
.hstats{{display:flex;gap:24px;margin-left:auto;flex-wrap:wrap}}
.stat{{display:flex;flex-direction:column;gap:1px}}
.stat-n{{font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:600;line-height:1}}
.stat-l{{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.1em}}
.tab-bar{{display:flex;background:rgba(8,14,26,.95);border-bottom:1px solid var(--bdr);flex-shrink:0;padding:0 36px;overflow-x:auto}}
.tab-bar::-webkit-scrollbar{{height:0}}
.tab{{display:flex;align-items:center;gap:6px;padding:11px 15px;font-size:11.5px;font-weight:600;color:var(--dim);background:none;border:none;border-bottom:2px solid transparent;cursor:pointer;letter-spacing:.04em;text-transform:uppercase;white-space:nowrap;transition:color .18s,border-color .18s;font-family:'Sora',sans-serif}}
.tab:hover{{color:var(--text)}}.tab.active{{color:var(--acc);border-color:var(--acc)}}
.content{{flex:1;overflow:hidden;position:relative}}
.pane{{display:none;height:100%;overflow-y:auto;padding:28px 36px}}
.pane.active{{display:block}}
.pane::-webkit-scrollbar{{width:5px}}.pane::-webkit-scrollbar-thumb{{background:var(--bdr);border-radius:3px}}
.ptitle{{font-size:15px;font-weight:700;color:#e8f4ff;margin-bottom:6px}}
.pdesc{{font-size:12px;color:var(--dim);margin-bottom:20px;max-width:700px;line-height:1.7}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:18px}}
.g3{{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}}
.card{{background:var(--surf);border:1px solid var(--bdr);border-radius:8px;padding:18px 22px}}
.ctitle{{font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dim);margin-bottom:14px}}
.legend{{display:flex;flex-wrap:wrap;gap:8px 14px;margin-top:10px}}
.li{{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--dim)}}
.ld{{width:10px;height:10px;border-radius:50%;flex-shrink:0}}
#heat-wrap{{overflow-x:auto;padding-bottom:6px}}
.hm-cell{{height:26px;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:10px;transition:transform .15s;cursor:default;font-family:'JetBrains Mono',monospace}}
.hm-cell:hover{{transform:scale(1.12);z-index:2}}
.hm-lbl{{font-size:10.5px;color:var(--dim);display:flex;align-items:center;padding-right:8px;white-space:nowrap;height:26px}}
.hm-yr{{font-family:'JetBrains Mono',monospace;font-size:10px;text-align:center;height:20px;display:flex;align-items:center;justify-content:center;font-weight:600;color:var(--acc)}}
#net-card{{padding:0;overflow:hidden;display:flex;flex-direction:column;height:calc(100vh - 195px);min-height:480px}}
#net-ctrl{{display:flex;flex-wrap:wrap;gap:7px;align-items:center;padding:10px 16px;border-bottom:1px solid var(--bdr);background:var(--surf2);flex-shrink:0}}
.cl{{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dim);white-space:nowrap}}
.cg{{display:flex;gap:5px;flex-wrap:wrap}}
.cb{{font-family:'JetBrains Mono',monospace;font-size:10px;padding:4px 10px;border-radius:14px;border:1px solid var(--bdr);background:var(--surf);color:var(--dim);cursor:pointer;transition:all .15s;white-space:nowrap}}
.cb:hover{{border-color:rgba(255,255,255,.3);color:#e8f4ff}}.cb.on{{color:#fff}}
.csep{{width:1px;height:20px;background:var(--bdr);margin:0 3px}}
#net-body{{display:flex;flex:1;overflow:hidden;min-height:0}}
#net-svg-wrap{{flex:1;overflow:hidden;position:relative}}
#net-svg-wrap svg{{width:100%;height:100%;display:block;cursor:grab}}
#net-svg-wrap svg:active{{cursor:grabbing}}
.lnk{{stroke-opacity:.1}}
#dp{{width:260px;flex-shrink:0;border-left:1px solid var(--bdr);background:var(--surf2);overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:8px}}
#dp::-webkit-scrollbar{{width:4px}}#dp::-webkit-scrollbar-thumb{{background:var(--bdr);border-radius:2px}}
.de{{color:var(--dim);font-size:11px;text-align:center;margin:auto;line-height:2.2;opacity:.7}}
.dy{{font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;margin-bottom:3px}}
.dt{{font-size:12px;color:#e8f4ff;line-height:1.5;font-weight:600}}
.ds{{font-size:10.5px;color:var(--dim);margin-top:3px;line-height:1.4}}
.dtags{{display:flex;flex-wrap:wrap;gap:4px;margin-top:8px}}
.dtag{{font-size:9.5px;font-family:'JetBrains Mono',monospace;padding:2px 8px;border-radius:12px}}
.dnbl{{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dim);margin:8px 0 5px}}
.nbc{{font-size:10px;color:var(--dim);background:var(--surf);border:1px solid var(--bdr);border-radius:4px;padding:4px 8px;cursor:pointer;transition:border-color .15s;line-height:1.4;margin-bottom:4px;display:block;text-align:left;width:100%;font-family:'Sora',sans-serif}}
.nbc:hover{{border-color:var(--acc);color:var(--text)}}
.filters{{display:flex;flex-wrap:wrap;gap:9px;margin-bottom:14px;align-items:center}}
.fb{{font-family:'JetBrains Mono',monospace;font-size:10px;padding:5px 11px;border-radius:20px;border:1px solid var(--bdr);background:var(--surf2);color:var(--dim);cursor:pointer;transition:all .15s;white-space:nowrap}}
.fb:hover{{border-color:var(--acc);color:var(--acc)}}.fb.on{{background:rgba(79,195,247,.15);border-color:var(--acc);color:var(--acc)}}
.sb{{flex:1;min-width:180px;max-width:320px;background:var(--surf2);border:1px solid var(--bdr);color:var(--text);padding:6px 14px;border-radius:20px;font-size:12px;font-family:'Sora',sans-serif;outline:none;transition:border-color .2s}}
.sb:focus{{border-color:var(--acc)}}.sb::placeholder{{color:var(--dim)}}
.tc{{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--dim);margin-left:auto}}
.tl{{display:flex;flex-direction:column;gap:5px;max-height:calc(100vh - 310px);overflow-y:auto;padding-right:3px}}
.tl::-webkit-scrollbar{{width:4px}}.tl::-webkit-scrollbar-thumb{{background:var(--bdr);border-radius:2px}}
.ti{{background:var(--surf2);border:1px solid var(--bdr);border-radius:6px;padding:8px 13px;display:grid;grid-template-columns:36px 1fr auto;gap:6px 10px;align-items:start;transition:border-color .15s}}
.ti:hover{{border-color:rgba(79,195,247,.3)}}
.ty{{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--acc);font-weight:600;line-height:1.8}}
.tt2{{font-size:12px;color:var(--text);line-height:1.5}}
.ts{{font-size:10.5px;color:var(--dim);margin-top:2px}}
.tgs{{display:flex;flex-wrap:wrap;gap:4px;justify-content:flex-end;max-width:220px}}
.tp{{font-size:9.5px;padding:2px 7px;border-radius:12px;white-space:nowrap;font-family:'JetBrains Mono',monospace}}
.gg{{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px}}
.gc{{background:var(--surf2);border:1px solid var(--bdr);border-radius:6px;padding:14px 16px;border-left:3px solid}}
.gc.low{{border-left-color:#f87171}}.gc.med{{border-left-color:#fbbf24}}
.gct{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--dim);margin-bottom:4px;letter-spacing:.05em}}
.gcn{{font-size:22px;font-weight:700;font-family:'JetBrains Mono',monospace}}
.gcn.low{{color:#f87171}}.gcn.med{{color:#fbbf24}}
.gcx{{font-size:11px;color:var(--dim);margin-top:4px}}
.tt3{{position:fixed;background:#0a1525;border:1px solid var(--acc);border-radius:6px;padding:8px 12px;font-size:11px;pointer-events:none;z-index:9999;max-width:260px;display:none;color:var(--text);box-shadow:0 8px 24px rgba(0,0,0,.6)}}
.tt3.on{{display:block}}
@media(max-width:800px){{
  header,main{{padding:16px 18px}}
  .tab-bar{{padding:0 12px}}
  .pane{{padding:20px 18px}}
  .g2,.g3{{grid-template-columns:1fr}}
  #net-body{{flex-direction:column}}#dp{{width:100%;height:200px;border-left:none;border-top:1px solid var(--bdr)}}
  #net-card{{height:auto}}
}}
</style>
</head>
<body>
<div class="tt3" id="tt"></div>
<header>
  <div class="logo">ESCMID Global &nbsp;<span>Programme Analysis</span></div>
  <div class="hstats">
    <div class="stat"><span class="stat-n" style="color:var(--acc)" id="hd-sessions">—</span><span class="stat-l">Sessions</span></div>
    <div class="stat"><span class="stat-n" style="color:#fbbf24" id="hd-years">—</span><span class="stat-l">Years</span></div>
    <div class="stat"><span class="stat-n" style="color:#34d399">12</span><span class="stat-l">ESCMID categories</span></div>
  </div>
</header>
<div class="tab-bar" id="tab-bar">
  <button class="tab active" data-tab="overview">📊 Overview</button>
  <button class="tab" data-tab="trends">📈 Trends</button>
  <button class="tab" data-tab="heatmap">🔥 Heatmap</button>
  <button class="tab" data-tab="network">🕸 Clustering</button>
  <button class="tab" data-tab="explore">🔍 Explore</button>
  <button class="tab" data-tab="gaps">🎯 Gaps</button>
</div>
<div class="content">
  <div class="pane active" id="pane-overview">
    <div class="ptitle">Programme Overview</div>
    <p class="pdesc">Sessions per year broken down by ESCMID category. Click a category in the legend to highlight it.</p>
    <div class="g2">
      <div class="card"><div class="ctitle">Sessions per year by category</div><div style="position:relative;height:280px"><canvas id="chtOverview"></canvas></div><div class="legend" id="lgOverview"></div></div>
      <div class="card"><div class="ctitle">Session type mix per year</div><div style="position:relative;height:280px"><canvas id="chtTypes"></canvas></div></div>
    </div>
  </div>
  <div class="pane" id="pane-trends">
    <div class="ptitle">Category Trends</div>
    <p class="pdesc">Select categories to compare their trajectory over time.</p>
    <div class="card" style="margin-bottom:16px">
      <div class="ctitle">Filter categories</div>
      <div id="trend-filters" style="display:flex;flex-wrap:wrap;gap:7px"></div>
    </div>
    <div class="card"><div class="ctitle">Session count per year (selected categories)</div><div style="position:relative;height:320px"><canvas id="chtTrends"></canvas></div></div>
  </div>
  <div class="pane" id="pane-heatmap">
    <div class="ptitle">Full Category × Year Heatmap</div>
    <p class="pdesc">Each cell = number of sessions in that category and year. Hover for the count and category name.</p>
    <div class="card"><div id="heat-wrap"><div id="hmc"></div></div></div>
  </div>
  <div class="pane" id="pane-network">
    <div class="ptitle" style="margin-bottom:6px">Clustering Network</div>
    <p class="pdesc">Nodes = sessions (major types only). Edges connect sessions in the same category. Coloured by ESCMID category. <strong>Click</strong> to inspect · <strong>scroll</strong> to zoom · <strong>drag background</strong> to pan.</p>
    <div class="card" id="net-card">
      <div id="net-ctrl">
        <span class="cl">Year</span><div class="cg" id="net-yr"></div>
        <div class="csep"></div>
        <span class="cl">Category</span><div class="cg" id="net-cat"></div>
        <div class="csep"></div>
        <button class="cb on" id="btn-edges">Edges on</button>
        <button class="cb on" id="btn-cluster">Cluster on</button>
        <button class="cb" id="btn-reset" style="margin-left:auto">↺ Reset</button>
      </div>
      <div id="net-body">
        <div id="net-svg-wrap"><svg id="net-svg"></svg></div>
        <div id="dp"><div class="de">Click any node<br>to see session details<br>and connections</div></div>
      </div>
    </div>
  </div>
  <div class="pane" id="pane-explore">
    <div class="ptitle">Explore All Sessions</div>
    <p class="pdesc">Filter by year or category, or search session titles. Sorted by year, most recent first.</p>
    <div class="card">
      <div class="filters">
        <input type="text" class="sb" id="search" placeholder="Search session titles…">
        <div id="ef-years" style="display:flex;gap:6px;flex-wrap:wrap"></div>
        <div id="ef-cats" style="display:flex;flex-wrap:wrap;gap:5px"></div>
        <span class="tc" id="etcount">— sessions</span>
      </div>
      <div class="tl" id="tlist"></div>
    </div>
  </div>
  <div class="pane" id="pane-gaps">
    <div class="ptitle">Under-represented Subcategories</div>
    <p class="pdesc">ESCMID subcategories with the fewest sessions across all years — potential gaps for 2027 proposals.</p>
    <div class="gg" id="gapgrid"></div>
  </div>
</div>

<script>
const D = {data_json};
const ROWS=D.rows, YEARS=D.years, CAT_YEAR=D.cat_year;
const CATS=D.categories, YEAR_TOTAL=D.year_total;
const YCOLORS = {year_colors_js};

function h2r(hex,a){{const r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);return `rgba(${{r}},${{g}},${{b}},${{a}})`}}
function catColor(cn){{return (CATS[cn]||{{}}).color||'#617d9b'}}

// Header stats
document.getElementById('hd-sessions').textContent = ROWS.length.toLocaleString();
document.getElementById('hd-years').textContent = YEARS.length;

// ── Tab switching ──
let netInit=false;
document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>{{
  document.querySelectorAll('.tab,.pane').forEach(x=>x.classList.remove('active'));
  b.classList.add('active');
  const pane=document.getElementById('pane-'+b.dataset.tab);
  pane.classList.add('active');
  if(b.dataset.tab==='network'){{if(!netInit){{buildNetwork();netInit=true;}}else{{window._netResize&&window._netResize();}}}}
}}));

// ── Chart defaults ──
Chart.defaults.color='#617d9b';Chart.defaults.font.family="'JetBrains Mono',monospace";Chart.defaults.font.size=10;
const GRID={{color:'rgba(30,52,80,.5)'}};

// ── Overview chart ──
const catNums=Object.keys(CATS).map(Number);
new Chart(document.getElementById('chtOverview'),{{type:'bar',
  data:{{labels:YEARS,datasets:catNums.map(cn=>({{\
    label:`${{cn}}. ${{CATS[cn].short}}`,
    data:YEARS.map(y=>(CAT_YEAR[cn]||{{}})[y]||0),
    backgroundColor:h2r(CATS[cn].color,.7),
    borderColor:CATS[cn].color,borderWidth:1,borderRadius:2
  }}))
  }},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},
    tooltip:{{backgroundColor:'#0f1d2e',borderColor:'#1e3450',borderWidth:1}}}},
    scales:{{x:{{stacked:true,grid:GRID}},y:{{stacked:true,grid:GRID,min:0}}}}
  }}
}});
const lgOv=document.getElementById('lgOverview');
catNums.forEach(cn=>lgOv.innerHTML+=`<div class="li"><div class="ld" style="background:${{CATS[cn].color}}"></div><span>${{cn}}. ${{CATS[cn].short}}</span></div>`);

// ── Session type chart ──
const stypes=['Oral Session','Symposium','Educational','ePoster Flash','Meet-the-Expert','Keynote','Late-Breaking','Pipeline','Case Session','Other'];
const stC=['#4fc3f7','#818cf8','#34d399','#fbbf24','#f472b6','#fb923c','#f87171','#a3e635','#6ee7b7','#94a3b8'];
const typeCounts = {{}};
ROWS.forEach(r=>{{typeCounts[r.type]=(typeCounts[r.type]||{{}});typeCounts[r.type][r.year]=(typeCounts[r.type][r.year]||0)+1;}});
new Chart(document.getElementById('chtTypes'),{{type:'bar',
  data:{{labels:YEARS,datasets:stypes.map((s,i)=>({{\
    label:s,data:YEARS.map(y=>(typeCounts[s]||{{}})[y]||0),
    backgroundColor:stC[i]+'cc',borderColor:stC[i],borderWidth:1,borderRadius:2
  }}))
  }},options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,padding:8,color:'#617d9b',font:{{size:9}}}}}},
      tooltip:{{backgroundColor:'#0f1d2e',borderColor:'#1e3450',borderWidth:1}}}},
    scales:{{x:{{stacked:true,grid:GRID}},y:{{stacked:true,grid:GRID,min:0}}}}
  }}
}});

// ── Trends ──
const activeTrendCats = new Set(catNums);
let trendChart;
function buildTrendChart(){{
  if(trendChart)trendChart.destroy();
  trendChart=new Chart(document.getElementById('chtTrends'),{{type:'line',
    data:{{labels:YEARS,datasets:[...activeTrendCats].map(cn=>({{\
      label:`${{cn}}. ${{CATS[cn].short}}`,
      data:YEARS.map(y=>(CAT_YEAR[cn]||{{}})[y]||0),
      borderColor:CATS[cn].color,backgroundColor:h2r(CATS[cn].color,.06),
      borderWidth:2,pointRadius:5,tension:.3,fill:true
    }}))
    }},options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,padding:10,color:'#617d9b',font:{{size:9}}}}}},
        tooltip:{{backgroundColor:'#0f1d2e',borderColor:'#1e3450',borderWidth:1}}}},
      scales:{{x:{{grid:GRID}},y:{{grid:GRID,min:0}}}}
    }}
  }});
}}
const tfc=document.getElementById('trend-filters');
catNums.forEach(cn=>{{
  const b=document.createElement('button');b.className='fb on';
  b.style.cssText=`background:${{h2r(CATS[cn].color,.2)}};border-color:${{CATS[cn].color}}66;color:${{CATS[cn].color}}`;
  b.textContent=`${{cn}}. ${{CATS[cn].short}}`;
  b.onclick=()=>{{
    activeTrendCats.has(cn)?(activeTrendCats.delete(cn),b.classList.remove('on'),b.style.background='transparent')
                           :(activeTrendCats.add(cn),b.classList.add('on'),b.style.background=h2r(CATS[cn].color,.2));
    buildTrendChart();
  }};tfc.appendChild(b);
}});
buildTrendChart();

// ── Heatmap ──
const hmc=document.getElementById('hmc');
hmc.style.cssText=`display:grid;grid-template-columns:200px repeat(${{YEARS.length}},70px);gap:3px`;
hmc.innerHTML='<div></div>'+YEARS.map(y=>`<div class="hm-yr">${{y}}</div>`).join('');
const tt=document.getElementById('tt');
const maxHM=Math.max(...catNums.flatMap(cn=>YEARS.map(y=>(CAT_YEAR[cn]||{{}})[y]||0)));
catNums.forEach(cn=>{{
  hmc.innerHTML+=`<div class="hm-lbl">${{cn}}. ${{CATS[cn].short}}</div>`;
  YEARS.forEach(y=>{{
    const v=(CAT_YEAR[cn]||{{}})[y]||0;
    const a=v===0?0:0.1+(v/maxHM)*0.8;
    const c=document.createElement('div');c.className='hm-cell';c.textContent=v||'';
    c.style.background=v>0?h2r(CATS[cn].color,a):'rgba(255,255,255,.03)';
    if(v>0)c.style.color=a>0.5?'#fff':'#c8d8ec';
    c.addEventListener('mousemove',e=>{{tt.className='tt3 on';
      tt.innerHTML=`<strong style="color:${{CATS[cn].color}}">${{y}}</strong> · ${{CATS[cn].name}}<br><span style="font-family:JetBrains Mono,monospace;font-size:14px;color:#e8f4ff">${{v}}</span> sessions`;
      tt.style.left=(e.clientX+12)+'px';tt.style.top=(e.clientY-36)+'px'}});
    c.addEventListener('mouseleave',()=>tt.className='tt3');
    hmc.appendChild(c);
  }});
}});

// ── NETWORK ──
function buildNetwork(){{
  const wrap=document.getElementById('net-svg-wrap');
  let W=wrap.getBoundingClientRect().width||900;
  const ctrlH=document.getElementById('net-ctrl').offsetHeight||44;
  const cardH=document.getElementById('net-card').offsetHeight||540;
  let H=cardH-ctrlH; if(H<380)H=480;

  const svg=d3.select('#net-svg').attr('width',W).attr('height',H);
  const mainG=svg.append('g');
  const zoom=d3.zoom().scaleExtent([.1,10]).on('zoom',e=>mainG.attr('transform',e.transform));
  svg.call(zoom);

  const nodes=D.net_rows.map((r,i)=>Object.assign({{}},r,{{x:W/2+(Math.random()-.5)*250,y:H/2+(Math.random()-.5)*250}}));
  const nodeById=Object.fromEntries(nodes.map(n=>[n.net_id,n]));
  const links=D.net_edges.map(e=>({{source:nodeById[e.s],target:nodeById[e.t],w:e.w}}));

  // Cluster centres — arrange 12 categories in a 4×3 grid-ish circle
  const cx=W/2,cy=H/2,cr=Math.min(W,H)*.32;
  const clusterC={{}};
  catNums.forEach((cn,i)=>{{
    const a=(i/12)*2*Math.PI - Math.PI/2;
    clusterC[cn]={{x:cx+cr*Math.cos(a),y:cy+cr*Math.sin(a)}};
  }});

  // Zone circles
  catNums.forEach(cn=>{{
    const c=clusterC[cn];
    mainG.append('circle').attr('cx',c.x).attr('cy',c.y).attr('r',Math.min(W,H)*.11)
      .attr('fill',h2r(CATS[cn].color,.03)).attr('stroke',h2r(CATS[cn].color,.15))
      .attr('stroke-width',1).attr('stroke-dasharray','4,3');
    mainG.append('text').attr('x',c.x).attr('y',c.y-Math.min(W,H)*.115)
      .attr('text-anchor','middle').attr('fill',h2r(CATS[cn].color,.22))
      .style('font-family','Sora,sans-serif').style('font-size','9px')
      .style('font-weight','700').style('letter-spacing','.07em')
      .style('text-transform','uppercase').text(`${{cn}}. ${{CATS[cn].short}}`);
  }});

  const linkSel=mainG.append('g').selectAll('line').data(links).enter().append('line')
    .attr('class','lnk').attr('stroke-width',1)
    .attr('stroke',d=>catColor(d.source.cat));

  const nodeSel=mainG.append('g').selectAll('g.nd').data(nodes).enter().append('g').attr('class','nd')
    .call(d3.drag().on('start',ds).on('drag',dd).on('end',de));

  nodeSel.append('circle').attr('r',9).attr('fill','none')
    .attr('stroke',d=>YCOLORS[d.year]||'#617d9b').attr('stroke-width',1.2).attr('opacity',.4);
  nodeSel.append('circle').attr('class','nc').attr('r',6)
    .attr('fill',d=>h2r(catColor(d.cat),.9)).attr('stroke',d=>catColor(d.cat))
    .attr('stroke-width',1.5).style('cursor','pointer');

  function ds(e,d){{if(!e.active)sim.alphaTarget(.3).restart();d.fx=d.x;d.fy=d.y}}
  function dd(e,d){{d.fx=e.x;d.fy=e.y}}
  function de(e,d){{if(!e.active)sim.alphaTarget(0);d.fx=null;d.fy=null}}

  nodeSel.on('mousemove',function(e,d){{
    const t=d.title.substring(0,80);
    tt.className='tt3 on';
    tt.innerHTML=`<strong style="color:${{catColor(d.cat)}}">${{d.year}} · ${{d.cat_short||''}}</strong><br>${{t}}…`;
    tt.style.left=(e.clientX+14)+'px';tt.style.top=(e.clientY-44)+'px';
  }}).on('mouseleave',()=>tt.className='tt3');

  let selId=null;
  nodeSel.on('click',function(e,d){{e.stopPropagation();selectNode(d)}});
  svg.on('click',()=>selectNode(null));

  function selectNode(d){{
    selId=d?d.net_id:null;
    const panel=document.getElementById('dp');
    if(!d){{nodeSel.attr('opacity',1);linkSel.attr('stroke-opacity',.1);
      panel.innerHTML='<div class="de">Click any node<br>to see session details<br>and connections</div>';return;}}
    const nbIds=new Set();
    links.forEach(l=>{{if(l.source.net_id===d.net_id)nbIds.add(l.target.net_id);
                       if(l.target.net_id===d.net_id)nbIds.add(l.source.net_id)}});
    const connIdx=new Set();links.forEach((l,i)=>{{if(l.source.net_id===d.net_id||l.target.net_id===d.net_id)connIdx.add(i)}});
    nodeSel.attr('opacity',n=>n.net_id===d.net_id||nbIds.has(n.net_id)?1:.06);
    linkSel.attr('stroke-opacity',(l,i)=>connIdx.has(i)?.5:.01);
    const pills=`<span class="dtag" style="background:${{h2r(catColor(d.cat),.15)}};color:${{catColor(d.cat)}};border:1px solid ${{h2r(catColor(d.cat),.3)}}">${{d.cat}}. ${{d.cat_short}}</span>${{d.subcat?`<span class="dtag" style="background:rgba(255,255,255,.06);color:var(--dim)">${{d.subcat}}: ${{d.subcat_name}}</span>`:''}}`; 
    const nbChips=[...nbIds].slice(0,5).map(nid=>{{const nb=nodeById[nid];return `<button class="nbc" onclick="window._selN(${{nid}})">${{nb.year}}: ${{nb.title.substring(0,55)}}…</button>`}}).join('');
    panel.innerHTML=`<div class="dy" style="color:${{catColor(d.cat)}}">${{d.year}} · ${{d.type}}</div><div class="dt">${{d.title}}</div><div class="dtags">${{pills}}</div>${{nbIds.size?`<div class="dnbl">Connected (${{nbIds.size}})</div>${{nbChips}}${{nbIds.size>5?`<div style="font-size:10px;color:var(--dim);margin-top:3px">+${{nbIds.size-5}} more…</div>`:''}}`:''}}`;
  }}
  window._selN=id=>selectNode(nodeById[id]);

  let clusterOn=true,edgesOn=true;
  function clusterForce(alpha){{
    if(!clusterOn)return;
    nodes.forEach(n=>{{const c=clusterC[n.cat];if(!c)return;n.vx-=(n.x-c.x)*.04*alpha;n.vy-=(n.y-c.y)*.04*alpha}});
  }}
  const sim=d3.forceSimulation(nodes)
    .force('link',d3.forceLink(links).id(d=>d.net_id).distance(35).strength(d=>d.w*.15))
    .force('charge',d3.forceManyBody().strength(-70))
    .force('collide',d3.forceCollide(11).strength(.8))
    .force('cluster',clusterForce)
    .force('center',d3.forceCenter(W/2,H/2).strength(.03))
    .on('tick',()=>{{
      linkSel.attr('x1',d=>d.source.x).attr('y1',d=>d.source.y).attr('x2',d=>d.target.x).attr('y2',d=>d.target.y);
      nodeSel.attr('transform',d=>`translate(${{d.x}},${{d.y}})`);
    }});

  // Year buttons
  const yrEl=document.getElementById('net-yr');
  let activeYrs=new Set(YEARS);
  YEARS.forEach(y=>{{
    const b=document.createElement('button');b.className='cb on';
    b.style.cssText=`background:${{h2r(YCOLORS[y],.22)}};border-color:${{YCOLORS[y]}}66;color:${{YCOLORS[y]}}`;
    b.textContent=y;
    b.onclick=()=>{{activeYrs.has(y)?(activeYrs.size>1&&(activeYrs.delete(y),b.classList.remove('on'),b.style.background='transparent')):(activeYrs.add(y),b.classList.add('on'),b.style.background=h2r(YCOLORS[y],.22));applyVis()}};
    yrEl.appendChild(b);
  }});

  // Cat buttons
  const catEl=document.getElementById('net-cat');
  let activeCats=new Set(catNums);
  catNums.forEach(cn=>{{
    const b=document.createElement('button');b.className='cb on';
    b.style.cssText=`background:${{h2r(CATS[cn].color,.22)}};border-color:${{CATS[cn].color}}66;color:${{CATS[cn].color}}`;
    b.textContent=`${{cn}}`;b.title=CATS[cn].name;
    b.onclick=()=>{{activeCats.has(cn)?(activeCats.size>1&&(activeCats.delete(cn),b.classList.remove('on'),b.style.background='transparent')):(activeCats.add(cn),b.classList.add('on'),b.style.background=h2r(CATS[cn].color,.22));applyVis()}};
    catEl.appendChild(b);
  }});

  function applyVis(){{
    nodeSel.attr('opacity',d=>activeYrs.has(d.year)&&activeCats.has(d.cat)?1:.03);
    linkSel.attr('stroke-opacity',d=>{{
      const vs=activeYrs.has(d.source.year)&&activeCats.has(d.source.cat);
      const vt=activeYrs.has(d.target.year)&&activeCats.has(d.target.cat);
      return vs&&vt&&edgesOn?.1:0;
    }});
  }}
  document.getElementById('btn-edges').onclick=function(){{edgesOn=!edgesOn;this.textContent=edgesOn?'Edges on':'Edges off';this.classList.toggle('on',edgesOn);linkSel.attr('stroke-opacity',edgesOn?.1:0)}};
  document.getElementById('btn-cluster').onclick=function(){{clusterOn=!clusterOn;this.textContent=clusterOn?'Cluster on':'Cluster off';this.classList.toggle('on',clusterOn);if(clusterOn)sim.alpha(.4).restart()}};
  document.getElementById('btn-reset').onclick=()=>{{svg.transition().duration(500).call(zoom.transform,d3.zoomIdentity);selectNode(null)}};
  window._netResize=()=>{{const nW=wrap.getBoundingClientRect().width;if(nW>50){{W=nW;svg.attr('width',W);sim.force('center',d3.forceCenter(W/2,H/2).strength(.03)).alpha(.2).restart()}}}};
  new ResizeObserver(()=>{{if(netInit)window._netResize&&window._netResize()}}).observe(wrap);
}}

// ── EXPLORE ──
const sortedRows=[...ROWS].sort((a,b)=>b.year.localeCompare(a.year)||a.title.localeCompare(b.title));
let eYear=null,eCats=new Set(),eSearch='';

const eyEl=document.getElementById('ef-years');
['All',...YEARS].forEach(y=>{{
  const b=document.createElement('button');b.className='fb'+(y==='All'?' on':'');b.textContent=y;
  b.onclick=()=>{{eYear=y==='All'?null:y;document.querySelectorAll('#ef-years .fb').forEach(x=>x.classList.remove('on'));b.classList.add('on');renderTalks()}};
  eyEl.appendChild(b);
}});

const ecEl=document.getElementById('ef-cats');
catNums.forEach(cn=>{{
  const b=document.createElement('button');b.className='fb';
  b.style.borderColor=h2r(CATS[cn].color,.4);
  b.textContent=`${{cn}}. ${{CATS[cn].short}}`;
  b.onclick=()=>{{eCats.has(cn)?(eCats.delete(cn),b.classList.remove('on'),b.style.background=''):(eCats.add(cn),b.classList.add('on'),b.style.background=h2r(CATS[cn].color,.18));renderTalks()}};
  ecEl.appendChild(b);
}});

function renderTalks(){{
  const q=eSearch.toLowerCase();
  const f=sortedRows.filter(r=>{{
    if(eYear&&r.year!==eYear)return false;
    if(eCats.size>0&&!eCats.has(r.cat))return false;
    if(q&&!r.title.toLowerCase().includes(q))return false;
    return true;
  }});
  document.getElementById('etcount').textContent=f.length.toLocaleString()+' sessions';
  document.getElementById('tlist').innerHTML=f.map(r=>{{
    const cat=CATS[r.cat]||{{}};
    const pill=`<span class="tp" style="background:${{h2r(cat.color||'#617d9b',.15)}};color:${{cat.color||'#617d9b'}};border:1px solid ${{h2r(cat.color||'#617d9b',.3)}}">${{r.cat}}. ${{cat.short||''}}</span>`;
    const spill=r.subcat?`<span class="tp" style="background:rgba(255,255,255,.05);color:var(--dim)">${{r.subcat}}</span>`:'';
    return `<div class="ti"><div class="ty">${{r.year}}</div><div class="tt2">${{r.title}}<div class="ts">${{r.type}}</div></div><div class="tgs">${{pill}}${{spill}}</div></div>`;
  }}).join('');
}}
renderTalks();
document.getElementById('search').addEventListener('input',e=>{{eSearch=e.target.value;renderTalks()}});

// ── GAPS ──
const subcatTotal={{}};
ROWS.forEach(r=>{{if(r.subcat)subcatTotal[r.subcat]=(subcatTotal[r.subcat]||0)+1}});
const gg=document.getElementById('gapgrid');
// All defined subcats sorted by count ascending
const allSubcats=[];
Object.entries(D.categories).forEach(([cn,cat])=>Object.entries(cat.subcats).forEach(([sc,sn])=>allSubcats.push([sc,sn,+cn,subcatTotal[sc]||0])));
allSubcats.sort((a,b)=>a[3]-b[3]);
allSubcats.slice(0,24).forEach(([sc,sn,cn,count])=>{{
  const cls=count<=2?'low':'med';
  gg.innerHTML+=`<div class="gc ${{cls}}"><div class="gct">${{sc}}: ${{sn}}</div><div class="gcn ${{cls}}">${{count}}</div><div class="gcx">in ${{CATS[cn].name}}</div></div>`;
}});
</script>
</body>
</html>"""


def generate_dashboard(sessions: list, path: Path):
    data = build_dashboard_data(sessions)
    html = generate_dashboard_html(json.dumps(data))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
    print(f"  HTML   → {path}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(description="ESCMID EPC Copilot")
    ap.add_argument("--years",        help="Comma-separated years, e.g. 2024,2025,2026")
    ap.add_argument("--skip-tagging", action="store_true", help="Use keyword fallback only (free, less accurate)")
    ap.add_argument("--api-extract",  action="store_true",
                    help="Use API for BOTH extraction clean-up AND classification (recommended)")
    ap.add_argument("--skip-excel",   action="store_true", help="Skip Excel output")
    ap.add_argument("--skip-dash",    action="store_true", help="Skip HTML dashboard")
    ap.add_argument("--programmes-dir", default=str(PROGRAMMES_DIR),
                    help=f"Directory containing PDFs (default: {PROGRAMMES_DIR})")
    ap.add_argument("--output-dir",   default=str(OUTPUT_DIR),
                    help=f"Output directory (default: {OUTPUT_DIR})")
    args = ap.parse_args()

    prog_dir = Path(args.programmes_dir)
    out_dir  = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Select years
    if args.years:
        selected = set(y.strip() for y in args.years.split(","))
    else:
        selected = set(PROGRAMME_FILES.keys())

    pdf_map = {
        yr: prog_dir / fname
        for yr, fname in PROGRAMME_FILES.items()
        if yr in selected
    }

    print("=" * 58)
    print(" ESCMID EPC Copilot")
    print("=" * 58)

    # ── 1. Extract ────────────────────────────────────────────
    print("\n[1/3] Extracting sessions from PDFs …")
    sessions = extract_all(pdf_map)
    print(f"  Total: {len(sessions)} sessions across {len(pdf_map)} years")

    if not sessions:
        print("  No sessions found. Check PDF paths and try again.")
        sys.exit(1)

    # ── 2. Tag ────────────────────────────────────────────────
    print("\n[2/3] Tagging sessions …")
    use_api = not args.skip_tagging and not args.api_extract
    sessions = tag_all(sessions, use_api=use_api, api_extract=args.api_extract)

    # Save raw JSON (useful for debugging / re-running without re-tagging)
    raw_path = out_dir / "sessions_raw.json"
    raw_path.write_text(json.dumps(sessions, indent=2))
    print(f"  Raw data → {raw_path}")

    # Category counts summary
    counts = Counter(s["escmid_cat"] for s in sessions)
    print("  Category breakdown:")
    for cn in sorted(counts):
        print(f"    {cn:2d}. {CATEGORIES[cn]['short']:<22} {counts[cn]:>4} sessions")

    # ── 3. Output ─────────────────────────────────────────────
    print("\n[3/3] Generating outputs …")
    if not args.skip_excel:
        generate_excel(sessions, out_dir / "ESCMID_Programmes.xlsx")
    if not args.skip_dash:
        generate_dashboard(sessions, out_dir / "ESCMID_Dashboard.html")

    print(f"\nDone — outputs in {out_dir}")


if __name__ == "__main__":
    main()
