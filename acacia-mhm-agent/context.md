# Acacia MHM Agent — Session Context

## What This Project Does

Local Python pipeline that reads 6 French MHM (Méthode Heuristique de Mathématiques) teacher guide PDFs and produces two branded Word documents:
- `Acacia_MHM_FR.docx` — full French content with Acacia brand styling
- `Acacia_MHM_EN.docx` — same structure, text translated to English

PDFs live at `../HMH/` (parent of the project dir). Output goes to `output/`.

---

## File Structure

```
acacia-mhm-agent/
├── main.py                     # CLI entry point
├── config.py                   # Colors, fonts, file paths
├── requirements.txt
├── .env                        # GEMINI_API_KEY=...
├── assets/
│   ├── logo-acacia.jpeg        # Full logo (used in header + cover)
│   └── logo-A.jpeg             # Small logo
├── extractors/
│   └── pdf_extractor.py        # PDF → JSON via Gemini OCR
├── ai/
│   └── gemini_agent.py         # FR→EN translation + glossary
├── generators/
│   └── docx_generator.py       # JSON → .docx dispatch
├── templates/
│   └── acacia_styles.py        # All python-docx styling helpers
├── output/
│   ├── extracted/<doc_name>/   # Per-PDF: .json + images/page_NNNN.jpg
│   └── *.docx
└── glossary.json               # MHM terminology, built during translation
```

---

## How to Run

```bash
cd acacia-mhm-agent
source .venv/bin/activate

python main.py --test              # intro PDF only, extract + translate + build both docs
python main.py --test --extract-only   # intro PDF only, extract + build FR only (no Gemini translation calls)
python main.py --all               # all 6 PDFs
```

After extraction, you can rebuild the docx from cached JSON without re-calling Gemini:
```python
from config import EXTRACTED_DIR, OUTPUT_DIR
from extractors.pdf_extractor import load_extracted
from generators.docx_generator import build_document

data = load_extracted("MHM PS MS - Teachers' guide (K1-K2) intro", EXTRACTED_DIR)
build_document(data['pages'], OUTPUT_DIR / 'test.docx', 'Title', lang='fr-FR')
```

---

## Pipeline Stages

1. **Extract** (`pdf_extractor.py`): renders each PDF page to JPEG (PyMuPDF), sends to Gemini 2.0 Flash inline (no file upload), receives structured JSON with `heading/paragraph/table/caption` blocks. Saves JSON + page JPEGs to `output/extracted/<doc>/`.

2. **Translate** (`gemini_agent.py`): sends 10-page chunks of JSON to Gemini, gets back translated JSON + new glossary terms. Builds `glossary.json` incrementally.

3. **Build** (`docx_generator.py` + `acacia_styles.py`): walks page blocks, dispatches to style helpers, saves `.docx`.

---

## Key Technical Decisions

### Why inline JPEG, not Files API
The Files API upload hangs indefinitely on large PDFs (21 MB+). Inline JPEG — one page at a time — is reliable. Each page renders to ~100 KB at 1.2× zoom, JPEG quality 70.

### Retry + lightweight fallback
`_call_with_retry` in `pdf_extractor.py`: 5 attempts, backoff `8s × attempt`. On attempt 3+, switches to grayscale/low-quality render (~25 KB) to avoid output token limits on image-heavy pages. Also switches to `_COMPACT_PROMPT` (minimal JSON structure) on later attempts.

### JSON cleaning
Gemini sometimes embeds raw `\n`/`\t` inside JSON string values. `_clean_json()` walks char-by-char and escapes control characters only inside JSON strings (same logic duplicated in `gemini_agent.py` as `_clean_json_strings()`).

### SDK
Uses `google-genai` (NOT the deprecated `google-generativeai`):
```python
from google import genai
from google.genai import types
client = genai.Client(api_key=..., http_options=types.HttpOptions(timeout=120_000))
```
Timeout is in milliseconds.

---

## Acacia Brand Colors (`config.py`)

```python
class Colors:
    BLUE   = (91,  200, 245)   # sky blue
    ORANGE = (245, 166,  35)   # warm orange
    YELLOW = (248, 200,  64)   # golden yellow
    CORAL  = (232, 120, 106)
    GREEN  = (141, 198,  63)
    PURPLE = (176, 124, 198)
    WHITE      = (255, 255, 255)
    DARK_GREY  = (64,   64,  64)
    LIGHT_GREY = (240, 240, 240)

    # MHM activity type mapping
    RITUALISED_ACTIVITIES = YELLOW
    GUIDED_LEARNING       = BLUE
    AUTONOMOUS_ACTIVITIES = GREEN
    PS_BADGE              = BLUE
    MS_BADGE              = ORANGE
    WEEK_ROW              = ORANGE

class Fonts:
    HEADING = "Mali"
    BODY    = "Nunito"
```

---

## Styling Rules (`acacia_styles.py`)

- **Tables**: full content-width (6.27"), column widths proportional to max cell content length via `_proportional_col_widths()`. Header row → purple. Week rows → orange. PS/MS badge cells → blue/orange. Activity-type columns get their category color.
- **Bullets**: paragraphs with `•` or `- ` lines are split into individual indented bullet paragraphs. Continuation lines (non-bullet lines after a bullet start) are joined to their parent bullet.
- **Soft hyphens**: `_fix_soft_hyphens()` joins `word-\nword` and `word- word` OCR artifacts. Regex: `r'-[\s\n]+([a-zàâäéèêëîïôùûüçœæ])'`
- **Term highlighting**: `_add_highlighted_runs()` splits text on `_HIGHLIGHT_RE` and applies color+bold to key MHM terms (activités ritualisées→yellow, guidées→blue, autonomes→green, évaluation formative/sommative→orange, Petite/Moyenne Section, MHM→orange). Works for both FR and EN.
- **Language**: `setup_document(doc, lang)` sets proofing language on Normal style + docDefaults. FR doc → `fr-FR`, EN doc → `en-US`. Per-run language set via `_set_run_lang()`.
- **Images**: pages that have `caption` blocks get their saved page JPEG embedded after the text blocks (from `page["image_path"]`).
- **Period banners**: heading text starting with `PÉRIODE` / `PERIOD` triggers `add_period_banner()` — a full-width orange banner.

---

## Known Issues / Next Steps

- **Page 3 extraction failed** (image-heavy Avant-propos page hit token limit even on retry). Logged as `[Extraction error on page 3: retry later]` in the JSON. May need manual re-extraction or a different prompt strategy for that page.
- **Images only appear from new extractions** — the existing JSON (before image-saving was added) doesn't have `image_path` keys. Re-run `--extract-only` to regenerate with images.
- **Only tested on intro PDF** — need to run `--all` for the full 6-PDF pipeline once layout is approved.
- **Translation not yet run** — `Acacia_MHM_EN.docx` hasn't been produced yet; only FR has been built and iterated on.
- **Output versioning** — current output is `Acacia_MHM_FR_v4.docx`. When ready for final run, use the standard names `Acacia_MHM_FR.docx` / `Acacia_MHM_EN.docx`.

---

## Gemini API Notes

- Model: `gemini-2.0-flash`
- Billing enabled on Google Cloud (THB 400 credit)
- Free tier was 15 RPM — caused rate-limit errors before billing was enabled
- Extraction: `temperature=0`, inline JPEG parts + text prompt
- Translation: 10-page chunks, system prompt in the user message (no system role in Gemini API)
