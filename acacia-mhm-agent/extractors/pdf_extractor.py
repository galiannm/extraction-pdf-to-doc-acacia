"""
PDF extraction using Gemini inline image batches.

Renders each page to a compressed JPEG (~100-150 KB) and sends batches of
PAGES_PER_BATCH pages directly in the API call — no file upload, no hanging.
"""
import json
import os
import time
import fitz  # PyMuPDF
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types
from rich.console import Console

load_dotenv()
console = Console()

PAGES_PER_BATCH  = 1    # 1 page at a time — most reliable, eliminates size issues
MAX_RETRIES      = 5
JPEG_QUALITY     = 70
INTER_BATCH_WAIT = 1    # seconds between batches
RETRY_WAIT_BASE  = 8    # seconds for first retry (×attempt each time)
REQUEST_TIMEOUT  = 120  # 2 minutes per page — plenty for dense content


def _extraction_prompt(start_page: int, n_pages: int, extra_instructions: str = "") -> str:
    page_range = (
        f"page {start_page}" if n_pages == 1
        else f"pages {start_page} to {start_page + n_pages - 1}"
    )
    extra = f"\n\nADDITIONAL INSTRUCTIONS:\n{extra_instructions.strip()}" if extra_instructions.strip() else ""
    return f"""You are extracting structured content from {page_range} of a French early-years mathematics teacher guide (MHM — Méthode Heuristique de Mathématiques).

Each image is one page, in order. Return ALL content as JSON:
{{
  "pages": [
    {{
      "page_num": {start_page},
      "blocks": [
        {{"type": "activity_title", "periode_num": 1, "week": "S1", "activity_label": "Activité ritualisée", "activity_type": "ritualisee", "classes": ["PS", "MS"]}},
        {{"type": "subtitle_badge", "badge_text": "CHAQUE JOUR", "title_text": "Le rituel de l'étiquette d'appel", "activity_type": "ritualisee"}},
        {{"type": "heading", "level": 1, "text": "..."}},
        {{"type": "heading", "level": 2, "text": "..."}},
        {{"type": "heading", "level": 3, "text": "..."}},
        {{"type": "paragraph", "text": "..."}},
        {{"type": "table", "data": [["h1","h2"],["r1c1","r1c2"]]}},
        {{"type": "caption", "text": "..."}}
      ]
    }}
  ]
}}

Block rules:
- activity_title: USE THIS ONLY for the visual banner element — a wide colored pill/banner that spans the page, with a white box on the left showing "Période N" (small) and "S[week]" (large) in the period color, and the activity type name in white on a colored background.
  Fields:
  • periode_num: integer (1–5)
  • week: string like "S1", "S2", etc.
  • activity_label: exact French label as written ("Activité ritualisée", "Apprentissages guidés", "Autonomie et semi-autonomie")
  • activity_type: one of "ritualisee", "guidee", "autonomie"
  • classes: array containing "PS" and/or "MS" — only include classes shown as badges on that specific banner
  ⚠️ DO NOT use activity_title for rows in Programmation/summary tables — those must remain as table blocks. The Programmation overview page shows a grid of all weeks (SEMAINE 1–5) × all activity types (ritualisées, guidés, autonomie) — extract the ENTIRE grid as ONE table block with column headers and week rows.
  ⚠️ DO NOT use activity_title when the activity type appears as a simple text heading without the visual banner design.
  ⚠️ Page running headers (small "Période N" text printed at the top of each page) must be IGNORED completely — do not extract them as content, and do not use them for periode_num. Only use the periode_num that is clearly part of the main visual activity banner on the page.
  ⚠️ DO NOT use activity_title for individual activity names like "Le coin marchande", "Un peu/beaucoup", "Activité 1", "Activité 2" — those are headings, not activity banners.

- subtitle_badge: USE THIS for day/timing badges — a small colored pill followed by a subtitle title on the same line.
  Fields:
  • badge_text: MUST be a timing word ONLY — "CHAQUE JOUR", "JOUR 1", "JOUR 2", "JOUR 3", "JOUR 4", "JOUR 5", or similar. NEVER put "PS", "MS", or "PS MS" here — if PS or MS appears in a badge-like position before a subtitle, use a heading level 3 block instead.
  • title_text: the activity name that follows the badge on the same line (do not include PS/MS in title_text either — those go in the parent activity_title classes field)
  • activity_type: must match the parent activity_title's activity_type ("ritualisee", "guidee", or "autonomie")

- heading 1: largest titles (period titles, major section names) — NOT used for activity banners
- heading 2: sub-section titles AND all pedagogical section labels — Avant-propos, SOMMAIRE, Objectifs, Déroulement, Matériel, Différenciation, Ce qu'il faut savoir, Ressources, Matériel de classe, Remarques
- heading 3: numbered sub-sections, PS/MS labels when standalone
- paragraph: body text; each visual paragraph in the original MUST be a separate block — NEVER merge multiple paragraphs into one
  BOLD: you MUST wrap every word or phrase that appears in bold font in the PDF with **double asterisks** — e.g. **Guide de la méthode**, **PS**, **un jeu par période**. This is critical: scan every line for bold text and mark it. Do not skip any bold words.
- table: full 2D array — preserve ALL rows/columns, PS/MS labels, week labels, sub-headers
- caption: text immediately below/beside an illustration

STYLE CAPTURE (add only when color is visually present — skip for plain black text on white):
- For any block with a visible colored background, add "bg_color": one of: "yellow", "blue", "green", "orange", "purple", "red", "grey", or a hex code like "#F5C842". Also add "text_color": "white" or "black".
- For table blocks, add "col_colors": ["color", ...] — one entry per column representing the column header background. Use "none" for uncolored columns.
- Example with style: {{"type": "heading", "level": 2, "text": "Activités ritualisées", "bg_color": "yellow", "text_color": "white"}}
- Example table: {{"type": "table", "data": [...], "col_colors": ["none", "yellow", "blue", "green"]}}

Rules:
- One page object per image, in order, starting from page_num {start_page}
- Preserve every word exactly as written (French, accents, special chars)
- Do NOT translate, simplify, or omit anything{extra}
- Return ONLY the JSON — no markdown fences"""


def extract_pdf(pdf_path: Path, extracted_dir: Path, extra_instructions: str = "", force: bool = False) -> dict:
    doc_name = pdf_path.stem
    out_dir  = extracted_dir / doc_name
    json_path = out_dir / f"{doc_name}.json"

    if not force and json_path.exists():
        console.print(f"[green]Extraction cache hit:[/green] {json_path.name}")
        with open(json_path, encoding="utf-8") as f:
            return json.load(f)

    images_dir = out_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    console.print(f"[cyan]Extracting:[/cyan] {pdf_path.name}")

    client = _make_client()
    doc    = fitz.open(str(pdf_path))
    total  = len(doc)
    all_pages: list[dict] = []

    for start in range(0, total, PAGES_PER_BATCH):
        end   = min(start + PAGES_PER_BATCH, total)
        batch = list(range(start, end))
        label = f"pages {start+1}–{end}/{total}"
        console.print(f"  Extracting {label}…", end=" ")

        # Render pages to JPEG bytes — save to disk and pass bytes for Gemini
        jpeg_list = []
        for i in batch:
            img_bytes = _render_page(doc, i)
            img_path  = images_dir / f"page_{i+1:04d}.jpg"
            img_path.write_bytes(img_bytes)
            jpeg_list.append(img_bytes)

        pages = _call_with_retry(client, doc, batch, jpeg_list, start_page=start + 1, extra_instructions=extra_instructions)

        for page_dict, i in zip(pages, batch):
            page_dict["image_path"] = str(images_dir / f"page_{i+1:04d}.jpg")

        all_pages.extend(pages)
        console.print(f"[green]✓[/green] ({len(pages)} pages, {sum(len(p['blocks']) for p in pages)} blocks)")
        if end < total:
            time.sleep(INTER_BATCH_WAIT)

    doc.close()

    result = {"document": doc_name, "total_pages": total, "pages": all_pages}

    json_path = out_dir / f"{doc_name}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    n_blocks = sum(len(p["blocks"]) for p in all_pages)
    console.print(f"  [green]Done[/green] — {len(all_pages)} pages, {n_blocks} blocks → {json_path.name}")
    return result


def load_extracted(doc_name: str, extracted_dir: Path) -> dict:
    json_path = extracted_dir / doc_name / f"{doc_name}.json"
    with open(json_path, encoding="utf-8") as f:
        return json.load(f)




# ── Rendering ─────────────────────────────────────────────────────────────────

def _render_page(doc: fitz.Document, page_index: int, lightweight: bool = False) -> bytes:
    """Render a page to JPEG. lightweight=True → grayscale + low quality for heavy pages."""
    page = doc[page_index]
    if lightweight:
        # Grayscale, smaller, very low quality — photos become noise, text stays readable
        pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0), colorspace=fitz.csGRAY)
        return pix.tobytes("jpeg", jpg_quality=35)
    pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
    return pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY)


# ── Gemini call with retry ────────────────────────────────────────────────────

def _call_with_retry(client, doc, batch: list[int], jpeg_list: list[bytes], start_page: int, extra_instructions: str = "") -> list:
    last_err = ""
    current_jpegs = jpeg_list
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            use_compact = attempt > 2
            return _call_gemini(client, current_jpegs, start_page, compact=use_compact, extra_instructions=extra_instructions)
        except Exception as e:
            last_err = str(e)
            err_short = last_err[:70]
            if attempt < MAX_RETRIES:
                wait = RETRY_WAIT_BASE * attempt
                console.print(f"\n    [yellow]Attempt {attempt} failed ({err_short}) — retry in {wait}s[/yellow]", end=" ")
                time.sleep(wait)
                # On attempt 3+, switch to lightweight grayscale renders to reduce size
                if attempt >= 2:
                    current_jpegs = [_render_page(doc, i, lightweight=True) for i in batch]
            else:
                console.print(f"\n    [red]Failed after {MAX_RETRIES} attempts:[/red] {err_short}")
                return [{"page_num": start_page + i, "blocks": [
                    {"type": "paragraph", "text": f"[Extraction error on page {start_page + i}: retry later]"}
                ]} for i in range(len(jpeg_list))]


_COMPACT_PROMPT = """Extract text from this page as JSON. Keep each visual heading and paragraph as a separate block.
Return ONLY:
{"pages": [{"page_num": PAGE, "blocks": [
  {"type": "heading", "level": 2, "text": "section title here"},
  {"type": "paragraph", "text": "paragraph text here"}
]}]}
Rules: heading for titles/section names, paragraph for body text. One block per visual paragraph. Preserve French exactly. No markdown fences."""


def _call_gemini(client, jpeg_list: list[bytes], start_page: int, compact: bool = False, extra_instructions: str = "") -> list:
    parts = []
    for jpeg in jpeg_list:
        parts.append(types.Part.from_bytes(data=jpeg, mime_type="image/jpeg"))
    prompt = _COMPACT_PROMPT.replace("PAGE", str(start_page)) if compact else _extraction_prompt(start_page, len(jpeg_list), extra_instructions)
    parts.append(types.Part.from_text(text=prompt))

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=parts,
        config=types.GenerateContentConfig(
            temperature=0,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
        ),
    )

    raw = response.text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]

    data = json.loads(_clean_json(raw))
    return data.get("pages", [])


# ── Client ────────────────────────────────────────────────────────────────────

def _clean_json(raw: str) -> str:
    """
    Escape control characters that appear inside JSON string values.
    Gemini sometimes embeds raw newlines/tabs in string content without escaping them.
    Walks the JSON character-by-character so we only touch chars inside strings.
    """
    result = []
    in_string = False
    escaped = False
    _ESCAPE = {'\n': '\\n', '\r': '\\r', '\t': '\\t'}

    for ch in raw:
        if escaped:
            result.append(ch)
            escaped = False
        elif ch == '\\' and in_string:
            result.append(ch)
            escaped = True
        elif ch == '"':
            result.append(ch)
            in_string = not in_string
        elif in_string and ord(ch) < 0x20:
            result.append(_ESCAPE.get(ch, ''))   # escape or drop
        else:
            result.append(ch)

    return ''.join(result)


def _make_client():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY not set in .env")
    return genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=REQUEST_TIMEOUT * 1000),
    )
