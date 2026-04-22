"""
Editorial cleanup stage: takes raw extracted pages and restructures them
for teacher readability using Gemini.

The stage groups blocks into logical sections (intro / SOMMAIRE / week units)
and sends each section to Gemini with an editorial prompt.  Tables and
toc_table blocks are passed through unchanged.  The stage returns a
single-page structure with all cleaned blocks, ready for build_document.
"""
import json
import os
import re
import time
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types
from rich.console import Console

load_dotenv()
console = Console()

# Section markers that delimit editorial chunks
_SEMAINE_RE  = re.compile(r'^(Semaine|Semame|Week)\s+\d+', re.IGNORECASE)
_PERIOD_RE   = re.compile(r'^(PÉRIODE|PERIOD)\s*\d+', re.IGNORECASE)
_SOMMAIRE_RE = re.compile(r'SOMMAIRE', re.IGNORECASE)

MAX_BLOCKS_PER_CHUNK = 40    # keep chunks small enough for reliable JSON output
RETRY_WAIT = 10              # seconds between retries
MAX_RETRIES = 3


# ── Public API ────────────────────────────────────────────────────────────────

def edit_pages(pages: list[dict], doc_name: str, edited_dir: Path, system_prompt: str | None = None, force: bool = False) -> list[dict]:
    """
    Restructure extracted pages for readability.

    Returns pages list with a single page containing all cleaned blocks.
    Result is cached to edited_dir so re-runs are instant.
    """
    cache_path = edited_dir / f"{doc_name}_edited.json"
    if force and cache_path.exists():
        cache_path.unlink()
    if cache_path.exists():
        console.print(f"[green]Editorial cache hit:[/green] {cache_path.name}")
        with open(cache_path, encoding="utf-8") as f:
            return json.load(f)

    client = _make_client()
    _active_system_prompt = system_prompt if system_prompt is not None else _SYSTEM_PROMPT
    all_blocks = [b for page in pages for b in page.get("blocks", [])]
    period_banner, all_blocks = _extract_period_banner(all_blocks)
    all_blocks = _deduplicate_period_headings(all_blocks)
    if period_banner:
        all_blocks = [period_banner] + all_blocks
    sections   = _split_into_sections(all_blocks)

    console.print(f"[cyan]Editorial cleanup:[/cyan] {len(sections)} sections, {len(all_blocks)} blocks total")

    cleaned: list[dict] = []
    for i, section in enumerate(sections):
        label = _section_label(section)
        console.print(f"  [{i+1}/{len(sections)}] {label} ({len(section)} blocks)…", end=" ")

        if _is_passthrough_section(section):
            cleaned.extend(section)
            console.print("[dim]pass-through[/dim]")
        else:
            edited = _edit_section(client, section, _active_system_prompt)
            cleaned.extend(edited)
            console.print(f"[green]✓[/green] → {len(edited)} blocks")

        time.sleep(1)

    cleaned = _sanitize_blocks(cleaned)
    result = [{"page_num": 0, "blocks": cleaned}]
    edited_dir.mkdir(parents=True, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    console.print(f"  [green]Saved →[/green] {cache_path.name}")
    return result


# ── Pre-processing ────────────────────────────────────────────────────────────

def _extract_period_banner(blocks: list[dict]) -> tuple[dict | None, list[dict]]:
    """Pull out the first PÉRIODE heading to use as the single banner."""
    for i, block in enumerate(blocks):
        if block.get("type") == "heading" and _PERIOD_RE.match((block.get("text") or "").strip()):
            text = (block.get("text") or "").strip()
            # Only use it if it's a pure period label (no extra content after)
            if not _PERIOD_PREFIX_RE.sub("", text).strip():
                return block, blocks[:i] + blocks[i+1:]
    return None, blocks


_PERIOD_PREFIX_RE = re.compile(
    r'^(PÉRIODE|PERIOD)\s*\d+[\s\n\-—]*',
    re.IGNORECASE,
)

def _deduplicate_period_headings(blocks: list[dict]) -> list[dict]:
    """
    Clean up repeated "PÉRIODE N" page-header noise.

    Two cases:
    1. Heading text is ONLY a period label (e.g. "Période 1") → drop it entirely;
       the period banner is already rendered once via add_period_banner.
    2. Heading text starts with a period label followed by real content
       (e.g. "Période 1 S4 Activités ritualisées MS") → strip the prefix,
       leaving just "S4 Activités ritualisées MS".
    """
    result = []
    for block in blocks:
        if block.get("type") != "heading":
            result.append(block)
            continue

        text = (block.get("text") or "").strip()

        if not _PERIOD_RE.match(text):
            result.append(block)
            continue

        # Strip the "Période N" prefix
        cleaned = _PERIOD_PREFIX_RE.sub("", text).strip()

        if not cleaned:
            # Was only a period label — drop it (banner already shown once)
            continue

        # Keep with prefix removed
        result.append({**block, "text": cleaned})

    return result


# ── Section splitting ─────────────────────────────────────────────────────────

def _split_into_sections(blocks: list[dict]) -> list[list[dict]]:
    """
    Split flat block list into logical sections.

    Boundaries: PÉRIODE banners, Semaine headings, SOMMAIRE heading.
    Each section is processed independently by the editorial LLM.
    """
    if not blocks:
        return []

    sections: list[list[dict]] = []
    current: list[dict] = []

    for block in blocks:
        text  = (block.get("text") or "").strip()
        btype = block.get("type", "")

        is_boundary = (
            btype == "heading" and (
                _PERIOD_RE.match(text) or
                _SEMAINE_RE.match(text) or
                _SOMMAIRE_RE.search(text)
            )
        )

        if is_boundary and current:
            sections.append(current)
            current = []

        current.append(block)

        # Hard cap: split oversized sections
        if len(current) >= MAX_BLOCKS_PER_CHUNK:
            sections.append(current)
            current = []

    if current:
        sections.append(current)

    return sections


def _section_label(section: list[dict]) -> str:
    for block in section[:3]:
        text = (block.get("text") or "").strip()
        if text:
            return text[:60]
    return "(untitled)"


def _is_passthrough_section(section: list[dict]) -> bool:
    """
    Skip editorial processing for:
    - SOMMAIRE / TOC sections (handled by _inject_toc_table)
    - Sections that are mostly tables
    - Very short sections (≤ 3 blocks)
    """
    if len(section) <= 2:
        return True

    for block in section[:3]:
        text = (block.get("text") or "").strip()
        if _SOMMAIRE_RE.search(text):
            return True

    table_count = sum(1 for b in section if b.get("type") == "table")
    if table_count / len(section) > 0.5:
        return True

    return False


# ── Editorial LLM call ────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
Tu es un graphiste éditorial professionnel spécialisé dans la conception de ressources pédagogiques pour l'école maternelle (kindergarten), travaillant pour une école bilingue internationale haut de gamme.

Tu reçois du contenu brut extrait d'un guide pédagogique MHM (Méthode Heuristique de Mathématiques) pour PS/MS, sous forme de blocs JSON.

MISSION : Restructure et réécris ce contenu pour en faire un guide enseignant clair, lisible et professionnel, prêt à être publié dans une école internationale premium (IB / Cambridge / AEFE).

RÈGLES DE STRUCTURE :
• Pour chaque séance/activité, utilise cette hiérarchie :
  - heading level 2 : titre de la semaine ou de l'activité (ex. "Semaine 1 — Activité ritualisée")
  - heading level 3 : labels de sections en majuscules : OBJECTIF, MATÉRIEL, DÉROULEMENT, DIFFÉRENCIATION, NOTES
  - paragraph : contenu en paragraphes courts ou listes à puces (• item)
• Ne crée PAS de heading level 1 sauf pour les titres de période majeurs.

RÈGLES ÉDITORIALES :
• Paragraphes courts (max 3-4 lignes)
• Listes à puces (•) pour le matériel, les étapes et les objectifs
• **Gras** autour des mots-clés importants (utilise les doubles astérisques)
• Supprime les formulations redondantes et les artefacts d'extraction (numéros de page isolés, URLs, textes parasites)
• Langage professionnel, clair, concis, pédagogique

ESPACES VISUELS :
• Là où un visuel améliorerait la compréhension, insère un bloc image_placeholder
• Format : {"type": "image_placeholder", "description": "description courte du visuel (ex: enfants manipulant des cubes)"}
• Maximum 1-2 image_placeholder par section

IMPORTANT :
• Ne supprime AUCUNE information pédagogique
• Réécris et restructure l'intégralité — ne copie pas le texte brut tel quel
• Préserve toutes les activités, objectifs et consignes
• Les blocs de type "table" doivent être reproduits IDENTIQUEMENT dans la sortie

FORMAT DE SORTIE — retourne UNIQUEMENT ce JSON, sans balises markdown :
{
  "blocks": [
    {"type": "heading", "level": 2, "text": "..."},
    {"type": "heading", "level": 3, "text": "OBJECTIF"},
    {"type": "paragraph", "text": "• objectif 1\\n• objectif 2"},
    {"type": "image_placeholder", "description": "enfants triant des objets par taille"},
    {"type": "table", "data": [["col1","col2"],["r1","r2"]]},
    {"type": "paragraph", "text": "Suite du contenu..."}
  ]
}
"""

_INPUT_TEMPLATE = """\
Voici les blocs bruts à restructurer :

{blocks_json}

Restructure et réécris l'intégralité de ce contenu selon les règles ci-dessus.
Retourne UNIQUEMENT le JSON — aucun commentaire, aucune balise markdown."""


def _edit_section(client, section: list[dict], system_prompt: str = _SYSTEM_PROMPT) -> list[dict]:
    blocks_json = json.dumps(section, ensure_ascii=False, indent=2)
    prompt = _INPUT_TEMPLATE.format(blocks_json=blocks_json)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=[
                    types.Part.from_text(text=system_prompt),
                    types.Part.from_text(text=prompt),
                ],
                config=types.GenerateContentConfig(temperature=0.2),
            )
            raw = response.text.strip()
            if raw.startswith("```"):
                raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
            data = json.loads(_clean_json(raw))
            blocks = _sanitize_blocks(data.get("blocks", []))
            if blocks:
                return blocks
            # Empty result — fall through to retry
        except Exception as e:
            err = str(e)[:80]
            if attempt < MAX_RETRIES:
                console.print(f"\n    [yellow]Attempt {attempt} failed ({err}) — retry in {RETRY_WAIT}s[/yellow]", end=" ")
                time.sleep(RETRY_WAIT)
            else:
                console.print(f"\n    [red]Editorial failed after {MAX_RETRIES} attempts — keeping original blocks[/red]")
                return section

    return section


# ── Block sanitization ────────────────────────────────────────────────────────

def _sanitize_blocks(blocks: list) -> list[dict]:
    """Ensure all blocks have string text/description fields (LLM sometimes returns dicts)."""
    clean = []
    for block in blocks:
        if not isinstance(block, dict):
            continue
        btype = block.get("type", "paragraph")
        if btype in ("heading", "paragraph", "caption", "qr_label", "image_placeholder"):
            # Coerce text/description to string
            for key in ("text", "description"):
                val = block.get(key)
                if val is not None and not isinstance(val, str):
                    block = dict(block)
                    block[key] = " ".join(str(v) for v in val.values()) if isinstance(val, dict) else str(val)
        elif btype == "table":
            data = block.get("data", [])
            if not isinstance(data, list):
                continue
            # Ensure all cells are strings
            block = dict(block)
            block["data"] = [
                [str(cell) if not isinstance(cell, str) else cell for cell in row]
                for row in data if isinstance(row, list)
            ]
        clean.append(block)
    return clean


# ── JSON cleaning (same as pdf_extractor) ────────────────────────────────────

def _clean_json(raw: str) -> str:
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
            result.append(_ESCAPE.get(ch, ''))
        else:
            result.append(ch)
    return ''.join(result)


# ── Client ────────────────────────────────────────────────────────────────────

def _make_client():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY not set in .env")
    return genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=120_000),
    )
