"""
Translate Acacia_MHM_FR.docx → Acacia_MHM_EN.docx by:
1. Copying the FR docx verbatim (all styling preserved)
2. Translating all body text in-place
3. Re-applying English MHM term highlighting on body paragraphs
"""
import json
import os
import re
import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from dotenv import load_dotenv
from google import genai
from google.genai import types as gtypes
from rich.console import Console

from config import Fonts
from templates.acacia_styles import _add_highlighted_runs

load_dotenv()
console = Console()

BATCH_SIZE = 40

_PROMPT = """You are a native English-speaking academic expert in early childhood education and pedagogy.
Your task is to produce a high-quality English version of each French text below.

IMPORTANT:
- This is NOT a literal translation
- This is NOT a simplification
- You must preserve ALL key ideas, concepts, and pedagogical intentions

TRANSLATION APPROACH:
- Rewrite the content in natural, fluent, professional English
- Ensure it reads as if it was originally written by a native English-speaking education expert
- Use appropriate educational terminology used in international schools
- Maintain structure, sections, and pedagogical clarity

STYLE REQUIREMENTS:
- Academic but accessible tone
- Clear, precise, and professional language
- Natural flow (no "translated feeling")
- Consistent vocabulary throughout

CRITICAL GOAL:
The final English version must give the impression that:
→ it was originally written in English
→ by a native expert in early childhood education
NOT translated from French

RULES:
- Return ONLY valid JSON: {"translations": ["...", "...", ...]}
- The array must contain EXACTLY the same number of items as the input array
- Preserve bullet symbols (•, -) at the start of lines exactly as they appear
- Preserve tab characters (\\t) where they appear in the input
- Preserve newlines (\\n) in the output where they appear in the input
- Do NOT translate: codes like N1, G2, E4, P1–P5, the letter X, page numbers, or URLs
- If an item is already English, a code, or a symbol — return it unchanged
- Translate "Semaine" → "Week", "Période" → "Period", "Petite Section" → "Kindergarten 1", "Moyenne Section" → "Kindergarten 2"
"""


def translate_docx(
    fr_path: Path,
    en_path: Path,
    title: str,
    subtitle: str,
    fr_title: str = "Guide des séances MHM",
    fr_subtitle: str = "Petite Section / Moyenne Section — Acacia International Pre-school",
) -> Path:
    """Copy FR docx → translate all body text in-place → save as EN docx."""
    console.print(f"[cyan]Translating docx:[/cyan] {fr_path.name} → {en_path.name}")
    shutil.copy(str(fr_path), str(en_path))
    doc = Document(str(en_path))
    client = _make_client()

    items = _collect_items(doc)
    console.print(f"  {len(items)} text segments collected")

    texts = [text for _, text in items]
    translated_texts = _batch_translate(client, texts)

    for (para, _), new_text in zip(items, translated_texts):
        if new_text and new_text.strip():
            _apply_to_para(para, new_text)

    _patch_cover(doc, fr_title, title, fr_subtitle, subtitle)

    doc.save(str(en_path))
    console.print(f"  [green]Saved →[/green] {en_path}")
    return en_path


# ── Collection ────────────────────────────────────────────────────────────────

def _collect_items(doc: Document) -> list[tuple]:
    items = []
    seen = set()

    def add(para):
        pid = id(para)
        if pid in seen:
            return
        seen.add(pid)
        text = para.text
        if _should_translate(text):
            items.append((para, text))

    for para in doc.paragraphs:
        add(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    add(para)

    return items


def _should_translate(text: str) -> bool:
    t = text.strip()
    if not t or len(t) <= 2:
        return False
    if re.fullmatch(r'[NnGgEe]\d+', t):   # competence codes
        return False
    if re.fullmatch(r'\d+', t):            # page numbers
        return False
    if t.lower().startswith('http'):
        return False
    if re.fullmatch(r'[Xx]', t):           # X marks in tables
        return False
    return True


# ── Translation ───────────────────────────────────────────────────────────────

def _batch_translate(client, texts: list[str]) -> list[str]:
    results = []
    total = len(texts)
    for i in range(0, total, BATCH_SIZE):
        batch = texts[i : i + BATCH_SIZE]
        n = len(results) + len(batch)
        console.print(f"  Batch {i // BATCH_SIZE + 1}: items {i+1}–{n}/{total}…")
        translated = _call_gemini(client, batch)
        results.extend(translated)
    return results


def _call_gemini(client, texts: list[str]) -> list[str]:
    payload = json.dumps({"texts": texts}, ensure_ascii=False)
    prompt = f"{_PROMPT}\n\nTexts to translate:\n{payload}"

    for attempt in range(1, 4):
        try:
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt,
                config=gtypes.GenerateContentConfig(temperature=0),
            )
            raw = response.text.strip()
            if raw.startswith("```"):
                raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
            data = json.loads(_clean_json(raw))
            result = data.get("translations", [])
            if len(result) == len(texts):
                return result
            console.print(f"  [yellow]Length mismatch ({len(result)} vs {len(texts)}), retry {attempt}…[/yellow]")
        except Exception as e:
            console.print(f"  [red]Attempt {attempt} failed:[/red] {e}")

    console.print("  [yellow]Returning originals as fallback[/yellow]")
    return texts


# ── Application ───────────────────────────────────────────────────────────────

def _apply_to_para(para, new_text: str) -> None:
    runs = para.runs
    if not runs:
        return

    font_name = runs[0].font.name

    if font_name == Fonts.HEADING:
        # Heading (Mali): simple replace, keep existing run styling
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""
    else:
        # Body / table cell: clear runs, re-render with EN term highlighting
        size_emu = runs[0].font.size
        font_size_pt = int(size_emu / 12700) if size_emu else 11
        for r_el in para._p.findall(qn('w:r')):
            para._p.remove(r_el)
        _add_highlighted_runs(para, new_text, font_size_pt, 'en-US')


def _patch_cover(doc, fr_title, en_title, fr_subtitle, en_subtitle) -> None:
    for para in doc.paragraphs:
        t = para.text.strip()
        if t == fr_title:
            _simple_replace(para, en_title)
        elif t == fr_subtitle:
            _simple_replace(para, en_subtitle)


def _simple_replace(para, new_text: str) -> None:
    runs = para.runs
    if not runs:
        return
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


# ── Helpers ───────────────────────────────────────────────────────────────────

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


def _make_client():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY not set in .env")
    return genai.Client(
        api_key=api_key,
        http_options=gtypes.HttpOptions(timeout=120_000),
    )
