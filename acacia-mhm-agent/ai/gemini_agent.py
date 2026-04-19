"""
Translates extracted page blocks from French to English using Gemini.
Builds and maintains a MHM terminology glossary across documents.
"""
import json
import os
import copy
from pathlib import Path

from google import genai
from dotenv import load_dotenv
from rich.console import Console

from config import GLOSSARY_PATH

load_dotenv()
console = Console()

_SYSTEM_PROMPT = """You are an expert educational translator specialising in early-years mathematics pedagogy.
Your task is to translate French MHM (Méthode Heuristique de Mathématiques) teacher guide content into English.

Rules:
- Translate ALL text fields faithfully. Do NOT simplify, summarise, or omit any content.
- Preserve every structural marker (type, level, data keys) exactly.
- Use the provided glossary for consistent terminology. Add new MHM terms you encounter.
- Adapt only minimally for an international classroom (e.g. replace purely French cultural references only when unavoidable, and note the change).
- Return ONLY valid JSON — no markdown fences, no extra text.

Response format:
{
  "pages": [ <same structure as input, with text fields translated> ],
  "new_glossary_terms": { "french_term": "english_translation", ... }
}"""


def translate_pages(pages: list[dict]) -> list[dict]:
    """Translate a list of page dicts FR→EN, updating the glossary."""
    client = _make_client()
    glossary = _load_glossary()

    translated_pages = []
    chunk_size = 10
    for i in range(0, len(pages), chunk_size):
        chunk = pages[i : i + chunk_size]
        console.print(
            f"  Translating pages {chunk[0]['page_num']}–{chunk[-1]['page_num']}…"
        )
        translated_chunk, new_terms = _translate_chunk_with_retry(client, chunk, glossary)
        translated_pages.extend(translated_chunk)
        if new_terms:
            glossary.update(new_terms)
            _save_glossary(glossary)
            console.print(f"    [dim]+ {len(new_terms)} glossary terms[/dim]")

    return translated_pages


def _translate_chunk_with_retry(
    client, chunk: list[dict], glossary: dict
) -> tuple[list[dict], dict]:
    """Try chunk as-is; on failure retry page-by-page."""
    translated, new_terms = _translate_chunk(client, chunk, glossary)
    # Detect fallback: if any page text is still French (unchanged from input)
    if translated is chunk and len(chunk) > 1:
        console.print("  [yellow]Retrying page-by-page…[/yellow]")
        all_pages: list[dict] = []
        all_terms: dict = {}
        for page in chunk:
            page_result, page_terms = _translate_chunk(client, [page], glossary)
            all_pages.extend(page_result)
            all_terms.update(page_terms)
            glossary.update(page_terms)
        return all_pages, all_terms
    return translated, new_terms


def _translate_chunk(
    client, chunk: list[dict], glossary: dict
) -> tuple[list[dict], dict]:
    payload = json.dumps({"pages": chunk}, ensure_ascii=False)
    glossary_str = json.dumps(glossary, ensure_ascii=False, indent=2)

    prompt = f"""{_SYSTEM_PROMPT}

Current glossary:
{glossary_str}

Translate this content:
{payload}"""

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt,
        )
        raw = response.text.strip()
        if raw.startswith("```"):
            raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
        raw = _clean_json_strings(raw)
        data = json.loads(raw)
        return data.get("pages", chunk), data.get("new_glossary_terms", {})
    except Exception as e:
        console.print(f"  [red]Translation error:[/red] {e}")
        return chunk, {}


def _clean_json_strings(raw: str) -> str:
    """Same fix as in pdf_extractor — escape control chars inside JSON strings."""
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
    from google.genai import types as gtypes
    return genai.Client(
        api_key=api_key,
        http_options=gtypes.HttpOptions(timeout=120_000),  # 120s for translation
    )


def _load_glossary() -> dict:
    if GLOSSARY_PATH.exists():
        with open(GLOSSARY_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_glossary(glossary: dict) -> None:
    with open(GLOSSARY_PATH, "w", encoding="utf-8") as f:
        json.dump(glossary, f, ensure_ascii=False, indent=2)
    console.print(f"  [dim]Glossary saved ({len(glossary)} terms)[/dim]")
