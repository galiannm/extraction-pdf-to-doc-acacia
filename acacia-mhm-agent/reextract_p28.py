"""Re-extract page 28 with a focused prompt for the competences table."""
import json
import os
import fitz
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

PDF_PATH = Path(__file__).parent.parent / "HMH" / "MHM PS MS - Teachers' guide (K1-K2) intro.pdf"
JSON_PATH = (
    Path(__file__).parent
    / "output/extracted/MHM PS MS - Teachers' guide (K1-K2) intro"
    / "MHM PS MS - Teachers' guide (K1-K2) intro.json"
)

PROMPT = """Extract page 28 of this French early-years mathematics teacher guide as JSON.

This page contains a competences table titled 'Les compétences travaillées MS'.
The table has 7 columns: code (e.g. N1, G1, E1) | description text | P1 | P2 | P3 | P4 | P5
The table also has sub-section header rows with no code (e.g. 'Utiliser les nombres', 'Explorer des formes...').

Return ONLY this JSON (no markdown fences):
{
  "pages": [{
    "page_num": 28,
    "blocks": [
      {"type": "heading", "level": 1, "text": "..."},
      {"type": "paragraph", "text": "..."},
      {"type": "table", "data": [["header1","header2",...],["r1c1","r1c2",...],...]}
    ]
  }]
}

Rules:
- The header row of the competences table must be: ["", "", "P1", "P2", "P3", "P4", "P5"]
- Sub-section rows (category headers within the table) must be: ["", "Utiliser les nombres", "", "", "", "", ""]
- Each competence row must be: ["N1", "description text", "x or empty", "x or empty", ...]
- Use lowercase x for checked cells, empty string for unchecked
- Preserve French text exactly
- Do NOT extract table rows as paragraphs — ALL competence rows must be inside the table"""


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


def main():
    doc = fitz.open(str(PDF_PATH))
    page = doc[27]  # 0-indexed → page 28
    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
    img_bytes = pix.tobytes("jpeg", jpg_quality=80)
    doc.close()

    client = genai.Client(
        api_key=os.getenv("GEMINI_API_KEY"),
        http_options=types.HttpOptions(timeout=180_000),
    )

    parts = [
        types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"),
        types.Part.from_text(text=PROMPT),
    ]

    print("Calling Gemini for page 28…")
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=parts,
        config=types.GenerateContentConfig(temperature=0),
    )

    raw = response.text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]

    data = json.loads(_clean_json(raw))
    new_page = data["pages"][0]
    print(f"Extracted {len(new_page['blocks'])} blocks")
    for b in new_page["blocks"]:
        if b["type"] == "table":
            print(f"  Table: {len(b['data'])} rows × {max(len(r) for r in b['data'])} cols")
            print(f"  Header row: {b['data'][0]}")
            print(f"  Row 1: {b['data'][1] if len(b['data']) > 1 else '(none)'}")
        else:
            print(f"  {b['type']}: {b.get('text','')[:70]}")

    # Patch the JSON
    with open(JSON_PATH, encoding="utf-8") as f:
        full = json.load(f)

    for i, p in enumerate(full["pages"]):
        if p["page_num"] == 28:
            new_page["image_path"] = p.get("image_path", "")
            full["pages"][i] = new_page
            print("Patched page 28 in JSON.")
            break

    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(full, f, ensure_ascii=False, indent=2)
    print(f"Saved → {JSON_PATH.name}")


if __name__ == "__main__":
    main()
