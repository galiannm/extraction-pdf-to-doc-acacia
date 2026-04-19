"""Re-extract programmation pages (2-3) of Period 1 with a focused table prompt."""
import json
import os
import fitz
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

PDF_PATH = Path(__file__).parent.parent / "HMH" / "MHM PS MS - Teachers' guide (K1-K2) period 1.pdf"
JSON_PATH = (
    Path(__file__).parent
    / "output/extracted/MHM PS MS - Teachers' guide (K1-K2) period 1"
    / "MHM PS MS - Teachers' guide (K1-K2) period 1.json"
)

# Pages to re-extract (1-indexed)
TARGET_PAGES = [2, 3]

PROMPT_TEMPLATE = """Extract page {page_num} of this French early-years mathematics teacher guide (MHM) as JSON.

This page shows a weekly programming overview table (Programmation / Période {period}) for PS and MS students.

The table structure is:
- Optional title row at the top (e.g. "PÉRIODE 1 — PS MS")
- A header row with 4 columns: [row label column (empty or "SEMAINE N") | "Activités ritualisées" | "Apprentissages guidés" | "Autonomie et semi-autonomie"]
- For each SEMAINE (week), there is a spanning label row "SEMAINE N" followed by sub-rows for "PS MS", "PS", and/or "MS"
- Each cell contains short text descriptions of activities

Return ONLY this JSON (no markdown fences):
{{
  "pages": [{{
    "page_num": {page_num},
    "blocks": [
      {{"type": "heading", "level": 1, "text": "any page title if present"}},
      {{"type": "table", "data": [
        ["", "Activités ritualisées", "Apprentissages guidés", "Autonomie et semi-autonomie"],
        ["SEMAINE 1", "", "", ""],
        ["PS MS", "activity text", "activity text", "activity text"],
        ["PS", "activity text", "activity text", "activity text"],
        ["MS", "activity text", "activity text", "activity text"],
        ...
      ]}}
    ]
  }}]
}}

Rules:
- Extract the ENTIRE programmation table as a SINGLE table block — do NOT split it into paragraphs or headings
- The first column contains either "SEMAINE N" (week header spanning the row), "PS MS", "PS", or "MS" (level labels)
- For SEMAINE N rows, put the week label in column 0 and empty strings in columns 1-3
- Preserve all French text exactly as written, including bold markers if any
- If there are any introductory paragraphs or headings before or after the table, include them as heading/paragraph blocks
- Do NOT convert table rows into separate heading or paragraph blocks"""


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


def reextract_page(client, doc, page_num: int) -> dict:
    """Re-extract a single page. page_num is 1-indexed."""
    page = doc[page_num - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
    img_bytes = pix.tobytes("jpeg", jpg_quality=85)

    prompt = PROMPT_TEMPLATE.format(page_num=page_num, period=1)
    parts = [
        types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"),
        types.Part.from_text(text=prompt),
    ]

    print(f"Calling Gemini for page {page_num}…")
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=parts,
        config=types.GenerateContentConfig(temperature=0),
    )

    raw = response.text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]

    data = json.loads(_clean_json(raw))
    return data["pages"][0]


def main():
    doc = fitz.open(str(PDF_PATH))

    client = genai.Client(
        api_key=os.getenv("GEMINI_API_KEY"),
        http_options=types.HttpOptions(timeout=180_000),
    )

    with open(JSON_PATH, encoding="utf-8") as f:
        full = json.load(f)

    for page_num in TARGET_PAGES:
        new_page = reextract_page(client, doc, page_num)
        print(f"  Extracted {len(new_page['blocks'])} blocks for page {page_num}")
        for b in new_page["blocks"]:
            if b["type"] == "table":
                rows = b["data"]
                cols = max(len(r) for r in rows) if rows else 0
                print(f"    Table: {len(rows)} rows × {cols} cols")
                print(f"    Header: {rows[0] if rows else '(empty)'}")
                if len(rows) > 1:
                    print(f"    Row 1:  {rows[1]}")
            else:
                print(f"    {b['type']}: {b.get('text', '')[:80]}")

        # Patch the JSON — preserve image_path from original
        for i, p in enumerate(full["pages"]):
            if p["page_num"] == page_num:
                new_page["image_path"] = p.get("image_path", "")
                full["pages"][i] = new_page
                print(f"  Patched page {page_num} in JSON.")
                break

    doc.close()

    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(full, f, ensure_ascii=False, indent=2)
    print(f"Saved → {JSON_PATH.name}")
    print()
    print("Now rebuild the FR doc:")
    print("  cd acacia-mhm-agent && python main.py --stage build_fr --period 1")


if __name__ == "__main__":
    main()
