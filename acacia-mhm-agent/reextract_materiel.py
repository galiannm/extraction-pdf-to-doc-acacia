"""Re-extract only page 3 MATÉRIEL section with PS/MS split."""
from dotenv import load_dotenv; load_dotenv()
import json
from config import EXTRACTED_DIR, SOURCE_PDFS
from extractors.pdf_extractor import _make_client, _render_page
from google.genai import types
import fitz

pdf = SOURCE_PDFS[1]
doc = fitz.open(str(pdf))
client = _make_client()

PROMPT = '''You are extracting page 3 of a French MHM teacher guide (Période 1, Semaine 1).

This page has TWO sections:
1. A Programmation table (week schedule) — extract as a normal table block
2. A MATÉRIEL section — this has a LEFT column for PS and a RIGHT column for MS

For the MATÉRIEL section, use EXACTLY this structure to keep PS and MS content separate:
{"type": "heading", "level": 2, "text": "MATÉRIEL"}
{"type": "heading", "level": 3, "text": "PS — Ressources"}
{"type": "paragraph", "text": "• item1\\n• item2\\n..."}
{"type": "heading", "level": 3, "text": "MS — Ressources"}
{"type": "paragraph", "text": "• item1\\n..."}
{"type": "heading", "level": 3, "text": "PS — Matériel de classe"}
{"type": "paragraph", "text": "• item1\\n..."}
{"type": "heading", "level": 3, "text": "MS — Matériel de classe"}
{"type": "paragraph", "text": "• item1\\n..."}

CRITICAL: PS = LEFT column of the PDF, MS = RIGHT column. Do NOT merge them.
Extract ALL bullet points from each column separately.
Return ONLY: {"pages": [{"page_num": 3, "blocks": [...]}]}'''

img = _render_page(doc, 2)
parts = [
    types.Part.from_bytes(data=img, mime_type='image/jpeg'),
    types.Part.from_text(text=PROMPT)
]
response = client.models.generate_content(
    model='gemini-2.5-flash', contents=parts,
    config=types.GenerateContentConfig(
        temperature=0,
        thinking_config=types.ThinkingConfig(thinking_budget=0)
    ),
)
raw = response.text.strip()
if raw.startswith('```'):
    raw = raw.split('\n', 1)[1].rsplit('```', 1)[0]

# Clean JSON
result = []; in_string = False; escaped = False
_ESC = {'\n': '\\n', '\r': '\\r', '\t': '\\t'}
for ch in raw:
    if escaped: result.append(ch); escaped = False
    elif ch == '\\' and in_string: result.append(ch); escaped = True
    elif ch == '"': result.append(ch); in_string = not in_string
    elif in_string and ord(ch) < 0x20: result.append(_ESC.get(ch, ''))
    else: result.append(ch)
new_blocks = json.loads(''.join(result))['pages'][0]['blocks']

print('New blocks:')
for b in new_blocks:
    print(f'  [{b["type"]}]', str(b.get('text', ''))[:70])

# Patch: keep old Programmation table, replace MATÉRIEL section
json_path = EXTRACTED_DIR / pdf.stem / f'{pdf.stem}.json'
with open(json_path, encoding='utf-8') as f:
    full_data = json.load(f)

old_blocks = full_data['pages'][2]['blocks']
keep_old = [b for b in old_blocks if b.get('type') in ('activity_title', 'table')]
mat_start = next(
    (i for i, b in enumerate(new_blocks)
     if b.get('type') == 'heading' and 'MAT' in (b.get('text') or '').upper()),
    None
)
patched = keep_old + new_blocks[mat_start:] if mat_start is not None else new_blocks
full_data['pages'][2]['blocks'] = patched

with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(full_data, f, ensure_ascii=False, indent=2)

print(f'\nPatched page 3: kept {len(keep_old)} old + {len(patched)-len(keep_old)} new MATÉRIEL blocks. Saved.')
doc.close()
