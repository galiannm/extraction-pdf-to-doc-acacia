"""
Helpers to apply Acacia brand styles to a python-docx Document.
"""
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_PERIOD_ENTRY_RE = re.compile(r'^(PÉRIODE|PERIOD)\s*\d+', re.IGNORECASE)
_WEEK_ENTRY_RE   = re.compile(r'^(Semaine|Semame|Week)\s+', re.IGNORECASE)

from config import Colors, Fonts, LOGO_PATH

# ── Page layout ──────────────────────────────────────────────────────────────

PAGE_WIDTH   = Inches(8.27)   # A4
PAGE_HEIGHT  = Inches(11.69)
MARGIN       = Inches(1.0)
CONTENT_WIDTH = Inches(6.27)  # PAGE_WIDTH - 2*MARGIN


def setup_document(doc: Document, lang: str = 'fr-FR') -> None:
    """Apply page size, margins, base paragraph font, and proofing language."""
    section = doc.sections[0]
    section.page_width  = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, MARGIN)

    normal = doc.styles["Normal"]
    normal.font.name = Fonts.BODY
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(*Colors.DARK_GREY)
    _set_style_lang(normal, lang)
    _set_doc_default_lang(doc, lang)


# ── Header / Footer ───────────────────────────────────────────────────────────

def add_header(doc: Document) -> None:
    section = doc.sections[0]
    section.different_first_page_header_footer = False
    header = section.header
    for p in header.paragraphs:
        p.clear()
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run()
    if LOGO_PATH.exists():
        run.add_picture(str(LOGO_PATH), height=Pt(30))


def add_footer(doc: Document) -> None:
    section = doc.sections[0]
    footer = section.footer
    for p in footer.paragraphs:
        p.clear()
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("www.acacia-education.com    ")
    run.font.name = Fonts.BODY
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(*Colors.DARK_GREY)

    _add_page_number_field(p, size=9)


def _add_page_number_field(para, size: int = 9) -> None:
    """Append a Word PAGE field run to an existing paragraph."""
    run = para.add_run()
    run.font.name = Fonts.BODY
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(*Colors.DARK_GREY)
    r = run._r
    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    r.append(begin)
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    r.append(instr)
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    r.append(end)


# ── Cover page ────────────────────────────────────────────────────────────────

def add_cover_page(doc: Document, title: str, subtitle: str = "") -> None:
    """Acacia-branded cover: logo centred, title, optional subtitle."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(60)
    if LOGO_PATH.exists():
        run = p.add_run()
        run.add_picture(str(LOGO_PATH), width=Inches(2.5))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(40)
    run = p.add_run(title)
    run.font.name = Fonts.HEADING
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(*Colors.ORANGE)

    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        run = p.add_run(subtitle)
        run.font.name = Fonts.BODY
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(*Colors.DARK_GREY)

    doc.add_page_break()


# ── Period cover banner ───────────────────────────────────────────────────────

def add_period_banner(doc: Document, period_num: int, date_range: str = "") -> None:
    """Coloured banner row to mark a new period."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.cell(0, 0)
    _shade_cell(cell, _rgb_hex(Colors.ORANGE))
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(8)
    run = p.add_run(f"Period {period_num}")
    run.font.name  = Fonts.HEADING
    run.font.size  = Pt(28)
    run.font.bold  = True
    run.font.color.rgb = RGBColor(*Colors.WHITE)
    if date_range:
        run2 = p.add_run(f"  {date_range}")
        run2.font.name  = Fonts.BODY
        run2.font.size  = Pt(14)
        run2.font.color.rgb = RGBColor(*Colors.WHITE)
    _remove_table_borders(table)
    doc.add_paragraph()


# ── Language helpers ─────────────────────────────────────────────────────────

def _set_style_lang(style, lang: str) -> None:
    rPr = style.element.get_or_add_rPr()
    lang_el = rPr.find(qn('w:lang'))
    if lang_el is None:
        lang_el = OxmlElement('w:lang')
        rPr.append(lang_el)
    lang_el.set(qn('w:val'), lang)


def _set_doc_default_lang(doc: Document, lang: str) -> None:
    styles_el = doc.styles.element
    docDefaults = styles_el.find(qn('w:docDefaults'))
    if docDefaults is None:
        return
    rPrDefault = docDefaults.find(qn('w:rPrDefault'))
    if rPrDefault is None:
        rPrDefault = OxmlElement('w:rPrDefault')
        docDefaults.append(rPrDefault)
    rPr = rPrDefault.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        rPrDefault.append(rPr)
    lang_el = rPr.find(qn('w:lang'))
    if lang_el is None:
        lang_el = OxmlElement('w:lang')
        rPr.append(lang_el)
    lang_el.set(qn('w:val'), lang)


def _set_run_lang(run, lang: str) -> None:
    rPr = run._r.get_or_add_rPr()
    lang_el = rPr.find(qn('w:lang'))
    if lang_el is None:
        lang_el = OxmlElement('w:lang')
        rPr.append(lang_el)
    lang_el.set(qn('w:val'), lang)


# ── Term highlighting ─────────────────────────────────────────────────────────

# Each entry: (regex_pattern, color_rgb, bold)
_HIGHLIGHT_TERMS = [
    # Activity types — FR
    (r'activités?\s+ritualisées?',           Colors.RITUALISED_ACTIVITIES, True),
    (r'activités?\s+guidées?',               Colors.GUIDED_LEARNING,       True),
    (r'activités?\s+semi-autonomes?',        Colors.AUTONOMOUS_ACTIVITIES, True),
    (r'activités?\s+autonomes?',             Colors.AUTONOMOUS_ACTIVITIES, True),
    # Activity types — EN
    (r'ritualiz(?:ed)?\s+activit(?:y|ies)',  Colors.RITUALISED_ACTIVITIES, True),
    (r'guided\s+(?:learning|activit\w*)',    Colors.GUIDED_LEARNING,       True),
    (r'semi-autonomous\s+activit\w*',        Colors.AUTONOMOUS_ACTIVITIES, True),
    (r'autonomous\s+activit\w*',             Colors.AUTONOMOUS_ACTIVITIES, True),
    # Assessment — FR & EN
    (r'évaluation\s+formative',              Colors.ORANGE, True),
    (r'évaluation\s+sommative',              Colors.ORANGE, True),
    (r'formative\s+assessment',              Colors.ORANGE, True),
    (r'summative\s+assessment',              Colors.ORANGE, True),
    # Grade labels — FR
    (r'Petite\s+Section',                    Colors.PS_BADGE,  True),
    (r'Moyenne\s+Section',                   Colors.MS_BADGE,  True),
    # Method name
    (r'\bMHM\b',                             Colors.ORANGE,    True),
    (r'\bheuristique\b',                     Colors.ORANGE,    True),
]

# Compile once
_HIGHLIGHT_RE = re.compile(
    '(' + '|'.join(p for p, _, _ in _HIGHLIGHT_TERMS) + ')',
    re.IGNORECASE,
)
_TERM_COLOR: dict[str, tuple] = {}  # cache lowercase → (color, bold)
_BOLD_RE = re.compile(r'\*\*(.+?)\*\*', re.DOTALL)


def _highlight_lookup(matched: str) -> tuple[tuple, bool]:
    key = matched.lower()
    if key not in _TERM_COLOR:
        for pattern, color, bold in _HIGHLIGHT_TERMS:
            if re.fullmatch(pattern, matched, re.IGNORECASE):
                _TERM_COLOR[key] = (color, bold)
                break
        else:
            _TERM_COLOR[key] = (None, False)
    return _TERM_COLOR[key]


def _parse_bold_segments(text: str) -> list[tuple[str, bool]]:
    """Split text on **...** markers → list of (segment, is_bold)."""
    segments: list[tuple[str, bool]] = []
    last = 0
    for m in _BOLD_RE.finditer(text):
        if m.start() > last:
            segments.append((text[last:m.start()], False))
        segments.append((m.group(1), True))
        last = m.end()
    if last < len(text):
        segments.append((text[last:], False))
    return segments or [(text, False)]


def _add_highlighted_runs(p, text: str, base_size: int, lang: str) -> None:
    """Add runs to paragraph `p`, respecting **bold** markup and MHM term colours."""
    for seg_text, is_markup_bold in _parse_bold_segments(text):
        for chunk in _HIGHLIGHT_RE.split(seg_text):
            if not chunk:
                continue
            run = p.add_run(chunk)
            run.font.name = Fonts.BODY
            run.font.size = Pt(base_size)
            color, _ = _highlight_lookup(chunk)
            if color:
                run.font.color.rgb = RGBColor(*color)
            if is_markup_bold:
                run.font.bold = True
            _set_run_lang(run, lang)


# ── Text helpers ──────────────────────────────────────────────────────────────

def _fix_soft_hyphens(text: str) -> str:
    """Join words split by PDF line-break hyphens: 'manipu- lations' or 'manipu-\nlations' → 'manipulations'."""
    # Match hyphen + optional spaces/newline + lowercase letter (not a real hyphen like "enseignant-e")
    return re.sub(r'-[\s\n]+([a-zàâäéèêëîïôùûüçœæ])', r'\1', text)


def _is_bullet_line(line: str) -> bool:
    s = line.strip()
    return s.startswith('•') or bool(re.match(r'^-\s', s))


# ── Text blocks ───────────────────────────────────────────────────────────────

def add_heading(doc: Document, text: str, level: int, lang: str = 'fr-FR') -> None:
    sizes  = {1: 19, 2: 16, 3: 13}
    colors = {1: Colors.ORANGE, 2: Colors.BLUE, 3: Colors.DARK_GREY}
    clean = _fix_soft_hyphens(text).replace('\n', ' ').strip()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(clean)
    run.font.name  = Fonts.HEADING
    run.font.size  = Pt(sizes.get(level, 13))
    run.font.bold  = True
    run.font.color.rgb = RGBColor(*colors.get(level, Colors.DARK_GREY))
    _set_run_lang(run, lang)


def add_paragraph(doc: Document, text: str, lang: str = 'fr-FR') -> None:
    text = _fix_soft_hyphens(text)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if not lines:
        return

    has_bullets = any(_is_bullet_line(l) for l in lines)

    if not has_bullets:
        _add_body_para(doc, ' '.join(lines), lang)
        return

    # Group lines: each bullet + continuations → one bullet chunk;
    # non-bullet runs → body chunk.
    chunks: list[tuple[bool, str]] = []
    current_bullet = None
    current_lines: list[str] = []

    for line in lines:
        if _is_bullet_line(line):
            if current_lines:
                chunks.append((bool(current_bullet), ' '.join(current_lines)))
            current_bullet = True
            current_lines = [re.sub(r'^[•\-]\s*', '', line)]
        else:
            if current_bullet is True:
                current_lines.append(line)
            else:
                if current_bullet is False:
                    current_lines.append(line)
                else:
                    current_bullet = False
                    current_lines = [line]

    if current_lines:
        chunks.append((bool(current_bullet), ' '.join(current_lines)))

    for is_bullet, chunk_text in chunks:
        if is_bullet:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(18)
            p.paragraph_format.first_line_indent = Pt(-12)
            p.paragraph_format.space_after = Pt(3)
            _add_highlighted_runs(p, '•  ' + chunk_text, 11, lang)
        else:
            _add_body_para(doc, chunk_text, lang)


def _add_body_para(doc: Document, text: str, lang: str = 'fr-FR', left_indent: int = 0) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    if left_indent:
        p.paragraph_format.left_indent = Pt(left_indent)
    _add_highlighted_runs(p, text, 11, lang)


def add_indented_paragraph(doc: Document, text: str, lang: str = 'fr-FR') -> None:
    """Body paragraph indented to sit under a preceding bullet item."""
    text = _fix_soft_hyphens(text)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if lines:
        _add_body_para(doc, ' '.join(lines), lang, left_indent=18)


# ── TOC table ─────────────────────────────────────────────────────────────────

def add_toc_table(doc: Document, entries: list[dict]) -> None:
    """Render sommaire as a 2-column table: Title | Page."""
    if not entries:
        return

    table = doc.add_table(rows=len(entries), cols=2)
    table.style = "Table Grid"

    content_twips = int(CONTENT_WIDTH * 1440 / 914400)
    page_col_w  = int(0.55 * 1440)   # ~0.55" for page numbers
    title_col_w = content_twips - page_col_w
    _set_table_col_widths(table, [title_col_w, page_col_w])

    for r_idx, entry in enumerate(entries):
        level    = entry.get('level', 'item')
        title    = entry.get('title', '')
        page_num = entry.get('page')

        t_cell = table.cell(r_idx, 0)
        p_cell = table.cell(r_idx, 1)
        t_cell.text = ""
        p_cell.text = ""

        tp = t_cell.paragraphs[0]
        pp = p_cell.paragraphs[0]

        for para in (tp, pp):
            para.paragraph_format.space_before = Pt(3)
            para.paragraph_format.space_after  = Pt(3)

        pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        if level == 'period':
            _shade_cell(t_cell, _rgb_hex(Colors.ORANGE))
            _shade_cell(p_cell, _rgb_hex(Colors.ORANGE))
            run = tp.add_run(title.upper())
            run.font.name  = Fonts.HEADING
            run.font.size  = Pt(11)
            run.font.bold  = True
            run.font.color.rgb = RGBColor(*Colors.WHITE)

        elif level == 'week':
            _shade_cell(t_cell, _rgb_hex(Colors.BLUE))
            _shade_cell(p_cell, _rgb_hex(Colors.BLUE))
            tp.paragraph_format.left_indent = Pt(14)
            run = tp.add_run(title)
            run.font.name  = Fonts.HEADING
            run.font.size  = Pt(10)
            run.font.bold  = True
            run.font.color.rgb = RGBColor(*Colors.WHITE)

        else:  # item
            tp.paragraph_format.left_indent = Pt(28)
            run = tp.add_run(title)
            run.font.name  = Fonts.BODY
            run.font.size  = Pt(10)
            run.font.color.rgb = RGBColor(*Colors.DARK_GREY)

            if page_num is not None:
                pr = pp.add_run(str(page_num))
                pr.font.name  = Fonts.BODY
                pr.font.size  = Pt(10)
                pr.font.color.rgb = RGBColor(*Colors.DARK_GREY)

    doc.add_paragraph()


# ── Tables ─────────────────────────────────────────────────────────────────────

def add_table(doc: Document, data: list[list[str]]) -> None:
    """Render an extracted table with Acacia colour coding and proper column widths."""
    if not data:
        return

    num_cols = max(len(row) for row in data)
    if num_cols == 0:
        return

    table = doc.add_table(rows=len(data), cols=num_cols)
    table.style = "Table Grid"

    content_twips = int(CONTENT_WIDTH * 1440 / 914400)
    _set_table_col_widths(table, _proportional_col_widths(data, num_cols, content_twips))

    for r_idx, row in enumerate(data):
        for c_idx in range(num_cols):
            cell_text = str(row[c_idx]) if c_idx < len(row) else ""
            cell = table.cell(r_idx, c_idx)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)

            # Render cell text — handle embedded bullets and newlines
            cell_lines = _fix_soft_hyphens(cell_text).split('\n') if cell_text else ['']
            first = True
            for line in cell_lines:
                line = line.strip()
                if not line:
                    continue
                if not first:
                    p.add_run().add_break()
                run = p.add_run(line)
                run.font.name = Fonts.BODY
                run.font.size = Pt(9)
                first = False

            fill = _cell_fill(r_idx, c_idx, cell_text, data)
            if fill:
                _shade_cell(cell, fill)
                if r_idx == 0 or _is_week_row(cell_text):
                    for run in p.runs:
                        run.font.color.rgb = RGBColor(*Colors.WHITE)
                        run.font.bold = True

    doc.add_paragraph()


def _proportional_col_widths(data: list, num_cols: int, content_twips: int) -> list[int]:
    """Distribute column widths proportional to max cell content length per column."""
    col_len = [0] * num_cols
    for row in data:
        for c, cell in enumerate(row):
            if c < num_cols:
                col_len[c] = max(col_len[c], len(str(cell or '')))

    # Clamp each column: min 4 chars, max 120 chars equivalent
    col_len = [max(4, min(120, l)) for l in col_len]
    total = sum(col_len)

    # Minimum column: 6% of content width
    min_twips = max(int(content_twips * 0.06), 600)
    raw = [int(content_twips * l / total) for l in col_len]

    # Enforce minimums, redistribute surplus
    result = [max(r, min_twips) for r in raw]
    diff = content_twips - sum(result)
    # Distribute rounding diff to the widest column
    if result:
        result[result.index(max(result))] += diff
    return result


def _set_table_col_widths(table, twips: list) -> None:
    """Set explicit column widths (in twips) via OOXML for reliable Word rendering."""
    tbl = table._tbl

    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(sum(twips)))
    tblW.set(qn('w:type'), 'dxa')

    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is None:
        tblGrid = OxmlElement('w:tblGrid')
        tbl.insert(1, tblGrid)
    else:
        for child in list(tblGrid):
            tblGrid.remove(child)
    for w in twips:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(w))
        tblGrid.append(gridCol)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(twips):
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW')
                    tcPr.append(tcW)
                tcW.set(qn('w:w'), str(twips[i]))
                tcW.set(qn('w:type'), 'dxa')


def _cell_fill(r_idx: int, c_idx: int, text: str, data: list) -> str | None:
    t = text.lower().strip()

    if _is_week_row(text):
        return _rgb_hex(Colors.WEEK_ROW)

    if r_idx == 0:
        if any(k in t for k in ("ritualis", "ritualized", "ritual")):
            return _rgb_hex(Colors.RITUALISED_ACTIVITIES)
        if any(k in t for k in ("apprentiss", "guided", "learning")):
            return _rgb_hex(Colors.GUIDED_LEARNING)
        if any(k in t for k in ("autonomie", "autonomous", "semi")):
            return _rgb_hex(Colors.AUTONOMOUS_ACTIVITIES)
        if any(k in t for k in ("objectif", "objective")):
            return _rgb_hex(Colors.OBJECTIVES_HEADER)
        if t in ("ps", "ms"):
            return _rgb_hex(Colors.PS_BADGE if t == "ps" else Colors.MS_BADGE)
        return _rgb_hex(Colors.PURPLE)

    if t == "ps":
        return _rgb_hex(Colors.PS_BADGE)
    if t == "ms":
        return _rgb_hex(Colors.MS_BADGE)

    if data and c_idx < len(data[0]):
        header = data[0][c_idx].lower()
        if any(k in header for k in ("ritualis", "ritualized")):
            return _rgb_hex(Colors.LIGHT_GREY)
        if any(k in header for k in ("apprentiss", "guided")):
            return _rgb_hex(Colors.LIGHT_GREY)

    return None


def _is_week_row(text: str) -> bool:
    t = text.strip().upper()
    return t.startswith("SEMAINE") or t.startswith("WEEK")


# ── Images ────────────────────────────────────────────────────────────────────

def add_image_placeholder(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Pt(18)
    run = p.add_run(text)
    run.font.name   = Fonts.BODY
    run.font.size   = Pt(10)
    run.font.italic = True
    run.font.color.rgb = RGBColor(*Colors.DARK_GREY)


def add_image(doc: Document, abs_path: str, caption: str = "") -> None:
    from pathlib import Path
    img = Path(abs_path)
    if not img.exists():
        doc.add_paragraph(f"[Image not found: {img.name}]")
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    try:
        run.add_picture(str(img), width=Inches(4))
    except Exception:
        p.add_run(f"[Could not embed: {img.name}]")
    if caption:
        cp = doc.add_paragraph(caption)
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in cp.runs:
            r.font.size = Pt(9)
            r.font.italic = True


# ── XML helpers ───────────────────────────────────────────────────────────────

def _rgb_hex(rgb: tuple) -> str:
    return "{:02X}{:02X}{:02X}".format(*rgb)


def _shade_cell(cell, hex_color: str) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _remove_table_borders(table) -> None:
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "none")
        tblBorders.append(border)
    tblPr.append(tblBorders)
