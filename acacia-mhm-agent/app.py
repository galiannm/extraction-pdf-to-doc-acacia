"""
Acacia MHM Agent — Streamlit UI

Run with:
    streamlit run app.py
"""
import json
import os
import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv()

from config import SOURCE_PDFS, EXTRACTED_DIR, EDITED_DIR, OUTPUT_DIR
from generators.doc_editor import _SYSTEM_PROMPT as DEFAULT_EDIT_PROMPT
from generators.docx_translator import _PROMPT as DEFAULT_TRANSLATE_PROMPT

# ── Constants ─────────────────────────────────────────────────────────────────

PROMPTS_FILE = Path(__file__).parent / "prompts.json"
FR_DOCX       = OUTPUT_DIR / "Acacia_MHM_FR.docx"
EN_DOCX       = OUTPUT_DIR / "Acacia_MHM_EN.docx"
FAITHFUL_DOCX = OUTPUT_DIR / "MHM_faithful.docx"

PDF_OPTIONS = [
    ("Intro",    SOURCE_PDFS[0]),
    ("Period 1", SOURCE_PDFS[1]),
    ("Period 2", SOURCE_PDFS[2]),
    ("Period 3", SOURCE_PDFS[3]),
    ("Period 4", SOURCE_PDFS[4]),
    ("Period 5", SOURCE_PDFS[5]),
]

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

STAGE_OPTIONS = ["All stages", "Extraction (OCR)", "Editorial cleanup", "Translation (FR→EN)"]

# ── Prompt persistence ────────────────────────────────────────────────────────

def _load_prompts() -> dict:
    if PROMPTS_FILE.exists():
        return json.loads(PROMPTS_FILE.read_text(encoding="utf-8"))
    return {
        "global":        "",
        "extract_extra": "",
        "edit":          DEFAULT_EDIT_PROMPT,
        "translate":     DEFAULT_TRANSLATE_PROMPT,
    }

def _save_prompts(p: dict) -> None:
    PROMPTS_FILE.write_text(json.dumps(p, ensure_ascii=False, indent=2), encoding="utf-8")

# ── Cache helpers ─────────────────────────────────────────────────────────────

def _extracted_json(pdf_path: Path) -> Path | None:
    p = EXTRACTED_DIR / pdf_path.stem / f"{pdf_path.stem}.json"
    return p if p.exists() else None

def _edited_json(pdf_paths: list[Path]) -> Path | None:
    doc_name = "_".join(p.stem[:20] for p in pdf_paths)
    p = EDITED_DIR / f"{doc_name}_edited.json"
    return p if p.exists() else None

def _with_global(global_p: str, stage_p: str) -> str:
    if global_p.strip():
        return f"{global_p.strip()}\n\n---\n\n{stage_p}"
    return stage_p

# ── AI prompt rewriter ────────────────────────────────────────────────────────

def _rewrite_prompt(current: str, user_request: str, stage_name: str) -> str:
    """Convert the user's request into a concise instruction and append it to the existing prompt."""
    import time
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

    # Ask the model to distill the request into one clear instruction sentence
    meta = f"""Convert this user request into a single clear instruction suitable for an AI prompt.

User request: "{user_request}"
Stage: {stage_name}

Return ONLY the instruction as 1-2 concise sentences. No explanation, no markdown, no quotes."""

    wait_seconds = [30, 60]
    for attempt, wait in enumerate([-1] + wait_seconds):
        if wait > 0:
            placeholder = st.empty()
            for remaining in range(wait, 0, -1):
                placeholder.info(f"Rate limited — retrying in {remaining}s…")
                time.sleep(1)
            placeholder.empty()
        try:
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=meta,
                config=types.GenerateContentConfig(
                    temperature=0.2,
                    thinking_config=types.ThinkingConfig(thinking_budget=0),
                ),
            )
            instruction = resp.text.strip()
            break
        except Exception as e:
            if "429" not in str(e) and "RESOURCE_EXHAUSTED" not in str(e):
                raise
            if attempt == len(wait_seconds):
                raise

    # Append to the existing prompt — never overwrite it
    if "ADDITIONAL INSTRUCTIONS:" in current:
        return current + f"\n- {instruction}"
    else:
        return current + f"\n\nADDITIONAL INSTRUCTIONS:\n- {instruction}"

# ── Session state init ────────────────────────────────────────────────────────

st.set_page_config(page_title="Acacia MHM Agent", page_icon="📚", layout="wide")

# Bigger fonts + chat input sizing via CSS
st.markdown("""
<style>
    /* Base font size bump */
    html, body, [class*="css"] { font-size: 16px; }

    /* Paragraphs, labels, captions */
    p, li, span, label, div, .stMarkdown { font-size: 16px !important; }

    /* Sidebar text */
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] span { font-size: 15px !important; }

    /* Headings */
    h1 { font-size: 2rem !important; }
    h2 { font-size: 1.5rem !important; }
    h3 { font-size: 1.25rem !important; }

    /* Text areas */
    textarea { font-size: 15px !important; line-height: 1.6 !important; }

    /* Buttons */
    .stButton > button { font-size: 16px !important; padding: 0.5rem 1.2rem !important; }

    /* Chat messages */
    [data-testid="stChatMessage"] p { font-size: 16px !important; }

    /* Download buttons */
    .stDownloadButton > button { font-size: 15px !important; }

    /* Selectbox */
    [data-testid="stSelectbox"] label,
    [data-testid="stSelectbox"] div { font-size: 16px !important; }

    /* Expander labels */
    [data-testid="stExpander"] summary p { font-size: 16px !important; font-weight: 500; }

    /* Caption / small text — slightly smaller but still readable */
    .stCaption, [data-testid="stCaptionContainer"] p { font-size: 14px !important; }
</style>
""", unsafe_allow_html=True)

if "prompts_loaded" not in st.session_state:
    saved = _load_prompts()
    st.session_state.global_prompt    = saved.get("global",        "")
    st.session_state.extract_extra    = saved.get("extract_extra", "")
    st.session_state.edit_prompt      = saved.get("edit",          DEFAULT_EDIT_PROMPT)
    st.session_state.translate_prompt = saved.get("translate",     DEFAULT_TRANSLATE_PROMPT)
    st.session_state.chat_history   = []
    st.session_state.prev_prompts   = None
    st.session_state.prompts_loaded = True


UPLOADS_DIR = OUTPUT_DIR / "uploads"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)


def _save_uploads(uploaded_files) -> list[Path]:
    """Save Streamlit uploaded file objects to disk and return their paths."""
    paths = []
    for f in uploaded_files:
        dest = UPLOADS_DIR / f.name
        dest.write_bytes(f.read())
        paths.append(dest)
    return sorted(paths)


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("Acacia MHM Agent")
    st.divider()

    st.subheader("Upload PDFs")
    uploaded_files = st.file_uploader(
        "Drop your PDF files here",
        type="pdf",
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    # Save uploads to disk and discover already-uploaded files
    if uploaded_files:
        saved_paths = _save_uploads(uploaded_files)
    else:
        saved_paths = sorted(UPLOADS_DIR.glob("*.pdf"))

    st.subheader("Documents")
    selected: list[Path] = []
    for path in saved_paths:
        cached = "✓" if _extracted_json(path) else "○"
        if st.checkbox(f"{cached}  {path.name}", key=f"sel_{path.stem}"):
            selected.append(path)

    if not saved_paths:
        st.caption("Upload PDFs above to get started.")

    st.divider()
    st.subheader("Stages")
    do_extract   = st.checkbox("1. Extract (OCR)",   value=True)
    do_edit      = st.checkbox("2. Edit (cleanup)",  value=True)
    do_translate = st.checkbox("3. Translate FR→EN", value=True)
    force_extract = st.checkbox("Force re-extract")
    force_edit    = st.checkbox("Force re-edit")

    st.divider()
    run_all = st.button("▶  Run Pipeline", type="primary", use_container_width=True, disabled=not selected)

# ── Main area ─────────────────────────────────────────────────────────────────

st.title("MHM Document Pipeline")
st.caption("Select documents in the sidebar, then describe what you want or click Run Pipeline.")

# ── Downloads (always visible if files exist) ─────────────────────────────────

if FR_DOCX.exists() or EN_DOCX.exists() or FAITHFUL_DOCX.exists():
    st.markdown("""
<div style="
    background: linear-gradient(135deg, #f0fff4, #e6ffee);
    border: 2.5px solid #38a169;
    border-radius: 16px;
    padding: 20px 28px 16px 28px;
    margin-bottom: 24px;
">
<div style="font-size:1.3rem; font-weight:700; color:#276749; margin-bottom:14px;">
    ✅ Documents ready — click to download
</div>
""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        if FAITHFUL_DOCX.exists():
            st.download_button("📄 Download faithful reproduction (.docx)", FAITHFUL_DOCX.read_bytes(),
                               "MHM_faithful.docx", DOCX_MIME, key="dl_faithful_top", use_container_width=True)
    with c2:
        if FR_DOCX.exists():
            st.download_button("📄 Download French version (.docx)", FR_DOCX.read_bytes(),
                               "Acacia_MHM_FR.docx", DOCX_MIME, key="dl_fr_top", use_container_width=True)
    with c3:
        if EN_DOCX.exists():
            st.download_button("📄 Download English version (.docx)", EN_DOCX.read_bytes(),
                               "Acacia_MHM_EN.docx", DOCX_MIME, key="dl_en_top", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.divider()

# ── Chat — primary interaction ────────────────────────────────────────────────

st.markdown("""
<div style="
    background: linear-gradient(135deg, #fffbf5, #fff8ee);
    border: 2.5px solid #F5A623;
    border-radius: 16px;
    padding: 28px 32px 20px 32px;
    margin: 8px 0 0 0;
">
<div style="font-size:1.5rem; font-weight:700; color:#d4820a; margin-bottom:6px;">
    💬 Tell the AI what you want to change
</div>
<div style="font-size:1rem; color:#555; margin-bottom:0;">
    Describe any adjustment in plain language — the AI will update the right settings for you.<br>
    <span style="color:#888;">
    e.g. <em>"The translation should sound more academic"</em> &nbsp;·&nbsp;
    <em>"Pay more attention to bold text during extraction"</em> &nbsp;·&nbsp;
    <em>"Use shorter paragraphs in the editorial cleanup"</em>
    </span>
</div>
</div>
""", unsafe_allow_html=True)

st.write("")  # spacing

col_stage, _ = st.columns([1, 3])
with col_stage:
    target = st.selectbox("Applies to", STAGE_OPTIONS, label_visibility="visible")

with st.form("chat_form", clear_on_submit=True):
    user_input_text = st.text_area(
        "Your request",
        height=100,
        label_visibility="collapsed",
        placeholder='Type your request here, e.g. "Make the translation sound more formal and natural"',
    )
    send_col, revert_col, _ = st.columns([1, 1, 4])
    with send_col:
        send_btn = st.form_submit_button("Apply change", type="primary", use_container_width=True)
    with revert_col:
        revert_btn = st.form_submit_button(
            "↩ Undo last change",
            disabled=st.session_state.prev_prompts is None,
            use_container_width=True,
        )

# Chat history
if st.session_state.chat_history:
    st.write("")
    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

if revert_btn and st.session_state.prev_prompts is not None:
    snap = st.session_state.prev_prompts
    st.session_state.global_prompt    = snap["global_prompt"]
    st.session_state.extract_extra    = snap["extract_extra"]
    st.session_state.edit_prompt      = snap["edit_prompt"]
    st.session_state.translate_prompt = snap["translate_prompt"]
    st.session_state.prev_prompts     = None
    st.session_state.chat_history.append({"role": "assistant", "content": "Last change reverted."})
    st.rerun()

if send_btn and user_input_text.strip():
    user_input = user_input_text.strip()
    # Snapshot current prompts before overwriting
    st.session_state.prev_prompts = {
        "global_prompt":    st.session_state.global_prompt,
        "extract_extra":    st.session_state.extract_extra,
        "edit_prompt":      st.session_state.edit_prompt,
        "translate_prompt": st.session_state.translate_prompt,
    }
    st.session_state.chat_history.append({"role": "user", "content": user_input})

    if target == "All stages":
        targets_to_update = [("global_prompt", "Global Instructions")]
    else:
        stage_map = {
            "Extraction (OCR)":    ("extract_extra",    "PDF Extraction"),
            "Editorial cleanup":   ("edit_prompt",      "Editorial Cleanup"),
            "Translation (FR→EN)": ("translate_prompt", "Translation"),
        }
        key, name = stage_map[target]
        targets_to_update = [(key, name)]

    updated_names = []
    try:
        for state_key, stage_name in targets_to_update:
            with st.spinner(f"Updating {stage_name} prompt…"):
                updated = _rewrite_prompt(st.session_state[state_key], user_input, stage_name)
            st.session_state[state_key] = updated
            updated_names.append(stage_name)
    except Exception as e:
        # Roll back snapshot — nothing was saved
        st.session_state.prev_prompts = None
        st.session_state.chat_history.pop()  # remove the user message we just added
        if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
            st.error("The AI is temporarily rate-limited. Please wait 30–60 seconds and try again.")
        else:
            st.error(f"Something went wrong: {e}")
        st.stop()

    label = " and ".join(updated_names)
    reply = (
        f"Done — updated the **{label}** prompt based on your request.\n\n"
        f"> _{user_input}_\n\n"
        "The change is saved. You can review or fine-tune it in the Advanced settings below."
    )
    st.session_state.chat_history.append({"role": "assistant", "content": reply})
    st.rerun()

# ── Advanced settings (collapsed) ────────────────────────────────────────────

st.divider()
with st.expander("Advanced — edit prompts directly", expanded=False):
    st.caption("You can edit the AI prompts directly here. Use the chat above for most changes.")

    st.markdown("**Global instructions** *(applied to every stage)*")
    st.text_area(
        "Global",
        key="global_prompt",
        height=90,
        label_visibility="collapsed",
        placeholder="e.g. This document is for Acacia International Pre-school, targeting teachers of children aged 3–6.",
    )

    st.markdown("---")
    st.markdown("**Stage 1 — Extraction: extra instructions** *(appended to built-in prompt)*")
    st.text_area("Extract extra", key="extract_extra", height=110, label_visibility="collapsed",
                 placeholder="e.g. Pay extra attention to table column headers.")

    st.markdown("---")
    c1, c2 = st.columns([5, 1])
    with c1:
        st.markdown("**Stage 2 — Editorial cleanup prompt**")
    with c2:
        if st.button("Reset", key="reset_edit"):
            st.session_state.edit_prompt = DEFAULT_EDIT_PROMPT
            st.rerun()
    st.text_area("Edit prompt", key="edit_prompt", height=320, label_visibility="collapsed")

    st.markdown("---")
    c3, c4 = st.columns([5, 1])
    with c3:
        st.markdown("**Stage 3 — Translation prompt**")
    with c4:
        if st.button("Reset", key="reset_translate"):
            st.session_state.translate_prompt = DEFAULT_TRANSLATE_PROMPT
            st.rerun()
    st.text_area("Translate prompt", key="translate_prompt", height=320, label_visibility="collapsed")

    # Per-stage JSON downloads
    st.markdown("---")
    st.markdown("**Intermediate outputs**")
    dl_cols = st.columns(len(selected) + 1)
    for i, path in enumerate(selected):
        jp = _extracted_json(path)
        if jp:
            dl_cols[i].download_button(
                f"{path.stem[:18]}.json",
                data=jp.read_bytes(),
                file_name=f"{path.stem}_extracted.json",
                mime="application/json",
                key=f"dl_ext_{path.stem}",
            )
    ep = _edited_json(selected) if selected else None
    if ep:
        dl_cols[-1].download_button(
            "edited.json", data=ep.read_bytes(),
            file_name="edited.json", mime="application/json", key="dl_edited",
        )

# ── Auto-save prompts ─────────────────────────────────────────────────────────

_save_prompts({
    "global":        st.session_state.global_prompt,
    "extract_extra": st.session_state.extract_extra,
    "edit":          st.session_state.edit_prompt,
    "translate":     st.session_state.translate_prompt,
})

# ── Pipeline stage runners ────────────────────────────────────────────────────

def _run_extract(pdf_paths: list[Path]) -> list[dict]:
    from extractors.pdf_extractor import extract_pdf
    extra = _with_global(st.session_state.global_prompt, st.session_state.extract_extra)
    all_pages = []
    for path in pdf_paths:
        with st.spinner(f"Extracting {path.name}…"):
            result = extract_pdf(path, EXTRACTED_DIR, extra_instructions=extra, force=force_extract)
        all_pages.extend(result["pages"])
    return all_pages

def _run_edit(pages: list[dict], pdf_paths: list[Path]) -> list[dict]:
    from generators.doc_editor import edit_pages
    doc_name = "_".join(p.stem[:20] for p in pdf_paths)
    prompt   = _with_global(st.session_state.global_prompt, st.session_state.edit_prompt)
    with st.spinner("Running editorial cleanup…"):
        return edit_pages(pages, doc_name, EDITED_DIR, system_prompt=prompt, force=force_edit)

def _run_build_faithful(pages: list[dict]) -> None:
    from generators.docx_generator import build_faithful_document
    OUTPUT_DIR.mkdir(exist_ok=True)
    with st.spinner("Building faithful reproduction…"):
        build_faithful_document(
            pages=pages,
            output_path=FAITHFUL_DOCX,
            title="Guide MHM — Reproduction fidèle",
            subtitle="Petite Section / Moyenne Section",
            lang="fr-FR",
        )

def _run_build_fr(pages: list[dict]) -> None:
    from generators.docx_generator import build_document
    OUTPUT_DIR.mkdir(exist_ok=True)
    with st.spinner("Building French document…"):
        build_document(
            pages=pages,
            output_path=FR_DOCX,
            title="Guide des séances MHM",
            subtitle="Petite Section / Moyenne Section — Acacia International Pre-school",
            lang="fr-FR",
        )

def _run_translate() -> None:
    from generators.docx_translator import translate_docx
    prompt = _with_global(st.session_state.global_prompt, st.session_state.translate_prompt)
    with st.spinner("Translating to English…"):
        translate_docx(
            fr_path=FR_DOCX,
            en_path=EN_DOCX,
            title="MHM Teachers' Guide",
            subtitle="Kindergarten 1 / Kindergarten 2 — Acacia International Pre-school",
            translation_prompt=prompt,
        )

def _load_pages_from_cache(pdf_paths: list[Path]) -> list[dict] | None:
    pages = []
    for p in pdf_paths:
        jp = _extracted_json(p)
        if not jp:
            st.error(f"No extraction found for **{p.name}**. Run the pipeline with extraction enabled first.")
            return None
        data = json.loads(jp.read_text(encoding="utf-8"))
        pages.extend(data["pages"])
    return pages

# ── Full pipeline runner ──────────────────────────────────────────────────────

if run_all and selected:
    st.divider()
    st.subheader("Running pipeline…")

    all_pages: list[dict] = []

    if do_extract:
        all_pages = _run_extract(selected)
        st.success(f"Stage 1 complete — {len(all_pages)} pages extracted")
    else:
        p = _load_pages_from_cache(selected)
        if p is None:
            st.stop()
        all_pages = p

    _run_build_faithful(all_pages)
    st.success("Stage 2 complete — faithful reproduction built")

    build_pages = all_pages

    if do_edit:
        build_pages = _run_edit(all_pages, selected)
        st.success("Stage 3 complete — editorial cleanup done")

    _run_build_fr(build_pages)
    st.success("Stage 4 complete — French document built")

    if do_translate:
        _run_translate()
        st.success("Stage 5 complete — English document translated")

    st.rerun()
