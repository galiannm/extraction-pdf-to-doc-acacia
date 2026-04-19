"""
Acacia MHM Agent — main entry point.

Usage:
    python main.py --test          # process intro PDF only
    python main.py --all           # process all 6 PDFs
    python main.py --extract-only  # stop after extraction (no editorial, no translation)
    python main.py --no-edit       # skip editorial cleanup stage
"""
import argparse
import sys
from pathlib import Path
from dotenv import load_dotenv
from rich.console import Console
from rich.rule import Rule

from config import SOURCE_PDFS, TEST_PDF, EXTRACTED_DIR, EDITED_DIR, OUTPUT_DIR
from extractors.pdf_extractor import extract_pdf, load_extracted
from generators.doc_editor import edit_pages
from generators.docx_generator import build_document
from generators.docx_translator import translate_docx

load_dotenv()
console = Console()


def run(pdfs: list[Path], extract_only: bool = False, no_edit: bool = False) -> None:
    console.print(Rule("[bold orange3]Acacia MHM Agent[/bold orange3]"))
    console.print(f"Processing {len(pdfs)} document(s)\n")

    all_fr_pages: list[dict] = []

    for pdf_path in pdfs:
        if not pdf_path.exists():
            console.print(f"[red]File not found:[/red] {pdf_path}")
            sys.exit(1)

        # ── Stage 1: Extract ────────────────────────────────────────────────
        extracted = extract_pdf(pdf_path, EXTRACTED_DIR)
        all_fr_pages.extend(extracted["pages"])

    if extract_only:
        console.print("\n[yellow]--extract-only: stopping after extraction[/yellow]")
        return

    # ── Stage 2: Editorial cleanup ─────────────────────────────────────────
    if no_edit:
        console.print("[yellow]--no-edit: skipping editorial cleanup[/yellow]")
        build_pages = all_fr_pages
        fr_suffix   = "Acacia_MHM_FR.docx"
        en_suffix   = "Acacia_MHM_EN.docx"
    else:
        console.print(Rule("Editorial cleanup"))
        doc_name = "_".join(p.stem[:20] for p in pdfs)
        build_pages = edit_pages(all_fr_pages, doc_name, EDITED_DIR)
        fr_suffix   = "Acacia_MHM_FR.docx"
        en_suffix   = "Acacia_MHM_EN.docx"

    # ── Stage 3: Build FR document ─────────────────────────────────────────
    console.print(Rule("Building French document"))
    fr_path = OUTPUT_DIR / fr_suffix
    build_document(
        pages=build_pages,
        output_path=fr_path,
        title="Guide des séances MHM",
        subtitle="Petite Section / Moyenne Section — Acacia International Pre-school",
        lang="fr-FR",
    )

    # ── Stage 4: Build EN by translating FR docx in-place ─────────────────
    console.print(Rule("Building English document"))
    en_path = OUTPUT_DIR / en_suffix
    translate_docx(
        fr_path=fr_path,
        en_path=en_path,
        title="MHM Teachers' Guide",
        subtitle="Kindergarten 1 / Kindergarten 2 — Acacia International Pre-school",
    )

    console.print(Rule("[green]Complete[/green]"))
    console.print(f"  FR → {fr_path}")
    console.print(f"  EN → {en_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Acacia MHM document pipeline")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--test",  action="store_true", help="Run on intro PDF only")
    group.add_argument("--all",   action="store_true", help="Run on all 6 PDFs")
    parser.add_argument("--extract-only", action="store_true", help="Stop after extraction")
    parser.add_argument("--no-edit",      action="store_true", help="Skip editorial cleanup")
    args = parser.parse_args()

    pdfs = [TEST_PDF] if args.test else SOURCE_PDFS
    run(pdfs, extract_only=args.extract_only, no_edit=args.no_edit)


if __name__ == "__main__":
    main()
