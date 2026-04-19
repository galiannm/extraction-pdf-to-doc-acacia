"""
Acacia MHM Agent — main entry point.

Usage:
    python main.py --test          # process intro PDF only
    python main.py --all           # process all 6 PDFs
    python main.py --extract-only  # extract + build FR doc, skip translation
"""
import argparse
import sys
from pathlib import Path
from dotenv import load_dotenv
from rich.console import Console
from rich.rule import Rule

from config import SOURCE_PDFS, TEST_PDF, EXTRACTED_DIR, OUTPUT_DIR
from extractors.pdf_extractor import extract_pdf, load_extracted
from generators.docx_generator import build_document
from generators.docx_translator import translate_docx

load_dotenv()
console = Console()


def run(pdfs: list[Path], extract_only: bool = False) -> None:
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

    # ── Stage 2: Build FR document ─────────────────────────────────────────
    console.print(Rule("Building French document"))
    fr_path = OUTPUT_DIR / "Acacia_MHM_FR.docx"
    build_document(
        pages=all_fr_pages,
        output_path=fr_path,
        title="Guide des séances MHM",
        subtitle="Petite Section / Moyenne Section — Acacia International Pre-school",
        lang="fr-FR",
    )

    if extract_only:
        console.print("\n[yellow]--extract-only: skipping EN document[/yellow]")
        return

    # ── Stage 3: Build EN by translating FR docx in-place ─────────────────
    console.print(Rule("Building English document"))
    en_path = OUTPUT_DIR / "Acacia_MHM_EN.docx"
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
    parser.add_argument(
        "--extract-only",
        action="store_true",
        help="Extract + build FR doc only (no Gemini translation)",
    )
    args = parser.parse_args()

    pdfs = [TEST_PDF] if args.test else SOURCE_PDFS
    run(pdfs, extract_only=args.extract_only)


if __name__ == "__main__":
    main()
