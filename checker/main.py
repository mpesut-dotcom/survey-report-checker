"""
Main orchestrator — 2-phase report QC pipeline.

Phase 0: Parse Excel tables into structured datasets.
Phase 1: Two-pass Pro verification (match + verify) for all slides.
Phase 2: Generate report.

Usage:
    python -m checker.main <pptx_path> [excel1.xlsx excel2.xlsx ...]
    python -m checker.main <pptx_path> --phase 0-1 --slides 5-20
    python -m checker.main <pptx_path> --dry-run
"""
import argparse
import json
import sys
import time
from pathlib import Path

from checker.config import BASE_DIR
from checker.utils.gemini_client import GeminiClient
from checker.utils.pptx_utils import extract_slide_texts
from checker.utils.image_utils import convert_pptx_to_images
from checker.utils.json_utils import save_json, load_json
from checker.excel_parser import parse_all_excels
from checker.verifier import verify_all_slides
from checker.report_generator import generate_report

sys.stdout.reconfigure(encoding="utf-8")


def main():
    args = parse_args()

    pptx_path = Path(args.pptx).resolve()
    work_dir = pptx_path.parent
    output_dir = work_dir / "_output"
    output_dir.mkdir(exist_ok=True)
    cache_dir = work_dir / "_cache"
    cache_dir.mkdir(exist_ok=True)

    # Resolve Excel files
    if args.excel:
        excel_paths = [Path(e).resolve() for e in args.excel]
    else:
        excel_paths = sorted(work_dir.glob("*.xlsx"))
        if not excel_paths:
            print(f"Nema .xlsx datoteka u {work_dir}")
            sys.exit(1)

    # Parse phase range (0-2)
    phase_start, phase_end = _parse_phase_range(args.phase)

    # Parse slide filter
    slide_filter = _parse_slide_filter(args.slides)

    print(f"  PPTX: {pptx_path.name}")
    print(f"  Excel: {', '.join(p.name for p in excel_paths)}")
    print(f"  Faze: {phase_start}-{phase_end}")
    if slide_filter:
        print(f"  Slajdovi: {slide_filter}")
    print()

    client = GeminiClient()
    t0 = time.time()

    # ── Phase 0: Excel parsing ──
    excel_datasets = None
    excel_json_path = output_dir / "phase0_excel.json"

    if phase_start <= 0:
        print("═══ FAZA 0: Parsiranje Excel tablica ═══")
        from checker.models.excel_dataset import ExcelDataset
        excel_datasets = parse_all_excels(excel_paths)
        save_json([ds.model_dump() for ds in excel_datasets], excel_json_path)
        print(f"  -> {len(excel_datasets)} pitanja ukupno")
        print(f"  -> Spremljeno: {excel_json_path.name}")
        print()

    if phase_end < 1:
        _finish(t0, client)
        return

    # Load Excel data if starting from later phase
    if excel_datasets is None:
        raw = load_json(excel_json_path)
        if raw is None:
            print(f"GREŠKA: {excel_json_path.name} ne postoji. Pokreni fazu 0 prvo.")
            sys.exit(1)
        from checker.models.excel_dataset import ExcelDataset
        excel_datasets = [ExcelDataset(**d) for d in raw]
        print(f"  Učitano {len(excel_datasets)} pitanja iz {excel_json_path.name}")

    # Prepare slide texts and images
    slide_texts = extract_slide_texts(pptx_path)
    slide_images = convert_pptx_to_images(pptx_path, cache_dir / "slides")
    print(f"  -> {len(slide_texts)} slajdova, {len(slide_images)} slika")

    # ── Phase 1: Two-pass Pro verification ──
    verifications = None
    verification_json_path = output_dir / "phase1_verifications.json"

    if phase_start <= 1 and phase_end >= 1:
        print("═══ FAZA 1: Verifikacija (Pro 2-pass) ═══")
        print(f"  Excel dataseta: {len(excel_datasets)}")

        if args.dry_run:
            print("  [DRY RUN] Preskačem LLM pozive")
        else:
            verifications = verify_all_slides(
                excel_datasets,
                slide_texts, slide_images, client,
                slide_filter=slide_filter, verbose=args.verbose,
            )
            save_json(
                [v.model_dump() for v in verifications],
                verification_json_path,
            )
            errors = sum(1 for v in verifications if v.overall_status == "error")
            warnings = sum(1 for v in verifications if v.overall_status == "warning")
            print(f"  -> {len(verifications)} slajdova verificirano")
            print(f"  -> {errors} grešaka, {warnings} upozorenja")
            print(f"  -> Spremljeno: {verification_json_path.name}")
        print()

    if phase_end < 2:
        _finish(t0, client)
        return

    # Load verifications if needed
    if verifications is None:
        raw = load_json(verification_json_path)
        if raw is None:
            print(f"GREŠKA: {verification_json_path.name} ne postoji. Pokreni fazu 1 prvo.")
            sys.exit(1)
        from checker.models.verification import SlideVerification
        verifications = [SlideVerification(**d) for d in raw]
        print(f"  Učitano {len(verifications)} verifikacija iz {verification_json_path.name}")

    # ── Phase 2: Report generation ──
    print("═══ FAZA 2: Generiranje izvještaja ═══")
    report_name = f"provjera_{pptx_path.stem}.docx"
    report_path = work_dir / report_name

    generate_report(
        verifications=verifications,
        output_path=report_path,
        pptx_name=pptx_path.name,
        cost_info=client.estimated_cost(),
    )
    print(f"  -> Izvještaj: {report_path.name}")

    _finish(t0, client)


def _finish(t0: float, client: GeminiClient):
    elapsed = time.time() - t0
    print(f"\n  Ukupno vrijeme: {elapsed:.1f}s")
    client.print_cost_summary()


def _parse_phase_range(phase_str: str | None) -> tuple[int, int]:
    """Parse phase range string like '0-2', '1', '0-1'."""
    if not phase_str:
        return 0, 2
    if "-" in phase_str:
        parts = phase_str.split("-")
        return int(parts[0]), int(parts[1])
    n = int(phase_str)
    return n, n


def _parse_slide_filter(slides_str: str | None) -> list[int] | None:
    """Parse slide filter string like '5-20', '1,3,5', '10'."""
    if not slides_str:
        return None
    result = []
    for part in slides_str.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            result.extend(range(int(a), int(b) + 1))
        else:
            result.append(int(part))
    return sorted(set(result))


def parse_args():
    parser = argparse.ArgumentParser(
        description="Pipeline v3 — QC provjera market research prezentacija",
    )
    parser.add_argument("pptx", help="Putanja do PPTX prezentacije")
    parser.add_argument("excel", nargs="*",
                        help="Excel datoteke (ako se ne navedu, auto-detect u PPTX folderu)")
    parser.add_argument("--phase", "-p", default=None,
                        help="Raspon faza: 0-2, 1, 0-1 (default: 0-2)")
    parser.add_argument("--slides", "-s", default=None,
                        help="Filter slajdova: 5-20, 1,3,5, 10")
    parser.add_argument("--dry-run", "-d", action="store_true",
                        help="Preskači LLM pozive")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Detaljniji ispis")
    return parser.parse_args()


if __name__ == "__main__":
    main()
