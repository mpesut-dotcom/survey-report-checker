"""Inspect current pipeline outputs (phase1 verifications + optional Word report)."""
import json
import re
from pathlib import Path


ROOT = Path("primjer7")
VER_PATH = ROOT / "_output" / "phase1_verifications.json"
DOC_PATH = ROOT / "provjera_Izvjestaj_HTZ_Brand Tracking 2026_AUT_v0.0_gotov.docx"


def _print_slide(v: dict):
    sn = v["slide_number"]
    st = v.get("overall_status", "?")
    n_data = len(v.get("data_issues", []))
    n_err = sum(1 for d in v.get("data_issues", []) if d.get("severity") == "error")
    n_text = len(v.get("text_issues", []))
    n_vis = len(v.get("visual_issues", []))
    n_src = len(v.get("match_sources", []))
    n_fail = len(v.get("match_failures", []))

    print(f"\n{'=' * 90}")
    print(
        f"SLIDE {sn}: {st.upper()} "
        f"(data={n_data} err={n_err}, text={n_text}, vis={n_vis}, src={n_src}, fail={n_fail})"
    )
    print(f"{'=' * 90}")
    print(f"Summary: {v.get('summary', '')}")
    print(
        "Pass1: "
        f"type={v.get('pass1_slide_type', '?')}, "
        f"candidates={v.get('pass1_total_candidates', 0)}, "
        f"datasets={v.get('pass1_total_datasets', 0)}, "
        f"confident={v.get('pass1_confident_datasets', 0)}"
    )

    if v.get("match_sources"):
        print("\n  Match sources:")
        for src in v["match_sources"]:
            conf = src.get("confidence")
            conf_str = f"{conf:.2f}" if isinstance(conf, (int, float)) else "-"
            print(
                "   "
                f"{src.get('excel_id', '?')} ({src.get('question_code', '?')}) "
                f"via {src.get('included_via', '?')}/{src.get('resolved_by', '?')} conf={conf_str}"
            )

    if v.get("match_failures"):
        print("\n  Match failures:")
        for mf in v["match_failures"]:
            conf = mf.get("confidence")
            conf_str = f"{conf:.2f}" if isinstance(conf, (int, float)) else "-"
            print(
                "   "
                f"id={mf.get('excel_id', '') or '(prazno)'} "
                f"q={mf.get('question_code', '?')} "
                f"reason={mf.get('reason', '?')} conf={conf_str}"
            )

    for di in v.get("data_issues", []):
        sv = di.get("slide_value", "")
        ev = di.get("excel_value", "")
        print(f"\n  DATA [{di.get('severity', '?').upper()}] {di.get('issue_type', '')}")
        print(f"    {di.get('detail', '')}")
        if sv or ev:
            print(f"    slide={sv} excel={ev}")

    for ti in v.get("text_issues", []):
        print(f"\n  TEXT [{ti.get('severity', '?').upper()}] {ti.get('issue_type', '')}")
        print(f"    {ti.get('detail', '')}")

    for vi in v.get("visual_issues", []):
        print(f"\n  VIS [{vi.get('severity', '?').upper()}] {vi.get('issue_type', '')}")
        print(f"    {vi.get('detail', '')}")


def _print_doc_summary(doc_path: Path):
    if not doc_path.exists():
        print(f"\nDOCX ne postoji: {doc_path}")
        return

    try:
        from docx import Document
    except ImportError:
        print("\npython-docx nije dostupan, preskačem DOCX analizu.")
        return

    doc = Document(str(doc_path))
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    heads = [p for p in paras if re.match(r"^Slajd\s+\d+\s+—", p)]

    print("\n" + "-" * 90)
    print("DOCX SAŽETAK")
    print("-" * 90)
    print(f"Putanja: {doc_path}")
    print(f"Paragrafa: {len(paras)}")
    print(f"Slajd headinga: {len(heads)}")
    if heads:
        print(f"Prvi heading: {heads[0]}")
        print(f"Zadnji heading: {heads[-1]}")


def main():
    if not VER_PATH.exists():
        print(f"Nedostaje verifikacija: {VER_PATH}")
        return

    with open(VER_PATH, encoding="utf-8") as f:
        verifications = json.load(f)

    print(f"Učitano verifikacija: {len(verifications)} iz {VER_PATH}")
    status_counts: dict[str, int] = {}
    for v in verifications:
        st = v.get("overall_status", "?")
        status_counts[st] = status_counts.get(st, 0) + 1
    print(f"Statusi: {status_counts}")

    for v in sorted(verifications, key=lambda x: x["slide_number"]):
        _print_slide(v)

    _print_doc_summary(DOC_PATH)


if __name__ == "__main__":
    main()
