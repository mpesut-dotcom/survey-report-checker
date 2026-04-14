"""
Rich diagnostics for pipeline run — evaluates quality of each phase.
Usage: python _diagnostics.py [primjer_folder]
"""
import json
import sys
from collections import Counter
from pathlib import Path


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "primjer7"
    out = Path(folder) / "_output"

    if not out.exists():
        print(f"GREŠKA: {out} ne postoji"); return

    excel = load(out / "phase0_excel.json")
    extractions = load(out / "phase1_extract_match.json")
    verifications = load(out / "phase2_verifications.json")

    excel_by_id = {d["id"]: d for d in excel}

    # Build flat match list from extractions (for backward-compatible logic)
    matches = []
    for ext in extractions:
        sn = ext["slide_number"]
        for di, ds in enumerate(ext.get("datasets", [])):
            matches.append({
                "slide_number": sn,
                "dataset_index": di,
                "matched_excel_id": ds.get("matched_excel_id"),
                "confidence": ds.get("confidence", 0.0),
                "match_reasoning": ds.get("match_reasoning", ""),
            })

    print("=" * 90)
    print(f"  DIJAGNOSTIKA: {folder}")
    print("=" * 90)

    # ═══════════════════════════════════════════════════════════════
    # PHASE 0: Excel coverage
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  FAZA 0: Excel parsing")
    print(f"{'─'*90}")
    print(f"  Ukupno Excel pitanja: {len(excel)}")

    by_file = Counter(d["file_key"] for d in excel)
    for fk, cnt in by_file.most_common():
        print(f"    {fk}: {cnt} pitanja")

    by_type = Counter(d["type"] for d in excel)
    print(f"  Po tipu: {dict(by_type)}")

    has_cross = sum(1 for d in excel if any(c.get("breakdowns") for c in d.get("categories", [])))
    has_mean = sum(1 for d in excel if (d.get("derived_metrics") or {}).get("mean") is not None)
    print(f"  S cross-tab breakdownom: {has_cross}")
    print(f"  S MEAN vrijednošću: {has_mean}")

    # ═══════════════════════════════════════════════════════════════
    # PHASE 1: Extraction quality
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  FAZA 1: Ekstrakcija + Matchiranje")
    print(f"{'─'*90}")

    total_ds = 0
    empty_slides = 0
    chart_types = Counter()
    ds_per_slide = []

    for ext in sorted(extractions, key=lambda e: e["slide_number"]):
        sn = ext["slide_number"]
        ds = ext.get("datasets", [])
        n = len(ds)
        total_ds += n
        if n == 0:
            empty_slides += 1

        ds_per_slide.append((sn, n))
        for d in ds:
            chart_types[d.get("chart_type", "?")] += 1

    print(f"  Slajdova: {len(extractions)}, praznih: {empty_slides}")
    print(f"  Ukupno dataseta: {total_ds}")
    print(f"  Po tipu grafa: {dict(chart_types)}")
    print(f"\n  Dataseti po slajdu:")
    for sn, n in ds_per_slide:
        bar = "█" * n
        flag = " ⚠ PUNO" if n > 15 else ""
        print(f"    Slide {sn:2d}: {n:3d} {bar}{flag}")

    # Check for datasets with 0 data points
    empty_dp = 0
    for ext in extractions:
        for ds in ext.get("datasets", []):
            if not ds.get("data_points"):
                empty_dp += 1
    if empty_dp:
        print(f"\n  ⚠ Dataseta bez data_points (text_only): {empty_dp}")

    # ═══════════════════════════════════════════════════════════════
    # PHASE 2: Matching quality
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  FAZA 1 — Matching detalji")
    print(f"{'─'*90}")

    total_m = len(matches)
    matched = [m for m in matches if m.get("matched_excel_id")]
    unmatched = [m for m in matches if not m.get("matched_excel_id")]

    print(f"  Ukupno dataseta za matching: {total_m}")
    print(f"  Matchano: {len(matched)} ({100*len(matched)/max(total_m,1):.0f}%)")
    print(f"  Nematchano: {len(unmatched)} ({100*len(unmatched)/max(total_m,1):.0f}%)")

    # Confidence distribution
    confs = [m["confidence"] for m in matched]
    if confs:
        buckets = {"0.9-1.0": 0, "0.8-0.9": 0, "0.7-0.8": 0, "0.60-0.7": 0}
        for c in confs:
            if c >= 0.9: buckets["0.9-1.0"] += 1
            elif c >= 0.8: buckets["0.8-0.9"] += 1
            elif c >= 0.7: buckets["0.7-0.8"] += 1
            else: buckets["0.60-0.7"] += 1
        print(f"\n  Confidence distribucija (matchani):")
        for bk, cnt in buckets.items():
            bar = "█" * cnt
            print(f"    {bk}: {cnt:3d} {bar}")

    # Which Excel questions got matched
    excel_match_count = Counter(m["matched_excel_id"] for m in matched)
    print(f"\n  Excel pitanja korištena u matchu: {len(excel_match_count)}/{len(excel)}")
    print(f"  Top 5 najčešće matchana Excel pitanja:")
    for eid, cnt in excel_match_count.most_common(5):
        ed = excel_by_id.get(eid, {})
        qt = ed.get("question_text", "?")[:60]
        print(f"    {cnt}× {eid[:50]} → {qt}")

    # Unmatched reasons
    print(f"\n  Razlozi za nematchane datasete (sample):")
    reason_keywords = Counter()
    for m in unmatched:
        r = m.get("match_reasoning", "")
        if "agregirani" in r.lower() or "izvedeni" in r.lower() or "izračunat" in r.lower():
            reason_keywords["Agregirani/izvedeni podatak"] += 1
        elif "konverzij" in r.lower():
            reason_keywords["Konverzija (funnel)"] += 1
        elif "not returned" in r.lower():
            reason_keywords["LLM nije vratio odgovor"] += 1
        elif "ukupno" in r.lower() or "zbroj" in r.lower():
            reason_keywords["Ukupno/zbroj (nema u Excelu)"] += 1
        elif "ostali" in r.lower() or "spontan" in r.lower():
            reason_keywords["Ostali spontano (nema u Excelu)"] += 1
        elif "nedostaje" in r.lower() or "ne postoji" in r.lower():
            reason_keywords["Nema u Excelu"] += 1
        elif "metod" in r.lower() or "opis" in r.lower():
            reason_keywords["Metodološki/opisni"] += 1
        else:
            reason_keywords["Ostalo"] += 1
    for rk, cnt in reason_keywords.most_common():
        print(f"    {cnt:3d}× {rk}")

    # ═══════════════════════════════════════════════════════════════
    # PHASE 2 → 3 BRIDGE: Base mismatch detection
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  ANALIZA: Provjera baza (slide N vs Excel N)")
    print(f"{'─'*90}")

    ext_by_slide = {e["slide_number"]: e for e in extractions}
    base_ok = 0
    base_mismatch = 0
    base_unknown = 0
    mismatches_detail = []

    for m in matched:
        sn = m["slide_number"]
        di = m["dataset_index"]
        eid = m["matched_excel_id"]
        ed = excel_by_id.get(eid, {})
        excel_n = ed.get("base_n")

        ext = ext_by_slide.get(sn, {})
        ds_list = ext.get("datasets", [])
        if di < len(ds_list):
            slide_ds = ds_list[di]
            slide_n = slide_ds.get("base_n")
        else:
            slide_n = None

        if slide_n is None or excel_n is None:
            base_unknown += 1
        elif abs(slide_n - excel_n) <= 5:
            base_ok += 1
        else:
            base_mismatch += 1
            mismatches_detail.append({
                "slide": sn, "ds": di, "excel_id": eid,
                "slide_n": slide_n, "excel_n": excel_n,
                "subset": slide_ds.get("subset", ""),
            })

    print(f"  Baza OK (±5): {base_ok}")
    print(f"  Baza MISMATCH: {base_mismatch}")
    print(f"  Baza nepoznata: {base_unknown}")

    if mismatches_detail:
        print(f"\n  ⚠ BASE MISMATCH detalji (verifikacija ovih matcheva je NEPOUZDANA):")
        for mm in mismatches_detail:
            subset = f" subset='{mm['subset']}'" if mm["subset"] else ""
            print(f"    Slide {mm['slide']} ds[{mm['ds']}]: slide_N={mm['slide_n']} vs excel_N={mm['excel_n']}"
                  f"  ({mm['excel_id'][:50]}){subset}")

    # ═══════════════════════════════════════════════════════════════
    # PHASE 3: Verification quality
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  FAZA 2: Verifikacija")
    print(f"{'─'*90}")

    status_count = Counter(v["overall_status"] for v in verifications)
    print(f"  Statusni: {dict(status_count)}")
    print(f"  Ukupno verificiranih slajdova: {len(verifications)}")

    # Count issues by type + severity
    all_data_issues = []
    all_text_issues = []
    all_vis_issues = []
    for v in verifications:
        all_data_issues.extend(v.get("data_issues", []))
        all_text_issues.extend(v.get("text_issues", []))
        all_vis_issues.extend(v.get("visual_issues", []))

    print(f"\n  Data issues: {len(all_data_issues)}")
    di_sev = Counter(d["severity"] for d in all_data_issues)
    di_type = Counter(d.get("issue_type", "?") for d in all_data_issues)
    print(f"    Po severity: {dict(di_sev)}")
    print(f"    Po tipu: {dict(di_type)}")

    print(f"\n  Text issues: {len(all_text_issues)}")
    ti_sev = Counter(t["severity"] for t in all_text_issues)
    ti_type = Counter(t.get("issue_type", "?") for t in all_text_issues)
    print(f"    Po severity: {dict(ti_sev)}")
    print(f"    Po tipu: {dict(ti_type)}")

    print(f"\n  Visual issues: {len(all_vis_issues)}")
    vi_sev = Counter(v["severity"] for v in all_vis_issues)
    vi_type = Counter(v.get("issue_type", "?") for v in all_vis_issues)
    print(f"    Po severity: {dict(vi_sev)}")
    print(f"    Po tipu: {dict(vi_type)}")

    # ═══════════════════════════════════════════════════════════════
    # SUSPECT FINDINGS: potential false positives
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  SUSPECT: Potencijalni lažni pozitivi")
    print(f"{'─'*90}")

    suspect_count = 0

    # Check for data errors on base-mismatched slides
    mismatch_slides = {mm["slide"] for mm in mismatches_detail}
    for v in verifications:
        sn = v["slide_number"]
        for di in v.get("data_issues", []):
            if di["severity"] == "error" and sn in mismatch_slides:
                suspect_count += 1
                print(f"  ⚠ Slide {sn} DATA ERROR ali ima base mismatch → vjerojatno FALSE POSITIVE")
                print(f"    {di['detail'][:120]}")

    # Check for date/year complaints
    for v in verifications:
        sn = v["slide_number"]
        for ti in v.get("text_issues", []):
            detail_l = ti.get("detail", "").lower()
            if any(w in detail_l for w in ["budućnost", "tipfeler", "tipkarska", "future"]) and \
               any(w in detail_l for w in ["2025", "2026", "2027", "godin"]):
                suspect_count += 1
                print(f"  ⚠ Slide {sn} TEXT: godina označena kao greška → FALSE POSITIVE (prompt fix)")
                print(f"    {ti['detail'][:120]}")

    if suspect_count == 0:
        print("  ✓ Nema očitih lažnih pozitiva.")

    # ═══════════════════════════════════════════════════════════════
    # PER-SLIDE CHAIN VIEW
    # ═══════════════════════════════════════════════════════════════
    print(f"\n{'─'*90}")
    print("  LANAC PO SLAJDU: extraction → matching → verification")
    print(f"{'─'*90}")

    verif_by_slide = {v["slide_number"]: v for v in verifications}
    matches_by_slide = {}
    for m in matches:
        matches_by_slide.setdefault(m["slide_number"], []).append(m)

    for ext in sorted(extractions, key=lambda e: e["slide_number"]):
        sn = ext["slide_number"]
        n_ds = len(ext.get("datasets", []))
        slide_m = matches_by_slide.get(sn, [])
        n_matched = sum(1 for m in slide_m if m.get("matched_excel_id"))
        v = verif_by_slide.get(sn)
        v_status = v["overall_status"] if v else "N/A"
        n_errs = sum(1 for d in (v or {}).get("data_issues", []) if d["severity"] == "error")
        n_warns = sum(1 for d in (v or {}).get("data_issues", []) if d["severity"] == "warning")
        n_text = len((v or {}).get("text_issues", []))
        n_vis = len((v or {}).get("visual_issues", []))

        in_mismatch = "⚠BASE" if sn in mismatch_slides else ""

        status_icon = {"ok": "✓", "warning": "⚠", "error": "✗", "info": "ℹ"}.get(v_status, "?")

        print(f"  S{sn:2d}: {n_ds:2d} ds → {n_matched:2d} matched → "
              f"{status_icon} {v_status:8s} "
              f"(err={n_errs} warn={n_warns} txt={n_text} vis={n_vis}) {in_mismatch}")

    print(f"\n{'='*90}")
    print("  KRAJ DIJAGNOSTIKE")
    print(f"{'='*90}")


if __name__ == "__main__":
    main()
