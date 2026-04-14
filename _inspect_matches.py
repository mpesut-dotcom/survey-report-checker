"""Quick inspection of phase outputs for debugging."""
import json
import sys

def main():
    slide = int(sys.argv[1]) if len(sys.argv) > 1 else 13
    base = "primjer7/_output"

    # Phase 1 extraction
    with open(f"{base}/phase1_extractions.json", encoding="utf-8") as f:
        extractions = json.load(f)

    print(f"=== SLIDE {slide}: PHASE 1 EXTRACTION ===")
    for ext in extractions:
        if ext["slide_number"] == slide:
            for i, ds in enumerate(ext["datasets"]):
                pts = ds.get("data_points", [])
                labels = [p["label"] for p in pts[:5]]
                values = [p["value"] for p in pts[:5]]
                print(f"  ds[{i}] {ds['chart_type']:15s} | {ds['title'][:70]}")
                print(f"         series={ds.get('series_name')} subset={ds.get('subset')}")
                print(f"         labels={labels}")
                print(f"         values={values}")
                if len(pts) > 5:
                    print(f"         ... +{len(pts)-5} more")
                print()

    # Phase 2 matches
    with open(f"{base}/phase2_matches.json", encoding="utf-8") as f:
        matches = json.load(f)

    print(f"=== SLIDE {slide}: PHASE 2 MATCHES ===")
    matched_ids = []
    for m in matches["matches"]:
        if m["slide_number"] == slide:
            eid = m.get("matched_excel_id")
            print(f"  ds[{m['dataset_index']}] -> {eid}  conf={m['confidence']}")
            print(f"    reason: {m['match_reasoning'][:120]}")
            print()
            if eid:
                matched_ids.append(eid)

    # Phase 0 Excel data for matched IDs
    if matched_ids:
        with open(f"{base}/phase0_excel.json", encoding="utf-8") as f:
            excel = json.load(f)

        excel_by_id = {d["id"]: d for d in excel}
        print(f"=== MATCHED EXCEL DATA ===")
        for eid in set(matched_ids):
            ed = excel_by_id.get(eid)
            if not ed:
                print(f"  {eid}: NOT FOUND!")
                continue
            print(f"  [{eid}] {ed['question_code']}: {ed['question_text'][:80]}")
            print(f"    base_n={ed.get('base_n')} type={ed.get('type')}")
            cats = ed.get("categories", [])
            for c in cats[:10]:
                print(f"    {c['label']:30s} total={c['total']}")
            if len(cats) > 10:
                print(f"    ... +{len(cats)-10} more")
            print()

    # Phase 3 verification
    with open(f"{base}/phase3_verifications.json", encoding="utf-8") as f:
        verifications = json.load(f)

    print(f"=== SLIDE {slide}: PHASE 3 VERIFICATION ===")
    for v in verifications:
        if v["slide_number"] == slide:
            print(f"  status: {v['overall_status']}")
            print(f"  summary: {v.get('summary', '')[:200]}")
            for di in v.get("data_issues", []):
                print(f"  DATA [{di['severity']}] {di['detail'][:120]}")
            for ti in v.get("text_issues", []):
                print(f"  TEXT [{ti['severity']}] {ti['detail'][:120]}")
            for vi in v.get("visual_issues", []):
                print(f"  VIS  [{vi['severity']}] {vi['detail'][:120]}")

if __name__ == "__main__":
    main()
