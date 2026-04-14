import json, sys
sys.stdout.reconfigure(encoding="utf-8")

# Check Phase 1 extraction details
p1 = json.load(open("primjer7/_output/phase1_extract_match.json", encoding="utf-8"))

# Focus on slides where Phase 2 said "Pass 1 matchao ali Excel podaci nisu pronađeni"
# or "nema podataka za verifikaciju" despite having datasets
problem_slides = [4, 6, 7, 8, 9, 17]

for ext in sorted(p1, key=lambda x: x["slide_number"]):
    sn = ext["slide_number"]
    if sn > 20:
        continue
    ds_list = ext.get("datasets", [])
    if not ds_list:
        continue
    n_matched = sum(1 for ds in ds_list if ds.get("matched_excel_id"))
    print(f"\nSlide {sn}: {len(ds_list)} datasets, {n_matched} matched")
    for ds in ds_list[:5]:
        mid = ds.get("matched_excel_id", "NONE")
        conf = ds.get("match_confidence", 0)
        title = ds.get("title", ds.get("description", "?"))[:80]
        print(f"  excel_id={mid}, conf={conf}, title={title}")
