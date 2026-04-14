import json, sys
sys.stdout.reconfigure(encoding="utf-8")

# Now look at Phase 2 in detail - what did Pass 1 and Pass 2 actually do?
p2 = json.load(open("primjer7/_output/phase2_verifications.json", encoding="utf-8"))

# Count key patterns
total = len(p2)
no_data = sum(1 for v in p2 if "nema podataka" in v.get("summary", ""))
pass1_no_excel = sum(1 for v in p2 if "nisu pronađeni" in v.get("summary", ""))
pass1_no_match = sum(1 for v in p2 if "nijedan Excel" in v.get("summary", ""))
actual_verified = sum(1 for v in p2 if v.get("data_issues") or 
                     (v.get("overall_status") in ("ok", "warning", "error") and 
                      "nema podataka" not in v.get("summary", "") and
                      "nisu pronađeni" not in v.get("summary", "") and
                      "nijedan Excel" not in v.get("summary", "")))

print(f"Total slides: {total}")
print(f"Pass 1 said no data (title/separator): {no_data}")
print(f"Pass 1 matched but Excel not found: {pass1_no_excel}")
print(f"Pass 1 no confident match: {pass1_no_match}")
print(f"Actually verified with data: {actual_verified}")
print()

# Focus on "Pass 1 matchao ali Excel podaci nisu pronađeni" slides
for v in sorted(p2, key=lambda x: x["slide_number"]):
    if "nisu pronađeni" in v.get("summary", ""):
        print(f"Slide {v['slide_number']}: {v['summary']}")
