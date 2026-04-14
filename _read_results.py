import json, sys
sys.stdout.reconfigure(encoding="utf-8")

d = json.load(open("primjer7/_output/phase2_verifications.json", encoding="utf-8"))
print(f"Total verifications: {len(d)}\n")
for v in sorted(d, key=lambda x: x["slide_number"]):
    sn = v["slide_number"]
    status = v["overall_status"]
    n_data = len(v.get("data_issues", []))
    n_text = len(v.get("text_issues", []))
    n_vis = len(v.get("visual_issues", []))
    summary = v.get("summary", "")[:150]
    print(f"Slide {sn:2d}: [{status:7s}] data={n_data} text={n_text} vis={n_vis}")
    print(f"          {summary}")
    if n_data:
        for di in v["data_issues"][:3]:
            sev = di.get("severity", "?")
            detail = di.get("detail", "")[:120]
            print(f"          DATA [{sev}]: {detail}")
    if n_text:
        for ti in v["text_issues"][:2]:
            print(f"          TEXT [{ti.get('severity','?')}]: {ti.get('detail','')[:120]}")
    print()
