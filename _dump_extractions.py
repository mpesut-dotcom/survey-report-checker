"""Dump phase1 extractions for analysis."""
import json

with open("primjer7/_output/phase1_extractions.json", encoding="utf-8") as f:
    data = json.load(f)

for ext in data:
    sn = ext["slide_number"]
    ds = ext.get("datasets", [])
    te = ext.get("text_elements", [])
    print(f"=== SLIDE {sn}: {len(ds)} datasets, {len(te)} text_elements ===")
    for i, d in enumerate(ds):
        pts = d.get("data_points", [])
        labels = [p["label"] for p in pts[:8]]
        vals = [p.get("value") for p in pts[:8]]
        title = d.get("title", "?")[:60]
        ctype = d.get("chart_type", "?")
        unit = d.get("unit", "?")
        base = d.get("base_n", "?")
        series = d.get("series_name") or ""
        subset = d.get("subset") or ""
        print(f"  [{i}] {ctype:12s} | {title}")
        print(f"       {len(pts)}pts unit={unit} base={base} series='{series}' subset='{subset}'")
        print(f"       labels={labels}")
        print(f"       values={vals}")
    if not ds:
        print("  (no datasets)")
    print()
