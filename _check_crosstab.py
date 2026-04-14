"""Quick diagnostic: what cross-tab data do we actually have?"""
import json

with open("primjer7/_output/phase0_excel.json", encoding="utf-8") as f:
    data = json.load(f)

# Show a sample question with cross-tab data
print("=== SAMPLE QUESTION WITH CROSS-TAB ===")
for d in data:
    cats = d.get("categories", [])
    has_breakdown = any(c.get("breakdowns") for c in cats)
    if has_breakdown:
        print(f"ID: {d['id']}")
        print(f"Q: {d['question_code']}: {d['question_text'][:80]}")
        print(f"Type: {d['type']}, Base N: {d['base_n']}")
        print(f"segment_sizes: {d.get('segment_sizes', {})}")
        # Show first category's breakdown
        for c in cats[:1]:
            print(f"  '{c['label']}' total={c['total']} breakdowns={c['breakdowns']}")
        # Collect all segment names from breakdowns
        all_segs = set()
        for c in cats:
            for k in c.get("breakdowns", {}):
                all_segs.add(k)
        print(f"  All segment names from breakdowns: {sorted(all_segs)}")
        print("---")
        break

# Summary stats
total = len(data)
with_cross = sum(1 for d in data if d.get("segment_sizes"))
with_breakdown = sum(1 for d in data if any(c.get("breakdowns") for c in d.get("categories", [])))
print(f"\nTotal questions: {total}")
print(f"With segment_sizes: {with_cross}")
print(f"With category breakdowns: {with_breakdown}")

# Show unique segment name sets
seg_sets = set()
for d in data:
    segs = d.get("segment_sizes", {})
    if segs:
        seg_sets.add(tuple(sorted(segs.keys())))
print(f"\nUnique segment name sets: {len(seg_sets)}")
for ss in seg_sets:
    print(f"  {ss}")

# Show which files contributed cross-tab data
print("\n=== CROSS-TAB BY FILE ===")
from collections import Counter
cross_by_file = Counter()
for d in data:
    if d.get("segment_sizes"):
        cross_by_file[d["file_key"]] += 1
for fk, cnt in cross_by_file.most_common():
    print(f"  {fk}: {cnt} questions with cross-tab")
