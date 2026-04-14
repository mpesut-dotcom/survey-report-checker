import sys
sys.stdout.reconfigure(encoding="utf-8")
from checker.excel_parser import parse_all_excels
from checker.slide_extractor import filter_excel_candidates, _prepare_excel_metadata
from checker.prompts.verification import build_pass1_prompt
from pathlib import Path

ds_all = parse_all_excels(sorted(Path("primjer7").glob("*.xlsx")))

# Simulate Q7 slide
texts = ["Q7. Zemlje koje bi rado posjetili", "N=600", "Bugarska", "Hrvatska"]
candidates = filter_excel_candidates(texts, ds_all, max_candidates=30)
meta = _prepare_excel_metadata(candidates)

prompt, index_to_id = build_pass1_prompt(7, texts, meta)

print(f"Prompt size: {len(prompt)} chars")
print(f"Index mapping: {len(index_to_id)} entries")
print(f"First 5 mappings:")
for k, v in list(index_to_id.items())[:5]:
    print(f"  {k} → {v}")

# Show first 60 lines of prompt
lines = prompt.split("\n")
print(f"\n--- First 40 lines of prompt ---")
for line in lines[:40]:
    print(line)
