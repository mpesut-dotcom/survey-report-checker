"""Prompt templates for Phase 2: Semantic matching (Flash model)."""


SYSTEM_CONTEXT = """Ti si stručnjak za market research koji povezuje podatke s prezentacijskih slajdova
s izvornim podacima iz Excel tablica. Tvoj zadatak je semantičko matchiranje — 
povezivanje na temelju ZNAČENJA pitanja, labela i konteksta, NE na temelju brojčanih vrijednosti."""


def build_matching_prompt(
    slide_extractions: list[dict],
    excel_metadata: list[dict],
    *,
    expected_count: int | None = None,
    is_retry: bool = False,
) -> str:
    """
    Build matching prompt for a single slide's datasets.

    slide_extractions: list of {slide_number, datasets[{title, chart_type, data_points[{label}], ...}]}
    excel_metadata: list of {id, question_code, question_text, labels, base_n, type}
        NOTE: NO values — only metadata to prevent circular reasoning
    expected_count: if set, tells the LLM exactly how many entries to return
    """
    # Format slides for prompt
    slides_text = ""
    for se in slide_extractions:
        sn = se["slide_number"]
        slides_text += f"\n--- SLAJD {sn} ---\n"
        for ds in se.get("datasets", []):
            di = ds.get("dataset_index", "?")
            title = ds.get("title", "?")
            chart_type = ds.get("chart_type", "?")
            labels = [dp.get("label", "?") for dp in ds.get("data_points", [])]
            unit = ds.get("unit", "?")
            base_n = ds.get("base_n")
            subset = ds.get("subset")
            series = ds.get("series_name")
            slides_text += f"  Dataset {di}: '{title}' ({chart_type})\n"
            slides_text += f"    Labele: {labels}\n"
            slides_text += f"    Jedinica: {unit}, Baza: {base_n}\n"
            if subset:
                slides_text += f"    Podskup: {subset}\n"
            if series:
                slides_text += f"    Serija: {series}\n"

    # Format Excel metadata (NO values!)
    excel_text = ""
    for em in excel_metadata:
        eid = em["id"]
        qcode = em["question_code"]
        qtext = em["question_text"]
        labels = em.get("labels", [])
        base = em.get("base_n")
        qtype = em.get("type", "?")
        excel_text += f"  [{eid}] {qcode}: {qtext}\n"
        excel_text += f"    Labele: {labels}, Tip: {qtype}, Baza: {base}\n"

    count_instruction = ""
    if expected_count is not None:
        count_instruction = f"""\n\nBROJ DATASETA: Očekujem TOČNO {expected_count} entry-ja u JSON arrayu — po jedan za svaki dataset.
Ako vratiš manje od {expected_count}, to je greška. Provjeri da nisi preskočio nijedan dataset."""

    retry_instruction = ""
    if is_retry:
        retry_instruction = f"""\n\nKRITIČNO — OVO JE PONOVLJENI POZIV!
U prethodnom pokušaju si PRESKOČIO neke datasete. Sada ti šaljem SAMO one koje si propustio.
MORAŠ vratiti odgovor za SVAKI dataset ispod — točno {expected_count or 'sve'} entry-ja.
Ne preskači nijedan, čak i ako nemaš dobar match — stavi matched_excel_id: null i confidence: 0."""

    return f"""{SYSTEM_CONTEXT}

=== PODACI SA SLAJDA (izvučeni iz prezentacije) ===
{slides_text}

=== EXCEL PITANJA (metadata — BEZ vrijednosti) ===
{excel_text}

=== ZADATAK ===
Za svaki dataset sa slajda pronađi NAJBLIŽE odgovarajuće Excel pitanje.
Matchaj na temelju:
1. SEMANTIČKE sličnosti pitanja/naslova — je li ista tema?
2. LABELA — poklapaju li se opcije odgovora?
3. BAZE — je li isti uzorak ispitanika?
4. KONTEKSTA — podskup podataka, vremensko razdoblje

NE matchaj na temelju brojčanih vrijednosti (nemaš ih za Excel)!{count_instruction}{retry_instruction}

=== FORMAT ODGOVORA ===
Vrati ISKLJUČIVO validan JSON (bez code fence-ova):
[
  {{
    "slide_number": 5,
    "dataset_index": 0,
    "matched_excel_id": "file1__Q11",
    "confidence": 0.92,
    "match_reasoning": "Kratko obrazloženje zašto je ovo match"
  }},
  {{
    "slide_number": 5,
    "dataset_index": 1,
    "matched_excel_id": null,
    "confidence": 0.0,
    "match_reasoning": "Nema odgovarajućeg Excel pitanja — izvedeni/agregirani podatak"
  }}
]

PRAVILA:
- SVAKI dataset MORA imati entry u rezultatu — ne preskači nijedan!
- Ako nema dobrog matcha, stavi matched_excel_id: null i confidence: 0
- confidence: 0.0-1.0 (1.0 = savršen match, 0.7+ = dobar, <0.5 = slab)
- AKO NISI SIGURAN BAREM 60%, STAVI matched_excel_id: null! Krivi match je GORI od nikakvog.
- Brand funneli, konverzije, agregirani KPI-ji i izračunate metrike tipično NEMAJU direktan Excel match
- Isti Excel ID može biti matchan na više datasetova (npr. različite serije istog pitanja)
- dataset_index u odgovoru MORA odgovarati dataset_index iz inputa
"""
