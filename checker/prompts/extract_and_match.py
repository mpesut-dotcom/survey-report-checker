"""Prompt template for merged Phase 1: Extract data + match to Excel (Flash multimodal)."""


SYSTEM_CONTEXT = """Ti si stručnjak za analizu market research prezentacija.
Tvoj zadatak je:
1. Izvući strukturirane podatke sa slajda (iz SLIKE — ona je primarni izvor)
2. Povezati svaki izvučeni dataset s odgovarajućim Excel pitanjem (na temelju metadataka, NE vrijednosti)

BITNO: SLIKA slajda je primarni izvor podataka (brojevi, labele, legende).
PPTX tekst koristi samo kao dodatni kontekst — PowerPoint često ne exporta sve brojeve."""

# Max labels per Excel candidate in prompt to avoid bloating context
_MAX_LABELS_IN_PROMPT = 15


def build_extract_and_match_prompt(
    slide_number: int,
    slide_texts: list[str],
    excel_candidates: list[dict],
) -> str:
    """
    Build a merged extraction + matching prompt for one slide.

    slide_texts: raw text shapes from PPTX
    excel_candidates: pre-filtered list of {id, question_code, question_text, labels, base_n, type}
        NOTE: NO values — only metadata to prevent circular reasoning.
    """
    text_block = "\n".join(slide_texts) if slide_texts else "(nema teksta)"

    excel_block = ""
    if excel_candidates:
        for em in excel_candidates:
            eid = em["id"]
            qcode = em["question_code"]
            qtext = em["question_text"]
            labels = em.get("labels", [])
            if len(labels) > _MAX_LABELS_IN_PROMPT:
                labels = labels[:_MAX_LABELS_IN_PROMPT] + [f"... (+{len(labels) - _MAX_LABELS_IN_PROMPT} više)"]
            base = em.get("base_n")
            qtype = em.get("type", "?")
            excel_block += f"  [{eid}] {qcode}: {qtext}\n"
            excel_block += f"    Labele: {labels}, Tip: {qtype}, Baza: {base}\n"
    else:
        excel_block = "  (nema kandidata)"

    return f"""{SYSTEM_CONTEXT}

Analiziraj slajd #{slide_number}.

=== TEKST IZ PPTX SHAPE-OVA ===
{text_block}

=== EXCEL KANDIDATI (metadata — BEZ vrijednosti) ===
{excel_block}

=== ZADATAK ===
KORAK 1 — EKSTRAKCIJA: Za svaki dataset (graf, tablicu, KPI) na slajdu izvuci:
- Naslov/opis dataseta
- Kod pitanja ako je vidljiv na slajdu (npr. "Q5", "Q7a", "Q12") — ovo je KLJUČNO za matchiranje!
- Tip vizualizacije (bar_chart, pie_chart, line_chart, table, kpi_number, stacked_bar, text_only)
- Sve podatkovne točke s labelama i vrijednostima (ČITAJ S SLIKE!)
- Jedinicu mjere (%, broj, indeks, prosjek, HRK, EUR...)
- Bazu (N=...) ako je vidljiva
- Vremenski period, podskup podataka, seriju

KORAK 2 — MATCHIRANJE: Za svaki izvučeni dataset pronađi NAJBLIŽE Excel pitanje iz kandidata.
Matchaj na temelju:
1. KODA PITANJA — ako slajd prikazuje "Q5" ili "Q7a", to je najjači signal za match!
2. SEMANTIČKE sličnosti naslova/pitanja — ista tema?
3. LABELA — poklapaju li se opcije odgovora?
4. BAZE — isti uzorak ispitanika?
5. KONTEKSTA — podskup, vremensko razdoblje

NE matchaj na temelju brojčanih vrijednosti (nemaš ih za Excel)!
AKO NISI SIGURAN BAREM 60%, NE MATCHAJ — stavi matched_excel_id: null!
Krivi match je GORI od nikakvog.

=== FORMAT ODGOVORA ===
Vrati ISKLJUČIVO validan JSON (bez code fence-ova):
{{
  "slide_number": {slide_number},
  "datasets": [
    {{
      "title": "Naslov dataseta",
      "question_code": "Q5",
      "chart_type": "bar_chart|pie_chart|line_chart|table|kpi_number|stacked_bar|text_only",
      "data_points": [
        {{"label": "Opcija 1", "value": 45.2}},
        {{"label": "Opcija 2", "value": 32.1}}
      ],
      "unit": "%",
      "base_n": 450,
      "base_description": "Svi ispitanici",
      "time_period": "2024",
      "subset": null,
      "series_name": null,
      "note": null,
      "matched_excel_id": "file1__Q11",
      "confidence": 0.92,
      "match_reasoning": "Kratko obrazloženje zašto je ovo match"
    }},
    {{
      "title": "Agregirani KPI",
      "question_code": null,
      "chart_type": "kpi_number",
      "data_points": [{{"label": "Ukupno", "value": 78.5}}],
      "unit": "%",
      "base_n": 450,
      "base_description": null,
      "time_period": null,
      "subset": null,
      "series_name": null,
      "note": null,
      "matched_excel_id": null,
      "confidence": 0.0,
      "match_reasoning": "Agregirani podatak — nema direktnog Excel pitanja"
    }}
  ],
  "text_elements": [
    {{
      "type": "title|subtitle|footnote|source|annotation",
      "content": "Tekst elementa"
    }}
  ]
}}

PRAVILA ZA EKSTRAKCIJU:
- Ako na slajdu ima VIŠE grafova/tablica, svaki je zaseban dataset
- Ako graf ima više serija (npr. po godinama), svaka serija je zaseban dataset s series_name
- Brojeve zapiši TOČNO kako su na slajdu (npr. 45.2, ne zaokružuj)
- Ako je nešto nejasno ili nečitljivo, stavi null za tu vrijednost
- NE izmišljaj podatke koji nisu na slajdu
- Ako slajd nema nikakvih podataka (naslovni slajd, sadržaj...), vrati datasets: []

KRITIČNO ZA TABLICE S VIŠE ZEMALJA/SEGMENATA:
- Ako slajd sadrži tablicu gdje REDOVI su različite destinacije/zemlje/segmenti, a STUPCI su različita pitanja/metrike:
  → Svako PITANJE (stupac) je JEDAN dataset koji sadrži sve zemlje/segmente kao data_points
  → NE pravi zaseban dataset za svaku ćeliju ili za svaku zemlju!
- PRIMJER: Tablica s 5 zemalja × 6 pitanja = 6 dataseta (po pitanju), svaki s 5 data_points (po zemlji)

PRAVILA ZA MATCHIRANJE:
- SVAKI dataset MORA imati matched_excel_id i confidence polja
- Ako nema dobrog matcha → matched_excel_id: null, confidence: 0.0
- confidence: 0.0-1.0 (1.0=savršen, 0.7+=dobar, <0.6=ne matchaj)
- Brand funneli, konverzije, agregirani KPI-ji tipično NEMAJU direktan Excel match
- Isti Excel ID može biti matchan na više datasetova (npr. različite serije istog pitanja)
"""
