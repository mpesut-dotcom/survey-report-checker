"""Prompt templates for Phase 1: Slide data extraction (Flash model)."""


SYSTEM_CONTEXT = """Ti si stručnjak za analizu market research prezentacija.
Tvoj zadatak je izvući strukturirane podatke s jednog slajda prezentacije.
Pažljivo analiziraj sliku slajda i prateći tekst iz PPTX-a.
BITNO: Za tablice i grafove, SLIKA je primarni izvor podataka (brojevi, labele, legende).
PPTX tekst koristi samo kao dodatni kontekst — PowerPoint često ne exporta sve brojeve iz tablica u text shapes."""


def build_extraction_prompt(slide_number: int, slide_texts: list[str]) -> str:
    """
    Build the extraction prompt for a single slide.
    This will be paired with the slide image in a multimodal call.
    """
    text_block = "\n".join(slide_texts) if slide_texts else "(nema teksta)"

    return f"""{SYSTEM_CONTEXT}

Analiziraj slajd #{slide_number} i izvuci SVE podatke.

=== TEKST IZ PPTX SHAPE-OVA ===
{text_block}

=== ZADATAK ===
Za svaki dataset (graf, tablicu, KPI) na slajdu izvuci:
1. Naslov/opis dataseta
2. Tip vizualizacije (bar_chart, pie_chart, line_chart, table, kpi_number, stacked_bar, text_only)
3. Sve podatkovne točke s labelama i vrijednostima
4. Jedinicu mjere (%, broj, indeks, prosjek, HRK, EUR...)
5. Bazu (N=...) ako je vidljiva
6. Vremenski period ili godinu ako je navedena
7. Podskup podataka (npr. "Top 2 Box", "Samo korisnici branda X")

Za tekstualne elemente na slajdu (naslovi, bilješke, source napomene, legende):
- Izvuci ih posebno u text_elements

=== FORMAT ODGOVORA ===
Vrati ISKLJUČIVO validan JSON (bez code fence-ova):
{{
  "slide_number": {slide_number},
  "datasets": [
    {{
      "title": "Naslov dataseta",
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
      "note": null
    }}
  ],
  "text_elements": [
    {{
      "type": "title|subtitle|footnote|source|annotation",
      "content": "Tekst elementa"
    }}
  ]
}}

PRAVILA:
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
- PRIMJER: Tablica s 5 zemalja (Austrija, Grčka, Italija, Španjolska, Turska) × 6 pitanja 
  (Spontano poznavanje, Planiranje odmora, Poznavanje ponude, Atraktivnost, Povoljnost, Prvo spomenuta)
  = 6 dataseta (po pitanju), svaki s 5 data_points (po zemlji za istu godinu/seriju)
  → Dataset 0: "Spontano poznavanje" → data_points: [{{"label": "Austrija", "value": 82}}, {{"label": "Grčka", "value": 45}}, ...]
  → Dataset 1: "Planiranje odmora" → data_points: [{{"label": "Austrija", "value": 12}}, {{"label": "Grčka", "value": 8}}, ...]
  → itd.
- Ako tablica ima više godina/serija, svaka godina je zasebna serija unutar istog dataseta (series_name="2024", series_name="2025").
"""
