# Provjera izvještaja

Automatska provjera kvalitete PowerPoint izvještaja iz istraživanja tržišta.

Uspoređuje podatke na grafovima u PPTX prezentaciji s izvornim podacima u Excel tablicama i detektira:
- **Nepodudarnost graf ↔ Excel** — vrijednosti na grafu ne odgovaraju tablicama
- **Tekst ↔ podaci** — numerički navodi u tekstu ne odgovaraju podacima
- **Pogrešni zaključci** — tekst tvrdi nešto što podaci ne podržavaju
- **Tipfeleri i gramatika** — pravopisne i jezične greške

## Kako radi

1. **Faza 1** — Parsira XLSX tablice i PPTX slajdove
2. **Faza 2** — Programatski matchira grafove na Q-kodove (bez LLM-a)
3. **Faza 3** — LLM (Gemini 2.5 Flash) provjerava svaki slajd uz Excel podatke kao izvor istine
4. **Faza 4** — Generira Word izvještaj s nalazima

## Postavljanje

```bash
pip install -r requirements.txt
```

Kreiraj `.env` datoteku:
```
GEMINI_API_KEY=tvoj_api_kljuc
```

## Korištenje

```bash
python pipeline.py <putanja_do_excel> <putanja_do_pptx>
```

Primjer:
```bash
python pipeline.py "primjer1/tablice.xlsx" "primjer1/izvjestaj.pptx"
```

Output: `rezultati_provjere_v3.docx` u istom folderu gdje je PPTX.

## Struktura Excel tablica

Skripta očekuje standardni format tablica iz istraživanja tržišta:
- Sheet s Total podacima (automatski pronalazi `CROSS_ALL`, `Total`, ili prvi sheet)
- Q-kodovi u koloni A, opcije u koloni B, Total % u koloni C
