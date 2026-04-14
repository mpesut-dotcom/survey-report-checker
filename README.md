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
python pipeline.py [excel1.xlsx excel2.xlsx ...] <putanja_do_pptx>
```

Ako se ne navedu Excel datoteke, automatski se pronalaze sve `.xlsx` datoteke u istom folderu kao PPTX.

Primjeri:
```bash
# Jedan Excel izvor
python pipeline.py "primjer1/tablice.xlsx" "primjer1/izvjestaj.pptx"

# Više Excel izvora
python pipeline.py "primjer3/file1.xlsx" "primjer3/file2.xlsx" "primjer3/report.pptx"

# Auto-detect svih .xlsx u folderu PPTX-a
python pipeline.py "primjer3/Report_Granola_SRB_26-01-008B.pptx"
```

Output: `rezultati_provjere_v3.docx` u istom folderu gdje je PPTX.

## Struktura Excel tablica

Skripta očekuje standardni format tablica iz istraživanja tržišta:
- Sheet s Total podacima (automatski pronalazi `CROSS_ALL`, `Total`, ili prvi sheet)
- Q-kodovi u koloni A, opcije u koloni B, Total % u koloni C
