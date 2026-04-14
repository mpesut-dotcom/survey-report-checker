"""Prompt templates for Phase 2: Two-pass verification (Pro model).

Pass 1 — Smart Matcher: Pro sees slide image + lightweight metadata → identifies
    which Excel questions are shown, which segments/banners are used.
Pass 2 — Targeted Verify: Pro sees slide image + exact data for matched questions
    with only relevant segment columns → number-by-number verification.
"""

# ──────────────────────────────────────────────────────────────────────
# PASS 1: Smart matching — identify what's on the slide
# ──────────────────────────────────────────────────────────────────────

_MAX_LABELS_IN_METADATA = 12
_MAX_BANNERS_IN_METADATA = 20


def _banner_signature(banners: dict[str, list[str]]) -> tuple:
    """Create a hashable signature from a banners dict for dedup."""
    return tuple(sorted((k, tuple(v)) for k, v in banners.items()))


def _format_banner_set(banners: dict[str, list[str]]) -> str:
    """Format a banner set for the reference table."""
    lines = []
    banner_items = list(banners.items())
    shown = banner_items[:_MAX_BANNERS_IN_METADATA]
    for bname, segs in shown:
        seg_str = ", ".join(segs)
        lines.append(f"      {bname}: [{seg_str}]")
    if len(banner_items) > _MAX_BANNERS_IN_METADATA:
        lines.append(f"      (+{len(banner_items) - _MAX_BANNERS_IN_METADATA} više bannera)")
    return "\n".join(lines)


def build_pass1_prompt(
    slide_number: int,
    slide_texts: list[str],
    excel_metadata: list[dict],
) -> tuple[str, dict[str, str]]:
    """
    Build Pass 1 prompt: Pro sees image + metadata (no values) → identifies matches.

    Returns (prompt_text, index_to_id) where index_to_id maps "C1"→real_excel_id.

    excel_metadata: list of {id, question_code, question_text, labels, base_n, type, banners}
    banners: {banner_name: [segment1, segment2, ...]}
    """
    # Build numeric index → real ID mapping
    index_to_id: dict[str, str] = {}
    for i, em in enumerate(excel_metadata, 1):
        index_to_id[f"C{i}"] = em["id"]
    text_block = "\n".join(slide_texts) if slide_texts else "(nema teksta)"

    # ── Deduplicate banner sets: assign each unique set a letter ID ──
    sig_to_id: dict[tuple, str] = {}
    sig_to_banners: dict[tuple, dict] = {}
    for em in excel_metadata:
        banners = em.get("banners", {})
        if not banners:
            continue
        sig = _banner_signature(banners)
        if sig not in sig_to_id:
            sig_to_id[sig] = chr(ord("A") + len(sig_to_id))
            sig_to_banners[sig] = banners

    # Build reference table (only if dedup saves space: ≥2 candidates share a set)
    banner_ref_block = ""
    sig_counts: dict[tuple, int] = {}
    for em in excel_metadata:
        banners = em.get("banners", {})
        if banners:
            sig = _banner_signature(banners)
            sig_counts[sig] = sig_counts.get(sig, 0) + 1
    shared_sigs = {sig for sig, cnt in sig_counts.items() if cnt >= 2}

    if shared_sigs:
        banner_ref_block = "=== REFERENTNA TABLICA KRIŽANJA ===\n"
        banner_ref_block += "(Banneri koji se ponavljaju kod više kandidata — referencirani dolje po oznaci)\n\n"
        for sig in sorted(shared_sigs, key=lambda s: sig_to_id[s]):
            label = sig_to_id[sig]
            banner_ref_block += f"  Set {label} ({len(sig_to_banners[sig])} bannera):\n"
            banner_ref_block += _format_banner_set(sig_to_banners[sig]) + "\n\n"

    # ── Format metadata compactly ──
    meta_block = ""
    for i, em in enumerate(excel_metadata, 1):
        idx = f"C{i}"  # Short numeric index
        qcode = em.get("question_code", "?")
        qtext = em.get("question_text", "?")
        labels = em.get("labels", [])
        base_n = em.get("base_n")
        qtype = em.get("type", "?")
        banners = em.get("banners", {})

        # Truncate labels for compactness
        if len(labels) > _MAX_LABELS_IN_METADATA:
            labels_str = ", ".join(labels[:_MAX_LABELS_IN_METADATA]) + f" (+{len(labels) - _MAX_LABELS_IN_METADATA} više)"
        else:
            labels_str = ", ".join(labels) if labels else "(nema)"

        meta_block += f"  [{idx}] {qcode}: {qtext}\n"
        meta_block += f"    Tip: {qtype}, Baza: N={base_n}, Odgovori: [{labels_str}]\n"

        # Show banners: reference shared set or inline unique ones
        if banners:
            sig = _banner_signature(banners)
            if sig in shared_sigs:
                meta_block += f"    Križanja: → Set {sig_to_id[sig]}\n"
            else:
                banner_items = list(banners.items())
                shown = banner_items[:_MAX_BANNERS_IN_METADATA]
                meta_block += f"    Križanja ({len(banner_items)} bannera):\n"
                for bname, segs in shown:
                    seg_str = ", ".join(segs)
                    meta_block += f"      {bname}: [{seg_str}]\n"
                if len(banner_items) > _MAX_BANNERS_IN_METADATA:
                    meta_block += f"      (+{len(banner_items) - _MAX_BANNERS_IN_METADATA} više bannera)\n"
        else:
            meta_block += "    Križanja: (nema)\n"

    prompt = f"""Ti si stručnjak za analizu market research prezentacija.
Tvoj zadatak je pogledati sliku slajda i utvrditi KOJI Excel podaci su korišteni za izradu tog slajda.
Kandidati su označeni kratkim oznakama [C1], [C2], ... — koristi TE oznake u odgovoru.

=== SLAJD {slide_number} ===

Tekst iz PPTX:
{text_block}

{banner_ref_block}=== EXCEL KANDIDATI (metadata — BEZ vrijednosti) ===
{meta_block}

=== ZADATAK ===
Pogledaj sliku slajda i za svaki graf/tablicu/KPI na njemu utvrdi:
1. Koje Excel pitanje je izvor podataka (row question)?
2. Prikazuje li slajd TOTAL kolonu ili specifične SEGMENTE/KRIŽANJA?
3. Ako prikazuje križanje — koji BANNER je korišten i koji segmenti su prikazani?
   (Kopiraj TOČNA imena bannera i segmenata iz liste križanja iznad!)
4. Kolika je baza (N=) na slajdu?

RAZUMIJEVANJE KRIŽANJA (CROSS-TABOVA):
- Cross-tab = pitanje križano po nekom banneru (obilježju poput spol, dob, segment, destinacija...)
- Svaki BANNER ima svoje segmente: npr. banner "SPOL (Q1)" ima segmente [Muškarac, Žena]
- Na slajdu "Brand awareness po dobnim skupinama" →
    row_question je pitanje o awareness-u, banner je "DOB (Q2)", segmenti su [18-29, 30-39, ...]
- Stupci tablice (npr. "Muškarac", "Žena") odgovaraju segmentima unutar jednog bannera
- Slajd može prikazivati samo JEDAN segment (npr. samo "Zainteresirani") → view_type: "segment"
- Slajd može prikazivati VIŠE segmenata iz istog bannera usporedno → view_type: "crosstab"
- Slajd može prikazivati TOTAL → view_type: "total"

=== FORMAT ODGOVORA ===
Vrati ISKLJUČIVO validan JSON (bez code fence-ova):
{{
  "slide_number": {slide_number},
  "datasets": [
    {{
      "description": "Kratki opis dataseta na slajdu",
      "excel_id": "C3",
      "question_code": "Q11",
      "view_type": "total|segment|crosstab|derived",
      "banner": "SPOL (Q1)",
      "segments_shown": ["Muškarac", "Žena"],
      "base_n_on_slide": 500,
      "confidence": 0.95,
      "reasoning": "Kratko obrazloženje"
    }}
  ],
  "slide_type": "data|title|separator|methodology|contents"
}}

PRAVILA:
- view_type:
  - "total" = slajd prikazuje ukupne rezultate (Total kolona) → banner: null, segments_shown: []
  - "segment" = slajd prikazuje podatke za jedan specifičan segment → banner: ime bannera, segments_shown: [1 segment]
  - "crosstab" = slajd prikazuje više segmenata usporedno → banner: ime bannera, segments_shown: [segmenti]
  - "derived" = izvedeni/agregirani podatak → banner: null, segments_shown: []
- banner: TOČNO ime bannera iz liste križanja (npr. "SPOL (Q1)", "DOB (Q2)"). null za total/derived.
- segments_shown: TOČNA imena segmenata iz tog bannera (prazna lista za total/derived)
- Ako slajd nema podataka (naslovni, separator, metodologija) → datasets: [], slide_type odgovarajući
- AKO NISI SIGURAN za match, stavi confidence < 0.6
- excel_id: koristi KRATKU OZNAKU iz liste kandidata (C1, C2, C3...)! NE puni ID.
- excel_id NIKAD ne smije biti "N/A", "NA", "null", "none", "?" ili prazno.
- Ako postoji i najmanja sumnja, i dalje odaberi NAJBLIŽI kandidat Cx i smanji confidence (<0.6).
- Isti Excel ID se može pojaviti više puta (npr. isti Q, različiti banneri)
- Za "derived" — objasni u reasoning kako je podatak izveden, koji Q su uključeni
"""

    return prompt, index_to_id


# ──────────────────────────────────────────────────────────────────────
# PASS 2: Targeted verification — compare numbers
# ──────────────────────────────────────────────────────────────────────

def build_pass2_prompt(
    slide_number: int,
    slide_texts: list[str],
    excel_data_blocks: list[dict],
) -> str:
    """
    Build Pass 2 prompt: Pro sees image + precise Excel data → number-by-number check.

    excel_data_blocks: list of {
        excel_id, question_code, question_text,
        view_type: "total"|"segment"|"crosstab",
        data: {column_name: [{label, value}]},  # only relevant columns
        derived_metrics: {mean, top2box, ...},
        base_n: int,                # Total N
        segment_sizes: {name: N},   # N per shown segment
    }
    """
    text_block = "\n".join(slide_texts) if slide_texts else "(nema teksta)"

    # Format Excel data blocks (compact, only relevant data)
    excel_text = ""
    for block in excel_data_blocks:
        eid = block.get("excel_id", "?")
        qcode = block.get("question_code", "?")
        qtext = block.get("question_text", "?")

        excel_text += f"\n  [{eid}] {qcode}: {qtext}\n"

        # Print each data column (Total and/or relevant segments)
        data_cols = block.get("data", {})
        for col_name, rows in data_cols.items():
            col_n = block.get("segment_sizes", {}).get(col_name) if col_name != "Total" else block.get("base_n")
            n_str = f" (N={col_n})" if col_n else ""
            excel_text += f"    {col_name}{n_str}:\n"
            for row in rows:
                label = row.get("label", "?")
                value = row.get("value")
                if value is not None:
                    excel_text += f"      {label}: {value}%\n"

        dm = block.get("derived_metrics", {})
        if dm:
            if dm.get("mean") is not None:
                excel_text += f"    MEAN: {dm['mean']}\n"
            if dm.get("top2box") is not None:
                excel_text += f"    Top 2 Box: {dm['top2box']}%\n"

    return f"""Ti si QC stručnjak za market research prezentacije.
Tvoj zadatak je temeljito provjeriti točnost podataka na slajdu usporedbom sa izvornim Excel podacima.
Budi VRLO precizan i kritičan.

=== SLAJD {slide_number} ===

Tekst iz PPTX:
{text_block}

=== EXCEL IZVORNI PODACI ===
Ovo su TOČNI Excel podaci za ovaj slajd (Total i/ili relevantni segmenti s punim vrijednostima):
{excel_text}

=== ZADATAK ===
ČITAJ PODATKE DIREKTNO SA SLIKE SLAJDA. Usporedi svaki broj sa slike s Excel izvornim podacima.

Provjeri:
1. DATA ISSUES — vrijednosti na slajdu vs. Excel:
   - Svaki broj na grafu/tablici usporedi s odgovarajućom Excel vrijednošću
   - Tolerancija zaokruživanja: ±0.5 postotnih bodova za %, ±0.1 za prosjeke
   - Provjeri i N/bazu ako je prikazana
   - Provjeri redoslijed opcija (ako je bitno za graf)
   - Provjeri SVAKI prikazani broj — ne preskači nijedan

   IZRAČUNI — OBAVEZNO:
   Ako slajd prikazuje bilo koji IZVEDENI podatak, MORAŠ ga sam izračunati iz Excel podataka
   i usporediti s prikazanom vrijednošću. Ovo uključuje:
   - Top 2 Box = zbroj prva 2 odgovora (npr. "Vrlo zadovoljan" + "Uglavnom zadovoljan")
   - Ukupno spontano = TOM + Ostalo spontano (zbroj svih spontano navedenih)
   - Gap/razlika = vrijednost_A - vrijednost_B
   - Zbroj segmenata = segment1 + segment2 + ... (može ≠ Total zbog nezavisnog zaokruživanja)
   - Prosječan broj = MEAN iz podataka
   Primjer: Ako Excel kaže "Vrlo zadovoljan: 32.1%, Uglavnom zadovoljan: 16.7%"
            onda Top 2 Box = 32.1 + 16.7 = 48.8%. Ako slajd kaže 49% → OK (zaokruživanje).
            Ako slajd kaže 52% → ERROR.
   AKO NE MOŽEŠ IZRAČUNATI jer podaci nedostaju, prijavi to kao "missing_data" warning.

2. TEXT ISSUES — tekstualni problemi:
   - Pravopisne greške (posebno u hrvatskom: č/ć, dž/đ, ije/je)
   - Gramatičke greške
   - Terminološka nekonzistentnost
   - Činjenične tvrdnje u tekstu koje ne odgovaraju podacima

3. VISUAL ISSUES — vizualni problemi (iz slike):
   - Odsječen tekst ili grafovi
   - Nečitljive legende ili osi
   - Krivi tipovi grafova za podatke
   - Nedostajuće jedinice ili baze
   - Nekonzistentno formatiranje

=== FORMAT ODGOVORA ===
Vrati ISKLJUČIVO validan JSON (bez code fence-ova):
{{
  "slide_number": {slide_number},
  "overall_status": "ok|warning|error",
  "data_issues": [
    {{
      "severity": "error|warning|info",
      "issue_type": "wrong_value|missing_data|wrong_order|wrong_base|wrong_label|rounding|extra_data|subset_mismatch",
      "detail": "Opis problema — budi konkretan, navedi točne brojeve",
      "slide_value": "45.2",
      "excel_value": "43.8"
    }}
  ],
  "text_issues": [
    {{
      "severity": "error|warning|info",
      "issue_type": "spelling|grammar|terminology|factual_claim",
      "detail": "Opis problema"
    }}
  ],
  "visual_issues": [
    {{
      "severity": "error|warning|info",
      "issue_type": "truncated|unreadable|wrong_chart_type|missing_info|formatting",
      "detail": "Opis problema"
    }}
  ],
  "summary": "Kratki sažetak nalaza za ovaj slajd"
}}

PRAVILA:
- Budi PRECIZAN — navedi konkretne brojeve koji se ne poklapaju
- error = kritična greška (krivi podaci prikazani), warning = potencijalni problem, info = napomena
- Ako je sve u redu, vrati prazne liste i overall_status: "ok"
- overall_status: "error" ako ima barem 1 error, "warning" ako ima warning ali ne error
- NE izmišljaj probleme — samo prijavi stvarne razlike
- Zaokruživanje ±0.5pp NIJE greška za postotke
- NE komentiraj datume i godine u izvještaju kao "buduće" — izvještaj je aktualan
- Ako Excel podaci imaju više stupaca (Total + segmenti), utvrdi koji stupac odgovara slajdu po bazi (N) i labelama
"""


# ──────────────────────────────────────────────────────────────────────
# LITE verification (Flash) — for slides without data
# ──────────────────────────────────────────────────────────────────────

def build_lite_verification_prompt(slide_number: int, slide_texts: list[str]) -> str:
    """Lightweight check for slides without matched data — spelling, grammar, visual only."""
    text_block = "\n".join(slide_texts) if slide_texts else "(nema teksta)"

    return f"""Ti si QC stručnjak za market research prezentacije.
Ovaj slajd NEMA matchane izvorne podatke — ne provjeravaj točnost brojeva.
Provjeri SVE OSTALO.

=== SLAJD {slide_number} ===

Tekst iz PPTX:
{text_block}

=== ZADATAK ===
Provjeri slajd bez usporedbe podataka s Excelom. Fokusiraj se na:

1. TEXT ISSUES:
   - Pravopisne greške (posebno hrvatski: č/ć, dž/đ, ije/je)
   - Gramatičke greške
   - Nedostajuće dijakritike (npr. "Ceska" umjesto "Češka")
   - Terminološka nekonzistentnost (miješanje hr/en, različiti nazivi za isto)

2. VISUAL ISSUES (iz slike):
   - Odsječen tekst ili grafovi
   - Nečitljive legende, osi, labele
   - Nedostajuće jedinice ili baze (N=)
   - Nekonzistentno formatiranje

=== FORMAT ODGOVORA ===
Vrati ISKLJUČIVO validan JSON (bez code fence-ova):
{{
  "slide_number": {slide_number},
  "overall_status": "ok|warning|error",
  "data_issues": [],
  "text_issues": [
    {{
      "severity": "error|warning|info",
      "issue_type": "spelling|grammar|terminology|missing_source|inconsistency",
      "detail": "Opis problema"
    }}
  ],
  "visual_issues": [
    {{
      "severity": "error|warning|info",
      "issue_type": "truncated|unreadable|missing_info|formatting|overlap",
      "detail": "Opis problema"
    }}
  ],
  "summary": "Kratki sažetak nalaza"
}}

PRAVILA:
- NE prijavljuj data_issues — nemaš Excel za usporedbu
- Budi precizan — citiraj konkretni problematični tekst
- NE komentiraj datume i godine u izvještaju kao "buduće" — izvještaj je aktualan
"""
