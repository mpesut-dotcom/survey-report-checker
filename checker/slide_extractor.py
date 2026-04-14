"""
Phase 1 (merged): Extract data from slides AND match to Excel questions.
For each slide: send image + PPTX text + Excel candidates → get extraction + matches.
"""
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from checker.config import MAX_CONCURRENT_LLM
from checker.models.excel_dataset import ExcelDataset
from checker.models.slide_dataset import SlideExtraction, SlideDataset, DataPoint, TextElement
from checker.prompts.extract_and_match import build_extract_and_match_prompt
from checker.utils.gemini_client import GeminiClient
from checker.utils.json_utils import parse_json_response

_print_lock = threading.Lock()

# Confidence below this → force unmatched
CONFIDENCE_CUTOFF = 0.60
MAX_EXCEL_CANDIDATES = 50

# Regex to detect question codes like Q5, Q7a, Q12, EXTRA Q3
_Q_CODE_RE = re.compile(r'\b(?:EXTRA\s+)?Q\d+[a-z]?\b', re.IGNORECASE)
_NUMERIC_LABEL_RE = re.compile(r'^\d+\+?$', re.IGNORECASE)
_NUMERIC_RESERVE_SLOTS = 5


def _looks_like_numeric_distribution(labels: list[str]) -> bool:
    """Heuristic: labels like 0,1,2,3,... indicate count-distribution tables."""
    if not labels:
        return False
    cleaned = [str(l).strip() for l in labels if str(l).strip()]
    if len(cleaned) < 4:
        return False
    numeric_like = sum(1 for l in cleaned if _NUMERIC_LABEL_RE.match(l))
    return numeric_like >= 4


def _select_candidate_subset(
    scored: list[tuple[int, dict]],
    max_candidates: int,
) -> list[dict]:
    """Select candidates with a small diversity reserve for numeric distributions.

    This avoids monocultures where one dense question family crowds out other
    structurally relevant tables.
    """
    positive = [(score, eq) for score, eq in scored if score > 0]
    if len(positive) <= max_candidates:
        return [eq for _, eq in positive]

    reserve = min(_NUMERIC_RESERVE_SLOTS, max_candidates // 5)
    if reserve <= 0:
        return [eq for _, eq in positive[:max_candidates]]

    primary_quota = max_candidates - reserve
    selected = [eq for _, eq in positive[:primary_quota]]
    selected_ids = {eq.get("id") for eq in selected}

    numeric_pool = [
        eq for _, eq in positive
        if _looks_like_numeric_distribution([str(x) for x in eq.get("labels", [])])
    ]
    for eq in numeric_pool:
        if len(selected) >= max_candidates:
            break
        eq_id = eq.get("id")
        if eq_id in selected_ids:
            continue
        selected.append(eq)
        selected_ids.add(eq_id)

    if len(selected) < max_candidates:
        for _, eq in positive[primary_quota:]:
            if len(selected) >= max_candidates:
                break
            eq_id = eq.get("id")
            if eq_id in selected_ids:
                continue
            selected.append(eq)
            selected_ids.add(eq_id)

    return selected


def extract_and_match_all_slides(
    slide_texts: list[dict],
    slide_images: dict[int, Path],
    excel_datasets: list[ExcelDataset],
    client: GeminiClient,
    *,
    slide_filter: list[int] | None = None,
    verbose: bool = False,
) -> list[SlideExtraction]:
    """
    Phase 1 (merged): Extract structured data AND match to Excel for each slide.
    Processes slides concurrently using ThreadPoolExecutor.
    """
    excel_metadata = _prepare_excel_metadata(excel_datasets)

    results: list[SlideExtraction] = []
    total = len(slide_texts)

    # Separate skipped from processable slides
    slides_to_process = []
    for slide_info in slide_texts:
        sn = slide_info["slide_number"]

        if slide_filter and sn not in slide_filter:
            continue

        if sn not in slide_images:
            if verbose:
                print(f"    Slajd {sn}: nema slike, preskačem")
            results.append(SlideExtraction(
                slide_number=sn, datasets=[], text_elements=[],
                raw_texts=slide_info["all_texts"],
            ))
            continue

        slides_to_process.append(slide_info)

    def _process(slide_info):
        sn = slide_info["slide_number"]
        raw_texts = slide_info["all_texts"]

        # Pre-filter Excel candidates using raw PPTX text
        filtered_excel = _filter_excel_candidates_from_texts(raw_texts, excel_metadata)

        extraction = _extract_and_match_single_slide(
            slide_number=sn,
            texts=raw_texts,
            image_path=slide_images[sn],
            excel_candidates=filtered_excel,
            client=client,
        )
        n_ds = len(extraction.datasets)
        n_matched = sum(1 for ds in extraction.datasets if ds.matched_excel_id)
        with _print_lock:
            print(f"    Slajd {sn}/{total}... "
                  f"({n_ds} dataset{'a' if n_ds != 1 else ''}, "
                  f"{n_matched} matchano)")
        return extraction

    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_LLM) as executor:
        futures = [executor.submit(_process, si) for si in slides_to_process]
        for future in as_completed(futures):
            results.append(future.result())

    results.sort(key=lambda e: e.slide_number)
    return results


def _prepare_excel_metadata(datasets: list[ExcelDataset]) -> list[dict]:
    """Prepare Excel metadata for matching. NO values — only labels, codes, text, banner-grouped segments."""
    metadata = []
    for ds in datasets:
        labels = [c.label for c in ds.categories]
        # Collect banners with their segment names from grouped breakdowns
        banners: dict[str, list[str]] = {}
        for c in ds.categories:
            for banner_name, segs in c.breakdowns.items():
                if banner_name not in banners:
                    banners[banner_name] = sorted(segs.keys())
        metadata.append({
            "id": ds.id,
            "question_code": ds.question_code,
            "question_text": ds.question_text,
            "labels": labels,
            "base_n": ds.base_n,
            "type": ds.type,
            "banners": banners,  # {banner_name: [seg1, seg2, ...]}
        })
    return metadata


def _filter_excel_candidates_from_texts(
    slide_texts: list[str],
    all_excel: list[dict],
    max_candidates: int = MAX_EXCEL_CANDIDATES,
) -> list[dict]:
    """Pre-filter Excel questions based on token overlap with raw PPTX text.
    Uses raw slide texts (not LLM extraction) for speed and independence.
    Question codes (Q5, Q12, etc.) get high weight since they're the strongest signal.
    """
    if len(all_excel) <= max_candidates:
        return all_excel

    full_text = " ".join(slide_texts).lower()
    numeric_tokens_in_slide = len(re.findall(r'\b\d+\b', full_text))

    # Collect word tokens from raw PPTX text shapes (≥3 chars)
    slide_tokens: set[str] = set()
    for text in slide_texts:
        for word in text.lower().split():
            if len(word) >= 3:
                slide_tokens.add(word)

    # Extract question codes from slide text (Q5, Q7a, Q12, etc.)
    slide_q_codes: set[str] = set()
    for text in slide_texts:
        for m in _Q_CODE_RE.finditer(text):
            slide_q_codes.add(m.group().upper())

    # Score each Excel question by token overlap + question code match
    scored: list[tuple[int, dict]] = []
    for eq in all_excel:
        score = 0
        eq_code = eq.get("question_code", "").upper()
        eq_id = eq.get("id", "").upper()
        eq_text = f"{eq.get('question_code', '')} {eq.get('question_text', '')}".lower()
        eq_labels = " ".join(eq.get("labels", [])).lower()
        combined = eq_text + " " + eq_labels + " " + eq_id.lower()

        # Question code match = very strong signal
        # Must match exactly or as prefix followed by non-digit (Q5→Q5a OK, Q5→Q50 NOT OK)
        # Also check embedded Q-codes (e.g. "[SPO] q15_long" contains Q15)
        for sq in slide_q_codes:
            if eq_code == sq:
                score += 10
                break
            if eq_code.startswith(f"{sq}."):
                score += 6
                break
            if eq_code.startswith(sq) and (len(eq_code) == len(sq) or not eq_code[len(sq)].isdigit()):
                score += 8
                break
            # Embedded Q-code: eq_code contains the slide Q-code somewhere inside
            # e.g. eq_code="[SPO] Q15_LONG" and sq="Q15"
            idx = eq_code.find(sq)
            if idx > 0:
                after = idx + len(sq)
                if after == len(eq_code) or not eq_code[after].isdigit():
                    score += 8
                    break

            # Some derived/aggregated tables keep Q-code in internal dataset id only.
            if sq in eq_id:
                score += 5
                break

        # Word token overlap (question text + labels)
        for token in slide_tokens:
            if token in combined:
                score += 1

        # Generic structural boost: tables with numeric category labels are often
        # count distributions and can be highly relevant on number-dense slides.
        labels = [str(x) for x in eq.get("labels", [])]
        if _looks_like_numeric_distribution(labels) and numeric_tokens_in_slide >= 4:
            score += 6

        # Banner/segment name match — if slide text contains a banner name or segment name,
        # this question's cross-tab is relevant (+3 for banner name, +1 per segment, max +6)
        eq_banners = eq.get("banners", {})
        banner_bonus = 0
        for banner_name, seg_list in eq_banners.items():
            # Check banner name (e.g. "SPOL (Q1)") — extract readable part
            # Extract the descriptive part before the parenthesized Q code
            banner_desc = re.sub(r'\s*\([^)]*\)\s*$', '', banner_name).strip().lower()
            if banner_desc and len(banner_desc) >= 3 and banner_desc in full_text:
                banner_bonus += 3
            # Check individual segment names
            for seg_name in seg_list:
                if len(seg_name) >= 3 and seg_name.lower() in full_text:
                    banner_bonus += 1
        score += min(banner_bonus, 6)

        scored.append((score, eq))

    scored.sort(key=lambda x: x[0], reverse=True)
    selected = _select_candidate_subset(scored, max_candidates)

    # Safety fallback: never return an empty candidate list.
    # Some slides have little/no extractable text (image-only), which can lead to
    # zero token overlap and an empty positive set.
    if not selected:
        selected = [eq for _, eq in scored[:max_candidates]]

    return selected


def filter_excel_candidates(
    slide_texts: list[str],
    excel_datasets: list[ExcelDataset],
    max_candidates: int = MAX_EXCEL_CANDIDATES,
) -> list[ExcelDataset]:
    """Public API: pre-filter Excel datasets by relevance to slide texts.
    Returns full ExcelDataset objects (not metadata dicts).
    Used by verifier to build context for Pro model.
    """
    metadata = _prepare_excel_metadata(excel_datasets)
    filtered_meta = _filter_excel_candidates_from_texts(slide_texts, metadata, max_candidates)
    filtered_ids = {m["id"] for m in filtered_meta}
    return [ds for ds in excel_datasets if ds.id in filtered_ids]


def _extract_and_match_single_slide(
    slide_number: int,
    texts: list[str],
    image_path: Path,
    excel_candidates: list[dict],
    client: GeminiClient,
) -> SlideExtraction:
    """Extract data and match to Excel for a single slide."""
    prompt = build_extract_and_match_prompt(slide_number, texts, excel_candidates)
    image_part = client.make_image_part(image_path)

    raw_response = client.call_flash_multimodal([prompt, image_part])

    try:
        data = parse_json_response(raw_response)
    except ValueError:
        print(f"    WARN: Slajd {slide_number} — LLM odgovor nije validan JSON, preskačem")
        return SlideExtraction(
            slide_number=slide_number, datasets=[], text_elements=[],
            raw_texts=texts,
        )

    # Parse datasets (with match fields)
    datasets: list[SlideDataset] = []
    for ds_raw in data.get("datasets", []):
        try:
            data_points = []
            for dp in ds_raw.get("data_points", []):
                raw_val = dp.get("value")
                if isinstance(raw_val, str):
                    try:
                        raw_val = float(raw_val.strip().rstrip("%"))
                    except (ValueError, TypeError):
                        raw_val = None
                data_points.append(DataPoint(
                    label=dp.get("label") or "",
                    value=raw_val,
                ))

            # Parse match fields
            matched_id = ds_raw.get("matched_excel_id")
            confidence = round(ds_raw.get("confidence") or 0.0, 2)

            # Apply confidence cutoff
            if matched_id is not None and confidence < CONFIDENCE_CUTOFF:
                matched_id = None
                confidence = 0.0

            datasets.append(SlideDataset(
                title=ds_raw.get("title") or "",
                question_code=ds_raw.get("question_code"),
                chart_type=ds_raw.get("chart_type") or "unknown",
                data_points=data_points,
                unit=ds_raw.get("unit") or "%",
                base_n=ds_raw.get("base_n"),
                base_description=ds_raw.get("base_description"),
                time_period=ds_raw.get("time_period"),
                subset=ds_raw.get("subset"),
                series_name=ds_raw.get("series_name"),
                note=ds_raw.get("note"),
                matched_excel_id=matched_id,
                confidence=confidence,
                match_reasoning=ds_raw.get("match_reasoning") or "",
            ))
        except Exception as e:
            print(f"    WARN: Greška pri parsiranju dataseta na slajdu {slide_number}: {e}")

    # Parse text elements
    text_elements: list[TextElement] = []
    for te_raw in data.get("text_elements", []):
        try:
            text_elements.append(TextElement(
                type=te_raw.get("type") or "annotation",
                content=te_raw.get("content") or "",
            ))
        except Exception:
            pass

    return SlideExtraction(
        slide_number=slide_number,
        datasets=_deduplicate_datasets(datasets),
        text_elements=text_elements,
        raw_texts=texts,
    )


def _deduplicate_datasets(datasets: list[SlideDataset]) -> list[SlideDataset]:
    """Merge duplicate datasets that Flash fragmented from the same table.

    Groups by (title, chart_type). If a group has multiple entries with identical
    data_points, keeps only one (preferring the one with a match).
    If data_points differ (e.g. different series), they are kept as separate datasets.
    """
    if len(datasets) <= 1:
        return datasets

    from collections import defaultdict
    groups: dict[tuple, list[SlideDataset]] = defaultdict(list)
    for ds in datasets:
        key = (ds.title.strip().lower(), ds.chart_type)
        groups[key].append(ds)

    result: list[SlideDataset] = []
    for key, group in groups.items():
        if len(group) == 1:
            result.append(group[0])
            continue

        # Within the group, dedup by data_points fingerprint
        # When duplicates exist, prefer the one with a match
        seen_fingerprints: dict[str, SlideDataset] = {}
        for ds in group:
            fp = _dataset_fingerprint(ds)
            existing = seen_fingerprints.get(fp)
            if existing is None:
                seen_fingerprints[fp] = ds
            elif ds.matched_excel_id and not existing.matched_excel_id:
                # New one is matched, old one isn't — replace
                seen_fingerprints[fp] = ds
            elif ds.matched_excel_id and existing.matched_excel_id and ds.confidence > existing.confidence:
                # Both matched, prefer higher confidence
                seen_fingerprints[fp] = ds
        result.extend(seen_fingerprints.values())

    return result


def _dataset_fingerprint(ds: SlideDataset) -> str:
    """Create a dedup fingerprint from labels + values + series_name."""
    parts = [ds.series_name or ""]
    for dp in sorted(ds.data_points, key=lambda p: p.label):
        parts.append(f"{dp.label}={dp.value}")
    return "|".join(parts)
