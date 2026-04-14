"""
Phase 2: Semantic matching — connect slide datasets to Excel questions.
Uses Flash model with metadata only (NO values) to prevent circular reasoning.
Processes one slide at a time + retries lost datasets.
"""
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

from checker.config import MAX_CONCURRENT_LLM
from checker.models.excel_dataset import ExcelDataset
from checker.models.slide_dataset import SlideExtraction
from checker.models.match_result import MatchResult, MatchingOutput
from checker.prompts.matching import build_matching_prompt
from checker.utils.gemini_client import GeminiClient
from checker.utils.json_utils import parse_json_response

# Confidence below this → force unmatched (wrong match worse than no match)
CONFIDENCE_CUTOFF = 0.55
MAX_RETRY_LOST = 1  # max retries for datasets LLM forgot to return
MAX_EXCEL_CANDIDATES = 50  # max Excel questions sent per slide matching call

_print_lock = threading.Lock()


def match_slides_to_excel(
    extractions: list[SlideExtraction],
    excel_datasets: list[ExcelDataset],
    client: GeminiClient,
    *,
    verbose: bool = False,
) -> MatchingOutput:
    """
    Phase 2: Match extracted slide data to Excel questions using semantic matching.
    Sends only metadata (no values) to prevent circular reasoning.
    Processes slides concurrently using ThreadPoolExecutor.
    """
    excel_metadata = _prepare_excel_metadata(excel_datasets)
    slides_with_data = [e for e in extractions if e.datasets]

    if not slides_with_data:
        return MatchingOutput(matches=[], unmatched_datasets=0)

    total_slides = len(slides_with_data)
    all_matches: list[MatchResult] = []

    def _match_slide_full(si: int, ext: SlideExtraction) -> list[MatchResult]:
        """Full match cycle for one slide: initial + retries + cutoff."""
        n_ds = len(ext.datasets)

        slide_matches = _match_one_slide(
            ext, excel_metadata, client, verbose=verbose,
            all_excel=excel_metadata,
        )

        # Retry lost datasets
        returned_indices = {m.dataset_index for m in slide_matches}
        expected_indices = {
            di for di, ds in enumerate(ext.datasets) if ds.data_points
        }
        missing = expected_indices - returned_indices

        retry = 0
        while missing and retry < MAX_RETRY_LOST:
            retry += 1
            retry_matches = _match_one_slide(
                ext, excel_metadata, client,
                only_indices=sorted(missing), verbose=verbose,
                is_retry=True, all_excel=excel_metadata,
            )
            slide_matches.extend(retry_matches)
            returned_indices = {m.dataset_index for m in slide_matches}
            missing = expected_indices - returned_indices

        for di in sorted(missing):
            slide_matches.append(MatchResult(
                slide_number=ext.slide_number,
                dataset_index=di,
                matched_excel_id=None,
                confidence=0.0,
                match_reasoning="Not returned by LLM after retries",
            ))

        for m in slide_matches:
            if m.matched_excel_id is not None and m.confidence < CONFIDENCE_CUTOFF:
                m.matched_excel_id = None
                m.unmatched = True

        n_matched = sum(1 for m in slide_matches if m.matched_excel_id is not None)
        with _print_lock:
            print(f"    Matching slajd {ext.slide_number} "
                  f"({si + 1}/{total_slides}, {n_ds} dataseta)... "
                  f"({n_matched}/{n_ds} matchano)")

        return slide_matches

    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_LLM) as executor:
        futures = [executor.submit(_match_slide_full, si, ext)
                   for si, ext in enumerate(slides_with_data)]
        for future in as_completed(futures):
            all_matches.extend(future.result())

    unmatched = sum(1 for m in all_matches if m.matched_excel_id is None)
    return MatchingOutput(matches=all_matches, unmatched_datasets=unmatched)


def _prepare_excel_metadata(datasets: list[ExcelDataset]) -> list[dict]:
    """
    Prepare Excel metadata for matching prompt.
    CRITICAL: NO values included — only labels, codes, question text.
    """
    metadata = []
    for ds in datasets:
        labels = [c.label for c in ds.categories]
        metadata.append({
            "id": ds.id,
            "question_code": ds.question_code,
            "question_text": ds.question_text,
            "labels": labels,
            "base_n": ds.base_n,
            "type": ds.type,
        })
    return metadata


def _match_one_slide(
    ext: SlideExtraction,
    excel_metadata: list[dict],
    client: GeminiClient,
    *,
    only_indices: list[int] | None = None,
    verbose: bool = False,
    is_retry: bool = False,
    all_excel: list[dict] | None = None,
) -> list[MatchResult]:
    """Match datasets from ONE slide against the full Excel index.
    If only_indices is set, include only those dataset indices in the prompt (retry).
    Pre-filters Excel candidates to reduce prompt size.
    """
    datasets_for_prompt = []
    index_map: list[int] = []  # prompt position → real dataset_index

    for di, ds in enumerate(ext.datasets):
        if only_indices is not None and di not in only_indices:
            continue
        # Skip datasets with no data points (text-only, empty extractions)
        if not ds.data_points:
            continue
        datasets_for_prompt.append({
            "title": ds.title,
            "chart_type": ds.chart_type,
            "data_points": [{"label": dp.label} for dp in ds.data_points],
            "unit": ds.unit,
            "base_n": ds.base_n,
            "subset": ds.subset,
            "series_name": ds.series_name,
            "dataset_index": di,  # pass real index so LLM returns it
        })
        index_map.append(di)

    slide_data = [{
        "slide_number": ext.slide_number,
        "datasets": datasets_for_prompt,
    }]

    # Pre-filter Excel candidates to reduce prompt size
    filtered_excel = _filter_excel_candidates(datasets_for_prompt, all_excel or excel_metadata)

    prompt = build_matching_prompt(
        slide_data, filtered_excel,
        expected_count=len(datasets_for_prompt),
        is_retry=is_retry,
    )
    raw_response = client.call_flash(prompt)

    try:
        matches_raw = parse_json_response(raw_response)
    except ValueError:
        return [
            MatchResult(
                slide_number=ext.slide_number,
                dataset_index=di,
                matched_excel_id=None,
                confidence=0.0,
                match_reasoning="LLM response failed to parse",
            )
            for di in index_map
        ]

    # Unwrap dict-wrapped responses
    if isinstance(matches_raw, dict):
        for key in ("matches", "results", "data"):
            if key in matches_raw and isinstance(matches_raw[key], list):
                matches_raw = matches_raw[key]
                break

    valid_indices = set(index_map)
    results: list[MatchResult] = []
    if isinstance(matches_raw, list):
        for m in matches_raw:
            di = m.get("dataset_index")
            if di is None:
                continue  # Skip malformed results without dataset_index
            if di not in valid_indices:
                continue  # LLM hallucinated a non-existent index
            results.append(MatchResult(
                slide_number=m.get("slide_number") or ext.slide_number,
                dataset_index=di,
                matched_excel_id=m.get("matched_excel_id"),
                confidence=round(m.get("confidence") or 0.0, 2),
                match_reasoning=m.get("match_reasoning") or "",
            ))

    return results


def _filter_excel_candidates(
    slide_datasets: list[dict],
    all_excel: list[dict],
) -> list[dict]:
    """Pre-filter Excel questions to top N candidates based on label/title overlap.
    Sends fewer candidates to LLM → less noise, better matching, cheaper.
    """
    if len(all_excel) <= MAX_EXCEL_CANDIDATES:
        return all_excel

    # Collect all tokens from slide datasets (titles + labels, lowercased)
    slide_tokens: set[str] = set()
    for ds in slide_datasets:
        title = (ds.get("title") or "").lower()
        for word in title.split():
            if len(word) >= 3:
                slide_tokens.add(word)
        for dp in ds.get("data_points", []):
            label = (dp.get("label") or "").lower()
            for word in label.split():
                if len(word) >= 3:
                    slide_tokens.add(word)
        if ds.get("subset"):
            for word in ds["subset"].lower().split():
                if len(word) >= 3:
                    slide_tokens.add(word)

    # Score each Excel question by token overlap
    scored: list[tuple[int, dict]] = []
    for eq in all_excel:
        score = 0
        eq_text = f"{eq.get('question_code', '')} {eq.get('question_text', '')}".lower()
        eq_labels = " ".join(eq.get("labels", [])).lower()
        combined = eq_text + " " + eq_labels
        for token in slide_tokens:
            if token in combined:
                score += 1
        scored.append((score, eq))

    # Sort by score descending, take top N (only candidates with score > 0)
    scored.sort(key=lambda x: x[0], reverse=True)
    result = [eq for score, eq in scored[:MAX_EXCEL_CANDIDATES] if score > 0]

    return result
