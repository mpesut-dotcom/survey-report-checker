"""
Two-pass deep verification using Pro model.

Pass 1 — Smart Match:  image + metadata → Pro identifies which Excel questions
    are on the slide, view type (total/segment/crosstab), and which segments.
Pass 2 — Targeted Verify: image + exact data (only relevant columns) → Pro
    verifies numbers, text, and visuals.
"""
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import re

from checker.config import MAX_CONCURRENT_LLM
from checker.models.excel_dataset import ExcelDataset
from checker.models.verification import (
    SlideVerification, DataIssue, TextIssue, VisualIssue, MatchSource, MatchFailure,
)
from checker.prompts.verification import (
    build_pass1_prompt, build_pass2_prompt, build_lite_verification_prompt,
)
from checker.slide_extractor import filter_excel_candidates, _prepare_excel_metadata
from checker.utils.gemini_client import GeminiClient
from checker.utils.json_utils import parse_json_response

_print_lock = threading.Lock()

# Max candidates for Pass 1 metadata (lightweight, can be generous)
MAX_PASS1_CANDIDATES = 50
# Confidence cutoff for Pass 1 match
PASS1_CONFIDENCE_CUTOFF = 0.60


_INVALID_EXCEL_IDS = {"", "N/A", "NA", "NONE", "NULL", "?", "-"}


def _normalize_q_code(value: str) -> str:
    return re.sub(r"\s+", "", (value or "").upper())


def _tokenize_text(value: str) -> set[str]:
    return {t for t in re.findall(r"[A-Za-z0-9ČĆŽŠĐčćžšđ]+", (value or "").lower()) if len(t) >= 3}


def _dataset_has_numeric_distribution(ds: ExcelDataset) -> bool:
    labels = [str(c.label).strip() for c in ds.categories if c.label]
    if len(labels) < 4:
        return False
    numeric_like = sum(1 for label in labels if re.fullmatch(r"\d+\+?", label))
    return numeric_like >= 4


def _find_related_family_datasets(
    seed: ExcelDataset,
    candidates: list[ExcelDataset],
    max_related: int = 8,
) -> list[ExcelDataset]:
    """Find closely related datasets that should be verified together.

    Generic criteria: same source file + high question-text overlap and/or shared
    numeric-distribution structure.
    """
    seed_tokens = _tokenize_text(seed.question_text)
    seed_numeric = _dataset_has_numeric_distribution(seed)

    related: list[tuple[int, float, ExcelDataset]] = []
    for ds in candidates:
        if ds.id == seed.id:
            continue
        if ds.file_key != seed.file_key:
            continue

        ds_tokens = _tokenize_text(ds.question_text)
        union = seed_tokens | ds_tokens
        overlap = len(seed_tokens & ds_tokens)
        jaccard = (overlap / len(union)) if union else 0.0

        ds_numeric = _dataset_has_numeric_distribution(ds)
        similar_text = overlap >= 3 and jaccard >= 0.25
        similar_structure = seed_numeric and ds_numeric and overlap >= 1

        if similar_text or similar_structure:
            related.append((overlap, jaccard, ds))

    related.sort(key=lambda item: (item[0], item[1], item[2].question_text), reverse=True)
    return [ds for _, _, ds in related[:max_related]]


def _resolve_dataset_from_pass1_match(
    match: dict,
    index_to_id: dict[str, str],
    excel_by_id: dict[str, ExcelDataset],
    candidate_by_id: dict[str, ExcelDataset],
) -> tuple[ExcelDataset | None, str]:
    """Resolve Pass 1 match to ExcelDataset with robust fallbacks.

    Priority:
    1) C-index mapping (C1, C2...) from prompt metadata
    2) Direct real Excel ID (if model ignored instruction)
    3) question_code match (candidate set first, then global)
    4) description/question_text similarity against candidate question text
    """
    raw_id = str(match.get("excel_id") or "").strip()
    norm_id = raw_id.upper()

    # 1) Expected path: C-index provided by model.
    if raw_id in index_to_id:
        real_id = index_to_id[raw_id]
        ds = excel_by_id.get(real_id)
        if ds:
            return ds, f"index:{raw_id}"

    # 2) Model provided real Excel ID directly.
    if norm_id not in _INVALID_EXCEL_IDS and raw_id in excel_by_id:
        return excel_by_id[raw_id], f"direct_id:{raw_id}"

    # 3) Fallback via question_code.
    q_code = _normalize_q_code(str(match.get("question_code") or ""))
    if q_code:
        for ds in candidate_by_id.values():
            if _normalize_q_code(ds.question_code) == q_code:
                return ds, f"q_code_candidate:{q_code}"
        for ds in excel_by_id.values():
            if _normalize_q_code(ds.question_code) == q_code:
                return ds, f"q_code_global:{q_code}"

    # 4) Last-resort text similarity against candidate question_text.
    desc = str(match.get("description") or "")
    desc_tokens = _tokenize_text(desc)
    if desc_tokens:
        best_ds = None
        best_overlap = 0
        for ds in candidate_by_id.values():
            q_tokens = _tokenize_text(ds.question_text)
            overlap = len(desc_tokens & q_tokens)
            if overlap > best_overlap:
                best_overlap = overlap
                best_ds = ds
        if best_ds and best_overlap >= 2:
            return best_ds, f"desc_overlap:{best_overlap}"

    return None, "unresolved"


def verify_all_slides(
    excel_datasets: list[ExcelDataset],
    slide_texts: list[dict],
    slide_images: dict[int, Path],
    client: GeminiClient,
    *,
    slide_filter: list[int] | None = None,
    verbose: bool = False,
) -> list[SlideVerification]:
    """
    Two-pass verification of each slide.
    Pass 1: Pro identifies Excel matches from metadata.
    Pass 2: Pro verifies numbers from precise data.
    Processes ALL slides with images — Pro decides which have data.
    """
    # Build lookup structures
    excel_by_id = {ds.id: ds for ds in excel_datasets}
    texts_by_slide = {s["slide_number"]: s for s in slide_texts}

    # All slides with images go through Pro 2-pass verification
    slides_to_verify = sorted(slide_images.keys())

    if slide_filter:
        slides_to_verify = [sn for sn in slides_to_verify if sn in slide_filter]

    total = len(slides_to_verify)
    results: list[SlideVerification] = []

    def _verify_slide(idx_sn: tuple[int, int]) -> SlideVerification | None:
        idx, sn = idx_sn
        texts = texts_by_slide.get(sn, {}).get("all_texts", [])

        try:
            verification = _verify_single_slide_2pass(
                slide_number=sn,
                excel_datasets=excel_datasets,
                excel_by_id=excel_by_id,
                texts=texts,
                image_path=slide_images[sn],
                client=client,
                verbose=verbose,
            )
        except Exception as exc:
            with _print_lock:
                print(f"    ERROR slajd {sn}: {exc!r}")
            verification = SlideVerification(
                slide_number=sn,
                overall_status="error",
                data_issues=[DataIssue(
                    severity="error", issue_type="crash",
                    detail=f"Verifikacija pala s greškom: {exc}",
                )],
                summary=f"Crash: {exc}",
            )

        status = verification.overall_status
        n_issues = (len(verification.data_issues) +
                    len(verification.text_issues) +
                    len(verification.visual_issues))
        with _print_lock:
            print(f"    Verifikacija {idx}/{total} — slajd {sn}... "
                  f"[{status}] ({n_issues} nalaza)")
        return verification

    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_LLM) as executor:
        futures = [executor.submit(_verify_slide, (idx, sn))
                   for idx, sn in enumerate(slides_to_verify, 1)]
        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)

    # Sort results by slide number
    results.sort(key=lambda v: v.slide_number)
    return results


# ──────────────────────────────────────────────────────────────────────
# Two-pass verification for a single slide
# ──────────────────────────────────────────────────────────────────────

def _verify_single_slide_2pass(
    slide_number: int,
    excel_datasets: list[ExcelDataset],
    excel_by_id: dict[str, ExcelDataset],
    texts: list[str],
    image_path: Path,
    client: GeminiClient,
    *,
    verbose: bool = False,
) -> SlideVerification:
    """Two-pass verification: smart match → targeted verify."""

    image_part = client.make_image_part(image_path)

    # ── PASS 1: Smart Match ──
    # Pre-filter candidates (metadata only, lightweight)
    candidates = filter_excel_candidates(texts, excel_datasets, MAX_PASS1_CANDIDATES)
    if not candidates:
        candidates = excel_datasets[:MAX_PASS1_CANDIDATES]
    candidate_by_id = {ds.id: ds for ds in candidates}
    candidate_metadata = _prepare_excel_metadata(candidates)

    if verbose:
        with _print_lock:
            print(f"      [slajd {slide_number}] PASS 1 start: {len(candidates)} kandidata")

    pass1_prompt, index_to_id = build_pass1_prompt(slide_number, texts, candidate_metadata)
    raw_pass1 = client.call_pro_multimodal([pass1_prompt, image_part])

    try:
        pass1_data = parse_json_response(raw_pass1)
    except ValueError:
        return SlideVerification(
            slide_number=slide_number,
            overall_status="error",
            pass1_slide_type="parse_error",
            pass1_total_candidates=len(candidates),
            data_issues=[DataIssue(
                severity="error", issue_type="parse_error",
                detail=f"Pass 1 LLM odgovor nije validan JSON: {raw_pass1[:200]}",
            )],
            summary="Greška u parsiranju Pass 1 odgovora.",
        )

    # Check if slide has data
    slide_type = pass1_data.get("slide_type", "data")
    pass1_datasets = pass1_data.get("datasets", [])

    match_failures: list[MatchFailure] = []
    confident_matches: list[dict] = []
    for d in pass1_datasets:
        raw_id = str(d.get("excel_id") or "").strip()
        q_code = str(d.get("question_code") or "")
        desc = str(d.get("description") or "")
        try:
            conf = float(d.get("confidence", 0) or 0)
        except (TypeError, ValueError):
            conf = 0.0

        if not raw_id:
            match_failures.append(MatchFailure(
                excel_id="",
                question_code=q_code,
                confidence=conf,
                reason="missing_excel_id",
                description=desc,
            ))
            continue

        if conf < PASS1_CONFIDENCE_CUTOFF:
            match_failures.append(MatchFailure(
                excel_id=raw_id,
                question_code=q_code,
                confidence=conf,
                reason="low_confidence",
                description=desc,
            ))
            continue

        confident_matches.append(d)

    if verbose:
        with _print_lock:
            print(f"      [slajd {slide_number}] PASS 1 done: "
                  f"slide_type={slide_type}, datasets={len(pass1_datasets)}")

    if slide_type != "data" or not pass1_datasets:
        # No data on slide — do lite verification (spelling/visual only)
        lite = _verify_lite(slide_number, texts, image_path, client, verbose=verbose)
        lite.pass1_slide_type = slide_type
        lite.pass1_total_candidates = len(candidates)
        lite.pass1_total_datasets = len(pass1_datasets)
        lite.pass1_confident_datasets = len(confident_matches)
        lite.match_failures = match_failures
        return lite

    if not confident_matches:
        # Data slide but no confident match — still do lite check (spelling/visual)
        lite = _verify_lite(slide_number, texts, image_path, client, verbose=verbose)
        lite.summary = ("Pass 1: nijedan Excel kandidat nije matchan s dovoljnom "
                        "sigurnošću. " + (lite.summary or ""))
        lite.pass1_slide_type = slide_type
        lite.pass1_total_candidates = len(candidates)
        lite.pass1_total_datasets = len(pass1_datasets)
        lite.pass1_confident_datasets = len(confident_matches)
        lite.match_failures = match_failures
        return lite

    if verbose:
        with _print_lock:
            print(f"      [slajd {slide_number}] PASS 1 confident match: {len(confident_matches)}")

    # ── BUILD DATA FOR PASS 2 ──
    excel_data_blocks = []
    match_sources: list[MatchSource] = []
    unresolved_matches: list[str] = []
    for md in confident_matches:
        raw_id = str(md.get("excel_id") or "")
        q_code = str(md.get("question_code") or "")
        desc = str(md.get("description") or "")
        try:
            conf = float(md.get("confidence", 0) or 0)
        except (TypeError, ValueError):
            conf = 0.0

        excel_ds, resolved_by = _resolve_dataset_from_pass1_match(
            md,
            index_to_id,
            excel_by_id,
            candidate_by_id,
        )

        if excel_ds:
            if resolved_by not in ("", "unresolved") and not resolved_by.startswith("index:"):
                with _print_lock:
                    print(f"    WARN slajd {slide_number}: excel_id '{raw_id}' "
                          f"resolved via {resolved_by} -> {excel_ds.id}")
        else:
            unresolved_matches.append(f"id='{raw_id}', q='{q_code}'")
            match_failures.append(MatchFailure(
                excel_id=raw_id,
                question_code=q_code,
                confidence=conf,
                reason="unresolved_after_pass1",
                description=desc,
            ))
            with _print_lock:
                print(f"    WARN slajd {slide_number}: excel_id '{raw_id}' unresolved "
                      f"(q_code='{q_code}')")
            continue

        view_type = md.get("view_type", "total")
        banner_name = md.get("banner")  # e.g. "SPOL (Q1)" or null
        segments_shown = md.get("segments_shown", [])

        # Build data columns: always include Total + requested segments
        data_cols = {}

        # Total column
        data_cols["Total"] = [
            {"label": c.label, "value": c.total}
            for c in excel_ds.categories
        ]

        # Add requested segment columns from the identified banner
        if banner_name and segments_shown:
            for seg_name in segments_shown:
                col_data = []
                for c in excel_ds.categories:
                    banner_data = c.breakdowns.get(banner_name, {})
                    seg_val = banner_data.get(seg_name)
                    if seg_val is not None:
                        col_data.append({"label": c.label, "value": seg_val})
                if col_data:
                    data_cols[seg_name] = col_data

            # Fuzzy fallback: if segment wasn't found by exact name, try substring match
            for seg_name in segments_shown:
                if seg_name in data_cols:
                    continue
                target = seg_name.lower()
                # Search within the specified banner
                first_cat = excel_ds.categories[0] if excel_ds.categories else None
                if first_cat and banner_name in first_cat.breakdowns:
                    for actual_seg in first_cat.breakdowns[banner_name]:
                        if target in actual_seg.lower() or actual_seg.lower() in target:
                            col_data = []
                            for c in excel_ds.categories:
                                sv = c.breakdowns.get(banner_name, {}).get(actual_seg)
                                if sv is not None:
                                    col_data.append({"label": c.label, "value": sv})
                            if col_data:
                                data_cols[actual_seg] = col_data
                            break

        elif segments_shown and not banner_name:
            # Fallback: Pro didn't specify banner — search across all banners
            for seg_name in segments_shown:
                for c in excel_ds.categories:
                    for bn, segs in c.breakdowns.items():
                        if seg_name in segs:
                            banner_name = bn  # lock to first found banner
                            break
                    if banner_name:
                        break
                if banner_name:
                    break
            if banner_name:
                for seg_name in segments_shown:
                    col_data = []
                    for c in excel_ds.categories:
                        seg_val = c.breakdowns.get(banner_name, {}).get(seg_name)
                        if seg_val is not None:
                            col_data.append({"label": c.label, "value": seg_val})
                    if col_data:
                        data_cols[seg_name] = col_data

        # Collect segment sizes for shown segments from the banner
        seg_sizes = {}
        if banner_name and banner_name in excel_ds.segment_sizes:
            banner_sizes = excel_ds.segment_sizes[banner_name]
            for seg_name in data_cols:
                if seg_name != "Total" and seg_name in banner_sizes:
                    seg_sizes[seg_name] = banner_sizes[seg_name]

        dm = {}
        if excel_ds.derived_metrics:
            dm = {
                "mean": excel_ds.derived_metrics.mean,
                "top2box": excel_ds.derived_metrics.top2box,
                "net": excel_ds.derived_metrics.net,
            }

        match_sources.append(MatchSource(
            excel_id=excel_ds.id,
            question_code=excel_ds.question_code,
            question_text=excel_ds.question_text,
            view_type=str(view_type or ""),
            pass1_excel_id=raw_id,
            pass1_question_code=q_code,
            confidence=conf,
            resolved_by=resolved_by,
            included_via="pass1",
            banner=banner_name,
            segments_shown=[str(s) for s in (segments_shown or [])],
        ))

        excel_data_blocks.append({
            "excel_id": excel_ds.id,
            "question_code": excel_ds.question_code,
            "question_text": excel_ds.question_text,
            "view_type": view_type,
            "data": data_cols,
            "derived_metrics": dm,
            "base_n": excel_ds.base_n,
            "segment_sizes": seg_sizes,
        })

    if not excel_data_blocks:
        # Matched IDs couldn't resolve — still do lite check (spelling/visual)
        lite = _verify_lite(slide_number, texts, image_path, client, verbose=verbose)
        unresolved_text = "; ".join(unresolved_matches[:3])
        if len(unresolved_matches) > 3:
            unresolved_text += f" (+{len(unresolved_matches) - 3} više)"
        detail = f" Neuspješni match detalji: {unresolved_text}." if unresolved_text else ""
        lite.summary = ("Pass 1 matchao ali Excel podaci nisu pronađeni." + detail + " "
                        + (lite.summary or ""))
        lite.pass1_slide_type = slide_type
        lite.pass1_total_candidates = len(candidates)
        lite.pass1_total_datasets = len(pass1_datasets)
        lite.pass1_confident_datasets = len(confident_matches)
        lite.match_sources = match_sources
        lite.match_failures = match_failures
        return lite

    # Expand context: if one dataset from a known family matched, include siblings too.
    selected_ids = {b["excel_id"] for b in excel_data_blocks}
    extra_blocks: list[dict] = []
    candidate_pool = list(candidate_by_id.values())
    for block in list(excel_data_blocks):
        seed_ds = excel_by_id.get(block.get("excel_id", ""))
        if not seed_ds:
            continue
        for rel_ds in _find_related_family_datasets(seed_ds, candidate_pool):
            if rel_ds.id in selected_ids:
                continue

            dm = {}
            if rel_ds.derived_metrics:
                dm = {
                    "mean": rel_ds.derived_metrics.mean,
                    "top2box": rel_ds.derived_metrics.top2box,
                    "net": rel_ds.derived_metrics.net,
                }

            extra_blocks.append({
                "excel_id": rel_ds.id,
                "question_code": rel_ds.question_code,
                "question_text": rel_ds.question_text,
                "view_type": "derived",
                "data": {
                    "Total": [
                        {"label": c.label, "value": c.total}
                        for c in rel_ds.categories
                    ]
                },
                "derived_metrics": dm,
                "base_n": rel_ds.base_n,
                "segment_sizes": {},
            })
            match_sources.append(MatchSource(
                excel_id=rel_ds.id,
                question_code=rel_ds.question_code,
                question_text=rel_ds.question_text,
                view_type="derived",
                pass1_excel_id="",
                pass1_question_code="",
                confidence=None,
                resolved_by="family_expansion",
                included_via="related_family",
                banner=None,
                segments_shown=[],
            ))
            selected_ids.add(rel_ds.id)

    if extra_blocks:
        with _print_lock:
            print(f"    INFO slajd {slide_number}: dodano {len(extra_blocks)} "
                  f"povezanih tablica u Pass 2 context")
        excel_data_blocks.extend(extra_blocks)

    if verbose:
        with _print_lock:
            print(f"      [slajd {slide_number}] PASS 2 start: {len(excel_data_blocks)} data blokova")

    # ── PASS 2: Targeted Verification ──
    pass2_prompt = build_pass2_prompt(slide_number, texts, excel_data_blocks)
    raw_pass2 = client.call_pro_multimodal([pass2_prompt, image_part])

    try:
        data = parse_json_response(raw_pass2)
    except ValueError:
        return SlideVerification(
            slide_number=slide_number,
            overall_status="error",
            pass1_slide_type=slide_type,
            pass1_total_candidates=len(candidates),
            pass1_total_datasets=len(pass1_datasets),
            pass1_confident_datasets=len(confident_matches),
            match_sources=match_sources,
            match_failures=match_failures,
            data_issues=[DataIssue(
                severity="error", issue_type="parse_error",
                detail=f"Pass 2 LLM odgovor nije validan JSON: {raw_pass2[:200]}",
            )],
            summary="Greška u parsiranju Pass 2 odgovora.",
        )

    # Parse verification result
    data_issues = [
        DataIssue(
            severity=d.get("severity") or "info",
            issue_type=d.get("issue_type") or "unknown",
            detail=d.get("detail") or "",
            slide_value=d.get("slide_value"),
            excel_value=d.get("excel_value"),
            dataset_index=d.get("dataset_index"),
        )
        for d in data.get("data_issues", [])
    ]
    text_issues = [
        TextIssue(
            severity=t.get("severity") or "info",
            issue_type=t.get("issue_type") or "unknown",
            detail=t.get("detail") or "",
        )
        for t in data.get("text_issues", [])
    ]
    visual_issues = [
        VisualIssue(
            severity=v.get("severity") or "info",
            issue_type=v.get("issue_type") or "unknown",
            detail=v.get("detail") or "",
        )
        for v in data.get("visual_issues", [])
    ]

    return SlideVerification(
        slide_number=slide_number,
        overall_status=data.get("overall_status") or "ok",
        pass1_slide_type=slide_type,
        pass1_total_candidates=len(candidates),
        pass1_total_datasets=len(pass1_datasets),
        pass1_confident_datasets=len(confident_matches),
        data_issues=data_issues,
        text_issues=text_issues,
        visual_issues=visual_issues,
        match_sources=match_sources,
        match_failures=match_failures,
        summary=data.get("summary", ""),
    )


# ──────────────────────────────────────────────────────────────────────
# Lite verification (Flash, no data check)
# ──────────────────────────────────────────────────────────────────────

def _verify_lite(
    slide_number: int,
    texts: list[str],
    image_path: Path,
    client: GeminiClient,
    *,
    verbose: bool = False,
) -> SlideVerification:
    """Lite verification using Flash — no data check, just spelling/visual/consistency."""
    if verbose:
        with _print_lock:
            print(f"      [slajd {slide_number}] LITE verify (Flash)")

    prompt = build_lite_verification_prompt(slide_number, texts)
    image_part = client.make_image_part(image_path)

    raw_response = client.call_flash_multimodal([prompt, image_part])

    try:
        data = parse_json_response(raw_response)
    except ValueError:
        return SlideVerification(
            slide_number=slide_number,
            overall_status="info",
            summary="Lite provjera — LLM odgovor nije parsiran.",
        )

    text_issues = [
        TextIssue(
            severity=t.get("severity") or "info",
            issue_type=t.get("issue_type") or "unknown",
            detail=t.get("detail") or "",
        )
        for t in data.get("text_issues", [])
    ]
    visual_issues = [
        VisualIssue(
            severity=v.get("severity") or "info",
            issue_type=v.get("issue_type") or "unknown",
            detail=v.get("detail") or "",
        )
        for v in data.get("visual_issues", [])
    ]

    return SlideVerification(
        slide_number=slide_number,
        overall_status=data.get("overall_status") or "ok",
        data_issues=[],
        text_issues=text_issues,
        visual_issues=visual_issues,
        summary=data.get("summary", ""),
    )
