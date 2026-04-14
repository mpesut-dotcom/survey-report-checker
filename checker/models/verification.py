"""Pydantic models for verification results."""
from pydantic import BaseModel, Field


class DataIssue(BaseModel):
    """A data-related issue found during verification."""
    severity: str  # error, warning, info
    issue_type: str = "unknown"  # wrong_value, wrong_base, wrong_label, missing_category, wrong_calculation, wrong_order, sum_mismatch
    dataset_question: str = ""
    detail: str = ""
    slide_value: float | str | None = None
    excel_value: float | str | None = None
    dataset_index: int | None = None
    location: str = ""


class TextIssue(BaseModel):
    """A text-related issue found during verification."""
    severity: str  # error, warning
    issue_type: str = "unknown"  # spelling, grammar, terminology, factual_claim
    detail: str = ""
    text: str = ""
    correction: str = ""
    location: str = ""


class VisualIssue(BaseModel):
    """A visual consistency issue."""
    severity: str  # warning, info
    issue_type: str = "unknown"  # chart_proportion, legend_mismatch, axis_error
    detail: str = ""


class MatchSource(BaseModel):
    """Resolved data source used for pass 2 verification."""
    excel_id: str = ""
    question_code: str = ""
    question_text: str = ""
    view_type: str = ""
    pass1_excel_id: str = ""
    pass1_question_code: str = ""
    confidence: float | None = None
    resolved_by: str = ""
    included_via: str = "pass1"  # pass1 | related_family
    banner: str | None = None
    segments_shown: list[str] = Field(default_factory=list)


class MatchFailure(BaseModel):
    """Pass 1 match that could not be used in pass 2."""
    excel_id: str = ""
    question_code: str = ""
    confidence: float | None = None
    reason: str = ""
    description: str = ""


class SlideVerification(BaseModel):
    """Full verification result for one slide."""
    slide_number: int
    overall_status: str = "ok"  # ok, warning, error
    pass1_slide_type: str = "unknown"
    pass1_total_candidates: int = 0
    pass1_total_datasets: int = 0
    pass1_confident_datasets: int = 0
    data_issues: list[DataIssue] = Field(default_factory=list)
    text_issues: list[TextIssue] = Field(default_factory=list)
    visual_issues: list[VisualIssue] = Field(default_factory=list)
    match_sources: list[MatchSource] = Field(default_factory=list)
    match_failures: list[MatchFailure] = Field(default_factory=list)
    summary: str = ""
