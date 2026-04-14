"""Pydantic models for matching results."""
from pydantic import BaseModel, Field


class MatchResult(BaseModel):
    """One slide dataset → Excel dataset match."""
    slide_number: int
    dataset_index: int
    slide_question: str = ""
    matched_excel_id: str | None = None
    matched_excel_ids: list[str] = Field(default_factory=list)
    confidence: float = 0.0
    match_reasoning: str = ""
    unmatched: bool = False


class MatchingOutput(BaseModel):
    """Full output of the matching phase."""
    matches: list[MatchResult] = Field(default_factory=list)
    unmatched_datasets: int = 0
