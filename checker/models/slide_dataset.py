"""Pydantic models for slide extraction structures."""
from pydantic import BaseModel, Field


class DataPoint(BaseModel):
    """Single data point extracted from a slide."""
    label: str
    value: float | None = None
    unit: str = "%"  # %, abs, mean, index
    group: str | None = None


class SlideDataset(BaseModel):
    """One logical dataset extracted from a slide by LLM."""
    title: str = ""
    question_code: str | None = None
    question_text: str | None = None
    base_description: str | None = None
    base_n: int | None = None
    chart_type: str = "unknown"
    data_points: list[DataPoint] = Field(default_factory=list)
    unit: str | None = "%"
    time_period: str | None = None
    subset: str | None = None
    series_name: str | None = None
    note: str | None = None
    derived_metrics: dict[str, float | None] = Field(default_factory=dict)
    flags: list[str] = Field(default_factory=list)

    # Match fields (populated during merged extract+match phase)
    matched_excel_id: str | None = None
    confidence: float = 0.0
    match_reasoning: str = ""


class TextElement(BaseModel):
    """Non-data text element on a slide."""
    type: str  # title, subtitle, footnote, comment, callout
    content: str


class SlideExtraction(BaseModel):
    """Full extraction result for one slide."""
    slide_number: int
    slide_title: str = ""
    datasets: list[SlideDataset] = Field(default_factory=list)
    text_elements: list[TextElement] = Field(default_factory=list)
    raw_texts: list[str] = Field(default_factory=list)  # All text from pptx shapes
