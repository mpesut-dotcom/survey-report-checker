"""Pydantic models for Excel dataset structures."""
from pydantic import BaseModel, Field


class CategoryBreakdown(BaseModel):
    """Single category/option from an Excel question."""
    label: str
    total: float | None = None
    breakdowns: dict[str, dict[str, float | None]] = Field(default_factory=dict)
    # breakdowns[banner_name][segment_name] = value
    # e.g. {"SPOL (Q1)": {"Muškarac": 0.7, "Žena": 0.3}, "DOB (Q2)": {"18-29": 0.8, ...}}


class DerivedMetrics(BaseModel):
    """Derived/aggregated metrics for a question."""
    mean: float | None = None
    top2box: float | None = None
    bottom2box: float | None = None
    net: float | None = None


class ExcelDataset(BaseModel):
    """One parsed question/table from Excel."""
    id: str = ""
    file_key: str = ""
    sheet_name: str
    question_code: str
    question_text: str
    base_description: str | None = None
    base_n: int | None = None
    type: str = "single_choice"  # single_choice | multi_choice | scale | open | grid
    categories: list[CategoryBreakdown] = Field(default_factory=list)
    derived_metrics: DerivedMetrics | None = Field(default_factory=DerivedMetrics)
    segment_sizes: dict[str, dict[str, int]] = Field(default_factory=dict)
    # segment_sizes[banner_name][segment_name] = N
    # e.g. {"SPOL (Q1)": {"Muškarac": 285, "Žena": 315}, ...}

    # Internal tracking
    source_file: str = ""
    excel_row: int = 0
