from dataclasses import dataclass, field
from typing import Optional


@dataclass
class WordItem:
    page_index: int
    page_number: int
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    block_no: int
    line_no: int
    word_no: int

    @property
    def y_center(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class CodeOccurrence:
    code: str
    page_number: int
    x0: float
    y0: float
    x1: float
    y1: float
    nearby_amount_text: str = ""
    nearby_context: str = ""
    note: str = ""
    confidence_score: int = 0
    selected: bool = False

    @property
    def location_text(self) -> str:
        return f"Page {self.page_number}, x={round(self.x0, 2)}, y={round(self.y0, 2)}"


@dataclass
class FieldSpec:
    expression: str
    label: str
    excel_cell: str
    codes: list[str] = field(default_factory=list)


@dataclass
class FieldResult:
    expression: str
    label: str
    excel_cell: str
    display_value: str
    numeric_value: Optional[float]
    should_write_blank: bool
    status: str
    difficulty: str
    notes: str
    component_details: str
