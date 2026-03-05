from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class PrefixMode(str, Enum):
    NONE = "none"
    SHEET_NAME = "sheet_name"
    CUSTOM = "custom"


@dataclass(slots=True)
class ColumnMapping:
    title_col: str | None
    description_col: str | None
    hours_col: str | None
    activity_col: str | None


@dataclass(slots=True)
class SheetConfig:
    sheet_name: str
    enabled: bool
    header_row: int
    prefix_mode: PrefixMode
    custom_prefix: str
    exclude_summary_rows: bool
    mapping: ColumnMapping
