from __future__ import annotations

import re
import unicodedata

import pandas as pd

from src.domain.models import PrefixMode, SheetConfig

AZURE_COLUMNS = [
    "Work Item Type",
    "Title",
    "Description",
    "Original Estimate",
    "Remaining Work",
    "Activity",
]

def to_float_or_blank(value: object) -> float | str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return ""
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return ""


def infer_activity(title: str, default_activity: str = "Development") -> str:
    if not isinstance(title, str):
        return default_activity

    normalized = title.lower()
    rules = [
        (r"\b(prueba|testing|qa|test|e2e)\b", "Testing"),
        (r"\b(doc|documentaci\w*|handoff)\b", "Documentation"),
        (r"\b(deploy|release|desplieg)\b", "Deployment"),
        (r"\b(spike|research|investig)\b", "Other"),
    ]
    for pattern, activity in rules:
        if re.search(pattern, normalized):
            return activity
    return default_activity


def guess_column_by_candidates(columns: list[str], candidates: list[str]) -> str | None:
    lower_map = {col.lower(): col for col in columns}
    for candidate in candidates:
        for col_lower, original in lower_map.items():
            if candidate in col_lower:
                return original
    return None


def _build_prefixed_title(raw_title: str, config: SheetConfig) -> str:
    title = raw_title.strip()
    if not title:
        return ""

    prefix = ""
    if config.prefix_mode == PrefixMode.SHEET_NAME:
        prefix = config.sheet_name.strip()
    elif config.prefix_mode == PrefixMode.CUSTOM:
        prefix = config.custom_prefix.strip()

    if prefix:
        return f"{prefix} - {title}"
    return title


def _safe_series(df: pd.DataFrame, column_name: str | None) -> pd.Series:
    if not column_name or column_name not in df.columns:
        return pd.Series([""] * len(df), index=df.index)
    return df[column_name].fillna("").astype(str).str.strip()


def _normalize_text(value: str) -> str:
    normalized = str(value).strip().lower()
    normalized = unicodedata.normalize("NFD", normalized)
    normalized = "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def _is_summary_token(value: str) -> bool:
    normalized = _normalize_text(value)
    if not normalized:
        return False
    return normalized in {"total", "subtotal", "resumen"}


def _is_summary_from_final_title(value: str) -> bool:
    normalized = _normalize_text(value)
    if not normalized:
        return False
    blocks = [block.strip() for block in re.split(r"[-:|]", normalized) if block.strip()]
    if not blocks:
        return False
    return _is_summary_token(blocks[-1])


def build_tasks_from_sheet(
    df: pd.DataFrame,
    config: SheetConfig,
    default_work_item_type: str,
) -> pd.DataFrame:
    title_series = _safe_series(df, config.mapping.title_col)
    description_series = _safe_series(df, config.mapping.description_col)
    hours_series = _safe_series(df, config.mapping.hours_col)
    activity_series = _safe_series(df, config.mapping.activity_col)
    keep_raw_mask = (
        ~title_series.apply(_is_summary_token)
        if config.exclude_summary_rows
        else pd.Series([True] * len(df), index=df.index)
    )

    output = pd.DataFrame(index=df.index)
    output["Work Item Type"] = default_work_item_type
    output["Title"] = title_series.apply(lambda title: _build_prefixed_title(title, config))
    output["Description"] = description_series

    parsed_hours = hours_series.apply(to_float_or_blank)
    output["Original Estimate"] = parsed_hours
    output["Remaining Work"] = parsed_hours
    output["Activity"] = [
        activity if activity else infer_activity(title)
        for title, activity in zip(output["Title"], activity_series)
    ]

    output = output[keep_raw_mask].copy()
    output = output[output["Title"].astype(str).str.len() > 0].copy()
    if config.exclude_summary_rows:
        output = output[~output["Title"].apply(_is_summary_from_final_title)].copy()
    return output[AZURE_COLUMNS]
