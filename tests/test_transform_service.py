import pandas as pd

from src.application.transform_service import (
    build_tasks_from_sheet,
    infer_activity,
    to_float_or_blank,
)
from src.domain.models import ColumnMapping, PrefixMode, SheetConfig


def _sheet_config(prefix_mode: PrefixMode = PrefixMode.NONE, custom_prefix: str = "") -> SheetConfig:
    return SheetConfig(
        sheet_name="Sprint 1",
        enabled=True,
        header_row=1,
        prefix_mode=prefix_mode,
        custom_prefix=custom_prefix,
        exclude_summary_rows=True,
        mapping=ColumnMapping(
            title_col="titulo",
            description_col="descripcion",
            hours_col="horas",
            activity_col="actividad",
        ),
    )


def test_to_float_or_blank_parses_common_formats() -> None:
    assert to_float_or_blank("1,5") == 1.5
    assert to_float_or_blank("2.25") == 2.25
    assert to_float_or_blank("") == ""
    assert to_float_or_blank("abc") == ""


def test_infer_activity_uses_keywords() -> None:
    assert infer_activity("Hacer prueba e2e") == "Testing"
    assert infer_activity("Escribir documentación") == "Documentation"
    assert infer_activity("Sin keyword") == "Development"


def test_build_tasks_from_sheet_applies_custom_prefix() -> None:
    df = pd.DataFrame(
        {
            "titulo": ["Crear endpoint"],
            "descripcion": ["detalle"],
            "horas": ["1,5"],
            "actividad": [""],
        }
    )
    cfg = _sheet_config(prefix_mode=PrefixMode.CUSTOM, custom_prefix="API")
    out = build_tasks_from_sheet(df, cfg, default_work_item_type="Task")

    assert len(out) == 1
    assert out.iloc[0]["Title"] == "API - Crear endpoint"
    assert out.iloc[0]["Original Estimate"] == 1.5
    assert out.iloc[0]["Remaining Work"] == 1.5


def test_build_tasks_from_sheet_removes_empty_titles() -> None:
    df = pd.DataFrame(
        {
            "titulo": ["", "Válido"],
            "descripcion": ["desc 1", "desc 2"],
            "horas": [1, 2],
            "actividad": ["", ""],
        }
    )
    cfg = _sheet_config()
    out = build_tasks_from_sheet(df, cfg, default_work_item_type="Task")

    assert len(out) == 1
    assert out.iloc[0]["Title"] == "Válido"


def test_build_tasks_from_sheet_excludes_summary_titles_by_default() -> None:
    df = pd.DataFrame(
        {
            "titulo": ["Total", " Subtotal ", "Resúmen", "Implementar login"],
            "descripcion": ["", "", "", ""],
            "horas": [1, 2, 3, 4],
            "actividad": ["", "", "", ""],
        }
    )
    cfg = _sheet_config()
    out = build_tasks_from_sheet(df, cfg, default_work_item_type="Task")

    assert len(out) == 1
    assert out.iloc[0]["Title"] == "Implementar login"


def test_build_tasks_from_sheet_keeps_non_exact_total_phrases() -> None:
    df = pd.DataFrame(
        {
            "titulo": ["Exponer total de pedidos creados", "Implementar login"],
            "descripcion": ["", ""],
            "horas": [1, 4],
            "actividad": ["", ""],
        }
    )
    cfg = _sheet_config()
    out = build_tasks_from_sheet(df, cfg, default_work_item_type="Task")

    assert len(out) == 2


def test_build_tasks_from_sheet_excludes_prefixed_legacy_total() -> None:
    df = pd.DataFrame(
        {
            "titulo": ["Modulo A - Total", "Implementar login"],
            "descripcion": ["", ""],
            "horas": [1, 4],
            "actividad": ["", ""],
        }
    )
    cfg = _sheet_config()
    out = build_tasks_from_sheet(df, cfg, default_work_item_type="Task")

    assert len(out) == 1
    assert out.iloc[0]["Title"] == "Implementar login"


def test_build_tasks_from_sheet_keeps_summary_titles_when_disabled() -> None:
    df = pd.DataFrame(
        {
            "titulo": ["Total", "Implementar login"],
            "descripcion": ["", ""],
            "horas": [1, 4],
            "actividad": ["", ""],
        }
    )
    cfg = _sheet_config()
    cfg.exclude_summary_rows = False
    out = build_tasks_from_sheet(df, cfg, default_work_item_type="Task")

    assert len(out) == 2
