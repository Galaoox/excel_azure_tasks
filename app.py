from __future__ import annotations

from dataclasses import asdict
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st

from config.task_types import TASK_TYPES
from src.application.transform_service import (
    AZURE_COLUMNS,
    build_tasks_from_sheet,
    guess_column_by_candidates,
)
from src.domain.models import ColumnMapping, PrefixMode, SheetConfig
from src.infrastructure.csv_writer import dataframe_to_azure_csv_bytes
from src.infrastructure.excel_reader import ExcelReader
from src.infrastructure.profile_store import ProfileStore


LANGUAGE_LABELS = {"Español": "es", "English": "en"}

TRANSLATIONS = {
    "title": {"es": "Excel a CSV para Azure Tasks", "en": "Excel to Azure Tasks CSV"},
    "subtitle": {
        "es": "Convierte hojas de Excel en un CSV compatible con carga masiva de Azure.",
        "en": "Convert Excel sheets into an Azure bulk import compatible CSV.",
    },
    "language": {"es": "Idioma", "en": "Language"},
    "upload": {"es": "Sube un archivo .xlsx", "en": "Upload an .xlsx file"},
    "no_file": {"es": "Sube un Excel para comenzar.", "en": "Upload an Excel file to start."},
    "sheet_select": {"es": "Hojas a procesar", "en": "Sheets to process"},
    "sheet_config": {"es": "Configuración por hoja", "en": "Sheet configuration"},
    "sheet_tab": {"es": "Hoja", "en": "Sheet"},
    "sheet_settings": {"es": "Configuración de hoja", "en": "Sheet settings"},
    "column_mapping_section": {"es": "Mapeo de columnas", "en": "Column mapping"},
    "header_row": {"es": "Fila de encabezado", "en": "Header row"},
    "prefix_mode": {"es": "Modo de prefijo", "en": "Prefix mode"},
    "prefix_custom": {"es": "Prefijo personalizado", "en": "Custom prefix"},
    "prefix_none": {"es": "Sin prefijo", "en": "No prefix"},
    "prefix_sheet_name": {"es": "Nombre de hoja", "en": "Sheet name"},
    "prefix_custom_mode": {"es": "Custom", "en": "Custom"},
    "exclude_summary_rows": {
        "es": "Omitir filas Total/Subtotal/Resumen",
        "en": "Exclude Total/Subtotal/Summary rows",
    },
    "col_map": {"es": "Mapeo de columnas", "en": "Column mapping"},
    "map_title": {"es": "Título", "en": "Title"},
    "map_description": {"es": "Descripción", "en": "Description"},
    "map_hours": {"es": "Horas", "en": "Hours"},
    "map_activity": {"es": "Activity", "en": "Activity"},
    "none": {"es": "(ninguna)", "en": "(none)"},
    "build_table": {"es": "Construir tabla de tareas", "en": "Build task table"},
    "tasks_ready": {"es": "Tabla consolidada lista", "en": "Consolidated table ready"},
    "search": {"es": "Buscar en título", "en": "Search in title"},
    "bulk_type": {"es": "Tipo de tarea masivo (filtradas)", "en": "Bulk task type (filtered)"},
    "apply_bulk": {"es": "Aplicar tipo a filtradas", "en": "Apply type to filtered"},
    "download": {"es": "Descargar CSV Azure", "en": "Download Azure CSV"},
    "profiles": {"es": "Perfiles de mapeo", "en": "Mapping profiles"},
    "profile_name": {"es": "Nombre del perfil", "en": "Profile name"},
    "save_profile": {"es": "Guardar perfil", "en": "Save profile"},
    "load_profile": {"es": "Cargar perfil", "en": "Load profile"},
    "rows_visible": {"es": "Filas visibles", "en": "Visible rows"},
    "total_rows": {"es": "Total filas", "en": "Total rows"},
    "default_type": {"es": "Tipo de tarea por defecto", "en": "Default Work Item Type"},
    "select_one_sheet": {"es": "Selecciona al menos una hoja.", "en": "Select at least one sheet."},
    "profile_saved": {"es": "Perfil guardado", "en": "Profile saved"},
    "profile_loaded": {"es": "Perfil cargado", "en": "Profile loaded"},
    "build_section": {"es": "Construcción", "en": "Build"},
    "include_in_build": {"es": "Incluir en construcción", "en": "Include in build"},
    "select_all": {"es": "Seleccionar todas", "en": "Select all"},
    "clear_selection": {"es": "Limpiar selección", "en": "Clear selection"},
    "no_sheet_selected_build": {
        "es": "Selecciona al menos una hoja para construir.",
        "en": "Select at least one sheet to build.",
    },
    "source_sheet_mapping": {"es": "Hoja origen de mapeo", "en": "Mapping source sheet"},
    "apply_mapping_others": {"es": "Aplicar mapeo a otras hojas", "en": "Apply mapping to other sheets"},
    "mapping_applied": {"es": "Mapeo aplicado", "en": "Mapping applied"},
    "mapping_skipped": {"es": "Mapeo omitido", "en": "Mapping skipped"},
    "built_with_sheets": {"es": "Hojas usadas", "en": "Sheets used"},
    "help_header_row": {
        "es": "Fila donde están los nombres de columnas en esta hoja.",
        "en": "Row containing column names for this sheet.",
    },
    "help_prefix_mode": {
        "es": "Define cómo se construye el prefijo del título.",
        "en": "Defines how title prefix is generated.",
    },
    "help_custom_prefix": {
        "es": "Texto que se antepone al título cuando el modo es Custom.",
        "en": "Text added before title when mode is Custom.",
    },
    "help_map_title": {
        "es": "Columna del Excel que representa el título de la tarea.",
        "en": "Excel column used as task title.",
    },
    "help_map_description": {
        "es": "Columna del Excel que representa la descripción de la tarea.",
        "en": "Excel column used as task description.",
    },
    "help_map_hours": {
        "es": "Columna del Excel con horas estimadas.",
        "en": "Excel column with estimated hours.",
    },
    "help_map_activity": {
        "es": "Columna de actividad. Si queda vacía se infiere automáticamente.",
        "en": "Activity column. If empty, it is inferred automatically.",
    },
    "help_include_sheet": {
        "es": "Si está activa, la hoja se incluye al construir la tabla final.",
        "en": "When enabled, this sheet is included in final table build.",
    },
    "help_source_sheet_mapping": {
        "es": "Hoja de referencia cuyo mapeo se intentará copiar al resto.",
        "en": "Reference sheet whose mapping will be copied to others.",
    },
    "help_exclude_summary_rows": {
        "es": "Filtra filas con título de resumen (total, subtotal o resumen).",
        "en": "Filters summary title rows (total, subtotal, or summary).",
    },
    "build_summary": {"es": "Resumen de construcción", "en": "Build summary"},
    "build_sheet_col": {"es": "Hoja", "en": "Sheet"},
    "build_rows_col": {"es": "Tareas generadas", "en": "Generated tasks"},
    "warning_missing_title_map": {
        "es": "No se construyó por mapeo de título faltante o inválido",
        "en": "Skipped due to missing or invalid title mapping",
    },
    "warning_zero_rows": {
        "es": "No generó tareas (revisa títulos vacíos o mapeo)",
        "en": "Generated 0 tasks (check empty titles or mapping)",
    },
}


TITLE_CANDIDATES = ["titulo", "título", "title", "task"]
DESCRIPTION_CANDIDATES = ["descripcion", "descripción", "description", "detalle"]
HOURS_CANDIDATES = ["estimacion", "estimación", "estimate", "hours", "horas"]
ACTIVITY_CANDIDATES = ["activity", "actividad"]


def t(key: str, lang: str) -> str:
    return TRANSLATIONS.get(key, {}).get(lang, key)


def ensure_state() -> None:
    if "sheet_configs" not in st.session_state:
        st.session_state.sheet_configs = {}
    if "tasks_df" not in st.session_state:
        st.session_state.tasks_df = None
    if "last_file_key" not in st.session_state:
        st.session_state.last_file_key = None
    if "pending_mapping_sync_targets" not in st.session_state:
        st.session_state.pending_mapping_sync_targets = []
    if "mapping_feedback_success" not in st.session_state:
        st.session_state.mapping_feedback_success = None
    if "mapping_feedback_warnings" not in st.session_state:
        st.session_state.mapping_feedback_warnings = []
    if "build_summary_rows" not in st.session_state:
        st.session_state.build_summary_rows = []
    if "build_summary_warnings" not in st.session_state:
        st.session_state.build_summary_warnings = []


def _clear_sheet_widget_state() -> None:
    prefixes = (
        "header_row_",
        "prefix_mode_",
        "custom_prefix_",
        "title_col_",
        "description_col_",
        "hours_col_",
        "activity_col_",
        "include_sheet_",
        "exclude_summary_rows_",
    )
    for key in list(st.session_state.keys()):
        if key.startswith(prefixes):
            del st.session_state[key]


def _apply_pending_mapping_widget_sync() -> None:
    targets: list[str] = list(st.session_state.get("pending_mapping_sync_targets", []))
    if not targets:
        return
    for sheet_name in targets:
        for key_prefix in ("title_col_", "description_col_", "hours_col_", "activity_col_"):
            key = f"{key_prefix}{sheet_name}"
            if key in st.session_state:
                del st.session_state[key]
    st.session_state.pending_mapping_sync_targets = []


def _default_mapping_from_columns(columns: list[str]) -> dict[str, str | None]:
    return {
        "title_col": guess_column_by_candidates(columns, TITLE_CANDIDATES),
        "description_col": guess_column_by_candidates(columns, DESCRIPTION_CANDIDATES),
        "hours_col": guess_column_by_candidates(columns, HOURS_CANDIDATES),
        "activity_col": guess_column_by_candidates(columns, ACTIVITY_CANDIDATES),
    }


def _ensure_sheet_configs_initialized(reader: ExcelReader, selected_sheets: list[str]) -> None:
    for sheet_name in selected_sheets:
        existing = _config_from_state(sheet_name)
        header_row = int(existing.get("header_row", 1))
        preview_df = reader.read_sheet(sheet_name, header_row)
        columns = [str(col).strip() for col in preview_df.columns]
        default_mapping = _default_mapping_from_columns(columns)
        existing_mapping = existing.get("mapping", {})

        normalized_cfg = {
            "sheet_name": sheet_name,
            "enabled": bool(existing.get("enabled", True)),
            "header_row": header_row,
            "prefix_mode": existing.get("prefix_mode", PrefixMode.NONE.value),
            "custom_prefix": str(existing.get("custom_prefix", "")),
            "exclude_summary_rows": bool(existing.get("exclude_summary_rows", True)),
            "mapping": {
                "title_col": existing_mapping.get("title_col", default_mapping["title_col"]),
                "description_col": existing_mapping.get(
                    "description_col", default_mapping["description_col"]
                ),
                "hours_col": existing_mapping.get("hours_col", default_mapping["hours_col"]),
                "activity_col": existing_mapping.get(
                    "activity_col", default_mapping["activity_col"]
                ),
            },
        }
        st.session_state.sheet_configs[sheet_name] = normalized_cfg


def _safe_options(columns: list[str], none_label: str) -> list[str]:
    return [none_label, *columns]


def _default_choice(options: list[str], preferred: str | None, fallback: str) -> str:
    if preferred and preferred in options:
        return preferred
    return fallback if fallback in options else options[0]


def _config_from_state(sheet_name: str) -> dict[str, Any]:
    return st.session_state.sheet_configs.get(sheet_name, {})


def _update_sheet_config(sheet_name: str, config: dict[str, Any]) -> None:
    st.session_state.sheet_configs[sheet_name] = config


def _render_profile_section(
    profile_store: ProfileStore,
    selected_sheets: list[str],
    lang: str,
) -> None:
    st.subheader(t("profiles", lang))
    profile_name = st.text_input(t("profile_name", lang), key="profile_name_input")
    if st.button(t("save_profile", lang), use_container_width=True):
        if profile_name.strip():
            payload = {
                "version": 1,
                "sheets": {
                    sheet: _config_from_state(sheet) for sheet in selected_sheets
                },
            }
            profile_store.save_profile(profile_name.strip(), payload)
            st.success(f"{t('profile_saved', lang)}: '{profile_name.strip()}'.")

    profiles = profile_store.list_profiles()
    if profiles:
        to_load = st.selectbox(t("load_profile", lang), profiles, key="profile_select")
        if st.button(t("load_profile", lang), key="btn_load_profile", use_container_width=True):
            data = profile_store.load_profile(to_load)
            sheets_data = data.get("sheets", {})
            for sheet_name, cfg in sheets_data.items():
                st.session_state.sheet_configs[sheet_name] = cfg
            st.success(f"{t('profile_loaded', lang)}: '{to_load}'.")


def _build_sheet_config_ui(
    sheet_name: str,
    preview_df: pd.DataFrame,
    lang: str,
) -> SheetConfig:
    existing = _config_from_state(sheet_name)
    none_label = t("none", lang)
    custom_prefix_key = f"custom_prefix_{sheet_name}"

    st.markdown(f"### {sheet_name}")
    with st.expander(t("sheet_settings", lang), expanded=True):
        header_row = st.number_input(
            f"{t('header_row', lang)} ({sheet_name})",
            min_value=1,
            max_value=100,
            value=int(existing.get("header_row", 1)),
            step=1,
            key=f"header_row_{sheet_name}",
            help=t("help_header_row", lang),
        )

        prefix_labels = {
            PrefixMode.NONE.value: t("prefix_none", lang),
            PrefixMode.SHEET_NAME.value: t("prefix_sheet_name", lang),
            PrefixMode.CUSTOM.value: t("prefix_custom_mode", lang),
        }
        prefix_modes = [PrefixMode.NONE.value, PrefixMode.SHEET_NAME.value, PrefixMode.CUSTOM.value]
        prefix_mode = st.selectbox(
            f"{t('prefix_mode', lang)} ({sheet_name})",
            prefix_modes,
            format_func=lambda val: prefix_labels[val],
            index=prefix_modes.index(existing.get("prefix_mode", PrefixMode.NONE.value)),
            key=f"prefix_mode_{sheet_name}",
            help=t("help_prefix_mode", lang),
        )

        # Initialize custom prefix once per sheet:
        # - use value from profile/state when available
        # - otherwise default to sheet name
        if custom_prefix_key not in st.session_state:
            seeded_custom_prefix = existing.get("custom_prefix")
            st.session_state[custom_prefix_key] = (
                seeded_custom_prefix if seeded_custom_prefix is not None else sheet_name
            )

        custom_prefix = existing.get("custom_prefix", "")
        if prefix_mode == PrefixMode.CUSTOM.value:
            custom_prefix = st.text_input(
                f"{t('prefix_custom', lang)} ({sheet_name})",
                key=custom_prefix_key,
                help=t("help_custom_prefix", lang),
            )
        else:
            # Keep current value in session state so user edits persist if mode toggles.
            custom_prefix = st.session_state.get(custom_prefix_key, custom_prefix)

        exclude_summary_rows = st.checkbox(
            f"{t('exclude_summary_rows', lang)} ({sheet_name})",
            value=bool(existing.get("exclude_summary_rows", True)),
            key=f"exclude_summary_rows_{sheet_name}",
            help=t("help_exclude_summary_rows", lang),
        )

    columns = [str(col).strip() for col in preview_df.columns]
    options = _safe_options(columns, none_label)
    mapping_cfg = existing.get("mapping", {})

    default_title = guess_column_by_candidates(columns, TITLE_CANDIDATES)
    default_description = guess_column_by_candidates(columns, DESCRIPTION_CANDIDATES)
    default_hours = guess_column_by_candidates(columns, HOURS_CANDIDATES)
    default_activity = guess_column_by_candidates(columns, ACTIVITY_CANDIDATES)

    with st.expander(t("column_mapping_section", lang), expanded=True):
        st.markdown(f"**{t('col_map', lang)} ({sheet_name})**")
        title_col = st.selectbox(
            t("map_title", lang),
            options,
            index=options.index(
                _default_choice(options, mapping_cfg.get("title_col"), default_title or none_label)
            ),
            key=f"title_col_{sheet_name}",
            help=t("help_map_title", lang),
        )
        description_col = st.selectbox(
            t("map_description", lang),
            options,
            index=options.index(
                _default_choice(
                    options,
                    mapping_cfg.get("description_col"),
                    default_description or none_label,
                )
            ),
            key=f"description_col_{sheet_name}",
            help=t("help_map_description", lang),
        )
        hours_col = st.selectbox(
            t("map_hours", lang),
            options,
            index=options.index(
                _default_choice(options, mapping_cfg.get("hours_col"), default_hours or none_label)
            ),
            key=f"hours_col_{sheet_name}",
            help=t("help_map_hours", lang),
        )
        activity_col = st.selectbox(
            t("map_activity", lang),
            options,
            index=options.index(
                _default_choice(
                    options,
                    mapping_cfg.get("activity_col"),
                    default_activity or none_label,
                )
            ),
            key=f"activity_col_{sheet_name}",
            help=t("help_map_activity", lang),
        )

    col_mapping = ColumnMapping(
        title_col=None if title_col == none_label else title_col,
        description_col=None if description_col == none_label else description_col,
        hours_col=None if hours_col == none_label else hours_col,
        activity_col=None if activity_col == none_label else activity_col,
    )

    config = SheetConfig(
        sheet_name=sheet_name,
        enabled=bool(existing.get("enabled", True)),
        header_row=int(header_row),
        prefix_mode=PrefixMode(prefix_mode),
        custom_prefix=custom_prefix.strip(),
        exclude_summary_rows=bool(exclude_summary_rows),
        mapping=col_mapping,
    )
    _update_sheet_config(sheet_name, asdict(config))
    return config


def _build_consolidated_dataframe(
    excel_bytes: bytes,
    selected_sheets: list[str],
    configs: dict[str, SheetConfig],
    default_work_item_type: str,
) -> tuple[pd.DataFrame, dict[str, int], list[str], list[str]]:
    reader = ExcelReader(excel_bytes)
    frames: list[pd.DataFrame] = []
    per_sheet_counts: dict[str, int] = {}
    missing_title_mapping: list[str] = []
    zero_rows_sheets: list[str] = []
    for sheet_name in selected_sheets:
        cfg = configs[sheet_name]
        df = reader.read_sheet(sheet_name, cfg.header_row)
        if not cfg.mapping.title_col or cfg.mapping.title_col not in df.columns:
            per_sheet_counts[sheet_name] = 0
            missing_title_mapping.append(sheet_name)
            continue
        transformed = build_tasks_from_sheet(df, cfg, default_work_item_type)
        count = len(transformed)
        per_sheet_counts[sheet_name] = count
        if count == 0:
            zero_rows_sheets.append(sheet_name)
            continue
        frames.append(transformed)

    if not frames:
        return (
            pd.DataFrame(columns=AZURE_COLUMNS),
            per_sheet_counts,
            missing_title_mapping,
            zero_rows_sheets,
        )
    return (
        pd.concat(frames, ignore_index=True),
        per_sheet_counts,
        missing_title_mapping,
        zero_rows_sheets,
    )


def _copy_mapping_to_other_sheets(
    source_sheet: str,
    selected_sheets: list[str],
    included_sheets: list[str],
    columns_by_sheet: dict[str, list[str]],
    lang: str,
) -> tuple[int, list[str], list[str]]:
    source_cfg = _config_from_state(source_sheet)
    source_mapping = dict(source_cfg.get("mapping", {}))

    required_columns = [value for value in source_mapping.values() if value]
    applied_count = 0
    skipped_messages: list[str] = []
    synced_targets: list[str] = []

    for target_sheet in selected_sheets:
        if target_sheet == source_sheet or target_sheet not in included_sheets:
            continue

        target_columns = columns_by_sheet.get(target_sheet, [])
        missing = [col for col in required_columns if col not in target_columns]
        if missing:
            skipped_messages.append(f"{target_sheet}: {', '.join(missing)}")
            continue

        target_cfg = _config_from_state(target_sheet)
        target_cfg["mapping"] = source_mapping.copy()
        st.session_state.sheet_configs[target_sheet] = target_cfg
        applied_count += 1
        synced_targets.append(target_sheet)
    return applied_count, skipped_messages, synced_targets


def _configs_from_state(selected_sheets: list[str]) -> dict[str, SheetConfig]:
    configs: dict[str, SheetConfig] = {}
    for sheet_name in selected_sheets:
        raw = _config_from_state(sheet_name)
        mapping_raw = raw.get("mapping", {})
        configs[sheet_name] = SheetConfig(
            sheet_name=sheet_name,
            enabled=bool(raw.get("enabled", True)),
            header_row=int(raw.get("header_row", 1)),
            prefix_mode=PrefixMode(raw.get("prefix_mode", PrefixMode.NONE.value)),
            custom_prefix=str(raw.get("custom_prefix", "")),
            exclude_summary_rows=bool(raw.get("exclude_summary_rows", True)),
            mapping=ColumnMapping(
                title_col=mapping_raw.get("title_col"),
                description_col=mapping_raw.get("description_col"),
                hours_col=mapping_raw.get("hours_col"),
                activity_col=mapping_raw.get("activity_col"),
            ),
        )
    return configs


def _render_editor(lang: str) -> None:
    if st.session_state.tasks_df is None:
        return

    source_df: pd.DataFrame = st.session_state.tasks_df.copy()
    search_text = st.text_input(t("search", lang), key="search_text")
    mask = (
        source_df["Title"].astype(str).str.contains(search_text, case=False, na=False)
        if search_text
        else pd.Series([True] * len(source_df), index=source_df.index)
    )
    filtered_df = source_df.loc[mask].copy()

    bulk_col, bulk_btn_col = st.columns([2, 1])
    with bulk_col:
        bulk_type = st.selectbox(t("bulk_type", lang), TASK_TYPES, key="bulk_type")
    with bulk_btn_col:
        if st.button(t("apply_bulk", lang), use_container_width=True):
            source_df.loc[mask, "Work Item Type"] = bulk_type
            st.session_state.tasks_df = source_df
            st.rerun()

    editor_df = filtered_df.copy()
    editor_df.insert(0, "_row_id", editor_df.index)
    edited = st.data_editor(
        editor_df,
        use_container_width=True,
        hide_index=True,
        disabled=["_row_id"],
        num_rows="fixed",
        column_config={
            "_row_id": st.column_config.NumberColumn("_row_id"),
            "Work Item Type": st.column_config.SelectboxColumn(
                "Work Item Type",
                options=TASK_TYPES,
                required=True,
            ),
        },
        key="task_editor",
    )

    for _, row in edited.iterrows():
        row_id = int(row["_row_id"])
        for col in AZURE_COLUMNS:
            source_df.loc[row_id, col] = row[col]

    st.session_state.tasks_df = source_df

    st.caption(f"{t('rows_visible', lang)}: {len(filtered_df)} | {t('total_rows', lang)}: {len(source_df)}")
    csv_bytes = dataframe_to_azure_csv_bytes(source_df)
    st.download_button(
        t("download", lang),
        data=csv_bytes,
        file_name="azure_tasks.csv",
        mime="text/csv",
        use_container_width=True,
    )


def main() -> None:
    st.set_page_config(page_title="Excel Azure Tasks", layout="wide")
    ensure_state()
    _apply_pending_mapping_widget_sync()

    lang_label = st.selectbox(t("language", "es"), list(LANGUAGE_LABELS.keys()), index=0)
    lang = LANGUAGE_LABELS[lang_label]

    st.title(t("title", lang))
    st.write(t("subtitle", lang))

    uploaded_file = st.file_uploader(t("upload", lang), type=["xlsx"])
    if uploaded_file is None:
        st.info(t("no_file", lang))
        return

    file_bytes = uploaded_file.getvalue()
    file_key = f"{uploaded_file.name}:{len(file_bytes)}"
    if st.session_state.last_file_key != file_key:
        st.session_state.last_file_key = file_key
        st.session_state.sheet_configs = {}
        st.session_state.tasks_df = None
        st.session_state.pending_mapping_sync_targets = []
        st.session_state.mapping_feedback_success = None
        st.session_state.mapping_feedback_warnings = []
        st.session_state.build_summary_rows = []
        st.session_state.build_summary_warnings = []
        _clear_sheet_widget_state()

    reader = ExcelReader(file_bytes)
    sheets = reader.list_sheet_names()

    st.subheader(t("sheet_select", lang))
    selected_sheets = st.multiselect(
        t("sheet_select", lang),
        sheets,
        default=sheets,
        key="selected_sheets",
    )

    profile_store = ProfileStore()
    _render_profile_section(profile_store, selected_sheets, lang)

    if not selected_sheets:
        st.warning(t("select_one_sheet", lang))
        return
    _ensure_sheet_configs_initialized(reader, selected_sheets)

    st.subheader(t("sheet_config", lang))
    configs: dict[str, SheetConfig] = {}
    columns_by_sheet: dict[str, list[str]] = {}
    tabs = st.tabs([f"{t('sheet_tab', lang)}: {name}" for name in selected_sheets])
    for sheet_name, tab in zip(selected_sheets, tabs):
        with tab:
            cfg_state = _config_from_state(sheet_name)
            header_row = int(cfg_state.get("header_row", 1))
            preview_df = reader.read_sheet(sheet_name, header_row)
            columns_by_sheet[sheet_name] = [str(col).strip() for col in preview_df.columns]
            configs[sheet_name] = _build_sheet_config_ui(sheet_name, preview_df, lang)
            st.dataframe(preview_df.head(5), use_container_width=True)

    st.subheader(t("build_section", lang))
    action_col_1, action_col_2 = st.columns(2)
    with action_col_1:
        if st.button(t("select_all", lang), use_container_width=True):
            for sheet_name in selected_sheets:
                st.session_state[f"include_sheet_{sheet_name}"] = True
                cfg = _config_from_state(sheet_name)
                cfg["enabled"] = True
                st.session_state.sheet_configs[sheet_name] = cfg
            st.rerun()
    with action_col_2:
        if st.button(t("clear_selection", lang), use_container_width=True):
            for sheet_name in selected_sheets:
                st.session_state[f"include_sheet_{sheet_name}"] = False
                cfg = _config_from_state(sheet_name)
                cfg["enabled"] = False
                st.session_state.sheet_configs[sheet_name] = cfg
            st.rerun()

    for sheet_name in selected_sheets:
        cfg = _config_from_state(sheet_name)
        include_key = f"include_sheet_{sheet_name}"
        if include_key not in st.session_state:
            st.session_state[include_key] = bool(cfg.get("enabled", True))
        include_sheet = st.checkbox(
            f"{sheet_name} - {t('include_in_build', lang)}",
            key=include_key,
            help=t("help_include_sheet", lang),
        )
        cfg["enabled"] = include_sheet
        st.session_state.sheet_configs[sheet_name] = cfg
        if sheet_name in configs:
            configs[sheet_name].enabled = include_sheet

    included_sheets = [
        sheet_name
        for sheet_name in selected_sheets
        if bool(_config_from_state(sheet_name).get("enabled", True))
    ]

    mapping_col, mapping_btn_col = st.columns([2, 1])
    with mapping_col:
        source_sheet = st.selectbox(
            t("source_sheet_mapping", lang),
            selected_sheets,
            key="source_sheet_mapping",
            help=t("help_source_sheet_mapping", lang),
        )
    with mapping_btn_col:
        if st.button(t("apply_mapping_others", lang), use_container_width=True):
            applied_count, skipped_messages, synced_targets = _copy_mapping_to_other_sheets(
                source_sheet=source_sheet,
                selected_sheets=selected_sheets,
                included_sheets=included_sheets,
                columns_by_sheet=columns_by_sheet,
                lang=lang,
            )
            st.session_state.pending_mapping_sync_targets = synced_targets
            st.session_state.mapping_feedback_success = (
                f"{t('mapping_applied', lang)}: {applied_count}" if applied_count else None
            )
            st.session_state.mapping_feedback_warnings = [
                f"{t('mapping_skipped', lang)} ({msg})" for msg in skipped_messages
            ]
            st.rerun()

    if st.session_state.mapping_feedback_success:
        st.success(st.session_state.mapping_feedback_success)
        st.session_state.mapping_feedback_success = None
    for message in st.session_state.mapping_feedback_warnings:
        st.warning(message)
    st.session_state.mapping_feedback_warnings = []

    default_work_item_type = st.selectbox(t("default_type", lang), TASK_TYPES, index=0)
    if st.button(t("build_table", lang), type="primary", use_container_width=True):
        if not included_sheets:
            st.warning(t("no_sheet_selected_build", lang))
            return
        configs = _configs_from_state(selected_sheets)
        (
            st.session_state.tasks_df,
            per_sheet_counts,
            missing_title_sheets,
            zero_rows_sheets,
        ) = _build_consolidated_dataframe(
            file_bytes,
            included_sheets,
            configs,
            default_work_item_type,
        )
        st.session_state.build_summary_rows = [
            {
                t("build_sheet_col", lang): sheet,
                t("build_rows_col", lang): per_sheet_counts.get(sheet, 0),
            }
            for sheet in included_sheets
        ]
        summary_warnings = [
            f"{sheet}: {t('warning_missing_title_map', lang)}" for sheet in missing_title_sheets
        ]
        summary_warnings.extend(
            f"{sheet}: {t('warning_zero_rows', lang)}" for sheet in zero_rows_sheets
        )
        st.session_state.build_summary_warnings = summary_warnings
        st.success(
            f"{t('tasks_ready', lang)}. {t('built_with_sheets', lang)}: {len(included_sheets)}"
        )

    if st.session_state.build_summary_rows:
        st.markdown(f"**{t('build_summary', lang)}**")
        st.dataframe(pd.DataFrame(st.session_state.build_summary_rows), use_container_width=True)
    for warning in st.session_state.build_summary_warnings:
        st.warning(warning)

    _render_editor(lang)


if __name__ == "__main__":
    main()
