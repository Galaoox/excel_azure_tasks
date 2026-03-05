import argparse
import re
import pandas as pd


AZURE_COLUMNS = [
    "Work Item Type",
    "Title",
    "Description",
    "Original Estimate",
    "Remaining Work",
    "Activity",
]


def to_float_or_blank(x):
    """Convierte a float; soporta '1,5' y '1.5'. Si no puede, retorna vacío."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return ""
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return ""


def infer_activity(title: str, default_activity: str) -> str:
    """Heurística simple por keywords."""
    if not isinstance(title, str):
        return default_activity

    t = title.lower()
    rules = [
        (r"\b(prueba|testing|qa|test|e2e)\b", "Testing"),
        (r"\b(doc|documentaci|handoff)\b", "Documentation"),
        (r"\b(deploy|release|desplieg)\b", "Deployment"),
        (r"\b(spike|research|investig)\b", "Other"),
    ]
    for pattern, activity in rules:
        if re.search(pattern, t):
            return activity
    return default_activity


def main():
    ap = argparse.ArgumentParser(description="Excel -> CSV compatible con Azure Boards")
    ap.add_argument("--excel", required=True, help="Ruta al archivo .xlsx")
    ap.add_argument("--sheet", default=None, help="Nombre de la hoja (si no se indica usa la primera)")
    ap.add_argument("--out", required=True, help="Ruta salida .csv")

    # Mapeo de columnas en el Excel
    ap.add_argument("--title-col", default="Título", help="Nombre columna título en Excel")
    ap.add_argument("--desc-col", default="Descripción", help="Nombre columna descripción en Excel")
    ap.add_argument("--hours-col", default="Estimación (h)", help="Nombre columna horas en Excel")

    # Azure fields
    ap.add_argument("--work-item-type", default="Task", help="Task, Bug, etc.")
    ap.add_argument("--default-activity", default="Development", help="Activity por defecto")
    ap.add_argument("--activity-col", default="Activity", help="Columna opcional de Activity en Excel")
    ap.add_argument("--use-activity-col", action="store_true", help="Usar columna Activity si existe")

    # Filtro por marker
    ap.add_argument("--after-title", default=None, help="Exportar solo filas DESPUÉS de este título (marker)")

    args = ap.parse_args()

    df = pd.read_excel(args.excel, sheet_name=args.sheet)
    df.columns = [str(c).strip() for c in df.columns]

    # Validaciones
    for col in [args.title_col, args.desc_col, args.hours_col]:
        if col not in df.columns:
            raise SystemExit(f"ERROR: No existe la columna '{col}' en el Excel. Columnas: {list(df.columns)}")

    # Filtro desde marker
    if args.after_title:
        titles = df[args.title_col].astype(str).fillna("").str.strip()
        matches = titles.eq(args.after_title.strip())
        if matches.any():
            idx = matches.idxmax()
            df = df.loc[idx + 1 :].copy()
        else:
            print(f"WARNING: Marker '{args.after_title}' no encontrado. Se exportará todo.")

    out = pd.DataFrame()
    out["Work Item Type"] = args.work_item_type
    out["Title"] = df[args.title_col].astype(str).fillna("").str.strip()
    out["Description"] = df[args.desc_col].astype(str).fillna("").str.strip()

    hours = df[args.hours_col].apply(to_float_or_blank)
    out["Original Estimate"] = hours
    out["Remaining Work"] = hours

    # Activity
    activities = pd.Series([""] * len(out), index=out.index)
    if args.use_activity_col and args.activity_col in df.columns:
        activities = df[args.activity_col].astype(str).fillna("").str.strip()
        activities = activities.replace({"nan": ""})

    final_acts = []
    for i, row in out.iterrows():
        act = str(activities.loc[i]).strip() if i in activities.index else ""
        if act.lower() == "nan":
            act = ""
        if not act:
            act = infer_activity(row["Title"], args.default_activity)
        final_acts.append(act)

    out["Activity"] = final_acts

    # Limpieza: quitar filas sin título
    out = out[out["Title"].astype(str).str.len() > 0].copy()

    # Export: UTF-8 con BOM para acentos
    out.to_csv(args.out, index=False, encoding="utf-8-sig", columns=AZURE_COLUMNS)
    print(f"OK: CSV generado en {args.out} ({len(out)} items)")


if __name__ == "__main__":
    main()