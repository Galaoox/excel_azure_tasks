from __future__ import annotations

import pandas as pd

from src.application.transform_service import AZURE_COLUMNS


def dataframe_to_azure_csv_bytes(df: pd.DataFrame) -> bytes:
    ordered = df.copy()
    for column in AZURE_COLUMNS:
        if column not in ordered.columns:
            ordered[column] = ""
    csv_text = ordered.to_csv(index=False, encoding="utf-8-sig", columns=AZURE_COLUMNS)
    return csv_text.encode("utf-8-sig")
