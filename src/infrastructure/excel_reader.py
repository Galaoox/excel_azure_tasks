from __future__ import annotations

from io import BytesIO

import pandas as pd


class ExcelReader:
    def __init__(self, excel_bytes: bytes) -> None:
        self._excel_bytes = excel_bytes

    def list_sheet_names(self) -> list[str]:
        with pd.ExcelFile(BytesIO(self._excel_bytes), engine="openpyxl") as excel:
            return list(excel.sheet_names)

    def read_sheet(self, sheet_name: str, header_row: int) -> pd.DataFrame:
        if header_row < 1:
            raise ValueError("header_row must be >= 1")
        return pd.read_excel(
            BytesIO(self._excel_bytes),
            sheet_name=sheet_name,
            header=header_row - 1,
            engine="openpyxl",
        )
