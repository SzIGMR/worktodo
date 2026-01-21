"""Excel writing helpers."""

from __future__ import annotations

from datetime import date
import json
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font


def write_sheet(workbook: Workbook, title: str, headers: Iterable[str], rows: Iterable[Iterable[Any]]) -> None:
    ws = workbook.create_sheet(title)
    ws.append(list(headers))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append([_convert_cell(value) for value in row])
    ws.auto_filter.ref = ws.dimensions


def _convert_cell(value: Any) -> Any:
    if isinstance(value, date):
        return value
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    if isinstance(value, (list, tuple, set)):
        serializable = sorted(value) if isinstance(value, set) else list(value)
        return json.dumps(serializable, ensure_ascii=False)
    return value
