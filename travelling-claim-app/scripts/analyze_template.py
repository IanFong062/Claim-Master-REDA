from __future__ import annotations

import hashlib
import json
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "Travelling claim form - updated version.xlsx"


def analyze() -> dict:
    wb = openpyxl.load_workbook(TEMPLATE, data_only=False)
    result = {
        "template": str(TEMPLATE),
        "sha256": hashlib.sha256(TEMPLATE.read_bytes()).hexdigest(),
        "sheets": [],
    }
    for ws in wb.worksheets:
        non_empty = []
        formulas = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                item = {"cell": cell.coordinate, "value": str(cell.value)}
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append(item)
                else:
                    non_empty.append(item)
        result["sheets"].append(
            {
                "name": ws.title,
                "usedRange": ws.calculate_dimension(),
                "mergedCells": [str(rng) for rng in ws.merged_cells.ranges],
                "nonEmptyCells": non_empty,
                "formulaCells": formulas,
                "printArea": str(ws.print_area or ""),
                "pageSetup": {
                    "orientation": ws.page_setup.orientation,
                    "paperSize": ws.page_setup.paperSize,
                    "fitToPage": bool(ws.sheet_properties.pageSetUpPr and ws.sheet_properties.pageSetUpPr.fitToPage),
                },
            }
        )
    return result


if __name__ == "__main__":
    print(json.dumps(analyze(), ensure_ascii=False, indent=2))
