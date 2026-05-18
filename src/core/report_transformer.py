from dataclasses import dataclass
from pathlib import Path
from typing import List

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


@dataclass
class ExportData:
    """Data rows from exported Finish Status Report (header excluded)."""

    rows: List[List[str]]
    row_count: int
    col_count: int
    tsv: str

    @property
    def clear_through_row(self) -> int:
        """Last row to clear in SharePoint sheet (header is row 1)."""
        return max(self.row_count + 500, 5000)


def read_export_data(source_path: Path) -> ExportData:
    """Read all data rows below the header, skipping rows without valid IDs."""
    wb = load_workbook(source_path, data_only=True)
    ws = wb.active
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    if max_row < 2:
        wb.close()
        return ExportData(rows=[], row_count=0, col_count=max_col, tsv="")

    data_rows: List[List[str]] = []
    for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col, values_only=True):
        # Get first cell value
        first_cell = row[0] if row else None
        first_cell_str = str(first_cell).strip() if first_cell else ""

        # Only include rows where first column is a number (ID)
        if first_cell_str and first_cell_str.isdigit():
            # Clean cell values: remove newlines and carriage returns to prevent row splits in TSV/Excel paste
            data_rows.append([
                "" if cell is None else str(cell).replace("\r", "").replace("\n", " ").strip()
                for cell in row
            ])

    wb.close()

    col_count = max((len(r) for r in data_rows), default=max_col)
    tsv = values_to_tsv(data_rows)
    return ExportData(
        rows=data_rows,
        row_count=len(data_rows),
        col_count=col_count,
        tsv=tsv,
    )


def read_values_without_header(source_path: Path) -> List[List[str]]:
    return read_export_data(source_path).rows


def values_to_tsv(values: List[List[str]]) -> str:
    lines = ["\t".join(row) for row in values]
    return "\n".join(lines)


def clear_range_address(col_count: int, clear_through_row: int) -> str:
    col = get_column_letter(max(col_count, 1))
    return f"A2:{col}{clear_through_row}"
