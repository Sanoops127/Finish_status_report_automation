import html
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from src.utils.logger import logger

# Control characters that break TSV/clipboard paste into Excel (tab splits columns).
_CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


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


def sanitize_cell_value(cell) -> str:
    """Normalize a cell for SharePoint paste. Tabs/newlines split columns in Excel TSV paste."""
    if cell is None:
        return ""
    text = str(cell)
    text = text.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    text = text.replace("\t", " ")
    text = _CONTROL_CHAR_RE.sub("", text)
    return " ".join(text.split())


def prepare_finish_status_report_workbook(source_path: Path, target_path: Path) -> Path:
    """
    Transform source Depotnet export workbook into a clean SharePoint target workbook.
    - Preserves header row (row 1).
    - Filters data rows to include only valid numeric JOB IDs (column 1).
    - Appends formulas for Column AA (=IF(ISBLANK(M{r}), "", TEXT(M{r}, "dd-mm-yyyy")))
      and Column AB (=IF(ISBLANK(V{r}), "", TEXT(V{r}, "dd-mm-yyyy hh:mm:ss"))).
    - Uses streaming openpyxl read/write for maximum speed & minimal memory footprint on 32,000+ rows.
    """
    source_path = source_path.resolve()
    target_path = target_path.resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    if not source_path.is_file():
        raise FileNotFoundError(f"Source export workbook not found: {source_path}")

    logger.info("Preparing workbook %s from %s", target_path.name, source_path.name)

    wb_in = load_workbook(source_path, read_only=True, data_only=True)
    ws_in = wb_in.active

    wb_out = Workbook(write_only=True)
    ws_out = wb_out.create_sheet()

    total_read = 0
    total_written = 0

    for row_idx, row in enumerate(ws_in.iter_rows(values_only=True), start=1):
        total_read += 1
        if row_idx == 1:
            # Header row
            header = [sanitize_cell_value(cell) for cell in row] if row else []
            while len(header) < 28:
                header.append("")
            if not header[26]:
                header[26] = "Formatted Date"
            if not header[27]:
                header[27] = "Formatted DateTime"
            ws_out.append(header)
            total_written += 1
        else:
            first_cell = row[0] if row else None
            first_cell_str = str(first_cell).strip() if first_cell is not None else ""
            if first_cell_str and first_cell_str.isdigit():
                clean_row = [sanitize_cell_value(cell) for cell in row]
                while len(clean_row) < 28:
                    clean_row.append("")

                # Excel row number in target file (1-indexed)
                target_excel_row = total_written + 1
                clean_row[26] = f'=IF(ISBLANK(M{target_excel_row}), "", TEXT(M{target_excel_row}, "dd-mm-yyyy"))'
                clean_row[27] = f'=IF(ISBLANK(V{target_excel_row}), "", TEXT(V{target_excel_row}, "dd-mm-yyyy hh:mm:ss"))'

                ws_out.append(clean_row)
                total_written += 1

    wb_in.close()

    # Save to temp file first to prevent partial file writes
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=target_path.parent) as tmp:
        temp_path = Path(tmp.name)

    try:
        wb_out.save(temp_path)
        shutil.copy2(temp_path, target_path)
    finally:
        temp_path.unlink(missing_ok=True)

    logger.info(
        "Workbook %s prepared successfully: %s rows written (%s source rows processed)",
        target_path.name,
        total_written,
        total_read,
    )
    return target_path


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
            data_rows.append([sanitize_cell_value(cell) for cell in row])

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
    if not values:
        return ""
    col_count = max(len(row) for row in values)
    lines: list[str] = []
    for row in values:
        cells = [sanitize_cell_value(cell) for cell in row]
        while len(cells) < col_count:
            cells.append("")
        lines.append("\t".join(cells))
    return "\n".join(lines)


def rows_to_html_table(values: List[List[str]]) -> str:
    """HTML table paste keeps each value in one cell even if it contained tabs."""
    if not values:
        return "<table></table>"
    col_count = max(len(row) for row in values)
    parts = ["<table>"]
    for row in values:
        parts.append("<tr>")
        cells = [sanitize_cell_value(cell) for cell in row]
        while len(cells) < col_count:
            cells.append("")
        for cell in cells:
            parts.append(f"<td>{html.escape(cell)}</td>")
        parts.append("</tr>")
    parts.append("</table>")
    return "\n".join(parts)


def clear_range_address(col_count: int, clear_through_row: int) -> str:
    col = get_column_letter(max(col_count, 1))
    return f"A2:{col}{clear_through_row}"

