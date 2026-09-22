"""Prepare local workbook for SharePoint replace (exact copy of portal export)."""

import shutil
import tempfile
from pathlib import Path

import html

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from src.core.report_transformer import sanitize_cell_value
from src.utils.logger import logger


def _is_valid_xlsx(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size < 100:
        return False
    with path.open("rb") as handle:
        return handle.read(2) == b"PK"


def prepare_sharepoint_workbook(export_path: Path, target_path: Path) -> Path:
    """
    Prepare exported workbook for SharePoint replacement.
    Filters invalid rows, formats formulas for columns AA and AB, and writes streaming .xlsx output.
    """
    from src.core.report_transformer import prepare_finish_status_report_workbook
    return prepare_finish_status_report_workbook(export_path, target_path)



def workbook_to_html_table(export_path: Path, skip_header: bool = False) -> str:
    """
    Convert workbook to HTML table for pasting into Excel.
    If skip_header=True, excludes the first row (headers) and starts from row 2.
    """
    if not _is_valid_xlsx(export_path):
        raise InvalidFileException(f"Invalid .xlsx: {export_path}")

    wb = load_workbook(export_path, read_only=True, data_only=True)
    ws = wb.active

    html_rows = ["<table>"]
    start_row = 2 if skip_header else 1

    for row in ws.iter_rows(min_row=start_row, values_only=True):
        html_rows.append("<tr>")
        for cell_value in row:
            value = html.escape(sanitize_cell_value(cell_value))
            html_rows.append(f"<td>{value}</td>")
        html_rows.append("</tr>")

    html_rows.append("</table>")
    wb.close()

    return "\n".join(html_rows)
