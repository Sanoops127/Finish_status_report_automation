"""Prepare local workbook for SharePoint replace (exact copy of portal export)."""

import shutil
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from src.utils.logger import logger


def _is_valid_xlsx(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size < 100:
        return False
    with path.open("rb") as handle:
        return handle.read(2) == b"PK"


def prepare_sharepoint_workbook(export_path: Path, target_path: Path) -> Path:
    """
    Copy the portal export to the SharePoint target filename.
    The SharePoint file is replaced with this copy (same rows, columns, values).
    """
    export_path = export_path.resolve()
    target_path = target_path.resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    if not _is_valid_xlsx(export_path):
        raise InvalidFileException(
            f"Export is not a valid .xlsx: {export_path}. "
            "Check the portal download completed successfully."
        )

    with tempfile.NamedTemporaryFile(
        suffix=".xlsx", delete=False, dir=target_path.parent
    ) as tmp:
        temp_path = Path(tmp.name)

    try:
        shutil.copy2(export_path, temp_path)
        if not _is_valid_xlsx(temp_path):
            raise InvalidFileException(f"Failed to copy export to temp file: {temp_path}")
        shutil.copy2(temp_path, target_path)
    finally:
        temp_path.unlink(missing_ok=True)

    wb = load_workbook(export_path, read_only=True, data_only=True)
    max_row = wb.active.max_row or 0
    max_col = wb.active.max_column or 0
    wb.close()

    logger.info(
        "Prepared %s — exact copy of export (%s rows x %s columns)",
        target_path.name,
        max_row,
        max_col,
    )
    return target_path


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
            value = cell_value if cell_value is not None else ""
            html_rows.append(f"<td>{value}</td>")
        html_rows.append("</tr>")

    html_rows.append("</table>")
    wb.close()

    return "\n".join(html_rows)
