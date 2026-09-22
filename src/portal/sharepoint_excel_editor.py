import re
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlparse

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError

from src.core.report_transformer import rows_to_html_table, values_to_tsv
from src.core.workbook_sync import workbook_to_html_table
from src.utils.logger import logger


class SharePointExcelEditor:
    def __init__(self, page: Page):
        self.page = page

    def _is_direct_doc_url(self, url: str) -> bool:
        parsed = urlparse(url.lower())
        path = parsed.path
        query = parsed.query
        return (
            "doc.aspx" in path
            or "sourcedoc=" in query
            or "/:x:/" in path
            or "/:w:/" in path
            or "/:p:/" in path
            or path.endswith(".xlsx")
            or path.endswith(".xls")
        )

    def _go_to_cell(self, cell_address: str) -> bool:
        """Navigate to cell or range in Excel Online using the Name Box or Go To shortcut."""
        name_box_selectors = [
            "input#formulaBarNameBox",
            "input[data-automationid='NameBox']",
            "input[aria-label='Name Box']",
            "#NameBox",
            "input[title*='Name Box']",
            "div[data-automationid='NameBoxContainer'] input",
        ]
        for sel in name_box_selectors:
            try:
                box = self.page.locator(sel).first
                if box.is_visible(timeout=1000):
                    box.click(timeout=1500)
                    self.page.wait_for_timeout(200)
                    box.fill(cell_address)
                    self.page.keyboard.press("Enter")
                    self.page.wait_for_timeout(500)
                    return True
            except Exception:
                continue

        # Fallback to Go To shortcut (Ctrl+G)
        try:
            self.page.keyboard.press("Control+g")
            self.page.wait_for_timeout(500)
            dialog_input = self.page.locator("input[type='text']:visible").first
            if dialog_input.is_visible(timeout=1000):
                dialog_input.fill(cell_address)
                self.page.keyboard.press("Enter")
                self.page.wait_for_timeout(500)
                return True
        except Exception:
            pass

        return False

    def _write_to_clipboard(self, html_content: str, tsv_content: Optional[str] = None) -> None:
        """Write both text/html and text/plain to clipboard so Excel Online accepts paste."""
        try:
            self.page.context.grant_permissions(["clipboard-read", "clipboard-write"])
        except Exception:
            pass

        self.page.evaluate(
            """async ({ html, tsv }) => {
                const items = {};
                if (html) {
                    items["text/html"] = new Blob([html], { type: "text/html" });
                }
                if (tsv) {
                    items["text/plain"] = new Blob([tsv], { type: "text/plain" });
                } else if (html) {
                    const tmp = document.createElement("div");
                    tmp.innerHTML = html;
                    items["text/plain"] = new Blob([tmp.innerText || tmp.textContent || ""], { type: "text/plain" });
                }
                const item = new ClipboardItem(items);
                await navigator.clipboard.write([item]);
            }""",
            {"html": html_content, "tsv": tsv_content},
        )

    def update_file_values(
        self,
        sharepoint_home_url: str,
        file_name: str,
        *,
        rows: Optional[List[List[str]]] = None,
        data_tsv: Optional[str] = None,
        data_html: Optional[str] = None,
        source_workbook: Optional[Path] = None,
    ) -> None:
        """Open SharePoint Excel file and paste export data into web editor."""
        logger.info("Opening SharePoint location: %s", sharepoint_home_url)
        self.page.goto(sharepoint_home_url, timeout=120_000, wait_until="domcontentloaded")
        self.page.wait_for_timeout(5000)

        if self._is_direct_doc_url(sharepoint_home_url):
            logger.info("Direct SharePoint document URL detected; workbook loaded directly")
        else:
            logger.info("Opening SharePoint file: %s", file_name)
            if not self._open_excel_file(file_name):
                raise RuntimeError(f"Could not open SharePoint file: {file_name}")

        self._ensure_edit_mode()
        self._focus_workbook()

        total_rows = len(rows) if rows else 0

        if rows:
            logger.info("Pasting %s rows at A2 in chunks into Excel Online...", total_rows)
            self._paste_rows_in_chunks(rows, start_row_offset=2)
        elif data_html:
            logger.info("Pasting HTML table at A1: %s", file_name)
            self._paste_html_content(data_html)
        elif data_tsv:
            logger.info("Pasting TSV at A1: %s", file_name)
            self._paste_excel_content(data_tsv)
        else:
            raise ValueError("Provide rows, data_html, or data_tsv")

        self.page.wait_for_timeout(3000)
        self._apply_formulas_to_columns(total_rows=total_rows)
        self._save_file()
        logger.info("SharePoint file updated in web editor successfully")

    def update_from_workbook(
        self,
        sharepoint_home_url: str,
        file_name: str,
        source_export: Path,
        prepared_workbook: Optional[Path] = None,
    ) -> None:
        """Open SharePoint Excel and paste the export workbook data (without headers) at A2."""
        paste_source = (prepared_workbook or source_export).resolve()
        if not paste_source.is_file():
            raise FileNotFoundError(f"Export file not found: {paste_source}")

        logger.info("Opening SharePoint location: %s", sharepoint_home_url)
        self.page.goto(sharepoint_home_url, timeout=120_000, wait_until="domcontentloaded")
        self.page.wait_for_timeout(5000)

        if self._is_direct_doc_url(sharepoint_home_url):
            logger.info("Direct SharePoint document URL detected; workbook loaded directly")
        else:
            logger.info("Opening SharePoint file: %s", file_name)
            if not self._open_excel_file(file_name):
                raise RuntimeError(f"Could not open SharePoint file: {file_name}")

        self._ensure_edit_mode()
        self._focus_workbook()

        logger.info("Building paste table from export: %s", paste_source.name)
        html_table = workbook_to_html_table(paste_source, skip_header=True)
        self._paste_data_at_a2(html_table)
        self._save_file()
        logger.info("SharePoint workbook updated from export %s (values only, at A2)", paste_source.name)

    def _paste_rows_in_chunks(self, rows: List[List[str]], start_row_offset: int = 2, chunk_size: int = 4000) -> None:
        """Paste rows in manageable chunks (e.g. 4000 rows) so Excel Online processes each paste instantly."""
        total_rows = len(rows)
        for i in range(0, total_rows, chunk_size):
            chunk = rows[i : i + chunk_size]
            current_excel_row = start_row_offset + i
            cell_target = f"A{current_excel_row}"

            logger.info("Pasting chunk rows %s to %s at %s...", i + 1, min(i + chunk_size, total_rows), cell_target)

            chunk_html = rows_to_html_table(chunk)
            chunk_tsv = values_to_tsv(chunk)

            # Navigate to target cell (A2, A4002, etc.)
            if not self._go_to_cell(cell_target):
                self._focus_workbook()
                self.page.keyboard.press("Control+Home")
                for _ in range(current_excel_row - 1):
                    self.page.keyboard.press("ArrowDown")

            self._write_to_clipboard(chunk_html, chunk_tsv)
            self.page.keyboard.press("Control+v")
            self.page.wait_for_timeout(3000)

    def _open_excel_file(self, file_name: str) -> bool:
        stem = Path(file_name).stem
        normalized_names = [file_name]
        if not file_name.lower().endswith(".xlsx"):
            normalized_names.append(f"{file_name}.xlsx")

        readable_title = stem.replace("_", " ").title()
        for name_variant in (f"{readable_title}.xlsx", readable_title, stem):
            if name_variant not in normalized_names:
                normalized_names.append(name_variant)

        for candidate in normalized_names:
            if self._click_file_and_switch_page(candidate):
                return True
        return False

    def _click_file_and_switch_page(self, candidate_name: str) -> bool:
        exact_pattern = re.compile(rf"^{re.escape(candidate_name)}$", re.IGNORECASE)

        link = self.page.get_by_role("link", name=exact_pattern)
        try:
            link.first.wait_for(state="visible", timeout=3_000)
        except Exception:
            link = self.page.locator(
                f'a:text-is("{candidate_name}"), '
                f'[aria-label="{candidate_name}"], '
                f'[title="{candidate_name}"]'
            )
            try:
                link.first.wait_for(state="visible", timeout=3_000)
            except Exception:
                link = self.page.get_by_text(candidate_name, exact=True)
                try:
                    link.first.wait_for(state="visible", timeout=3_000)
                except Exception:
                    return False

        existing_pages = list(self.page.context.pages)
        current_url = self.page.url
        try:
            link.first.click(timeout=5_000)
        except Exception:
            return False

        self.page.wait_for_timeout(5000)

        if len(self.page.context.pages) > len(existing_pages):
            self.page = self.page.context.pages[-1]
            self.page.wait_for_load_state("domcontentloaded", timeout=120_000)
            self.page.wait_for_timeout(8000)
            return True

        if self.page.url != current_url:
            self.page.wait_for_load_state("domcontentloaded", timeout=120_000)
            self.page.wait_for_timeout(8000)
            return True

        try:
            link.first.dblclick(timeout=10_000)
            self.page.wait_for_timeout(5000)
            if len(self.page.context.pages) > len(existing_pages):
                self.page = self.page.context.pages[-1]
            self.page.wait_for_load_state("domcontentloaded", timeout=120_000)
            self.page.wait_for_timeout(8000)
            return True
        except Exception:
            return False

    def _ensure_edit_mode(self) -> None:
        for name in ("Edit workbook", "Editing", "Edit"):
            btn = self.page.get_by_role("button", name=name).first
            try:
                btn.wait_for(state="visible", timeout=4_000)
                btn.click(timeout=3_000)
                self.page.wait_for_timeout(3000)
                break
            except Exception:
                continue

    def _focus_workbook(self) -> None:
        selectors = [
            "div[role='grid']",
            "canvas[aria-label*='Grid']",
            "div[aria-label*='Worksheet']",
            "div[data-automationid='Sheet']",
        ]
        for selector in selectors:
            try:
                locator = self.page.locator(selector).first
                locator.wait_for(state="visible", timeout=4_000)
                locator.click(timeout=3_000)
                self.page.wait_for_timeout(600)
                return
            except Exception:
                continue
        self.page.click("body", timeout=5_000)
        self.page.wait_for_timeout(600)

    def _paste_html_content(self, html_table: str) -> None:
        """Clear sheet and paste HTML so each cell stays in one column."""
        self._focus_workbook()
        self.page.keyboard.press("Control+A")
        self.page.wait_for_timeout(400)
        self.page.keyboard.press("Delete")
        self.page.wait_for_timeout(400)
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(400)

        self._write_to_clipboard(html_table)
        self.page.wait_for_timeout(1000)
        self.page.keyboard.press("Control+v")
        # Give Excel Online sufficient time (25s) to parse and render 32,000 rows
        self.page.wait_for_timeout(25000)

    def _paste_excel_content(self, data_tsv: str) -> None:
        """Delete all values below header (starting at A2) and paste TSV data."""
        self._focus_workbook()
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(500)
        self.page.keyboard.press("ArrowDown")
        self.page.wait_for_timeout(300)

        self._write_to_clipboard("", data_tsv)
        self.page.wait_for_timeout(1000)
        self.page.keyboard.press("Control+v")
        self.page.wait_for_timeout(25000)

    def _paste_data_at_a2(self, html_table: str) -> None:
        """Navigate to A2 and paste data values (no headers)."""
        self._focus_workbook()
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(500)
        self.page.keyboard.press("ArrowDown")
        self.page.wait_for_timeout(300)

        self._write_to_clipboard(html_table)
        self.page.wait_for_timeout(1000)
        self.page.keyboard.press("Control+v")
        # Give Excel Online sufficient time (25s) to parse and render 32,000 rows
        self.page.wait_for_timeout(25000)


    def _apply_formulas_to_columns(self, total_rows: int = 0) -> None:
        """Apply date formatting formulas to columns AA and AB after paste."""
        try:
            logger.info("Applying formulas to columns AA and AB (total rows: %s)", total_rows)
            max_row = max(total_rows + 1, 1000)

            self._focus_workbook()
            self.page.wait_for_timeout(1000)

            # 1. Date formula in AA2
            logger.info("Navigating to AA2")
            if not self._go_to_cell("AA2"):
                self.page.keyboard.press("Control+Home")
                for _ in range(26):
                    self.page.keyboard.press("ArrowRight")
                self.page.keyboard.press("ArrowDown")

            formula_date = '=IF(ISBLANK(M2), "", TEXT(M2, "dd-mm-yyyy"))'
            logger.info("Entering date formula in AA2")
            self.page.keyboard.type(formula_date, delay=10)
            self.page.keyboard.press("Enter")
            self.page.wait_for_timeout(1000)

            # Select range AA2:AA{max_row} and fill down (Ctrl+D)
            fill_range_aa = f"AA2:AA{max_row}"
            logger.info("Selecting %s and filling down (Ctrl+D)", fill_range_aa)
            if not self._go_to_cell(fill_range_aa):
                self.page.keyboard.press("ArrowUp")
                self.page.keyboard.press("Control+Shift+End")

            self.page.keyboard.press("Control+d")
            self.page.wait_for_timeout(3000)

            # 2. DateTime formula in AB2
            logger.info("Navigating to AB2")
            if not self._go_to_cell("AB2"):
                self._go_to_cell("AA2")
                self.page.keyboard.press("ArrowRight")

            formula_datetime = '=IF(ISBLANK(V2), "", TEXT(V2, "dd-mm-yyyy hh:mm:ss"))'
            logger.info("Entering datetime formula in AB2")
            self.page.keyboard.type(formula_datetime, delay=10)
            self.page.keyboard.press("Enter")
            self.page.wait_for_timeout(1000)

            # Select range AB2:AB{max_row} and fill down (Ctrl+D)
            fill_range_ab = f"AB2:AB{max_row}"
            logger.info("Selecting %s and filling down (Ctrl+D)", fill_range_ab)
            if not self._go_to_cell(fill_range_ab):
                self.page.keyboard.press("ArrowUp")
                self.page.keyboard.press("Control+Shift+End")

            self.page.keyboard.press("Control+d")
            self.page.wait_for_timeout(3000)

            self._go_to_cell("A1")
            logger.info("Formulas applied successfully to columns AA and AB")
        except Exception as e:
            logger.warning("Error applying formulas: %s", e)

    def _save_file(self) -> None:
        self.page.keyboard.press("Control+s")
        try:
            self.page.get_by_text("Saved", exact=False).first.wait_for(timeout=10_000)
        except PlaywrightTimeoutError:
            self.page.wait_for_timeout(3000)
