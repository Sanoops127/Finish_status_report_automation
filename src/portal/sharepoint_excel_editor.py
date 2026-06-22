from pathlib import Path
from typing import Optional

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError

from src.core.workbook_sync import workbook_to_html_table
from src.utils.logger import logger


class SharePointExcelEditor:
    def __init__(self, page: Page):
        self.page = page

    def update_file_values(
        self,
        sharepoint_home_url: str,
        file_name: str,
        *,
        data_tsv: str | None = None,
        data_html: str | None = None,
    ) -> None:
        """Open SharePoint Excel file and paste export data (HTML preferred; TSV optional)."""
        if not data_html and not data_tsv:
            raise ValueError("Provide data_html or data_tsv")

        logger.info("Opening SharePoint home page")
        self.page.goto(sharepoint_home_url, timeout=120_000, wait_until="domcontentloaded")
        self.page.wait_for_timeout(5000)

        logger.info("Opening SharePoint file: %s", file_name)
        if not self._open_excel_file(file_name):
            raise RuntimeError(f"Could not open SharePoint file: {file_name}")

        self._ensure_edit_mode()
        self._focus_workbook()
        if data_html:
            logger.info("Pasting HTML table at A1: %s", file_name)
            self._paste_html_content(data_html)
        else:
            logger.info("Pasting TSV at A1: %s", file_name)
            self._paste_excel_content(data_tsv or "")
        self._apply_formulas_to_columns()
        self._save_file()
        logger.info("SharePoint file updated with export data")

    def update_from_workbook(
        self,
        sharepoint_home_url: str,
        file_name: str,
        source_export: Path,
        prepared_workbook: Optional[Path] = None,
    ) -> None:
        """Open SharePoint Excel and paste the export workbook data (without headers) at A2."""
        paste_source = source_export.resolve()
        if not paste_source.is_file():
            raise FileNotFoundError(f"Export file not found: {paste_source}")

        logger.info("Opening SharePoint home page")
        self.page.goto(sharepoint_home_url, timeout=120_000, wait_until="domcontentloaded")
        self.page.wait_for_timeout(5000)

        logger.info("Opening SharePoint file: %s", file_name)
        if not self._open_excel_file(file_name):
            raise RuntimeError(f"Could not open SharePoint file: {file_name}")

        # self._ensure_edit_mode()
        # self._focus_workbook()
        logger.info("Building paste table from export: %s", paste_source.name)
        self._paste_data_at_a2(workbook_to_html_table(paste_source, skip_header=True))
        self._save_file()
        logger.info("SharePoint workbook updated from export %s (values only, at A2)", paste_source.name)

    def _open_excel_file(self, file_name: str) -> bool:
        normalized_names = [file_name]
        if not file_name.lower().endswith(".xlsx"):
            normalized_names.append(f"{file_name}.xlsx")
        else:
            normalized_names.append(file_name[:-5])

        for candidate in normalized_names:
            if self._click_file_and_switch_page(candidate):
                return True
        return False

    def _click_file_and_switch_page(self, candidate_name: str) -> bool:
        link = self.page.get_by_role("link", name=candidate_name)
        try:
            link.first.wait_for(state="visible", timeout=5_000)
        except Exception:
            link = self.page.get_by_text(candidate_name, exact=False).first
            try:
                link.wait_for(state="visible", timeout=5_000)
            except Exception:
                return False

        existing_pages = list(self.page.context.pages)
        current_url = self.page.url
        try:
            link.click(timeout=5_000)
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
            link.dblclick(timeout=10_000)
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

        self.page.evaluate(
            """async (htmlContent) => {
                const blob = new Blob([htmlContent], { type: "text/html" });
                const item = new ClipboardItem({ "text/html": blob });
                await navigator.clipboard.write([item]);
            }""",
            html_table,
        )
        self.page.keyboard.press("Control+v")
        self.page.wait_for_timeout(3000)

    def _paste_excel_content(self, data_tsv: str) -> None:
        """Delete all values below header (starting at A2) and paste TSV data."""
        self._focus_workbook()

        # Go to A2 to preserve header at A1
        self.page.keyboard.press("Control+A")
        self.page.wait_for_timeout(400)
        self.page.keyboard.press("Delete")
        self.page.wait_for_timeout(400)

        # Select all from A2 to the end and delete
        # self.page.keyboard.press("Control+Shift+End")
        # self.page.wait_for_timeout(600)
        # self.page.keyboard.press("Delete")
        # self.page.wait_for_timeout(800)

        # Go back to A2 and paste
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(400)

        self.page.evaluate(
            """async (tsvContent) => {
                await navigator.clipboard.writeText(tsvContent);
            }""",
            data_tsv,
        )
        self.page.keyboard.press("Control+v")
        self.page.wait_for_timeout(5000)

    def _paste_tsv_at_a2(self, data_tsv: str) -> None:
        """Navigate to A1 and paste TSV data directly (like manual paste)."""
        self._focus_workbook()

        # Go to A1
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(400)

        # Paste TSV data directly without clearing
        clean_tsv = data_tsv.strip()

        self.page.evaluate(
            """async (tsvContent) => {
                await navigator.clipboard.writeText(tsvContent);
            }""",
            clean_tsv,
        )
        self.page.keyboard.press("Control+v")
        self.page.wait_for_timeout(3000)

    def _paste_data_at_a2(self, html_table: str) -> None:
        """Navigate to A2 and paste data values (no headers)."""
        self._focus_workbook()
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(400)
        self.page.keyboard.press("ArrowDown")
        self.page.wait_for_timeout(200)

        self.page.evaluate(
            """async (htmlContent) => {
                const blob = new Blob([htmlContent], { type: "text/html" });
                const item = new ClipboardItem({ "text/html": blob });
                await navigator.clipboard.write([item]);
            }""",
            html_table,
        )
        self.page.keyboard.press("Control+v")
        self.page.wait_for_timeout(3000)

    def _replace_sheet_with_html(self, html_table: str) -> None:
        """Clear sheet and paste HTML so columns/rows match the prepared file."""
        self._focus_workbook()
        self.page.keyboard.press("Control+Home")
        self.page.wait_for_timeout(400)
        self.page.keyboard.press("Control+a")
        self.page.wait_for_timeout(200)
        self.page.keyboard.press("Control+a")
        self.page.wait_for_timeout(400)
        self.page.keyboard.press("Delete")
        self.page.wait_for_timeout(800)

        self.page.evaluate(
            """async (htmlContent) => {
                const blob = new Blob([htmlContent], { type: "text/html" });
                const item = new ClipboardItem({ "text/html": blob });
                await navigator.clipboard.write([item]);
            }""",
            html_table,
        )
        self.page.keyboard.press("Control+v")
        self.page.wait_for_timeout(3000)

    def _apply_formulas_to_columns(self) -> None:
        """Apply date formatting formulas to columns AA and AB after paste."""
        try:
            logger.info("Applying formulas to columns AA and AB")
            self._focus_workbook()
            self.page.wait_for_timeout(1000)

            # Navigate to A2 first
            self.page.keyboard.press("Control+Home")
            self.page.wait_for_timeout(400)

            # Navigate to column AA (27th column) by pressing Right 26 times
            logger.info("Navigating to column AA")
            for _ in range(27):
                self.page.keyboard.press("ArrowRight")
                self.page.wait_for_timeout(50)
            self.page.wait_for_timeout(600)

            # Enter the date formula in AA2
            formula_date = '=IF(ISBLANK(M2), "", TEXT(M2, "dd-mm-yyyy"))'
            logger.info("Entering date formula in AA2")
            self.page.keyboard.type(formula_date, delay=10)
            self.page.keyboard.press("Enter")
            self.page.wait_for_timeout(800)

            # Go back to AA2 to copy formula down
            self.page.keyboard.press("ArrowUp")
            self.page.wait_for_timeout(300)

            # Select from AA2 to AA1000 using keyboard
            logger.info("Selecting AA2:AA1000 and filling down")
            self.page.keyboard.press("Control+Shift+End")
            self.page.wait_for_timeout(600)

            # Fill down using Ctrl+D
            self.page.keyboard.press("Control+d")
            self.page.wait_for_timeout(2000)

            # Navigate to AB2 (move right one column from current position)
            logger.info("Navigating to column AB")
            
            # Collapse the selection from column AA and move one column to the right
            self.page.keyboard.press("ArrowRight")
            self.page.wait_for_timeout(300)
            self.page.keyboard.press("ArrowDown")
            self.page.wait_for_timeout(600)

            # Enter the datetime formula in AB2
            formula_datetime = '=IF(ISBLANK(V2), "", TEXT(V2, "dd-mm-yyyy hh:mm:ss"))'
            logger.info("Entering datetime formula in AB2")
            self.page.keyboard.type(formula_datetime, delay=10)
            self.page.keyboard.press("Enter")
            self.page.wait_for_timeout(800)

            # Go back to AB2 to copy formula down
            self.page.keyboard.press("ArrowUp")
            self.page.wait_for_timeout(300)

            # Select from AB2 to AB1000 using keyboard
            logger.info("Selecting AB2:AB1000 and filling down")
            self.page.keyboard.press("Control+Shift+End")
            self.page.wait_for_timeout(600)

            # Fill down using Ctrl+D
            self.page.keyboard.press("Control+d")
            self.page.wait_for_timeout(1000)

            self.page.keyboard.press("Control+Home")
            self.page.wait_for_timeout(400)

            logger.info("Formulas applied successfully to columns AA and AB")
        except Exception as e:
            logger.warning("Error applying formulas: %s", e)

    def _save_file(self) -> None:
        self.page.keyboard.press("Control+s")
        try:
            self.page.get_by_text("Saved", exact=False).first.wait_for(timeout=10_000)
        except PlaywrightTimeoutError:
            self.page.wait_for_timeout(3000)
