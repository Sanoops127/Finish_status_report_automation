from pathlib import Path
from typing import List, Optional

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeoutError

from src.utils.logger import logger
from src.utils.retry_helper import with_retry

COLUMNS_TO_HIDE_BEFORE_EXPORT: List[str] = [
    "Previous Work Type",
    "Flags",
    "Field Job Reference",
    "Other Information",
    "Tech ID",
]


class FinishStatusReportPage:
    def __init__(self, page: Page):
        self.page = page

        self.reporting_menu = self.page.get_by_text("Reporting", exact=True)
        self.finish_status_entry = self.page.get_by_text("Finish Status Report", exact=True)
        self.filter_button = self.page.locator("div.dx-button-content:has(i.dx-icon-filter)").first
        self.columns_button = self.page.locator(
            'div.dx-button-content:has(i.dx-icon-columnchooser), '
            '[title="Columns"]'
        ).first
        self.refresh_button = self.page.locator(
            'div.dx-button-content:has(i.dx-icon-refresh), '
            '[title="Refresh"]'
        ).first
        self.export_button = self.page.locator("div.dx-button-content:has(i.dx-icon-exportxlsx)").first

    @with_retry(max_retries=3, base_delay=3)
    def open_finish_status_report(self):
        logger.info("Opening Reporting > Finish Status Report")
        self.page.wait_for_load_state("domcontentloaded", timeout=60000)
        self.page.wait_for_timeout(2000)

        self.reporting_menu.wait_for(state="visible", timeout=45000)
        self.reporting_menu.click(timeout=45000)
        self.page.wait_for_timeout(1500)

        self.finish_status_entry.wait_for(state="visible", timeout=45000)
        self.finish_status_entry.click(timeout=45000)

        # Allow heavier pages to complete key API rendering steps.
        self.page.wait_for_load_state("networkidle", timeout=60000)
        self.page.wait_for_timeout(3000)
        logger.info("Finish Status Report page opened")

    def _is_filter_enabled(self) -> bool:
        # DevExtreme toggles active state on the nearest button container.
        filter_button_container = self.filter_button.locator("xpath=ancestor::*[contains(@class,'dx-button')][1]")

        try:
            filter_button_container.wait_for(state="visible", timeout=20000)
        except PlaywrightTimeoutError:
            logger.warning("Filter button not visible yet; assuming filter is OFF")
            return False

        class_name = filter_button_container.get_attribute("class") or ""
        aria_pressed = filter_button_container.get_attribute("aria-pressed") or ""
        return "dx-state-active" in class_name or aria_pressed.lower() == "true"

    @with_retry(max_retries=2, base_delay=2)
    def disable_filter_if_enabled(self):
        logger.info("Checking filter status")
        if self._is_filter_enabled():
            logger.info("Filter is ON, switching it OFF")
            self.filter_button.click(timeout=20000)
            self.page.wait_for_timeout(1200)
        else:
            logger.info("Filter already OFF")

    def _columns_popup(self) -> Locator:
        return self.page.locator(
            ".dx-datagrid-column-chooser-list, .dx-datagrid-column-chooser-plain"
        ).last

    def _open_columns_chooser(self) -> Locator:
        logger.info("Opening Columns chooser")
        self.columns_button.wait_for(state="visible", timeout=45_000)
        self.columns_button.click(timeout=20_000)
        popup = self._columns_popup()
        popup.wait_for(state="visible", timeout=20_000)
        self.page.wait_for_timeout(800)
        return popup

    def _close_columns_chooser(self) -> None:
        close_btn = self.page.locator(".dx-closebutton").last
        try:
            close_btn.click(timeout=5_000)
        except Exception:
            self.page.keyboard.press("Escape")
        self.page.wait_for_timeout(800)

    def _uncheck_column(self, popup: Locator, column_name: str) -> None:
        # DevExtreme column chooser uses treeview nodes with aria-label on each column.
        node = popup.locator(f'li[role="treeitem"][aria-label="{column_name}"]')
        node.wait_for(state="attached", timeout=15_000)
        node.scroll_into_view_if_needed()
        node.wait_for(state="visible", timeout=10_000)

        checkbox = node.locator('[role="checkbox"]').first
        checkbox.wait_for(state="visible", timeout=10_000)
        aria_checked = (checkbox.get_attribute("aria-checked") or "").lower()
        classes = checkbox.get_attribute("class") or ""

        if aria_checked == "true" or "dx-checkbox-checked" in classes:
            logger.info("Unchecking column: %s", column_name)
            checkbox.locator(".dx-checkbox-container").click(timeout=10_000)
            self.page.wait_for_timeout(500)
            # Verify unchecked; retry once if still checked.
            aria_checked = (checkbox.get_attribute("aria-checked") or "").lower()
            if aria_checked == "true":
                checkbox.click(timeout=10_000)
                self.page.wait_for_timeout(500)
        else:
            logger.info("Column already unchecked: %s", column_name)

    def refresh_grid(self) -> None:
        logger.info("Clicking grid refresh button")
        self.refresh_button.wait_for(state="visible", timeout=45_000)
        self.refresh_button.click(timeout=20_000)
        self.page.wait_for_timeout(3000)
        logger.info("Grid refresh triggered")

    @with_retry(max_retries=2, base_delay=2)
    def hide_columns_before_export(
        self, columns: Optional[List[str]] = None,
    ) -> None:
        columns = columns or COLUMNS_TO_HIDE_BEFORE_EXPORT
        logger.info("Hiding %s columns before export", len(columns))
        popup = self._open_columns_chooser()
        for column_name in columns:
            self._uncheck_column(popup, column_name)
        self._close_columns_chooser()
        self.refresh_grid()
        logger.info("Column visibility updated for export")

    @with_retry(max_retries=3, base_delay=3)
    def export_report(self, download_dir: Path, filename: Optional[str] = None) -> Path:
        download_dir.mkdir(parents=True, exist_ok=True)
        logger.info("Clicking export button and waiting for download")
        self.export_button.wait_for(state="visible", timeout=55000)
        with self.page.expect_download(timeout=130_000) as download_info:
            self.export_button.click(timeout=55000)
        download = download_info.value
        target_name = filename or download.suggested_filename
        target_path = download_dir / target_name
        if target_path.exists():
            target_path.unlink()
        download.save_as(str(target_path))
        logger.info("Export saved to %s", target_path)
        return target_path
