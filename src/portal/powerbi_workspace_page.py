import re

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeoutError

from src.utils.logger import logger
from src.utils.retry_helper import with_retry


class PowerBiWorkspacePage:
    """Power BI service workspace list: refresh a semantic model (not the report with the same name)."""

    def __init__(self, page: Page):
        self.page = page

    @with_retry(max_retries=2, base_delay=5)
    def navigate_and_refresh_semantic_model(self, workspace_url: str, model_name: str) -> None:
        logger.info("Opening Power BI workspace list")
        self.page.goto(workspace_url, timeout=120_000, wait_until="domcontentloaded")
        self.page.wait_for_timeout(6000)

        self._apply_keyword_filter(model_name)
        row = self._find_semantic_model_row(model_name)
        row.scroll_into_view_if_needed()
        row.hover()
        self.page.wait_for_timeout(1000)
        self._click_refresh_now(row)
        self.page.wait_for_timeout(3000)
        logger.info("Power BI refresh triggered for semantic model: %s", model_name)

    def _apply_keyword_filter(self, model_name: str) -> None:
        candidates = [
            self.page.get_by_placeholder("Filter by keyword"),
            self.page.get_by_placeholder(re.compile(r"filter.*keyword", re.I)),
            self.page.locator('input[type="search"]').first,
        ]
        for loc in candidates:
            try:
                first = loc.first
                first.wait_for(state="visible", timeout=5_000)
                first.fill("")
                first.fill(model_name)
                self.page.wait_for_timeout(2500)
                logger.info("Applied workspace keyword filter")
                return
            except PlaywrightTimeoutError:
                continue
            except Exception:
                continue

    def _row_name_equals(self, row: Locator, model_name: str) -> bool:
        """Match the workspace **Name** column exactly, not substring (avoids Depotnet - Pre Enablement Job vs Pre Enablement Jobs)."""
        want = model_name.strip().lower()
        try:
            cells = row.locator('[role="gridcell"]')
            if cells.count() > 0:
                name_text = cells.nth(0).inner_text(timeout=2_000).strip()
                if name_text.lower() == want:
                    return True
        except Exception:
            pass
        try:
            link = row.get_by_role("link", name=re.compile(r"^" + re.escape(model_name) + r"$", re.I))
            if link.count() > 0:
                return True
        except Exception:
            pass
        try:
            first_line = row.inner_text(timeout=2_000).split("\n")[0].strip()
            if first_line.lower() == want:
                return True
        except Exception:
            pass
        return False

    def _row_is_semantic_model(self, row: Locator) -> bool:
        try:
            text = row.inner_text(timeout=2_000)
        except Exception:
            return False
        if not re.search(r"semantic", text, re.I):
            return False
        # Skip header row
        if re.search(r"\bName\b", text) and re.search(r"\bType\b", text):
            return False
        return True

    def _find_semantic_model_row(self, model_name: str) -> Locator:
        semantic = re.compile(r"semantic", re.I)

        # Scan data rows: require exact name match + semantic type (not Report).
        rows = self.page.locator('[role="row"]')
        try:
            n = rows.count()
        except Exception:
            n = 0

        for i in range(min(n, 120)):
            row = rows.nth(i)
            try:
                if not row.is_visible(timeout=500):
                    continue
            except Exception:
                continue
            if not self._row_is_semantic_model(row):
                continue
            if not self._row_name_equals(row, model_name):
                continue
            return row

        # Fallback: listitem rows (some Power BI builds)
        items = self.page.get_by_role("listitem")
        try:
            m = items.count()
        except Exception:
            m = 0
        for i in range(min(m, 80)):
            row = items.nth(i)
            try:
                text = row.inner_text(timeout=2_000)
            except Exception:
                continue
            if not semantic.search(text):
                continue
            if not self._row_name_equals(row, model_name):
                continue
            return row

        raise RuntimeError(
            f"Could not find a semantic model row with exact name {model_name!r}. "
            "Check POWERBI_SEMANTIC_MODEL_NAME matches the Name column exactly (e.g. Pre Enablement Jobs)."
        )

    def _click_refresh_now(self, row: Locator) -> None:
        # Row action icons often appear after hover.
        for locator in (
            row.get_by_role("button", name=re.compile(r"refresh\s*now", re.I)),
            row.locator('button[aria-label*="Refresh now"]'),
            row.locator('button[title*="Refresh now"]'),
            row.locator('[aria-label*="Refresh now"]'),
            row.locator('button[aria-label*="Refresh"]').first,
        ):
            try:
                target = locator.first
                target.wait_for(state="visible", timeout=5_000)
                target.click(timeout=15_000)
                return
            except Exception:
                continue

        raise RuntimeError(
            'Could not click "Refresh now" for the semantic model row. '
            "UI may have changed; try updating selectors in powerbi_workspace_page.py."
        )
