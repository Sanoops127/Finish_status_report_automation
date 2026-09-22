import re
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlparse

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError

from src.utils.logger import logger


class SharePointUploader:
    """Replace a SharePoint workbook with a local .xlsx (exact copy of portal export)."""

    def __init__(self, page: Page):
        self.page = page

    def replace_workbook(
        self,
        site_url: str,
        file_name: str,
        local_path: Path,
        library_url: Optional[str] = None,
    ) -> None:
        if not local_path.is_file():
            raise FileNotFoundError(f"Workbook not found: {local_path}")

        path_str = str(local_path.resolve())
        logger.info("Replacing SharePoint file %s with %s", file_name, local_path.name)

        urls = self._candidate_urls(site_url, library_url)
        last_error = None
        for url in urls:
            try:
                if self._try_replace_at_url(url, file_name, path_str):
                    logger.info("SharePoint file replaced successfully")
                    return
            except Exception as exc:
                last_error = exc
                logger.warning("Replace attempt failed at %s: %s", url, exc)

        raise RuntimeError(
            f"Could not replace {file_name} on SharePoint. "
            f"Set SHAREPOINT_LIBRARY_URL to the document library folder URL. Last error: {last_error}"
        )

    def _candidate_urls(self, site_url: str, library_url: Optional[str]) -> List[str]:
        urls: List[str] = []
        if library_url:
            urls.append(library_url)
        derived = self._derive_library_url(site_url)
        if derived:
            urls.append(derived)
        if site_url not in urls:
            urls.append(site_url)
        return urls

    def _derive_library_url(self, site_url: str) -> Optional[str]:
        # e.g. .../Skander%20Sandbox/SitePages/Home.aspx -> .../Shared Documents/Forms/AllItems.aspx
        base = re.split(r"/SitePages/", site_url, flags=re.I)[0]
        if base == site_url:
            parsed = urlparse(site_url)
            parts = parsed.path.strip("/").split("/")
            if len(parts) >= 2:
                base = f"{parsed.scheme}://{parsed.netloc}/{'/'.join(parts[:2])}"
        if not base:
            return None
        return f"{base}/Shared Documents/Forms/AllItems.aspx"

    def _try_replace_at_url(self, url: str, file_name: str, path_str: str) -> bool:
        logger.info("Opening SharePoint location: %s", url)
        self.page.goto(url, timeout=120_000, wait_until="domcontentloaded")
        self.page.wait_for_timeout(5000)

        if self._try_file_context_replace(file_name, path_str):
            return True
        if self._try_upload_with_replace(path_str):
            return True
        return False

    def _try_file_context_replace(self, file_name: str, path_str: str) -> bool:
        """File row ... menu -> Replace."""
        stem = Path(file_name).stem
        for candidate in (file_name, stem):
            row = self.page.get_by_role("row").filter(has_text=candidate).first
            try:
                row.wait_for(state="visible", timeout=8_000)
            except PlaywrightTimeoutError:
                continue

            try:
                row.hover()
                self.page.wait_for_timeout(500)
            except Exception:
                pass

            more = row.locator(
                'button[aria-label*="More"], button[data-automationid="moreActions"], '
                'button[title*="More"]'
            ).first
            try:
                more.click(timeout=5_000)
            except Exception:
                continue

            for label in ("Replace", "Upload new version", "Replace file"):
                item = self.page.get_by_role("menuitem", name=re.compile(label, re.I))
                try:
                    with self.page.expect_file_chooser(timeout=30_000) as fc:
                        item.first.click(timeout=5_000)
                    fc.value.set_files(path_str)
                    self.page.wait_for_timeout(5000)
                    self._confirm_replace_dialog()
                    logger.info("Replaced via file menu: %s", label)
                    return True
                except Exception:
                    continue
        return False

    def _try_upload_with_replace(self, path_str: str) -> bool:
        if self._try_upload_command_file_chooser(path_str):
            self._confirm_replace_dialog()
            return True
        if self._try_role_upload_file_chooser(path_str):
            self._confirm_replace_dialog()
            return True
        if self._try_hidden_file_inputs(path_str):
            self._confirm_replace_dialog()
            return True
        return False

    def _confirm_replace_dialog(self) -> None:
        for locator in (
            self.page.get_by_role("button", name=re.compile(r"^Replace$", re.I)),
            self.page.locator('[data-automationid="conflictReplaceButton"]'),
            self.page.get_by_text("Replace", exact=True),
        ):
            try:
                btn = locator.first
                btn.wait_for(state="visible", timeout=8_000)
                btn.click(timeout=10_000)
                self.page.wait_for_timeout(5000)
                logger.info("Confirmed SharePoint replace dialog")
                return
            except Exception:
                continue

    def _try_hidden_file_inputs(self, path_str: str) -> bool:
        inputs = self.page.locator('input[type="file"]')
        try:
            n = inputs.count()
        except PlaywrightTimeoutError:
            return False
        for i in range(n):
            candidate = inputs.nth(i)
            try:
                candidate.set_input_files(path_str, timeout=30_000)
                logger.info("Upload triggered via file input")
                self.page.wait_for_timeout(5000)
                return True
            except Exception:
                continue
        return False

    def _try_upload_command_file_chooser(self, path_str: str) -> bool:
        upload = self.page.locator('[data-automationid="uploadCommand"]').first
        try:
            upload.wait_for(state="visible", timeout=15_000)
        except PlaywrightTimeoutError:
            return False
        try:
            with self.page.expect_file_chooser(timeout=60_000) as fc:
                upload.click(timeout=30_000)
                files_item = self.page.get_by_role("menuitem", name="Files")
                if files_item.count() > 0:
                    try:
                        files_item.first.click(timeout=5_000)
                    except PlaywrightTimeoutError:
                        pass
            fc.value.set_files(path_str)
            logger.info("Upload triggered via Upload > Files")
            self.page.wait_for_timeout(5000)
            return True
        except Exception:
            return False

    def _try_role_upload_file_chooser(self, path_str: str) -> bool:
        try:
            btn = self.page.get_by_role("button", name="Upload").first
            btn.wait_for(state="visible", timeout=10_000)
        except PlaywrightTimeoutError:
            return False
        try:
            with self.page.expect_file_chooser(timeout=60_000) as fc:
                btn.click(timeout=30_000)
            fc.value.set_files(path_str)
            logger.info("Upload triggered via Upload button")
            self.page.wait_for_timeout(5000)
            return True
        except Exception:
            return False
