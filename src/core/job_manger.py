from pathlib import Path
from typing import Optional
from src.core.report_transformer import read_values_without_header, rows_to_html_table
from src.portal.sharepoint_excel_editor import SharePointExcelEditor
from src.portal.finish_status_report_page import FinishStatusReportPage
from src.portal.login_page import LoginPage
from src.portal.powerbi_workspace_page import PowerBiWorkspacePage
from src.utils.logger import logger


class JobManager:
    def __init__(
        self,
        page,
        export_dir: Path,
        sharepoint_site_url: Optional[str] = None,
        sharepoint_library_url: Optional[str] = None,
        export_filename: Optional[str] = None,
        sharepoint_target_filename: str = "test_finish_status_report.xlsx",
        powerbi_workspace_url: Optional[str] = None,
        powerbi_semantic_model_name: str = "Pre Enablement Jobs",
    ):
        self.page = page
        self.export_dir = export_dir
        self.sharepoint_site_url = sharepoint_site_url
        self.sharepoint_library_url = sharepoint_library_url
        self.export_filename = export_filename
        self.sharepoint_target_filename = sharepoint_target_filename
        self.powerbi_workspace_url = powerbi_workspace_url
        self.powerbi_semantic_model_name = powerbi_semantic_model_name
        self.login_page = LoginPage(page)
        self.finish_status_report_page = FinishStatusReportPage(page)

    def run_finish_status_report_job(self, url: str, username: str, password: str):
        logger.info("Starting Finish Status Report automation job")
        self.login_page.navigate(url)
        self.login_page.login(username, password)
        self.finish_status_report_page.open_finish_status_report()
        self.finish_status_report_page.disable_filter_if_enabled()
        self.finish_status_report_page.hide_columns_before_export()
        saved_path = self.finish_status_report_page.export_report(
            self.export_dir,
            filename=self.export_filename,
        )

        if self.sharepoint_site_url:
            rows = read_values_without_header(saved_path)
            # Filter out rows without JOB ID (first column)
            rows = [row for row in rows if row and row[0] and str(row[0]).strip()]
            data_html = rows_to_html_table(rows)
            sp_page = self.page.context.new_page()
            try:
                SharePointExcelEditor(sp_page).update_file_values(
                    sharepoint_home_url=self.sharepoint_site_url,
                    file_name=self.sharepoint_target_filename,
                    data_html=data_html,
                )
                logger.info("SharePoint online Excel update finished")
            finally:
                sp_page.close()
        else:
            logger.info("SHAREPOINT_SITE_URL not set; skipping SharePoint update")

        if self.powerbi_workspace_url:
            pbi_page = self.page.context.new_page()
            try:
                PowerBiWorkspacePage(pbi_page).navigate_and_refresh_semantic_model(
                    self.powerbi_workspace_url,
                    self.powerbi_semantic_model_name,
                )
                logger.info("Power BI semantic model refresh finished")
            finally:
                pbi_page.close()
        else:
            logger.info("POWERBI_WORKSPACE_URL not set; skipping Power BI refresh")

        logger.info("Finish Status Report automation completed")
