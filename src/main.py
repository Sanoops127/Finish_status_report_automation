from dotenv import load_dotenv
import os
import sys
import asyncio
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from src.core.job_manager import JobManager
from src.paths import EDGE_AUTOMATION_PROFILE_DIR, STATUS_REPORT_DIR
from src.utils.browser_launcher import launch_edge_persistent_context
from src.utils.logger import logger

STOP_FLAG = Path(os.getenv("STOP_FLAG_PATH", "stop.flag"))
CHECK_INTERVAL_SECONDS = 30  # how often to check the flag while "sleeping"


def should_stop():
    return STOP_FLAG.exists()


def sleep_with_stop_check(total_seconds):
    """Sleep in small chunks so we can exit quickly if stop.flag appears."""
    elapsed = 0
    while elapsed < total_seconds:
        if should_stop():
            return True  # signal caller: stop requested
        chunk = min(CHECK_INTERVAL_SECONDS, total_seconds - elapsed)
        time.sleep(chunk)
        elapsed += chunk
    return False


def main():
    load_dotenv()

    # Clear any stale flag left over from a previous run/day
    if STOP_FLAG.exists():
        logger.info("Clearing stale stop flag from previous run.")
        STOP_FLAG.unlink()

    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    url = os.getenv("DEPOTNET_URL", "https://uat-sn.depotnet.co.uk/")
    username = os.getenv("DEPOTNET_USERNAME")
    password = os.getenv("DEPOTNET_PASSWORD")

    if not username or not password:
        logger.error("Credentials not found in environment variables.")
        sys.exit(1)

    sharepoint_site_url = (
        os.getenv("SHAREPOINT_FILE_URL")
        or os.getenv("SHAREPOINT_SITE_URL")
        or os.getenv("SHAREPOINT_LIBRARY_URL")
        or None
    )
    sharepoint_document_library_url = os.getenv("SHAREPOINT_DOCUMENT_LIBRARY_URL") or None
    export_filename = os.getenv("EXPORT_FILENAME") or None
    sharepoint_target_filename = os.getenv("SHAREPOINT_TARGET_FILENAME", "Finish Status Report.xlsx")
    powerbi_workspace_url = os.getenv("POWERBI_WORKSPACE_URL") or None
    powerbi_semantic_model_name = os.getenv("POWERBI_SEMANTIC_MODEL_NAME", "Pre Enablement Jobs")
    interval_minutes = int(os.getenv("JOB_INTERVAL_MINUTES", "15"))
    run_once = os.getenv("RUN_ONCE", "").strip().lower() in ("1", "true", "yes")

    edge_profile = Path(os.getenv("EDGE_USER_DATA_DIR", str(EDGE_AUTOMATION_PROFILE_DIR)))

    run_index = 0
    while True:
        if should_stop():
            logger.info("Stop flag detected before run start — shutting down.")
            STOP_FLAG.unlink(missing_ok=True)
            break

        run_index += 1
        logger.info("Starting automation run #%s", run_index)

        try:
            with sync_playwright() as p:
                try:
                    context = launch_edge_persistent_context(p, edge_profile)
                except Exception as exc:
                    logger.error("Browser failed to launch: %s", exc)
                    if run_once:
                        sys.exit(1)
                    raise

                try:
                    page = context.pages[0] if context.pages else context.new_page()
                    manager = JobManager(
                        page,
                        export_dir=STATUS_REPORT_DIR,
                        sharepoint_site_url=sharepoint_site_url,
                        sharepoint_library_url=sharepoint_document_library_url,
                        export_filename=export_filename,
                        sharepoint_target_filename=sharepoint_target_filename,
                        powerbi_workspace_url=powerbi_workspace_url,
                        powerbi_semantic_model_name=powerbi_semantic_model_name,
                    )

                    manager.run_finish_status_report_job(url, username, password)
                finally:
                    context.close()
        except Exception as exc:
            logger.exception("Run #%s failed: %s", run_index, exc)

        if run_once:
            logger.info("RUN_ONCE is set; exiting after single run.")
            break

        logger.info(
            "Waiting %s minutes before next run (set JOB_INTERVAL_MINUTES or RUN_ONCE=1 to change).",
            interval_minutes,
        )
        stop_requested = sleep_with_stop_check(interval_minutes * 60)
        if stop_requested:
            logger.info("Stop flag detected during wait — shutting down.")
            STOP_FLAG.unlink(missing_ok=True)
            break


if __name__ == "__main__":
    main()