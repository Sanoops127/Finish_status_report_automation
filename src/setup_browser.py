"""
One-time bootstrap after cloning the repo on a new machine or production VM.

Usage:
    python -m src.setup_browser

This resets the automation profile, installs the Playwright driver, opens Edge,
and waits for you to finish the Microsoft Edge sign-in screen before exiting.
"""

import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

from src.paths import EDGE_AUTOMATION_PROFILE_DIR
from src.utils.browser_launcher import launch_edge_persistent_context, reset_edge_profile
from src.utils.logger import logger


def main() -> int:
    load_dotenv()
    profile_dir = Path(os.getenv("EDGE_USER_DATA_DIR", str(EDGE_AUTOMATION_PROFILE_DIR)))

    logger.info("Bootstrap: resetting profile at %s", profile_dir)
    reset_edge_profile(profile_dir)

    os.environ["FORCE_PLAYWRIGHT_INSTALL"] = "1"

    logger.info("Bootstrap: launching Edge — complete the sign-in / welcome screen, then press Enter here.")
    with sync_playwright() as playwright:
        context = launch_edge_persistent_context(playwright, profile_dir)
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("about:blank")
        print(
            "\nEdge is open. Finish the Microsoft Edge sign-in (or click Continue without signing in).\n"
            "When done, press Enter in this terminal to save the profile and exit.\n"
        )
        try:
            input()
        except KeyboardInterrupt:
            print()
        finally:
            context.close()

    logger.info("Bootstrap complete. Profile saved at %s", profile_dir)
    logger.info("You can now run: python -m src.main")
    return 0


if __name__ == "__main__":
    sys.exit(main())
