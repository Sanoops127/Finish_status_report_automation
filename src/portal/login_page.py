from playwright.sync_api import Page
from src.utils.logger import logger
from src.utils.retry_helper import with_retry


class LoginPage:
    def __init__(self, page: Page):
        self.page = page

        self.username_input = "input[name='username']"
        self.password_input = "input[name='password']"
        self.login_button = "button.signin-button"

    def navigate(self, url: str):
        self.page.goto(url, timeout=60000, wait_until="domcontentloaded")

    def is_logged_in(self) -> bool:
        """Checks if the user is already logged in by looking for specific dashboard elements."""
        try:
            # Check URL first
            if "dashboard" in self.page.url.lower():
                return True
            
            # Check for the Reporting menu which is visible after login
            reporting_menu = self.page.get_by_text("Reporting", exact=True)
            if reporting_menu.is_visible():
                return True
                
            return False
        except Exception:
            return False

    @with_retry(max_retries=3, base_delay=2)
    def login(self, username: str, password: str):
        logger.info(f"Checking login status for user: {username}")
        if self.is_logged_in():
            logger.info("Already logged in. Skipping login form.")
            return

        logger.info(f"Starting login process for user: {username}")
        try:
            self.page.wait_for_load_state("domcontentloaded", timeout=60000)
            self.page.fill(self.username_input, username)
            self.page.fill(self.password_input, password)
            self.page.click(self.login_button)

            # Slow portal: allow additional loading buffer after auth.
            self.page.wait_for_timeout(2000)
            self.page.wait_for_load_state("networkidle", timeout=50000)

            reporting_menu = self.page.get_by_text("Reporting", exact=True)
            if "dashboard" in self.page.url.lower() or reporting_menu.is_visible():
                logger.info("Login successful. Reached dashboard.")
            else:
                logger.warning(f"Login executed but did not reach dashboard. Current URL: {self.page.url}")

        except Exception as e:
            logger.error(f"Login failed due to an error: {e}")
            raise
