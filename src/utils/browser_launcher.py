import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import BrowserContext, Playwright

from src.paths import PROJECT_ROOT
from src.utils.logger import logger

_PROFILE_LOCK_FILES = (
    "SingletonLock",
    "SingletonCookie",
    "lockfile",
    "DevToolsActivePort",
)


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in ("1", "true", "yes", "on")


def reset_edge_profile(profile_dir: Path) -> None:
    if profile_dir.exists():
        shutil.rmtree(profile_dir, ignore_errors=True)
        logger.info("Removed Edge automation profile: %s", profile_dir)
    profile_dir.mkdir(parents=True, exist_ok=True)


def clear_profile_locks(profile_dir: Path) -> None:
    for name in _PROFILE_LOCK_FILES:
        path = profile_dir / name
        if not path.exists():
            continue
        try:
            path.unlink()
            logger.info("Removed stale profile lock: %s", path)
        except OSError as exc:
            logger.warning("Could not remove profile lock %s: %s", path, exc)


def terminate_stale_edge_for_profile(profile_dir: Path) -> None:
    """Kill msedge.exe processes that were started with this automation profile."""
    profile_token = str(profile_dir.resolve()).lower()
    try:
        result = subprocess.run(
            [
                "wmic",
                "process",
                "where",
                "name='msedge.exe'",
                "get",
                "ProcessId,CommandLine",
                "/format:csv",
            ],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        logger.warning("Could not inspect Edge processes: %s", exc)
        return

    pids: list[str] = []
    for line in result.stdout.splitlines():
        if "msedge.exe" not in line.lower():
            continue
        if profile_token not in line.lower():
            continue
        parts = [part.strip() for part in line.split(",") if part.strip()]
        for part in reversed(parts):
            if part.isdigit():
                pids.append(part)
                break

    for pid in pids:
        logger.info("Stopping stale Edge process for automation profile (pid=%s)", pid)
        subprocess.run(
            ["taskkill", "/F", "/PID", pid, "/T"],
            capture_output=True,
            check=False,
        )


def ensure_playwright_driver() -> None:
    """Install Playwright browser driver once after a fresh clone (pip install is not enough)."""
    if _env_bool("SKIP_PLAYWRIGHT_INSTALL"):
        return

    marker = PROJECT_ROOT / ".playwright-msedge-installed"
    if marker.exists() and not _env_bool("FORCE_PLAYWRIGHT_INSTALL"):
        return

    logger.info("Installing Playwright msedge driver (first run after clone)...")
    result = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "msedge"],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        logger.warning(
            "playwright install msedge returned %s: %s",
            result.returncode,
            (result.stderr or result.stdout or "").strip(),
        )
    else:
        marker.touch()
        logger.info("Playwright msedge driver is ready.")


def _launch_kwargs(profile_dir: Path) -> dict:
    headless = _env_bool("BROWSER_HEADLESS", default=False)
    kwargs = {
        "user_data_dir": str(profile_dir),
        "channel": os.getenv("EDGE_CHANNEL", "msedge").strip() or "msedge",
        "headless": headless,
        "accept_downloads": True,
        "permissions": ["clipboard-read", "clipboard-write"],
        "timeout": int(os.getenv("BROWSER_LAUNCH_TIMEOUT_MS", "120000")),
    }

    edge_executable = os.getenv("EDGE_EXECUTABLE_PATH", "").strip()
    if edge_executable and Path(edge_executable).is_file():
        kwargs["executable_path"] = edge_executable
        kwargs.pop("channel")

    return kwargs


def _try_launch(playwright: Playwright, profile_dir: Path) -> BrowserContext:
    terminate_stale_edge_for_profile(profile_dir)
    clear_profile_locks(profile_dir)
    return playwright.chromium.launch_persistent_context(**_launch_kwargs(profile_dir))


def launch_edge_persistent_context(
    playwright: Playwright,
    profile_dir: Path,
) -> BrowserContext:
    """
    Launch Edge with a persistent profile.

    After a fresh repo clone, failed launches can leave a half-built profile that
    blocks the normal Edge sign-in screen. We recover the same way as a manual
    profile delete: wipe the profile and launch again.
    """
    ensure_playwright_driver()
    profile_dir.mkdir(parents=True, exist_ok=True)

    if _env_bool("EDGE_PROFILE_RESET"):
        reset_edge_profile(profile_dir)

    try:
        return _try_launch(playwright, profile_dir)
    except Exception as first_error:
        error_text = str(first_error)
        logger.warning("First Edge launch failed: %s", first_error)

        recoverable = any(
            token in error_text
            for token in (
                "Browser window not found",
                "getWindowForTarget",
                "Browser closed",
                "Target page, context or browser has been closed",
                "Failed to launch",
            )
        )
        if not recoverable:
            raise

        logger.warning(
            "Recovering like a manual profile delete — wiping %s and retrying once.",
            profile_dir,
        )
        reset_edge_profile(profile_dir)
        time.sleep(2)

        try:
            return _try_launch(playwright, profile_dir)
        except Exception as second_error:
            logger.error(
                "Edge still failed after profile reset.\n"
                "On production after a fresh clone, run once:\n"
                "  python -m src.setup_browser\n"
                "Then start the automation and complete the Edge / Microsoft sign-in "
                "when the browser opens.\n"
                "Profile: %s",
                profile_dir,
            )
            raise second_error from first_error
