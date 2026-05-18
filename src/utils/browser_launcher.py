import os
import shutil
import time
from pathlib import Path
from typing import Any, Optional

from playwright.sync_api import BrowserContext, Playwright

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


def reset_edge_profile(profile_dir: Path) -> None:
    if profile_dir.exists():
        shutil.rmtree(profile_dir, ignore_errors=True)
        logger.warning("Edge automation profile reset: %s", profile_dir)
    profile_dir.mkdir(parents=True, exist_ok=True)


def resolve_edge_executable() -> Optional[str]:
    configured = os.getenv("EDGE_EXECUTABLE_PATH", "").strip()
    if configured:
        path = Path(configured)
        if path.is_file():
            return str(path)
        logger.error("EDGE_EXECUTABLE_PATH does not exist: %s", configured)
        return None

    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
        / "Microsoft"
        / "Edge"
        / "Application"
        / "msedge.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
        / "Microsoft"
        / "Edge"
        / "Application"
        / "msedge.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)

    return None


def _build_launch_kwargs(
    profile_dir: Path,
    *,
    headless: bool,
    edge_executable: Optional[str],
    channel: Optional[str],
) -> dict[str, Any]:
    args = [
        "--disable-dev-shm-usage",
        "--no-first-run",
        "--disable-extensions",
        "--disable-background-networking",
        "--disable-background-timer-throttling",
        "--disable-renderer-backgrounding",
    ]
    if headless:
        args.extend(["--disable-gpu", "--headless=new"])
    else:
        args.extend(["--start-maximized", "--window-size=1920,1080"])

    kwargs: dict[str, Any] = {
        "user_data_dir": str(profile_dir),
        "headless": headless,
        "accept_downloads": True,
        "permissions": ["clipboard-read", "clipboard-write"],
        "args": args,
        "timeout": int(os.getenv("BROWSER_LAUNCH_TIMEOUT_MS", "120000")),
        "viewport": None if not headless else {"width": 1920, "height": 1080},
    }

    if edge_executable:
        kwargs["executable_path"] = edge_executable
    elif channel:
        kwargs["channel"] = channel

    return kwargs


def launch_edge_persistent_context(
    playwright: Playwright,
    profile_dir: Path,
) -> BrowserContext:
    headless = _env_bool("BROWSER_HEADLESS", default=False)
    channel = os.getenv("EDGE_CHANNEL", "msedge").strip() or "msedge"
    retries = max(1, int(os.getenv("BROWSER_LAUNCH_RETRIES", "3")))

    profile_dir.mkdir(parents=True, exist_ok=True)
    if _env_bool("EDGE_PROFILE_RESET"):
        reset_edge_profile(profile_dir)
    else:
        clear_profile_locks(profile_dir)

    edge_executable = resolve_edge_executable()
    if edge_executable:
        logger.info("Using Edge executable: %s", edge_executable)
    else:
        logger.warning("Edge executable not found on disk; falling back to Playwright channel lookup.")

    launch_plans: list[dict[str, Any]] = []
    launch_plans.append(
        _build_launch_kwargs(
            profile_dir,
            headless=headless,
            edge_executable=edge_executable,
            channel=None if edge_executable else channel,
        )
    )

    if edge_executable:
        launch_plans.append(
            _build_launch_kwargs(
                profile_dir,
                headless=headless,
                edge_executable=None,
                channel=channel,
            )
        )

    if channel == "msedge":
        launch_plans.append(
            _build_launch_kwargs(
                profile_dir,
                headless=headless,
                edge_executable=None,
                channel="msedge-beta",
            )
        )

    if not headless and _env_bool("BROWSER_HEADLESS_FALLBACK", default=True):
        launch_plans.append(
            _build_launch_kwargs(
                profile_dir,
                headless=True,
                edge_executable=edge_executable,
                channel=None if edge_executable else channel,
            )
        )

    last_error: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        for plan_index, launch_kwargs in enumerate(launch_plans, start=1):
            label = launch_kwargs.get("executable_path") or launch_kwargs.get("channel", "chromium")
            logger.info(
                "Launching Edge (attempt %s/%s, plan %s/%s, headless=%s, target=%s)",
                attempt,
                retries,
                plan_index,
                len(launch_plans),
                launch_kwargs["headless"],
                label,
            )
            try:
                return playwright.chromium.launch_persistent_context(**launch_kwargs)
            except Exception as exc:
                last_error = exc
                logger.warning("Browser launch failed: %s", exc)

        if attempt < retries:
            clear_profile_locks(profile_dir)
            time.sleep(2)

    _log_launch_troubleshooting(headless=headless, profile_dir=profile_dir)
    assert last_error is not None
    raise last_error


def _log_launch_troubleshooting(*, headless: bool, profile_dir: Path) -> None:
    logger.error(
        "Browser failed to launch after all retries. Common production causes on Windows:\n"
        "  1) Task Scheduler running without an interactive desktop — use "
        "'Run only when user is logged on' or set BROWSER_HEADLESS=1.\n"
        "  2) Stale or corrupted profile — stop all msedge.exe processes, then set "
        "EDGE_PROFILE_RESET=1 once (you will need to sign in again).\n"
        "  3) Wrong Edge install — set EDGE_EXECUTABLE_PATH to the 64-bit msedge.exe "
        "under Program Files\\Microsoft\\Edge\\Application.\n"
        "  4) Outdated Playwright — run: pip install -U playwright && playwright install\n"
        "Profile directory: %s | headless=%s",
        profile_dir,
        headless,
    )
