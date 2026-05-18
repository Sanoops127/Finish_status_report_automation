import os
import shutil
import socket
import subprocess
import sys
import time
import winreg
from pathlib import Path
from typing import Any, Optional

from playwright.sync_api import Browser, BrowserContext, Playwright

from src.paths import PROJECT_ROOT
from src.utils.logger import logger

_PROFILE_LOCK_FILES = (
    "SingletonLock",
    "SingletonCookie",
    "lockfile",
    "DevToolsActivePort",
)

_RECOVERABLE_LAUNCH_ERRORS = (
    "Browser window not found",
    "getWindowForTarget",
    "Browser closed",
    "Target page, context or browser has been closed",
    "Failed to launch",
    "ECONNREFUSED",
)

_EDGE_REGISTRY_KEYS = (
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
    (
        winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
    ),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
)


class ManagedBrowserContext:
    """BrowserContext wrapper that also stops CDP-spawned Edge processes on close."""

    def __init__(
        self,
        context: BrowserContext,
        *,
        edge_process: Optional[subprocess.Popen] = None,
        cdp_browser: Optional[Browser] = None,
    ):
        self._context = context
        self._edge_process = edge_process
        self._cdp_browser = cdp_browser

    def __getattr__(self, name: str) -> Any:
        return getattr(self._context, name)

    def close(self) -> None:
        try:
            self._context.close()
        finally:
            if self._cdp_browser is not None:
                try:
                    self._cdp_browser.close()
                except Exception:
                    pass
            if self._edge_process is not None and self._edge_process.poll() is None:
                self._edge_process.terminate()
                try:
                    self._edge_process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self._edge_process.kill()


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in ("1", "true", "yes", "on")


def _run_cmd(*cmd: str, timeout: int = 60) -> subprocess.CompletedProcess:
    return subprocess.run(
        list(cmd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )


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


def _is_fresh_profile(profile_dir: Path) -> bool:
    return not (profile_dir / "Default").exists()


def _edge_from_registry() -> list[Path]:
    found: list[Path] = []
    for hive, subkey in _EDGE_REGISTRY_KEYS:
        try:
            with winreg.OpenKey(hive, subkey) as key:
                value, _ = winreg.QueryValueEx(key, "")
        except OSError:
            continue
        if value:
            path = Path(str(value).strip('"'))
            if path.is_file():
                found.append(path)
    return found


def _edge_from_where() -> list[Path]:
    result = _run_cmd("where.exe", "msedge.exe", timeout=15)
    found: list[Path] = []
    for line in result.stdout.splitlines():
        candidate = Path(line.strip().strip('"'))
        if candidate.is_file():
            found.append(candidate)
    return found


def _edge_from_glob_search() -> list[Path]:
    ps_script = (
        "Get-ChildItem -Path @("
        "'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe', "
        "'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'"
        ") -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName"
    )
    result = _run_cmd("powershell", "-NoProfile", "-Command", ps_script, timeout=20)
    found: list[Path] = []
    for line in result.stdout.splitlines():
        candidate = Path(line.strip())
        if candidate.is_file():
            found.append(candidate)
    return found


def _sort_edge_executables(paths: list[Path]) -> list[Path]:
    def rank(path: Path) -> tuple[int, str]:
        text = str(path).lower()
        if "program files (x86)" in text:
            return (1, text)
        if "program files" in text:
            return (0, text)
        return (2, text)

    seen: set[str] = set()
    unique: list[Path] = []
    for path in sorted(paths, key=rank):
        key = str(path.resolve())
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def resolve_edge_executables() -> list[Path]:
    configured = os.getenv("EDGE_EXECUTABLE_PATH", "").strip()
    if configured:
        path = Path(configured)
        return [path] if path.is_file() else []

    candidates: list[Path] = []
    candidates.extend(
        [
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        ]
    )
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(
            Path(local_app_data) / "Microsoft" / "Edge" / "Application" / "msedge.exe"
        )
    program_files = os.environ.get("ProgramFiles")
    program_files_x86 = os.environ.get("ProgramFiles(x86)")
    if program_files:
        candidates.append(Path(program_files) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
    if program_files_x86:
        candidates.append(
            Path(program_files_x86) / "Microsoft" / "Edge" / "Application" / "msedge.exe"
        )

    candidates.extend(_edge_from_registry())
    candidates.extend(_edge_from_where())
    candidates.extend(_edge_from_glob_search())

    return _sort_edge_executables([path for path in candidates if path.is_file()])


def terminate_stale_edge_for_profile(profile_dir: Path) -> None:
    """Kill msedge.exe processes that were started with this automation profile."""
    profile_path = str(profile_dir.resolve())
    ps_script = (
        "$profile = '{profile}'; "
        "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
        "Where-Object {{ $_.CommandLine -and ($_.CommandLine -like \"*$profile*\") }} | "
        "ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }}"
    ).format(profile=profile_path.replace("'", "''"))

    result = _run_cmd("powershell", "-NoProfile", "-Command", ps_script, timeout=20)
    if result.returncode != 0 and result.stderr.strip():
        logger.debug("PowerShell edge cleanup: %s", result.stderr.strip())


def ensure_playwright_driver() -> None:
    """Install Playwright browser driver once after a fresh clone (pip install is not enough)."""
    if _env_bool("SKIP_PLAYWRIGHT_INSTALL"):
        return

    marker = PROJECT_ROOT / ".playwright-msedge-installed"
    if marker.exists() and not _env_bool("FORCE_PLAYWRIGHT_INSTALL"):
        return

    logger.info("Installing Playwright msedge driver (first run after clone)...")
    result = _run_cmd(sys.executable, "-m", "playwright", "install", "msedge", timeout=180)
    if result.returncode != 0:
        logger.warning(
            "playwright install msedge returned %s: %s",
            result.returncode,
            (result.stderr or result.stdout or "").strip(),
        )
    else:
        marker.touch()
        logger.info("Playwright msedge driver is ready.")


def _persistent_launch_kwargs(profile_dir: Path, *, plan: dict[str, Any]) -> dict[str, Any]:
    headless = _env_bool("BROWSER_HEADLESS", default=False)
    kwargs: dict[str, Any] = {
        "user_data_dir": str(profile_dir),
        "headless": headless,
        "accept_downloads": True,
        "permissions": ["clipboard-read", "clipboard-write"],
        "timeout": int(os.getenv("BROWSER_LAUNCH_TIMEOUT_MS", "60000")),
    }
    kwargs.update(plan)
    return kwargs


def _launch_plans(edge_executables: list[Path]) -> list[tuple[str, dict[str, Any]]]:
    channel = os.getenv("EDGE_CHANNEL", "msedge").strip() or "msedge"
    plans: list[tuple[str, dict[str, Any]]] = []
    prefer_cdp = _env_bool("EDGE_USE_CDP", default=True)

    if prefer_cdp:
        for edge_exe in edge_executables:
            plans.append(
                (f"cdp:exe:{edge_exe}", {"executable_path": str(edge_exe), "_cdp": True})
            )

    for edge_exe in edge_executables:
        plans.append((f"persistent:exe:{edge_exe}", {"executable_path": str(edge_exe)}))

    if not prefer_cdp:
        for edge_exe in edge_executables:
            plans.append(
                (f"cdp:exe:{edge_exe}", {"executable_path": str(edge_exe), "_cdp": True})
            )

    plans.append((f"persistent:channel:{channel}", {"channel": channel}))

    if channel == "msedge":
        plans.append(("persistent:channel:msedge-beta", {"channel": "msedge-beta"}))

    return plans


def _pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _launch_via_cdp(
    playwright: Playwright,
    profile_dir: Path,
    edge_executable: str,
) -> ManagedBrowserContext:
    port = int(os.getenv("EDGE_CDP_PORT", "0")) or _pick_free_port()
    fresh_profile = _is_fresh_profile(profile_dir)

    cmd = [
        edge_executable,
        f'--user-data-dir={profile_dir}',
        f"--remote-debugging-port={port}",
        "--no-default-browser-check",
    ]
    if not fresh_profile:
        cmd.append("--no-first-run")

    logger.info(
        "Launching Edge via CDP port %s using %s (fresh_profile=%s)",
        port,
        edge_executable,
        fresh_profile,
    )
    proc = subprocess.Popen(cmd)

    cdp_url = f"http://127.0.0.1:{port}"
    last_error: Optional[Exception] = None
    for _ in range(45):
        if proc.poll() is not None:
            raise RuntimeError(
                f"Edge exited before CDP was ready (exit code {proc.returncode})"
            )
        try:
            browser = playwright.chromium.connect_over_cdp(cdp_url)
            context = browser.contexts[0] if browser.contexts else browser.new_context(
                accept_downloads=True
            )
            if not context.pages:
                context.new_page()
            logger.info("Connected to Edge over CDP on port %s", port)
            return ManagedBrowserContext(context, edge_process=proc, cdp_browser=browser)
        except Exception as exc:
            last_error = exc
            time.sleep(1)

    if proc.poll() is None:
        proc.kill()
    raise RuntimeError(f"Could not connect to Edge over CDP at {cdp_url}: {last_error}")


def _try_launch_plan(
    playwright: Playwright,
    profile_dir: Path,
    label: str,
    plan: dict[str, Any],
) -> BrowserContext:
    terminate_stale_edge_for_profile(profile_dir)
    clear_profile_locks(profile_dir)

    if plan.pop("_cdp", False):
        return _launch_via_cdp(playwright, profile_dir, plan["executable_path"])

    logger.info("Trying Edge launch plan: %s", label)
    return playwright.chromium.launch_persistent_context(
        **_persistent_launch_kwargs(profile_dir, plan=plan)
    )


def _is_recoverable(error: Exception) -> bool:
    text = str(error)
    return any(token in text for token in _RECOVERABLE_LAUNCH_ERRORS)


def launch_edge_persistent_context(
    playwright: Playwright,
    profile_dir: Path,
) -> BrowserContext:
    """
    Launch Edge with a persistent profile.

    On some Windows machines Playwright's remote-debugging-pipe fails with
    Browser.getWindowForTarget. We launch Edge ourselves with --remote-debugging-port
    and connect_over_cdp instead, which is the same as starting Edge manually.
    """
    ensure_playwright_driver()
    profile_dir.mkdir(parents=True, exist_ok=True)

    if _env_bool("EDGE_PROFILE_RESET"):
        reset_edge_profile(profile_dir)

    edge_executables = resolve_edge_executables()
    if edge_executables:
        logger.info("Found Edge executable(s): %s", ", ".join(str(p) for p in edge_executables))
    else:
        logger.error(
            "Could not locate msedge.exe on this machine. Set in .env:\n"
            "  EDGE_EXECUTABLE_PATH=C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
        )

    if not edge_executables:
        raise FileNotFoundError(
            "msedge.exe not found. Set EDGE_EXECUTABLE_PATH in .env to the full path shown "
            "when you run: where msedge.exe"
        )

    plans = _launch_plans(edge_executables)
    last_error: Optional[Exception] = None

    for round_index in range(1, 3):
        for label, plan in plans:
            try:
                return _try_launch_plan(playwright, profile_dir, label, dict(plan))
            except Exception as exc:
                last_error = exc
                logger.warning("Edge launch plan %s failed: %s", label, exc)

        if round_index == 1 and last_error and _is_recoverable(last_error):
            logger.warning(
                "All launch plans failed; wiping profile %s and trying once more.",
                profile_dir,
            )
            reset_edge_profile(profile_dir)
            time.sleep(2)
            continue

        break

    logger.error(
        "Edge could not be launched. On production run in PowerShell:\n"
        "  where.exe msedge.exe\n"
        "Then add the path to .env:\n"
        "  EDGE_EXECUTABLE_PATH=C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe\n"
        "  taskkill /F /IM msedge.exe\n"
        "  python -m src.setup_browser\n"
        "Profile: %s",
        profile_dir,
    )
    assert last_error is not None
    raise last_error
