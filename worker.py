from __future__ import annotations

"""Console worker used by the frozen GUI build.

Playwright deliberately defaults frozen applications to a browser directory inside
PyInstaller's temporary ``_MEI`` folder.  The real browser is stored beside the
application (portable build) or in the Windows Playwright cache, so the path must be
set before Playwright is imported.
"""

import os
import sys
from pathlib import Path



def _configure_utf8_stdio() -> None:
    """Use UTF-8 for GUI pipes regardless of the Windows OEM code page."""
    os.environ.setdefault("PYTHONUTF8", "1")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    os.environ.setdefault("PYTHONUNBUFFERED", "1")

    for name in ("stdin", "stdout", "stderr"):
        stream = getattr(sys, name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if not callable(reconfigure):
            continue
        kwargs = {"encoding": "utf-8", "errors": "replace"}
        if name in ("stdout", "stderr"):
            kwargs.update({"line_buffering": True, "write_through": True})
        try:
            reconfigure(**kwargs)
        except (OSError, ValueError, TypeError):
            pass

def _application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _configure_playwright_browser_path() -> Path:
    existing = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "").strip()
    if existing and existing != "0":
        selected = Path(existing).expanduser()
    else:
        app_dir = _application_dir()
        candidates = [app_dir / "ms-playwright"]
        local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
        if local_app_data:
            candidates.append(Path(local_app_data) / "ms-playwright")
        user_profile = os.environ.get("USERPROFILE", "").strip()
        if user_profile:
            candidates.append(Path(user_profile) / "AppData" / "Local" / "ms-playwright")

        selected = next((item for item in candidates if item.is_dir()), candidates[0])
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(selected.resolve())

    # Prevent Playwright from garbage-collecting the portable browser directory.
    os.environ.setdefault("PLAYWRIGHT_SKIP_BROWSER_GC", "1")
    return selected


_configure_utf8_stdio()
BROWSER_PATH = _configure_playwright_browser_path()

import optimized_transfer  # noqa: E402  (must be imported after environment setup)


if __name__ == "__main__":
    print(f"Playwright browsers: {BROWSER_PATH}")
    optimized_transfer.core.main()
