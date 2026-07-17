from __future__ import annotations

import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: prepare_dist.py <app_dir>", file=sys.stderr)
        return 2

    app_dir = Path(sys.argv[1]).resolve()
    app_dir.mkdir(parents=True, exist_ok=True)

    required_files = (
        "config_transfer.yaml",
        "исходник.xls",
        "итог.xls",
    )
    for name in required_files:
        source = ROOT / name
        if not source.is_file():
            print(f"Missing required file: {source}", file=sys.stderr)
            return 3
        shutil.copy2(source, app_dir / name)

    for folder_name in (
        "скачанные_проекты",
        "logs_transfer",
        "screenshots_transfer",
    ):
        source_dir = ROOT / folder_name
        target_dir = app_dir / folder_name
        if source_dir.is_dir() and folder_name == "скачанные_проекты":
            shutil.copytree(source_dir, target_dir, dirs_exist_ok=True)
        else:
            target_dir.mkdir(parents=True, exist_ok=True)

    for optional_name in ("README.md", "REPAIR_CONFIG_PATH.txt", "REPAIR_UTF8_OUTPUT.txt"):
        source = ROOT / optional_name
        if source.is_file():
            shutil.copy2(source, app_dir / optional_name)

    print(f"Runtime files copied to: {app_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
