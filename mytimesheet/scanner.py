"""File activity scanners."""

from __future__ import annotations

import datetime as dt
import os
import platform
import shutil
import subprocess
from pathlib import Path


EXCLUDE_SUBSTRINGS = (
    "/.git/",
    "/.venv/",
    "/venv/",
    "/node_modules/",
    "/.cache/",
    "/__pycache__/",
    "/.claude/",
    "/.gemini/",
    "/.codex/",
    "/.trash/",
    "/.metadata/",
    "/.vscode/",
    "/.settings/",
    "/tmp/",
    "/temp/",
)

EXCLUDE_FILENAMES = {".ds_store", "desktop.ini", "thumbs.db"}


def get_work_date(timestamp: dt.datetime, boundary: dt.time) -> dt.date:
    if timestamp.time() < boundary:
        return (timestamp - dt.timedelta(days=1)).date()
    return timestamp.date()


def should_skip(path: Path) -> bool:
    lower_path = f"/{str(path).lower().strip('/')}"
    return path.name.lower() in EXCLUDE_FILENAMES or any(
        token in lower_path for token in EXCLUDE_SUBSTRINGS
    )


def spotlight_available() -> bool:
    return platform.system() == "Darwin" and shutil.which("mdfind") is not None


def _spotlight_paths(directory: Path, start_dt: dt.datetime, end_dt: dt.datetime) -> list[Path]:
    start_str = start_dt.strftime("%Y-%m-%dT%H:%M:%S")
    end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%S")
    query = (
        f"kMDItemContentModificationDate >= $time.iso({start_str}) && "
        f"kMDItemContentModificationDate <= $time.iso({end_str})"
    )
    result = subprocess.run(
        ["mdfind", "-onlyin", str(directory), query],
        capture_output=True,
        text=True,
        check=True,
    )
    return [Path(line) for line in result.stdout.splitlines() if line]


def _walk_paths(directory: Path, start_dt: dt.datetime, end_dt: dt.datetime) -> list[Path]:
    paths: list[Path] = []
    for root, dirs, files in os.walk(directory):
        root_path = Path(root)
        dirs[:] = [name for name in dirs if not should_skip(root_path / name)]
        for filename in files:
            path = root_path / filename
            if should_skip(path):
                continue
            try:
                modified = dt.datetime.fromtimestamp(path.stat().st_mtime)
            except OSError:
                continue
            if start_dt <= modified <= end_dt:
                paths.append(path)
    return paths


def collect_modified_paths(
    scan_dirs: list[Path],
    start_dt: dt.datetime,
    end_dt: dt.datetime,
    scanner: str = "auto",
    quiet: bool = False,
) -> list[Path]:
    """Collect modified file paths with Spotlight when possible, else os.walk."""
    use_spotlight = scanner == "spotlight" or (scanner == "auto" and spotlight_available())
    if scanner == "spotlight" and not spotlight_available():
        raise RuntimeError("Spotlight scanner requested, but mdfind is not available")

    paths: list[Path] = []
    for directory in scan_dirs:
        if not directory.exists():
            if not quiet:
                print(f"  Directory not found, skipping: {directory}")
            continue

        if not quiet:
            backend = "Spotlight" if use_spotlight else "walk"
            print(f"  Scanning with {backend}: {directory}")

        try:
            if use_spotlight:
                paths.extend(_spotlight_paths(directory, start_dt, end_dt))
            else:
                paths.extend(_walk_paths(directory, start_dt, end_dt))
        except Exception as exc:
            if scanner == "spotlight":
                raise
            if not quiet:
                print(f"  Spotlight failed for {directory}; falling back to os.walk ({exc})")
            paths.extend(_walk_paths(directory, start_dt, end_dt))

    unique: dict[Path, Path] = {}
    for path in paths:
        try:
            resolved = path.resolve()
        except OSError:
            continue
        if resolved.is_file() and not should_skip(resolved):
            unique[resolved] = resolved
    return list(unique.values())
