"""File activity scanners."""

from __future__ import annotations

import datetime as dt
import os
import platform
import shutil
import subprocess
from pathlib import Path, PurePath


# Folder names whose contents are never counted as work activity. They are
# matched exactly (ignoring case) against the folders *below* the scan root,
# so a scan root that itself sits under a folder named tmp or venv still works,
# and Windows paths (backslashes) are handled the same as POSIX paths.
EXCLUDE_DIRS = frozenset(
    {
        ".git",
        ".venv",
        "venv",
        "node_modules",
        ".cache",
        "__pycache__",
        ".claude",
        ".gemini",
        ".codex",
        ".trash",
        ".metadata",
        ".vscode",
        ".settings",
        "tmp",
        "temp",
    }
)

EXCLUDE_FILENAMES = {".ds_store", "desktop.ini", "thumbs.db"}


def get_work_date(timestamp: dt.datetime, boundary: dt.time) -> dt.date:
    if timestamp.time() < boundary:
        return (timestamp - dt.timedelta(days=1)).date()
    return timestamp.date()


def _parts_below(path: PurePath, root: PurePath | None) -> tuple[str, ...]:
    """Path components below root (all components when root is None or not a parent)."""
    if root is not None:
        try:
            return path.relative_to(root).parts
        except ValueError:
            pass
    return path.parts


def should_skip(path: PurePath, root: PurePath | None = None, is_dir: bool = False) -> bool:
    """True for noise files and for anything inside an excluded folder below root.

    With is_dir=True, path is itself a folder and is skipped when its own name
    is excluded (this is how the walk prunes .git, node_modules and so on).
    """
    if not is_dir and path.name.lower() in EXCLUDE_FILENAMES:
        return True
    parts = _parts_below(path, root)
    folders = parts if is_dir else parts[:-1]
    return any(part.lower() in EXCLUDE_DIRS for part in folders)


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
        dirs[:] = [name for name in dirs if not should_skip(root_path / name, directory, is_dir=True)]
        for filename in files:
            path = root_path / filename
            if should_skip(path, directory):
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
    """Collect modified file paths with Spotlight when possible, else os.walk.

    Under scanner="auto", a folder for which Spotlight returns nothing (for
    example one outside the Spotlight index) is walked instead, so an
    unindexed folder does not silently give an empty timesheet.
    """
    use_spotlight = scanner == "spotlight" or (scanner == "auto" and spotlight_available())
    if scanner == "spotlight" and not spotlight_available():
        raise RuntimeError("Spotlight scanner requested, but mdfind is not available")

    found: list[tuple[Path, Path]] = []
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
                these = _spotlight_paths(directory, start_dt, end_dt)
                if not these and scanner == "auto":
                    if not quiet:
                        print(f"  Spotlight found nothing in {directory}; checking with walk")
                    these = _walk_paths(directory, start_dt, end_dt)
            else:
                these = _walk_paths(directory, start_dt, end_dt)
        except Exception as exc:
            if scanner == "spotlight":
                raise
            if not quiet:
                print(f"  Spotlight failed for {directory}; falling back to os.walk ({exc})")
            these = _walk_paths(directory, start_dt, end_dt)
        found.extend((path, directory) for path in these)

    unique: dict[Path, Path] = {}
    for path, directory in found:
        try:
            resolved = path.resolve()
            root = directory.resolve()
        except OSError:
            continue
        if resolved.is_file() and not should_skip(resolved, root):
            unique[resolved] = resolved
    return list(unique.values())
