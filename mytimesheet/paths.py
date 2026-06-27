"""Path discovery and display helpers."""

from __future__ import annotations

import os
from pathlib import Path


ENV_SCAN_DIRS = "MYTIMESHEET_SCAN_DIRS"


def split_scan_dirs(value: str | None) -> list[Path]:
    if not value:
        return []
    separator = ";" if ";" in value else os.pathsep
    return [Path(part).expanduser() for part in value.split(separator) if part.strip()]


def discover_default_scan_dirs() -> list[Path]:
    """Find useful machine-local roots without hard-coding a user account."""
    home = Path.home()
    candidates: list[Path] = []

    cloud_storage = home / "Library" / "CloudStorage"
    if cloud_storage.exists():
        for drive_root in sorted(cloud_storage.glob("GoogleDrive-*")):
            candidates.extend([drive_root / "My Drive", drive_root / "Shared drives"])

    candidates.extend([home / "Documents", home / "Desktop"])
    existing = [path for path in candidates if path.exists()]
    return existing or [home]


def build_scan_dirs(cli_dirs: list[str], semicolon_dirs: str | None) -> list[Path]:
    explicit = [Path(path).expanduser() for path in cli_dirs]
    explicit.extend(split_scan_dirs(semicolon_dirs))
    explicit.extend(split_scan_dirs(os.environ.get(ENV_SCAN_DIRS)))

    raw_dirs = explicit if explicit else discover_default_scan_dirs()
    clean_dirs: list[Path] = []
    seen: set[Path] = set()
    for path in raw_dirs:
        resolved = path.resolve()
        if resolved.exists() and resolved not in seen:
            clean_dirs.append(resolved)
            seen.add(resolved)
    return clean_dirs


def beautify_path(path: Path, scan_dirs: list[Path]) -> str:
    """Return a readable path relative to a scan root, Google Drive root, or home."""
    resolved = path.resolve()

    for root in scan_dirs:
        try:
            return str(resolved.relative_to(root.resolve()))
        except ValueError:
            continue

    cloud_storage = Path.home() / "Library" / "CloudStorage"
    if cloud_storage.exists():
        for drive_root in sorted(cloud_storage.glob("GoogleDrive-*")):
            try:
                return str(resolved.relative_to(drive_root.resolve()))
            except ValueError:
                continue

    try:
        return f"~/{resolved.relative_to(Path.home().resolve())}"
    except ValueError:
        return str(resolved)
