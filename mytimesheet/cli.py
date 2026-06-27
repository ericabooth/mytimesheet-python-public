"""Command line interface for mytimesheet."""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

try:
    from . import __version__
    from .generator import generate_timesheet
    from .periods import normalize_period, resolve_period
    from .paths import build_scan_dirs
except ImportError:
    __version__ = "0.2.0"
    from generator import generate_timesheet
    from periods import normalize_period, resolve_period
    from paths import build_scan_dirs


def parse_boundary(value: str) -> dt.time:
    """Parse an HH:MM workday boundary."""
    try:
        hour, minute = value.split(":", 1)
        return dt.time(int(hour), int(minute))
    except (ValueError, TypeError) as exc:
        raise argparse.ArgumentTypeError("boundary must be in HH:MM format") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="mytimesheet",
        description="Generate an Excel timesheet estimate from local file modification activity.",
    )
    parser.add_argument(
        "period",
        nargs="?",
        default="last-week",
        help=(
            "Date window to scan: last-week, last-month, week, month, or range. "
            "Use --start and --end with range."
        ),
    )
    parser.add_argument("--start", help="Start date for an explicit range, YYYY-MM-DD.")
    parser.add_argument("--end", help="End date for an explicit range, YYYY-MM-DD.")
    parser.add_argument(
        "--rolling",
        action="store_true",
        help="Use a trailing 7-day or 30-day window for last-week/last-month.",
    )
    parser.add_argument(
        "--output",
        default="timesheet.xlsx",
        help="Output workbook path. Defaults to timesheet.xlsx in the current directory.",
    )
    parser.add_argument(
        "--scan-dir",
        action="append",
        default=[],
        help="Directory to scan. May be supplied more than once.",
    )
    parser.add_argument(
        "--scan-dirs",
        help="Semicolon-separated directories to scan. Useful for Stata wrappers.",
    )
    parser.add_argument(
        "--boundary",
        type=parse_boundary,
        default=dt.time(2, 0),
        help="Workday boundary in HH:MM. Default is 02:00.",
    )
    parser.add_argument(
        "--scanner",
        choices=("auto", "spotlight", "walk"),
        default="auto",
        help="Scanner backend. auto uses Spotlight on macOS when available, otherwise os.walk.",
    )
    parser.add_argument(
        "--title",
        default="WORKDAY & TIMESHEET SUMMARY",
        help="Workbook title text.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Only print warnings and the final output path.",
    )
    parser.add_argument("--version", action="version", version=f"mytimesheet {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        period = normalize_period(args.period)
        start_date, end_date = resolve_period(
            period=period,
            start=args.start,
            end=args.end,
            rolling=args.rolling,
        )
        scan_dirs = build_scan_dirs(args.scan_dir, args.scan_dirs)
        if not scan_dirs:
            raise ValueError(
                "No scan directories were found. Pass --scan-dir or set MYTIMESHEET_SCAN_DIRS."
            )

        output_path = Path(args.output).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if not args.quiet:
            print(f"Time sheet period: {start_date} to {end_date}")
            print(
                "Scan window: "
                f"{dt.datetime.combine(start_date, args.boundary)} to "
                f"{dt.datetime.combine(end_date + dt.timedelta(days=1), args.boundary)}"
            )
            print("Scan directories:")
            for directory in scan_dirs:
                print(f"  - {directory}")
            print("----------------------------------------")

        generate_timesheet(
            start_date=start_date,
            end_date=end_date,
            output_path=output_path,
            scan_dirs=scan_dirs,
            boundary=args.boundary,
            scanner=args.scanner,
            title=args.title,
            quiet=args.quiet,
        )
        print(f"Saved workbook: {output_path}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
