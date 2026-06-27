"""Date-window helpers for the mytimesheet CLI."""

from __future__ import annotations

import datetime as dt


ALIASES = {
    "lastweek": "last-week",
    "last_week": "last-week",
    "week": "last-week",
    "weekly": "last-week",
    "lastmonth": "last-month",
    "last_month": "last-month",
    "month": "last-month",
    "monthly": "last-month",
    "custom": "range",
}


def normalize_period(period: str) -> str:
    normalized = ALIASES.get(period.strip().lower(), period.strip().lower())
    if normalized not in {"last-week", "last-month", "range"}:
        raise ValueError("period must be last-week, last-month, week, month, or range")
    return normalized


def parse_date(value: str, label: str) -> dt.date:
    try:
        return dt.datetime.strptime(value, "%Y-%m-%d").date()
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} date must be in YYYY-MM-DD format") from exc


def resolve_period(
    period: str,
    start: str | None = None,
    end: str | None = None,
    rolling: bool = False,
    today: dt.date | None = None,
) -> tuple[dt.date, dt.date]:
    """Resolve a named or explicit period to inclusive start/end dates."""
    if start or end or period == "range":
        if not start or not end:
            raise ValueError("range requires both --start and --end")
        start_date = parse_date(start, "start")
        end_date = parse_date(end, "end")
        if start_date > end_date:
            raise ValueError("start date cannot be after end date")
        return start_date, end_date

    today = today or dt.date.today()

    if period == "last-week":
        if rolling:
            return today - dt.timedelta(days=6), today
        this_monday = today - dt.timedelta(days=today.weekday())
        end_date = this_monday - dt.timedelta(days=1)
        start_date = end_date - dt.timedelta(days=6)
        return start_date, end_date

    if period == "last-month":
        if rolling:
            return today - dt.timedelta(days=29), today
        first_this_month = today.replace(day=1)
        end_date = first_this_month - dt.timedelta(days=1)
        start_date = end_date.replace(day=1)
        return start_date, end_date

    raise ValueError("unrecognized period")
