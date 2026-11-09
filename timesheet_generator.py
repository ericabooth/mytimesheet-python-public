#!/usr/bin/env python3
"""Compatibility entrypoint for the mytimesheet command line tool.

This file supports both the source-tree package layout and Stata's net-install
layout, which places Python ancillary files together in PLUS/py.
"""

from pathlib import Path
import sys

try:
    from mytimesheet.cli import main
except ModuleNotFoundError as exc:
    if exc.name != "mytimesheet":
        raise
    script_dir = Path(__file__).resolve().parent
    sys.path.insert(0, str(script_dir / "mytimesheet"))
    from cli import main


def _pop_done_file(argv: list) -> "str | None":
    """Remove --done-file PATH (used by the Stata wrapper) from argv."""
    if "--done-file" in argv:
        i = argv.index("--done-file")
        path = argv[i + 1] if i + 1 < len(argv) else None
        del argv[i:i + 2]
        return path
    return None


if __name__ == "__main__":
    _argv = sys.argv[1:]
    _done = _pop_done_file(_argv)
    _rc = main(_argv)
    if _rc == 0 and _done:
        # completion marker for mytimesheet.ado; -shell- does not pass back exit codes
        Path(_done).write_text("done\n", encoding="utf-8")
    raise SystemExit(_rc)
