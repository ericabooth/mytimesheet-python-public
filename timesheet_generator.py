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


if __name__ == "__main__":
    raise SystemExit(main())
