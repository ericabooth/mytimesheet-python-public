# mytimesheet

`mytimesheet` creates an Excel timesheet estimate from local file modification
activity. It is a Python package first, with a small Stata wrapper for users who
want to launch the Python workflow from Stata.

The package does not hard-code a person, machine, date range, or Google Drive
account. By default it discovers common local folders on the current machine:

- `~/Library/CloudStorage/GoogleDrive-*/My Drive`
- `~/Library/CloudStorage/GoogleDrive-*/Shared drives`
- `~/Documents`
- `~/Desktop`

You can override those scan roots from the command line, a Stata option, or the
`MYTIMESHEET_SCAN_DIRS` environment variable.
=======
1. **`timesheet_generator.py`**: The reusable Python script that scans your workspace directories to collect and aggregate file modifications.
2. **`timesheet.xlsx`**: The generated Excel workbook populated with your actual activity data during that month or week, driven by live formulas.
3. **`README.md`**: This guide.
>>>>>>> e5e1b0f08076543edae21648e6edd596a9a07eab

## Files

- `mytimesheet/`: reusable Python package.
- `timesheet_generator.py`: compatibility entrypoint for direct script use.
- `pyproject.toml`: package metadata and console-script entrypoint.
- `requirements.txt`: minimal Python dependency list.
- `mytimesheet.ado`: Stata wrapper that calls Python.
- `mytimesheet.sthlp`: Stata help file.
- `mytimesheet.pkg` and `stata.toc`: Stata package-index files.

## Install Python Requirements

From this folder:

```bash
python3 -m pip install -r requirements.txt
```

If you are using the Stata wrapper on a Mac with Homebrew Python, install into
the same Python Stata will call:

```bash
/opt/homebrew/bin/python3 -m pip install -r requirements.txt
```

For editable local development:

```bash
python3 -m pip install -e .
```

After editable install, the `mytimesheet` command should be available:

```bash
mytimesheet --version
```

You can also run without installing the console script:

```bash
python3 timesheet_generator.py --version
python3 -m mytimesheet --version
```

## Python Usage

Previous completed Monday-Sunday week:

```bash
python3 timesheet_generator.py last-week --output timesheet_last_week.xlsx
```

Previous completed calendar month:

```bash
python3 timesheet_generator.py last-month --output timesheet_last_month.xlsx
```

Trailing 7-day or 30-day window:

```bash
python3 timesheet_generator.py last-week --rolling --output trailing_7_days.xlsx
python3 timesheet_generator.py last-month --rolling --output trailing_30_days.xlsx
```

Explicit date range:

```bash
python3 timesheet_generator.py range --start 2026-06-01 --end 2026-06-30 --output june.xlsx
```

Scan specific folders:

```bash
python3 timesheet_generator.py last-week \
  --scan-dir "$HOME/Documents" \
  --scan-dir "$HOME/Desktop" \
  --output timesheet.xlsx
```

Use a semicolon-separated scan list:

```bash
MYTIMESHEET_SCAN_DIRS="$HOME/Documents;$HOME/Desktop" \
  python3 timesheet_generator.py last-week
```

## Stata Usage

The Stata wrapper is a helper that sends options to Python and creates the Excel
workbook. The Python package remains the source of the workbook logic.

The wrapper chooses Python in this order when `python()` is not supplied:

1. `/opt/homebrew/bin/python3`
2. `/usr/local/bin/python3`
3. `python3` from the shell path

The selected Python must have `openpyxl` installed. The wrapper prints the
selected interpreter each time it runs, for example:

```text
Running mytimesheet Python generator...
Python: /opt/homebrew/bin/python3
```

If the selected Python is wrong, pass `python("...")` explicitly.

### Install in Stata

```stata
net install mytimesheet, from("https://raw.githubusercontent.com/ericabooth/mytimesheet-stata/master/") replace
which mytimesheet
help mytimesheet
mytimesheet last-week, output("timesheet_last_week.xlsx")
```

### Stata Examples

Previous completed week:

```stata
mytimesheet
```

Previous completed calendar month:

```stata
mytimesheet last-month, output("timesheet_last_month.xlsx")
```

Explicit range:

```stata
mytimesheet range, start(2026-06-01) end(2026-06-30) output("june.xlsx")
```

Trailing 7 days:

```stata
mytimesheet last-week, rolling output("trailing_7_days.xlsx")
```

Custom scan folders:

```stata
mytimesheet last-week, scandirs("/Users/me/Documents;/Users/me/Desktop")
```

Use a specific Python executable:

```stata
mytimesheet last-month, python("/opt/homebrew/bin/python3") output("timesheet.xlsx")
```

Install `openpyxl` into a specific Python executable:

```bash
/opt/homebrew/bin/python3 -m pip install openpyxl
```

## Workbook Logic

The generated workbook has two sheets:

- `Summary`: daily rows with formulas for earliest start, latest end, estimated
  elapsed work span, workday flag, and weekly summary totals.
- `Activity`: folder-level evidence showing work date, folder name, folder path,
  earliest action, latest action, and file modification counts.

The default workday boundary is `02:00`. File changes before 2 AM count toward
the previous workday. You can change this with `--boundary HH:MM` in Python or
`boundary(HH:MM)` in Stata.

Estimated work hours are the elapsed span between the first and last observed
file activity for a workday. Treat this as an evidence-backed estimate, not a
literal time clock.

## Scanner Behavior

On macOS, `mytimesheet` uses Spotlight (`mdfind`) when available because it is
much faster for large Google Drive and local folders. If Spotlight is unavailable
or fails for a folder, the package falls back to a recursive filesystem walk.

You can force a scanner:

```bash
python3 timesheet_generator.py last-week --scanner spotlight
python3 timesheet_generator.py last-week --scanner walk
```

The Stata equivalent is:

```stata
mytimesheet last-week, scanner(walk)
```

## Notes and Caveats

- The output depends on local file modification timestamps and the directories
  the user chooses to scan.
- Some cloud-sync tools preserve server-side timestamps rather than local edit
  times. Review the `Activity` sheet before using the output for reporting.
- The package excludes common noisy folders such as `.git`, virtual
  environments, caches, `node_modules`, and system metadata files.
- If no activity is found, the workbook is still created with date rows and empty
  activity evidence.

## Troubleshooting

### `ModuleNotFoundError: No module named 'openpyxl'`

This means Stata launched a Python interpreter that does not have `openpyxl`.
Check the interpreter printed by the wrapper, then install `openpyxl` into that
interpreter:

```bash
/opt/homebrew/bin/python3 -m pip install openpyxl
```

Or point Stata to a Python that already has `openpyxl`:

```stata
mytimesheet last-week, python("/opt/homebrew/bin/python3")
```

if `/usr/bin/python3` does not have `openpyxl`, while
`/opt/homebrew/bin/python3` does. The ado wrapper now prefers the Homebrew
interpreter automatically when it exists.

### `which mytimesheet` finds the wrong file

Run:

```stata
which mytimesheet
adopath
```

If `which mytimesheet` returns something other than the path under
`PLUS` (typically `~/Library/Application Support/Stata/ado/plus/m/`),
an earlier copy from elsewhere on the adopath is shadowing the package.
`ado uninstall mytimesheet` and re-run `net install`.



## Author

Eric A. Booth, Sr Researcher, Texas2036.org (eric.a.booth@gmail.com).
