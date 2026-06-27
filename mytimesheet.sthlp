{smcl}
{* *! version 0.2.0 27jun2026}{...}
{title:Title}

{p2colset 5 20 22 2}{...}
{p2col:{hi:mytimesheet} {hline 2}}Generate an Excel timesheet estimate from local file activity{p_end}
{p2colreset}{...}

{title:Syntax}

{p 8 18 2}
{cmd:mytimesheet} [{it:period}]
[{cmd:,}
{cmd:start(}{it:YYYY-MM-DD}{cmd:)}
{cmd:end(}{it:YYYY-MM-DD}{cmd:)}
{cmd:output(}{it:path}{cmd:)}
{cmd:python(}{it:python_command_or_path}{cmd:)}
{cmd:scandirs(}{it:semicolon-separated paths}{cmd:)}
{cmd:boundary(}{it:HH:MM}{cmd:)}
{cmd:rolling}
{cmd:scanner(}{it:auto|spotlight|walk}{cmd:)}
{cmd:title(}{it:text}{cmd:)}
{cmd:quiet}]

{pstd}
Allowed periods are {cmd:last-week}, {cmd:last-month}, {cmd:week}, {cmd:month}, and
{cmd:range}. If no period is supplied, {cmd:last-week} is used.

{title:Description}

{pstd}
{cmd:mytimesheet} is a Stata wrapper around the Python {cmd:mytimesheet}
generator. It launches Python, scans local file modification activity, and writes
an Excel workbook with a daily summary and folder-level activity evidence.

{pstd}
The wrapper is only a convenience layer. The workbook logic lives in Python.

{title:Options}

{phang}
{cmd:start()} and {cmd:end()} specify an explicit inclusive date range. They are
required when {it:period} is {cmd:range}.

{phang}
{cmd:output()} specifies the Excel workbook to create. The default is
{cmd:timesheet.xlsx} in Stata's current working directory.

{phang}
{cmd:python()} specifies the Python executable. If omitted, the wrapper tries
{cmd:/opt/homebrew/bin/python3}, then {cmd:/usr/local/bin/python3}, then
{cmd:python3}. The selected Python is printed when the command runs.

{phang}
{cmd:scandirs()} passes semicolon-separated scan roots to Python. If omitted,
Python discovers common local roots such as Google Drive, Documents, and Desktop.

{phang}
{cmd:boundary()} sets the workday rollover time. The default is {cmd:02:00}, so
activity before 2 AM is counted on the previous workday.

{phang}
{cmd:rolling} makes {cmd:last-week} use the trailing 7 days and {cmd:last-month}
use the trailing 30 days. Without {cmd:rolling}, these use the previous completed
Monday-Sunday week or previous completed calendar month.

{phang}
{cmd:scanner()} controls file discovery. {cmd:auto} uses Spotlight on macOS when
available and falls back to a recursive walk.

{title:Python requirements}

{pstd}
The selected Python must have {cmd:openpyxl} installed. On many Macs,
{cmd:/usr/bin/python3} does not include user-installed packages, while Homebrew
Python at {cmd:/opt/homebrew/bin/python3} often does. The wrapper therefore
prefers Homebrew Python when it is present.

{pstd}
To install the required Python package into Homebrew Python, run this in Terminal:

{phang2}{cmd:/opt/homebrew/bin/python3 -m pip install openpyxl}{p_end}

{pstd}
To force a specific Python from Stata, use:

{phang2}{cmd:. mytimesheet last-week, python("/opt/homebrew/bin/python3")}{p_end}

{title:Examples}

{phang2}{cmd:. mytimesheet}{p_end}
{phang2}{cmd:. mytimesheet last-month, output("timesheet_last_month.xlsx")}{p_end}
{phang2}{cmd:. mytimesheet range, start(2026-06-01) end(2026-06-30) output("june.xlsx")}{p_end}
{phang2}{cmd:. mytimesheet last-week, scandirs("/Users/me/Documents;/Users/me/Desktop")}{p_end}

{title:Troubleshooting}

{pstd}
If Stata reports {cmd:ModuleNotFoundError: No module named 'openpyxl'}, install
{cmd:openpyxl} into the Python printed by the wrapper, or pass a Python that
already has {cmd:openpyxl} with {cmd:python()}.

{pstd}
If Stata finds an older copy, run:

{phang2}{cmd:. which mytimesheet}{p_end}
{phang2}{cmd:. adopath}{p_end}

{pstd}
The expected shared-drive adopath copy is in
{cmd:.../Shared drives/Data and Research Team/_codeshare/mytimesheet.ado}.

{title:Author}

{pstd}
Eric A. Booth. MIT licensed.
