{smcl}
{* *! version 0.2.2 09nov2026}{...}
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
{cmd:range}. {cmd:week} and {cmd:month} are short names for {cmd:last-week} and
{cmd:last-month} (they give the same dates). If no period is supplied,
{cmd:last-week} is used, or {cmd:range} when {cmd:start()} or {cmd:end()} is
given.

{title:Description}

{pstd}
{cmd:mytimesheet} is a Stata wrapper around the Python {cmd:mytimesheet}
generator. It launches Python, scans local file modification activity, and writes
an Excel workbook with a daily summary and folder-level activity evidence.

{pstd}
The wrapper is only a convenience layer. The workbook logic lives in Python.

{pstd}
The wrapper passes Python the full path of the installed
{cmd:timesheet_generator.py}, expanding a leading {cmd:~} (as reported under the
default PLUS folder) from your home folder; a leading {cmd:~} in {cmd:python()}
or {cmd:output()} is expanded the same way. If Python does not run, or the
generator stops with an error, {cmd:mytimesheet} stops with error 498 after
Python's own message, so a do-file does not carry on as if a workbook had been
written. On macOS and Linux every argument is passed to Python in single quotes,
so double quotes and {cmd:$} text in {cmd:title()} or in paths arrive unchanged.

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
{cmd:python3} ({cmd:python} on Windows). The selected Python is printed when the
command runs.

{phang}
{cmd:scandirs()} passes semicolon-separated scan roots to Python; spaces around
each {cmd:;} are ignored. If omitted, Python discovers common local roots such as
Google Drive, Documents, and Desktop. Folders named {cmd:.git}, {cmd:node_modules},
{cmd:venv}, {cmd:.venv}, {cmd:tmp}, {cmd:temp} and similar are skipped when they
appear below a scan root.

{phang}
{cmd:boundary()} sets the workday rollover time. The default is {cmd:02:00}, so
activity before 2 AM is counted on the previous workday.

{phang}
{cmd:rolling} makes {cmd:last-week} use the trailing 7 days and {cmd:last-month}
use the trailing 30 days. Without {cmd:rolling}, these use the previous completed
Monday-Sunday week or previous completed calendar month.

{phang}
{cmd:scanner()} controls file discovery. {cmd:auto} uses Spotlight on macOS when
available and walks a folder instead when Spotlight fails or finds nothing in it
(for example a folder outside the Spotlight index). {cmd:spotlight} uses
Spotlight only; {cmd:walk} always walks.

{phang}
{cmd:title()} sets the title in cell A1 of the {cmd:Summary} sheet.

{phang}
{cmd:quiet} suppresses Python's progress messages (the period, scan folders and
scanner used); errors and the saved-workbook line are still shown.

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
If Stata reports {cmd:ModuleNotFoundError: No module named 'openpyxl'} (followed
by error 498), install
{cmd:openpyxl} into the Python printed by the wrapper, or pass a Python that
already has {cmd:openpyxl} with {cmd:python()}.

{pstd}
If Stata finds an older copy, run:

{phang2}{cmd:. which mytimesheet}{p_end}
{phang2}{cmd:. adopath}{p_end}

{pstd}
The expected copy is the one under your PLUS folder ({cmd:sysdir}); if another
copy shadows it, {cmd:ado uninstall mytimesheet} and run {cmd:net install} again.

{title:Changes}

{pstd}0.2.2 (09nov2026): arguments keep double quotes and {cmd:$} text; Windows
default {cmd:python} and quoting; {cmd:scandirs()} ignores spaces around {cmd:;};
excluded folders are matched only below each scan root (a root under {cmd:/tmp}
or {cmd:Temp} no longer gives zero activity) and on Windows paths;
{cmd:scanner(auto)} walks a folder when Spotlight finds nothing.

{pstd}0.2.1 (09nov2026): {cmd:~} expansion; error 498 when the generator does not
finish.

{title:Author}

{pstd}
Eric A. Booth. MIT licensed.
