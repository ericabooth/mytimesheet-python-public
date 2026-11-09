*! version 0.2.2 09nov2026
*! Eric A. Booth
* 0.2.2 09nov2026: every argument is passed to Python in single quotes on
*       macOS and Linux, so title() and paths keep double quotes and $ text
*       (the shell used to drop or expand them); on Windows the default
*       interpreter is python and the shell line is wrapped so cmd.exe keeps
*       its quotes; spaces around the ; separators in scandirs() are ignored.
*       Python side: folders such as tmp, temp or venv are excluded only below
*       a scan root (a root inside such a folder used to give zero activity),
*       exclusions work with Windows paths, and scanner(auto) walks a folder
*       when Spotlight finds nothing in it.
* 0.2.1 09nov2026: expand a leading ~ in the generator script path (findfile
*       returns ~/... under the default macOS and Linux PLUS, which the shell
*       passed to Python unexpanded, so nothing ran and rc was 0), in python()
*       and in output(); stop with r(498) when the generator does not finish;
*       stop with r(601) when timesheet_generator.py is not installed
* 0.2.0 27jun2026
program define mytimesheet
    version 16.0
    syntax [anything(name=period)] [, START(string) END(string) OUTPUT(string) PYTHON(string) SCANDIRs(string) BOUNDARY(string) ROLLING SCANNER(string) TITLE(string) QUIET]

    if `"`python'"' == "" {
        if fileexists("/opt/homebrew/bin/python3") {
            local python "/opt/homebrew/bin/python3"
        }
        else if fileexists("/usr/local/bin/python3") {
            local python "/usr/local/bin/python3"
        }
        else if "`c(os)'" == "Windows" {
            local python "python"
        }
        else {
            local python "python3"
        }
    }
    if `"`period'"' == "" {
        if `"`start'`end'"' == "" {
            local period "last-week"
        }
        else {
            local period "range"
        }
    }
    if `"`output'"' == "" {
        local output "timesheet.xlsx"
    }
    if `"`boundary'"' == "" {
        local boundary "02:00"
    }
    if `"`scanner'"' == "" {
        local scanner "auto"
    }

    capture findfile timesheet_generator.py
    if _rc {
        di as error "timesheet_generator.py not found on the adopath; reinstall mytimesheet"
        exit 601
    }
    local script `"`r(fn)'"'

    * The shell does not expand ~ inside double quotes, and findfile returns
    * ~/... when PLUS is the default ~/Library/.../plus/ or ~/ado/plus/
    foreach opt in script python output {
        _mytimesheet_tilde `"``opt''"'
        local `opt' `"`r(path)'"'
    }

    * Python's exit status does not come back through -shell-, so the
    * generator writes this marker only when it succeeds
    tempfile donefile
    capture erase `"`donefile'"'

    * Each argument is quoted in Mata (single quotes for sh/zsh, so nothing
    * inside is expanded; double quotes for cmd.exe)
    local iswin = ("`c(os)'" == "Windows")
    local cmd ""
    foreach opt in python script period output boundary scanner donefile start end scandirs title {
        if `"`macval(`opt')'"' != "" {
            mata: _mytimesheet_q("`opt'", `iswin')
        }
    }
    local cmd `"`macval(q_python)' `macval(q_script)' `macval(q_period)' --output `macval(q_output)' --boundary `macval(q_boundary)' --scanner `macval(q_scanner)' --done-file `macval(q_donefile)'"'
    if `"`start'"' != "" {
        local cmd `"`macval(cmd)' --start `macval(q_start)'"'
    }
    if `"`end'"' != "" {
        local cmd `"`macval(cmd)' --end `macval(q_end)'"'
    }
    if `"`scandirs'"' != "" {
        local cmd `"`macval(cmd)' --scan-dirs `macval(q_scandirs)'"'
    }
    if `"`macval(title)'"' != "" {
        local cmd `"`macval(cmd)' --title `macval(q_title)'"'
    }
    if "`rolling'" != "" {
        local cmd `"`macval(cmd)' --rolling"'
    }
    if "`quiet'" != "" {
        local cmd `"`macval(cmd)' --quiet"'
    }
    * cmd.exe strips the first and last quote of a line that starts with a
    * quote and holds more than two, so wrap the whole line in one more pair
    if `iswin' {
        local cmd `""`macval(cmd)'""'
    }

    di as text "Running mytimesheet Python generator..."
    di as text `"Python: `python'"'
    shell `macval(cmd)'
    capture confirm file `"`donefile'"'
    if _rc {
        di as error "the Python generator did not finish; no timesheet was written"
        di as error `"check the messages above and python() (now `python')"'
        exit 498
    }
    capture erase `"`donefile'"'
end

program define _mytimesheet_tilde, rclass
    args path
    if substr(`"`path'"', 1, 1) == "~" & inlist(substr(`"`path'"', 2, 1), "", "/", "\\") {
        local home : env HOME
        if "`c(os)'" == "Windows" {
            local up : env USERPROFILE
            local hd : env HOMEDRIVE
            local hp : env HOMEPATH
            if `"`up'"' != "" local home `"`up'"'
            else if `"`hd'`hp'"' != "" local home `"`hd'`hp'"'
        }
        if `"`home'"' == "" {
            di as error "cannot expand ~: HOME is not set"
            exit 601
        }
        local rest = substr(`"`path'"', 2, .)
        local path `"`home'`rest'"'
    }
    return local path `"`path'"'
end

mata:
// Quote the caller's local called name for the shell and store it in the
// local q_name (q_ plus that name): single quotes for sh/zsh (with ' written as '\''), double
// quotes for cmd.exe (with " replaced by ').
void _mytimesheet_q(string scalar name, real scalar iswin)
{
    string scalar v, sq

    v = st_local(name)
    if (iswin) {
        v = char(34) + subinstr(v, char(34), char(39)) + char(34)
    }
    else {
        sq = char(39)
        v = sq + subinstr(v, sq, sq + char(92) + sq + sq) + sq
    }
    st_local("q_" + name, v)
}
end
