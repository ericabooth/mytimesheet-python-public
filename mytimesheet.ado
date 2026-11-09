*! version 0.2.1 09nov2026
*! Eric A. Booth
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

    local q = char(34)

    local cmd `"`q'`python'`q' `q'`script'`q' `period' --output `q'`output'`q' --boundary `q'`boundary'`q' --scanner `q'`scanner'`q' --done-file `q'`donefile'`q'"'

    if `"`start'"' != "" {
        local cmd `"`macval(cmd)' --start `q'`start'`q'"'
    }
    if `"`end'"' != "" {
        local cmd `"`macval(cmd)' --end `q'`end'`q'"'
    }
    if `"`scandirs'"' != "" {
        local cmd `"`macval(cmd)' --scan-dirs `q'`scandirs'`q'"'
    }
    if `"`title'"' != "" {
        local cmd `"`macval(cmd)' --title `q'`macval(title)'`q'"'
    }
    if "`rolling'" != "" {
        local cmd `"`macval(cmd)' --rolling"'
    }
    if "`quiet'" != "" {
        local cmd `"`macval(cmd)' --quiet"'
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
