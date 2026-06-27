*! version 0.2.0 27jun2026
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

    quietly findfile timesheet_generator.py
    local script `"`r(fn)'"'
    local q = char(34)

    local cmd "`q'`python'`q' `q'`script'`q' `period' --output `q'`output'`q' --boundary `q'`boundary'`q' --scanner `q'`scanner'`q'"

    if `"`start'"' != "" {
        local cmd "`cmd' --start `q'`start'`q'"
    }
    if `"`end'"' != "" {
        local cmd "`cmd' --end `q'`end'`q'"
    }
    if `"`scandirs'"' != "" {
        local cmd "`cmd' --scan-dirs `q'`scandirs'`q'"
    }
    if `"`title'"' != "" {
        local cmd "`cmd' --title `q'`title'`q'"
    }
    if "`rolling'" != "" {
        local cmd "`cmd' --rolling"
    }
    if "`quiet'" != "" {
        local cmd "`cmd' --quiet"
    }

    di as text "Running mytimesheet Python generator..."
    di as text "Python: `python'"
    shell `cmd'
end
