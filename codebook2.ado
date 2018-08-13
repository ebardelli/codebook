/*
 * Codebook
 *
 * Version 0.1
 * Emanuele Bardelli <bardelli@umich.edu>
 *
 */

version 15.1
cap program drop codebook2
program define codebook2
quietly {
    syntax [varlist] using/, [replace modify]
    cap putexcel close

    ** Sheet: Dataset
     * =========================================================================
    putexcel set "`using'", sheet("Dataset") `replace' `modify' open
    local row = 1

    describe, varlist
    local sorted = "`r(sortlist)'"

    ** Filename
    if regexm("`c(filename)'", "^(.*/)([^/]*)$") {
        local path = regexs(1)
        local filename = regexs(2)
    }
    putexcel A`row' = "Dataset" B`row' = "`filename'"

    ** Label
    local lab: data label
    putexcel A`++row' = "Label" B`row' = "`lab'"

    ** Notes
    local notes = ""
    putexcel A`++row' = "Notes"
    notes _count k : _dta
    if `k' > 0 {
        forvalues i = 1/`k' {
            notes _fetch n`i' : _dta `i'
            putexcel B`row' = "`n`i''"
            if `i' < `k' {
                local row = `row' + 1
            }
        }
    }

    ** Variales
    putexcel A`++row' = "Variables"
    putexcel B`row' = (`r(k)'), nformat(number_sep)
    putexcel A`++row' = "Observations"
    putexcel B`row' = (`r(N)'), nformat(number_sep)

    putexcel A`++row' = "Sorted by" B`row' = "`sorted'"

    putexcel A`++row' = "Last update" B`row' = "`c(filedate)'"
    putexcel A`++row' = "Path" B`row' = "`path'"

    putexcel close

    ** Sheet: Variable List
     * =========================================================================
    putexcel set "`using'", sheet("Variable List") modify open

    putexcel B1 = "Variable Name" C1 = "Variable Label" D1 = "Variable Type" ///
             E1 = "Value Code Label" F1 = "Notes"
    putexcel (A1:F1), border(bottom) bold

    local n = 0
    local r = 2
    foreach x of varlist _all {

        local r1 = `r'

        ** Variable number
        putexcel A`r' = `++n'

        ** Name
        putexcel B`r' = "`x'"

        ** Label
        local lab: variable label `x'
        putexcel C`r' = "`lab'"

        ** Variable type
         * Note: if we prefer to just say string or numeric, then just use regexm for the
         * local t and if regexm matches to "str" then have it for string and otherwise
         * have it put numeric
        local t : type `x'
        putexcel D`r' = "`t'"

        ** Value codes
        local lab_name : value label `x'
        putexcel E`r' = "`lab_name'"

        ** Variable notes
        local notes = ""
        notes _count k : "`x'"
        if `k' > 0 {
            forvalues i = 1/`k' {
                notes _fetch n`i' : `x' `i'
                if `i' == 1 {
                    local notes = "`n`i''"
                }
                else {
                    local notes = "`notes'" + " | " + "`n`i''"
                }
            }
            putexcel F`r1' = "`notes'"
        }

    local ++r
    }

    ** Go back one row
    local --r
    putexcel (A`r':F`r'), border(bottom)
    putexcel (F1:F`r'), border(right)

    putexcel close

    ** Sheet: Details
     * =========================================================================
    putexcel set "`using'", sheet("Details") modify open
    local r = 0

    foreach var of varlist `varlist' {
        local lbe: value label `var'

        ** Name
        putexcel A`++r' = "Variable" B`r' = "`var'", bold

        ** Label
        local lab: variable label `var'
        putexcel A`++r' = "Label" B`r' = "`lab'"
        ** Notes
        local notes = ""
        notes _count k : "`var'"
        if `k' > 0 {
            forvalues i = 1/`k' {
                notes _fetch n`i' : `var' `i'
                if `i' == 1 {
                    local notes = "`n`i''"
                }
                else {
                    local notes = "`notes'" + " | " + "`n`i''"
                }
            }
        }
        putexcel A`++r' = "Notes" B`r' = "`notes'"

        ** Missing values
        count if missing(`var')
        local miss = `r(N)'
        local p_miss : display `miss' / _N
        putexcel A`++r' = "Missing"
        putexcel B`r' = `miss', nformat(number_sep)
        putexcel C`r' = `p_miss', nformat(percent_d2)

        local t : type `var'
        if regexm("`t'", "int|float|double") & "`lbe'" == "" {
            summarize `var', detail format

            local min : display %9.2g `r(min)'
            local p25 : display %9.2g `r(p25)'
            local p50 : display %9.2g `r(p50)'
            local p75 : display %9.2g `r(p75)'
            local max : display %9.2g `r(max)'

            putexcel A`++r' = "Distribution"
            putexcel B`r' = "Min"   C`r' = "Quartile 1" D`r' = "Median" E`r' = "Quartile 3" F`r' = "Max"
            putexcel B`++r' = `min' C`r' = `p25'        D`r' = `p50'    E`r' = `p75'        F`r' = `max'

        }

        if "`lbe'"! = "" {
            ** Code from Stata Blog: https://blog.stata.com/2013/09/25/export-tables-to-excel/
            tabulate `var', matcell(freq) matrow(names)

            putexcel A`++r' = ("Labels") B`r' = ("Value") C`r' = ("Freq.") D`r' = ("Percent") E`r' = ("Cum.")

            local rows = rowsof(names)
            local row = 2
            local cum_percent = 0

            forvalues i = 1/`rows' {

                    local val = names[`i',1]
                    local val_lab : label (`var') `val'

                    local freq_val = freq[`i',1]

                    local percent_val = `freq_val'/`r(N)'*100
                    local percent_val : display %9.2f `percent_val'

                    local cum_percent : display %9.2f (`cum_percent' + `percent_val')

                    putexcel A`++r' = ("`val_lab'") B`r' = (`val')  C`r' = (`freq_val') D`r' = (`percent_val') E`r' = (`cum_percent')
            }

            putexcel A`++r' = ("Total") B`r'=(r(N)) C`r' = (100.00)
        }

        ** Add an empty line between variables
        local r = `r' + 1
    }
    putexcel close
}
end

