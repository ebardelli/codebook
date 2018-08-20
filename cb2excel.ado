/*
 * Codebook
 *
 * Version 0.1
 * Emanuele Bardelli <bardelli@umich.edu>
 *
 */

version 13
cap program drop cb2excel
program define cb2excel
quietly {
    syntax [varlist(defaul=none)] using/, [replace modify prefix(string)]
    cap putexcel close

    ** Sheet: Dataset
     * =========================================================================
    if `c(version)' < 15 {
        putexcel set "`using'", sheet("`prefix'Dataset", replace) `replace' `modify'
        putexcel A1 = "Test"
    }
    else {
        putexcel set "`using'", sheet("`prefix'Dataset", replace) `replace' `modify' open
    }
    local row = 1

    describe, varlist
    local sorted = "`r(sortlist)'"

    ** Filename
    putexcel A`row' = "Dataset" B`row' = "`c(filename)'"

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
    putexcel A`++row' = "Prepared by" B`row' = "`c(username)'"
    putexcel A`++row' = "Prepared on" B`row' = "`c(current_date)' `c(current_time)'"

    cap putexcel close

    ** Sheet: Variable List
     * =========================================================================
    if `c(version)' < 15 {
        putexcel set "`using'", sheet("`prefix'Variables", replace) modify
        putexcel B1 = "Test"
    }
    else {
        putexcel set "`using'", sheet("`prefix'Variables", replace) modify open
    }

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
        local labels = ""

        if "`lab_name'" != "" {
            quietly levelsof `x', local(var_levels)
            local i = 0

            foreach var_level of local var_levels {
                local value_label : label (`x') `var_level'

                if "`labels'" == "" {
                    local labels = "`var_level': `value_label'"
                }
                else {
                    local labels = "`labels'`=char(10)'`var_level': `value_label'"
                }

                local ++i
                if `i' > 25 {
                    local labels = "`labels'`=char(10)'(...)"
                    continue, break
                }
            }
        }
        putexcel E`r' = "`labels'"

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

    cap putexcel close

    ** Sheet: Details
     * =========================================================================
     if !missing("`varlist'") {
        if `c(version)' < 15 {
            putexcel set "`using'", sheet("`prefix'Details", replace) modify
            putexcel A1 = "Test"
        }
        else {
            putexcel set "`using'", sheet("`prefix'Details", replace) modify open
        }
        local r = 0

        foreach var of varlist `varlist' {
            local lbe: value label `var'

            ** Name
            putexcel A`++r' = "Variable" B`r' = "`var'", bold

            ** Label
            putexcel A`++r' = "Label" B`r' = "`lbe'"
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
                quietly count if !missing(`var')
                local N = `r(N)'
                local cum_percent = 0

                quietly levelsof `var', local(var_levels)
                local i = 0

                foreach var_level of local var_levels {
                    local value_label : label (`var') `var_level'

                    quietly count if `var' == `var_level'
                    local freq_val = `r(N)'
                    local percent_val : display `freq_val' / `N' * 100
                    local cum_percent : display `cum_percent' + `percent_val'

                    putexcel A`++r' = ("`value_label'")
                    putexcel B`r' = (`var_level')
                    putexcel C`r' = (`freq_val'), nformat(number_sep)
                    putexcel D`r' = (`percent_val') E`r' = (`cum_percent'), nformat(number_d2)

                    local i = `i' + 1

                    if `i' > 100 {
                        putexcel A`++r' = ("...")
                        continue, break
                    }
                }
                putexcel A`++r' = ("Total")
                putexcel C`r' = (`N'), nformat(number_sep)
                putexcel D`r' = (`cum_percent'), nformat(number_d2)
            }

            ** Add an empty line between variables
            local r = `r' + 1
        }
        cap putexcel close
    }
}
end

