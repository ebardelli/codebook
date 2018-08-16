/*
 * Codebook to Do File
 *
 * Version 0.1
 * Emanuele Bardelli <bardelli@umich.edu>
 *
 */

version 13
cap program drop cb2dofile
program define cb2dofile
quietly {
    syntax, Codebook(string) DOfile(string) [sheet(passthru) firstrow clear]
    cap putexcel close

    import excel "`codebook'", `sheet' `firstrow' `clear'
    drop if missing(VariableLabel)

    file open dofile using "`dofile'", write text replace
    local handle = "dofile"

    local N = _N

    forvalues i = 1/`N' {
        file write `handle' "label variable `=VariableName[`i']' "
        file write `handle' `"`=char(34)'"'
        file write `handle' `"`=VariableLabel[`i']'"'
        file write `handle' `"`=char(34)'"'
        file write `handle' _n
    }
    file close dofile  
}
end
