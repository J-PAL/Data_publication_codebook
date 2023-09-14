local search_directory = "" //INSERT THE PATH TO THE FOLDER YOU WOULD LIKE TO MAKE A CODEBOOK OF

	
************ Getting list of files to loop through ***********************
	tempfile file_list 
	filelist, directory(`search_directory') pattern("*.dta")
	gen temp="/"
	egen file_path = concat(dirname temp filename)
	keep file_path
	save `file_list'

	qui count
	local total_files = `r(N)'
	forvalues i=1/`r(N)' {
		local file_`i' = file_path[`i']
	}

	
	
********** Opening output file *************
	capture file close cb
	file open cb using codebook_temp.csv, write replace text
		foreach header in "dataset" "name" "min" "max" "median" "mean" "sd"{
		file write cb _char(34) `"`header'"' _char(34) ","
	}
	file write cb _n
	loc num_lab_tot = 0
	forvalues i = 1/`total_files'{
		use "`file_`i''", clear
			preserve
			uselabel, clear
			count
			loc num_lab = `r(N)'
			local num_lab_tot = `num_lab_tot' + `num_lab'
			if `num_lab' == 0{
				restore
			}
			else{
			drop trunc
			tempfile lab_temp
			save "`lab_temp'", replace
			restore
			if `i' == 1{
			preserve
			use "`lab_temp'", clear
			tempfile labl
			save "`labl'", replace
			restore
			}
			else{
			preserve
			use "`labl'", clear
			append using "`lab_temp'"
			duplicates drop
			save "`labl'", replace
			restore
			}
			}
		*** Exporting vars that come with built-in describe command
			preserve
			describe, replace clear
			drop if substr(name, 1, 2) == "__"
			tempfile cb_temp
			save "`cb_temp'", replace
			restore
			if `i' == 1{
			preserve
			use "`cb_temp'", clear
			tempfile cb 
			save "`cb'", replace
			restore
			}
			else{
			preserve
			use "`cb'", clear
			append using "`cb_temp'"
			tempfile cb
			duplicates drop name, force
			save "`cb'", replace
			restore
			}
		foreach var of varlist * {
			tempvar nm2
			qui egen `nm2' = total(!missing(`var'))
			if `nm2' == 0{
			drop `var'
			continue 
			}
			drop `nm2'
			***First column: Dataset name
			file write cb _char(34) `"`file_`i''"' _char(34) ","
			***Second column: Var name
			file write cb _char(34) `"`var'"' _char(34) ","
			capture decode `var', gen(_`var')
			if _rc==0{
				drop `var'
				ren _`var' `var'
				}
			capture confirm string var `var'
			if _rc==0 {
			forvalues iter = 1/5{
			file write cb _char(34) "N/A" _char(34) ","
			}
				}
			else{
				qui: sum `var', det
				local min = "`r(min)'"
				file write cb _char(34) "`min'" _char(34) ","
				local max = "`r(max)'"
				file write cb _char(34) "`max'" _char(34) ","
				local median = "`r(p50)'"
				file write cb _char(34) "`median'" _char(34) ","
				local mean = "`r(mean)'"
				file write cb _char(34) "`mean'" _char(34) ","
				local stdev = "`r(sd)'"	
				file write cb _char(34) "`stdev'" _char(34) ","
			}
		file write cb _n
			}
 di "File `i' done"
	}
	
	file close cb
	import delimited "codebook_temp.csv", varnames(1) clear
	duplicates drop name, force
	merge 1:1 name using "`cb'"
	drop if _merge == 2
	drop _merge
	drop position v8
	duplicates drop name, force
	order varlab vallab isnumeric type format, after(name)
	replace dataset = subinstr(dataset,"`search_directory'","",.)
	
	foreach x of varlist min-sd{
		qui: destring `x', force replace
		qui: replace `x' = round(`x', .001)
	}
	sort dataset 
	qui: count
	local var_n = `r(N)' + 1
	export excel using  "codebook.xlsx", firstrow(variables) sheet("Variables") locale(C) replace
	rm "codebook_temp.csv"
	
	if `num_lab_tot' >0 {
	use "`labl'",clear
	qui: count
	local val_n = `r(N)' + 1
	export excel using  "codebook.xlsx", firstrow(variables) sheet("Value labels") locale(C)
	}
	//Formatting
	putexcel set codebook, modify sheet("Variables")
	putexcel A1:L`var_n', overwritefmt border(all) 
	putexcel save
	if `num_lab_tot' >0 {
	putexcel set codebook, modify sheet("Value labels")
	putexcel A1:K`val_n', overwritefmt border(right) 
	putexcel save
	}

	mata:b = xl()
	mata:b.load_book("codebook.xlsx")
	mata:b.set_sheet("Variables")
	mata:b.set_column_width(2,2,33) //make title column widest
	mata:b.set_column_width(3,3,75) //make each column fit title
	mata:b.close_book()
	
	if `num_lab_tot' >0 {
	mata:b = xl()
	mata:b.load_book("codebook.xlsx")
	mata:b.set_sheet("Value labels")
	mata:b.set_column_width(1,1,33) //make title column widest
	mata:b.set_column_width(3,3,75) //make each column fit title
	mata:b.close_book()
	}

	
	
di ""
di "---------------------------------------------------------------------"
di ""
di "Finished"
di ""
di "---------------------------------------------------------------------"
di ""


