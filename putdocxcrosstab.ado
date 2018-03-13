*! Original author : Jan Brogger (jan@brogger.no)
*! Description     : Produces twoway tables with putdocx.
*! Maintained at   : https://github.com/janbrogger/putdocxcrosstab
capture program drop putdocxcrosstab
program define putdocxcrosstab
	version 15.1
	syntax varlist(min=2 max=2) , [noROWSum] [noCOLSum] [TItle(string)] [MIssing] [noFREQ] [row] [col]
	
	capture putdocx describe
	if _rc {
		di in smcl as error "ERROR: No active docx."
		exit = 119
	}
	tokenize "`varlist'"
	
	local var1 "`1'"	
	local var2  "`2'"	
	local varlab1 : variable label `var1'
	if "`varlab1'"=="" {
		local varlab1 "`var1'"
	}
	local varlab2 : variable label `var2'
	if "`varlab2'"=="" {
		local varlab1 "`var2'"
	}
	if "`title'"=="" {		
		local title `"Crosstabulation of `varlab1' by `varlab2' "'
	}
	
	if "`row'"!="" & "`col'"!="" {
		di as err "Specify only one of row or column options"
		error -1
	}
	
	tabulate `var1' `var2' , `missing'
	local nrows=`r(r)'+1
	local ncols=`r(c)'+1
	
	if "`rowsum'"!="norowsum" {
		local ncols=`ncols'+1
	}
	
	if "`colsum'"!="nocolsum" {
		local nrows=`nrows'+1
	}
	
	if `ncols'>63 {
		di as error "More than 63 columns are not allowed"
		error -1
	} 
	
	tempname mytable
	putdocx table `mytable' = (`nrows', `ncols') , title("`title'") 
	
	** Write out row headers
	qui levelsof `var1' , local(levels1) `missing'
	local vallab1 : value label `var1'
	local currentrow=3
	local currentcol=1
	foreach val1 in `levels1' {
		if "`vallab1'"!="" {
			local value1 : label `vallab1' `val1'			
		}
		else {
			local value1 `val1'
		}
				
		putdocx table `mytable'(`currentrow',`currentcol') = (`"`value1'"'), halign(left)
		local currentrow=`currentrow'+1
	}
	
	** Write out column headers
	qui levelsof `var2' , local(levels2) `missing'
	local vallab2 : value label `var2'
	local currentrow=2
	local currentcol=2
	foreach val2 in `levels2' {
		if "`vallab2'"!="" {
			local value2 : label `vallab2' `val2'
		}
		else {
			local value2 `val2'
		}
				
		putdocx table `mytable'(`currentrow',`currentcol') = (`"`value2'"'), halign(left)
		local currentcol=`currentcol'+1		
	}
	
	** Write out cell counts, percentages or both
	local sum = 0
	qui levelsof `var2' , local(levels2) `missing'
	local vallab2 : value label `var2'
	local startrow=3
	local startcol=2
	local currentrow=`startrow'
	local currentcol=`startcol'	
	foreach val1 in `levels1' {
		qui count if `var1'==`val1'
		local rowcount=`r(N)'
		foreach val2 in `levels2' {
			qui count if `var2'==`val2'
			local colcount=`r(N)'
			qui count if `var1'==`val1'	& `var2'==`val2'		
			local cellcount=`r(N)'
			local sum=`sum'+`cellcount'
			local rowperc=`cellcount'/`rowcount'*100
			local colperc=`cellcount'/`colcount'*100
			local rowpercf : di %3.1f `rowperc'
			local colpercf : di %3.1f `colperc'
			if "`freq'"=="nofreq" {
				if "`row'"=="row" {
					local cell "`rowpercf' %"
				}
				else if "`col'"=="col" {
					local cell "`colpercf' %"
				}
			}
			else {
				local cell "`r(N)'"
				if "`row'"=="row" {
					local cell "`cell' (`rowpercf'%)"
				}
				else if "`col'"=="col" {
					local cell "`cell' (`colpercf'%)"
				}
			}
			
			putdocx table `mytable'(`currentrow',`currentcol') = ("`cell'"), halign(left)
			local currentcol=`currentcol'+1									
		}
		local currentrow=`currentrow'+1
		local currentcol=`startcol'
	}
	
	** Write out row sums
	if "`rowsum'"!="norowsum" {
		qui levelsof `var1' , local(levels1) `missing'
		local currentrow=3
		local currentcol=`ncols'
		foreach val1 in `levels1' {						
			qui count if `var1'==`val1'	
			local rowcount=`r(N)'
			qui count 
			local totalcount= `r(N)'
			local rowperc=`rowcount'/`totalcount'*100
			local rowpercf : di %3.1f `rowperc'
			if "`freq'"=="nofreq" {
				if "`row'"=="row" {
					local cell "`rowpercf' %"
				}
				else if "`col'"=="col" {
					local cell "`colpercf' %"
				}
			}
			else {
				local cell "`r(N)'"
				if "`row'"=="row" {
					local cell "`cell' (`rowpercf'%)"
				}
				else if "`col'"=="col" {
					local cell "`cell' (`colpercf'%)"
				}
			}
			putdocx table `mytable'(`currentrow',`currentcol') = ("`r(N)'"), halign(left)
			local currentrow=`currentrow'+1
		}
		putdocx table `mytable'(2,`ncols') = ("Total"), halign(left)
	}
	
	** Write out column sums
	if "`colsum'"!="nocolsum" {
		qui levelsof `var2' , local(levels2) `missing'
		local currentrow=`nrows'+1
		local currentcol=2
		foreach val2 in `levels2' {
			qui count if `var2'==`val2'				
			local colcount=`r(N)'
			qui count 
			local totalcount= `r(N)'
			local colperc=`colcount'/`totalcount'*100
			local colpercf : di %3.1f `colperc'
			
			if "`freq'"=="nofreq" {
				if "`row'"=="row" {
					local cell "`rowpercf' %"
				}
				else if "`col'"=="col" {
					local cell "`colpercf' %"
				}
			}
			else {
				local cell "`r(N)'"
				if "`row'"=="row" {
					local cell "`cell' (`rowpercf'%)"
				}
				else if "`col'"=="col" {
					local cell "`cell' (`colpercf'%)"
				}
			}
			
			putdocx table `mytable'(`currentrow',`currentcol') = ("`r(N)'"), halign(left)
			local currentcol=`currentcol'+1
		}
		putdocx table `mytable'(`currentrow',1) = ("Total"), halign(left)
	}
	
	if "`colsum'"!="nocolsum" & "`rowsum'"!="norowsum" {
		local currentrow=`nrows'+1
		qui count 
		local totalcount= `r(N)'
		local sumperc=`totalcount'/`sum'*100
		local sumpercf : di %3.1f `sumperc'
		if "`freq'"=="nofreq" {
				if "`row'"=="row" {
					local cell "`sumpercf' %"
				}
				else if "`col'"=="col" {
					local cell "`sumpercf' %"
				}
			}
			else {
				local cell "`totalcount'"
				if "`row'"=="row" {
					local cell "`cell' (`sumpercf'%)"
				}
				else if "`col'"=="col" {
					local cell "`cell' (`sumpercf'%)"
				}
			}
		putdocx table `mytable'(`currentrow',`ncols') = ("`cell'"), halign(left)
	}
	
end
