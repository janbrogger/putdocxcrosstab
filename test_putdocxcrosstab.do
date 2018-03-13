
capture program drop putdocxcrosstab
sysuse auto , clear
egen pricecat = cut(price), at(0,5000,10000,999999) label
label variable pricecat "Price (categorical)"
tab pricecat foreign

putdocx clear
putdocx begin
set tr off
putdocxcrosstab pricecat  foreign

putdocxcrosstab pricecat  foreign , norows
putdocxcrosstab pricecat  foreign , nocols
putdocxcrosstab pricecat  foreign , norowsum nocolsum


putdocxcrosstab pricecat  foreign , nofreq row

putdocxcrosstab pricecat  foreign , row
putdocxcrosstab pricecat  foreign , col

shell taskkill /F /IM WinWord.exe
tempfile doc
putdocx save "`doc'.docx", replace
shell "`doc'.docx"

