# xlsx2csv
Small memory footprint xlsx2csv with useful options, built with nim lang.
`Small memory footprint` means using `stream` instead of `dom` when parsing xml. In nonofficial test, using less than 8MB to transfer 80000*64 cells into csv. showed by nim procedure `getTotalMem()`.

***usage***: 
 `xlsx2csv [src xlsx] [dest csv]`

If `src xlsx` is empty, choose from dialog

***options*** 
You can change config in xlsx2csv.ini in the same dir.

***example config***

#skipFirstRows = 1
#skipLastRows = 1
not include empty rows, as describe below

#skipCols = 0
skip first columns
 
#successiveEmptyRows = 3
when found successive Empty Rows, stopped and rollback to row before empty

#dateTypeCols = N,O,P
denote which col is date type, as internal value is number ,it will convert to date( no time part ,may add config to support it) 
note dates before 1970 cannot be supported now

Of couse, there are still something to impove, as I just studied [nimlang](http://nim-lang.org/) for a month.

Licence : MIT
