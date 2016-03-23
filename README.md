# xlsx2csv
Small memory footprint xlsx2csv with useful options, built with  [nimlang](http://nim-lang.org/) .

`Small memory footprint` means using `stream` instead of `dom` when parsing xml. In nonofficial test, using less than 8MB to transfer 80000*64 cells into csv. showed by nim procedure `getTotalMem()`.

**usage**: 
 `xlsx2csv [src xlsx] [dest csv]`

If `src xlsx` is empty, choose from file dialog

**options** 
You can change config in xlsx2csv.ini in the same dir.

***example config***

#skipFirstRows = 1
Skip first rows, can not be negative

#skipLastRows = 1
Can not be negative, not include empty rows, as describe below

#skipCols = 0
Skip first columns, can not be negative
 
#successiveEmptyRows = 3
Can not be negative, when found successive Empty Rows, stopped and rollback to row before empty , these empty rows would not be written

#dateTypeCols = N,O,P
Denote which column is date type, as internal value is number ,it will be converted to `date` string( no time part ,may add config to support it) .
Note dates before 1970 cannot be supported now

Of couse, there are still something to impove, as I just studied [nimlang](http://nim-lang.org/) for a month.

Licence : MIT
