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


**Build on win32**

Go to develop dir, run `nimble build` to download dependancy first.

You have to use `gcc` build on windows instead of `vcc` build. Before built, something has to be changed:
first, you have to comment three lines in `C:\Users\Administrator\.nimble\pkgs\zip-0.1.0\zip\private\libzip_all.c`
<code>
//#define HAVE_FSEEKO 1 --line 276 
//#define HAVE_FTELLO 1 --line 279
//#define HAVE_MKSTEMP 1 --line 291
</code>

then you have to change  `nim`  source a little:
add ',' in `SymChars{}`  in  `%NIM_HOME%\lib\pure\parsecfg.nim`

`nim c --cpu:i386  -d:profile xlsx2csv.nim`
`-d:profile` is optional

**Build on other platform**
Build on other platform has not been tested, but 32bit is recommended. On linux, changing libzip_all.c maybe not necessary.


Of couse, there are still something to impove, as I just studied [nimlang](http://nim-lang.org/) for a month.

Licence : MIT
