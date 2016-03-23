import strutils,hashes,nre,options,times

type 
  CELL_TYPE* = enum
    ctString, ctNumber,ctError , ctBlank

#不能直接使用tuple, 对部分xlsx有问题?   
type 
  Position* = ref PositionObj
  PositionObj = tuple[row:int,col:int] #0-based
 
type 
  Style* = object #reserved  
 

type 
   Dimension* = tuple[leftTop: Position, rightBottom:Position  ]       
   
  
type 
   Sheet* = ref sheetObj
   sheetObj = object
     worksheet*:Worksheet
     rows*: seq[Row]
     dimension*:Dimension
     activeCell*:Cell
     selected*: bool
      
   Worksheet* = ref worksheetObj
   worksheetObj = object
    sheets*: seq[Sheet]
                   
  
   Row* = ref rowObj
   rowObj = object
     sheet*: Sheet
     cells*: seq[Cell]
 
   Cell* = ref cellObj
   cellObj = object
     row* : Row
     value*: string
     ctype*: CELL_TYPE 
     position*: Position
     formula*: string
     style* : Style


# proc hash*(x: Position): Hash =
  # Piggyback on the already available string hash proc.
  #
  # Without this proc nothing works!
  # see http://nim-lang.org/docs/hashes.html#!& and ...!$
#   result = x.row.hash !& x.col.hash
#   result = !$result
  # Computes a Hash from `x`.
#   var h: Hash = 0
#   h = h !& hash(x.row)
#   h = h !& hash(x.col)
#   result = !$h
    

proc toColVal*(idx: int) :string =
     var excelColNum:int = 1 + idx
     var colRemain:int = excelColNum;
     var s:seq[string] = @[]
     while colRemain > 0 :
            var thisPart:int = colRemain mod 26
            if thisPart == 0 : 
                thisPart = 26
            else:
                colRemain = (colRemain - thisPart) div 26

            # The letter A is at 65
            var colChar:char = char(thisPart+64);
            s.insert( $colChar, 0);
     result = join(s) 

# AA=26, BA=52
const ABSOLUTE_REFERENCE_MARKER = '$'
proc toColIdx*(str: string) :int  {.raises: [ValueError].}=
        var retval:int=0 
 
        for k,thechar in str.toUpper:
            if thechar == ABSOLUTE_REFERENCE_MARKER:
                if k != 0 :
                    raise newException(ValueError ,"Bad col ref format '" & str & "'");
                
                continue;
               
            # Character is uppercase letter, find relative value to A
            retval = (retval * 26) + (thechar.ord - 'A'.ord) + 1;
         
        retval-1

proc `$`*(p:Position):string =
    $p.row & "," & $p.col 

proc newPosition*(row:int, col:int) : Position =
     let p = new(Position)  
     p.row = row
     p.col = col
     p
  
proc newPosition*(excelColRow: string) : Position =
    #  echo excelColRow
     if excelColRow == nil or excelColRow == "":
         raise newException(ValueError,"excelColRow is empty")  
 
     var row: string
     var col:string
     let od = excelColRow.match(re"([a-zA-Z]+)(\d+)")
     if od.isSome:
       col = od.get.captures[0]
       row = od.get.captures[1]  # maybe error for Dimension parsing!
        
    #  echo row,",", col , 
     newPosition(parseInt(row) - 1 ,toColIdx(col))
    #  (parseInt(row),toColIdx(col))
    
     

#A1:C3 
proc newDimension*(excelDimension : string) : Dimension  =
   var dim:Dimension
   if excelDimension.contains(":"):
     let twoPart = excelDimension.split(":")
     dim = ( newPosition(twoPart[0]), newPosition(twoPart[1]) )   
   else:
     raise newException(ValueError,"`:` is required for Dimension, please resave the xlsx")  
   dim

proc newCell*(p:Position,value:string, forumla:string = ""):Cell = 
  var cell:Cell = new(Cell)
  cell.position = p
  cell.value = value
  cell.formula = forumla
  
  result = cell
      
proc `$`*(c: Cell): string =
  if c.formula.len > 0:
    "(" & $c.position & ")[" & c.formula & "] "  & c.value
  else:
    "(" & $c.position & ") " & c.value

#EXCEL的日期是以序列数的形式存储的，即保存的日期实际是这个日期到1900-1-1相差的天数。
#当单元格格式为数值时，就会显示出这个差值，即2006-12-1与1900-1-1相差39052天。
#这里不能用const, 否则编译错误 cannot use 'importc' variable at runtime !
let Date_1970 = times.timeInfoToTime(parse("1970-1-2","yyyy-M-d")) #最早时间
#19000101~19700101 = 25567
let DATES_FROM_1900_TO_1970 = 25567 + 3
proc numberToDateString*(num:int): string =
   if num < 0 :
     return "ERROR"
   else:
     let days = initInterval(days=num-DATES_FROM_1900_TO_1970)
    #  const Date_1900 = parse("1900-1-1","yyyy-M-d")  
     return format(getLocalTime( Date_1970 + days), "yyyy-MM-dd")
     
when isMainModule:
    echo """toColVal(52)=""" & toColVal(52)
    echo """toColIdx("BA")=""" & $toColIdx("BA")
    assert($toColIdx("BA") == $52)
    echo numberToDateString(40557)   
     