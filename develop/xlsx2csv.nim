import strutils, strtabs,tables,encodings, os,parseXml, zip/zipfiles,times

import xlsx

from native_dialogs import callDialogFileOpen,callDialogFileSave

# 仅支持32位 只能GCC build
# 需要修改.nimble\pkgs\zip-0.1.0\zip\private\libzip_all.c
# 即使是32位系统也要--cpu:i386 -t:-m32 -l:-m32
# nim c --cincludes:include -l:zlib1.dll --cpu:i386 -t:-m32 -l:-m32 -t:-DWIN32  -d:profile xlsx2csv.nim

# Dependencies

#requires "nim >= 0.13.0"
#requires "zip"
#requires "nre"
#requires "native_dialogs" 

{.passC:"-Iinclude -m32 -DWIN32".}
{.passL:"zlib1.dll -m32".}

{.hint: "Compiled with `-Iinclude -m32 -DWIN32`".}
{.hint: "Linked with  `zlib1.dll -m32`".}
# nim c --cpu:i386  -d:profile xlsx2csv.nim

let encoding = getCurrentEncoding().toLower
echo "Current encoding: " &  encoding
let u2aConverter = open("gb2312", "utf-8") #仅限源代码 
let a2uConveter = open("utf-8","gb2312") #反向

proc u2a(src:string):string = 
  if encoding != "utf-8":
    u2aConverter.convert(src)
  else:
    src

proc a2u(src:string):string = 
  if encoding != "utf-8":
    a2uConveter.convert(src)
  else:
    src    
 

const SST_SI = "si"
const SST_DATA_TAG = "t"
 
const DIMENSION_TAG = "dimension"
const ROW_TAG = "row"
const CELL_DATA_TAG = "c"
const CELL_DATA_TAG_ROW_ATTR = "r"
const CELL_DATA_TAG_TYPE_ATTR = "t"
const CELL_VALUE_TAG = "v"
const CELL_FORMULA_TAG = "f"
 

#------------------------------------------
  
when isMainModule:
    import parsecfg,streams
    var skipFirstRows = 0
    var skipLastRows = 0
    var skipCols = 0
    var successiveEmptyRows = 0 #连续空行数退出
    var dateTypeCols:seq[string]
    var dateTypeColNums:seq[int] = @[]
    
    var (_, name, _) = splitFile(getAppFilename())
    var conf = newFileStream(name & ".ini" , fmRead)
    defer: conf.close()
    
    {.hint:"Ensure SymChars{} has `,` in  parsecfg.nim" .}
    var p: CfgParser ##需要修改parsecfg.nim symChar{} 增加 ','
    open(p, conf, name)
    while true:
        var e = next(p)
        case e.kind
        of cfgEof:
        #   echo("EOF!")
          break
        of cfgSectionStart:   ## a ``[section]`` has been parsed
        #   echo("new section: " & e.section)
           discard
        of cfgKeyValuePair:
        #   echo("[debug]key-value-pair: " & e.key & ": " & e.value)
          if e.value == nil or e.value.strip == "":
               discard
          case e.key.toLower
          of "skipFirstRows".toLower: 
            skipFirstRows = parseInt(e.value) #ValueError may occurred
            if skipFirstRows < 0 : skipFirstRows = 0
          of "skipLastRows".toLower:
            skipLastRows = parseInt(e.value) #ValueError may occurred
            if skipLastRows < 0 : skipLastRows = 0
          of "skipCols".toLower:
            skipCols = parseInt(e.value) #ValueError may occurred
            if skipCols < 0 : skipCols = 0 
            
          of "successiveEmptyRows".toLower:
            successiveEmptyRows = parseInt(e.value)  
            if successiveEmptyRows < 0 : successiveEmptyRows = 0
          of "dateTypeCols".toLower:
            #需要修改parsecfg.nim symChar{}
            dateTypeCols = e.value.split(",")
            
            for c in dateTypeCols:
            #   echo c,"," ,xlsx.toColIdx(c)
              dateTypeColNums.add(xlsx.toColIdx(c.strip))
              
          else:
            echo "[conf]dicard line:" & $p.getLine()
            discard
            
        of cfgOption:
           discard
        #   echo("command: " & e.key & ": " & e.value)
        of cfgError:
          echo(e.msg)
          close(p)
        else:
          echo("cannot open config file: " & name)
    
    echo u2a("路径不支持空格, 仅限xlsx文件有效")
    echo u2a("*忽略读取`首行`数:") & $skipFirstRows
    echo u2a("*忽略读取`尾行`数:") & $skipLastRows
    echo u2a("*忽略读取`首列`数:") & $skipCols
    echo u2a("*连续空行数退出设置:") & $successiveEmptyRows
    echo u2a("*日期列表(字母):") & $dateTypeCols
 
   #---------------------------------------------------------------------
    var srcFile:string #utf8
    var inputSrcFile:string #ansi
   
    var destFile:string #utf8
    var inputDestFile:string #ansi
    
    if os.paramCount() > 0:
     inputSrcFile = $paramStr(1) #TaintedString, 已经是utf-8 !
     srcFile = inputSrcFile
     inputSrcFile = u2a(inputSrcFile) #确保是ansi
     if os.paramCount() > 1:
       inputDestFile = $paramStr(2)
       destFile =  inputDestFile
       inputDestFile = u2a(inputDestFile) #确保是ansi
       
    else:
       inputSrcFile =  callDialogFileOpen("Open File")
       inputDestFile =  callDialogFileSave("Save File")
       srcFile = a2u(inputSrcFile) # for split file
       destFile =  a2u(inputDestFile) 
    
 
    if inputSrcFile == nil :
      echo u2a("输入文件为空")
      quit(-1)

    
    let sf = splitFile(srcFile)
#    echo sf
    if sf.ext == nil or sf.ext.toLower != ".xlsx":
     echo u2a("输入文件不是xlsx: ") & inputSrcFile
     quit(-1)
   
    if inputDestFile == nil:
      if sf.dir == nil or sf.dir == "":
        destFile = sf.name &  ".csv" 
      else: 
        destFile = sf.dir & os.DirSep &  sf.name &  ".csv"
 
     
     
    let st = getTime()
    
    
    var za:ZipArchive
    var suc = za.open(inputSrcFile)
    defer: za.close()
    
    if not suc:
      echo u2a("无法打开 xlsx文件: ") & inputSrcFile 
      quit(-1)
      
    else:
      echo u2a("读取 xlsx文件: ")  & inputSrcFile
    # You can extract to /exmple
    #let destDir = "example"
    #os.createDir(destDir)
    #za.extractAll(destDir)
  
 #=====================================================================       
    #get sst data
    let sst = za.getStream("xl/sharedStrings.xml")
    var x: XmlParser
    open(x, sst, "sharedStrings.xml")
    #放到seq表中...
    var sstdata: seq[string] = @[] 
    var sidata: seq[string]
    var lastElementName:string
    var counter:int
    #
    #  <sst count="1"><si><t>合并&#10;标题</t></si></sst>
    #  <si><r><t xml:space="preserve">   xxxx   </t></r>
    #      <r><t xml:space="preserve">   yyyyy   </t></r>
    #     <phoneticPr fontId="4" type="noConversion"/>
    #  </si> 
    while true:
        x.next()
        # echo x.kind
        case x.kind
        of xmlElementStart,xmlElementOpen: 
        #    echo x.elementName
           lastElementName = x.elementName.toLower 
          
        of xmlElementEnd:
           inc(counter)
        #    if counter > 20 : #for debug
        #       break
           
           if x.elementName.toLower == SST_SI:
               if sidata == nil: #may be trimed
                #  echo lastElementName, ": sidata == nil, " ,$counter
                 sstdata.add("") #placeholder
               else:
                 sstdata.add( sidata.join().strip )
            #    echo "sstdata length=" , sstdata.len , ",",u2a(sstdata[sstdata.len-1])
               sidata = nil
            
        of xmlCharData,xmlWhitespace,xmlCData,xmlSpecial: 
        #    echo lastElementName, "> DATA=", u2a(x.charData)
           if lastElementName == SST_DATA_TAG:
              if sidata == nil:
                sidata = newSeq[string]()
              sidata.add(x.charData) #no strip for xml:space="preserve"
           
        of xmlEof: break # end of file reached
        else: discard # ignore other events

    x.close()
    
    echo "sstdata length=" , sstdata.len
     
    when defined(profile): 
        echo "got All SST @"  & format(getLocalTime(getTime()), "mm:ss") 
        echo "total/occupied/free: ", formatSize(getTotalMem()) , "\t", getOccupiedMem(),"(bytes)\t",getFreeMem()
        
        # just for debug
        # let sstdatafile = open("profile_sstdata.txt",fmWrite)
        # defer: sstdatafile.close()
        # var c = 0
        # for v in sstdata:
        #    sstdatafile.write(c)
        #    sstdatafile.write("\t")
        #    sstdatafile.write(v)
        #    sstdatafile.write("\n")
        #    inc(c)

#============================================================================
    # get Dimension & data
    let ss = za.getStream("xl/worksheets/sheet1.xml")

    open(x, ss, "sheet1.xml") 

    
    var f = open(destFile, fmWrite)
    defer: f.close()
    
    echo u2a("CSV文件保存到: ") & u2a(destFile)
    echo u2a("***请确保同名CSV文件未被其他进程占用***\n")
    echo u2a("正在保存...\n")
    
    var quote:bool
    var v: Cell
    
    var dim: Dimension
    var pos: Position
    var valType: string
    var value:string
    var formula:string
    var cell : Cell
    # var dataTable: TableRef[Position, Cell] = newTable[Position, Cell]() 
    var dataTable = newTable[int, Cell]() #重构:保存仅仅一行数据?
   
    var effectRows = 0
    var currentRow:Row
    new(currentRow)
     
    var lastEmptyRow = -1
    var emptyRows=0
    
    # var firstWriteLine = true
    var bufferLines:seq[string] = @[] # range 0 .. skipLastRows, 确保最后写入
    var max_lines = max(5, max(skipLastRows, successiveEmptyRows))
    
    #[
      <dimension ref="A1:C3"/>
      ... 
      <sheetData>  
      <row r="1" spans="1:3">
        <c r="A1" t="s"><v>0</v></c>
        <c r="B1" s="1" t="s">
          <v>1</v>
        </c>
        <c r="C1" s="2"/>
      </row>
      ...
      </sheetData>
      
      xmlElementOpen --> xmlAttribute+ --> xmlElementClose  --> xmlElementStart? --> xmlCharData --> xmlElementEnd 
      xmlElementStart --> xmlCharData --> xmlElementEnd
    ]#
    while true:
        x.next()
        # echo "kind=" ,x.kind
        case x.kind
        of xmlElementStart, xmlElementOpen:
            # echo("<$1>" % x.elementName)
            lastElementName = x.elementName
            if  x.elementName.toLower == ROW_TAG:
              currentRow.cells = @[]  
              dataTable = newTable[int, Cell]()
             
        of xmlElementEnd:
            # echo("</$1>" % x.elementName)
            case x.elementName.toLower
            of ROW_TAG: 
              if effectRows >= skipFirstRows: #skip top rows
                #判断是否行是否有效
                var empty = true
                for c in currentRow.cells:
                    if c.value != nil and c.value != "":
                      empty = false
                      break 
                
                if empty :
                    if lastEmptyRow == -1:
                      emptyRows = 1
                    elif  effectRows - lastEmptyRow == 1:
                    #连续
                      inc(emptyRows)
                    else : 
                      emptyRows = 1
                    
                    lastEmptyRow = effectRows  
                    echo "**!**empty line at line " & $(effectRows+1) & ", successive=" & $emptyRows
                    if emptyRows >= successiveEmptyRows: 
                       effectRows = effectRows - emptyRows
                       #remove buffer lines
                       var rm = emptyRows
                       while bufferLines.len > 0 and rm > 0 :
                         bufferLines.del(bufferLines.len-1)  #remove last if any
                         rm.dec
                         
                       break
                #DO buffer here ! 
                var cols:seq[string] = @[]
                for j in skipCols .. dim.rightBottom.col: 
                      v = dataTable.getOrDefault( j) #newPosition(i,j)
                      if isNil(v):
                         continue
                         
                      quote = if v.value.contains(' ')  or v.value.contains(","): true else: false
                      if quote:
                         cols.add("\"" & v.value & "\"")
                      else:  
                         cols.add(v.value)
                # echo "adding row ... " & cols[2]         
                bufferLines.add(cols.join(","))  
                       
                #write buffer here
                if  bufferLines.len == max_lines :
                    # if not firstWriteLine:
                    #    f.write("\n")
                    # else:   
                    #    firstWriteLine = false
                       
                    f.write(bufferLines.join("\n"))
                    f.write("\n")
                    bufferLines = @[]
                     
              inc(effectRows)
            of CELL_DATA_TAG:  
              if effectRows >= skipFirstRows: #skip top rows
                if valType == "s":
                    #  echo "parsing value from sst[", value,"]"
                    value = sstdata[parseInt(value)] #index from 0
                    
                if value == nil:
                    continue   
                
                var col = currentRow.cells.len 
                #日期类型转换
                if dateTypeColNums.contains(col):
                  var num = -1
                  try: 
                    num = parseInt(value)
                    value = numberToDateString(num) 
                  except: discard
                  
                    
                cell = newCell(pos, value, formula) 
                # echo "table count:" & $dataTable.len & ", inserting=" , u2a($cell)
                dataTable.add(pos.col, cell)
                currentRow.cells.add(cell)  
                
                #   when defined(profile):
                #      if dataTable.len mod 10000 == 0: 
                #         echo dataTable.len,"> ", getTotalMem() , "(bytes)\t", getOccupiedMem(),"\t",getFreeMem()  
                
 
              valType = nil
              value = nil
              formula = nil
              cell  = nil         
 
            else:
              discard  
                   
        of xmlAttribute:
            # echo "<" & lastElementName & "> " & x.attrKey & "=" & x.attrValue
            
            case x.attrKey.toLower
            of CELL_DATA_TAG_ROW_ATTR:
               #pos ref
               #必须的,避免重复
               if lastElementName == CELL_DATA_TAG:
                  pos = newPosition(x.attrValue)
               
            of CELL_DATA_TAG_TYPE_ATTR:
                #必须的,避免重复
               if lastElementName == CELL_DATA_TAG:
                  valType = x.attrValue
                
            of "ref":
                 if lastElementName == DIMENSION_TAG:  
                   dim = newDimension(x.attrValue) #crrent attr
            else:
              discard  
        
        of xmlCharData,xmlWhitespace,xmlCData,xmlSpecial:   
            # echo "DATA=" & x.charData
            if lastElementName == CELL_VALUE_TAG:
               value = x.charData
            elif lastElementName == CELL_FORMULA_TAG:
               formula = x.charData
            else : discard
                   
        of xmlEof: break # end of file reached
        else: discard # ignore other events
    
    # if not firstWriteLine:
    #      f.write("\n")
    # else:      
    #      firstWriteLine = false
    #skipLastRows
    var r =  if skipLastRows > 0 : bufferLines.len - skipLastRows else: bufferLines.len
    f.write(bufferLines[0..r-1].join("\n"))
    # for i in 0..<r: 
    #   f.write(bufferLines[i])

    x.close()
    
    echo "Dimension=" & $dim
    echo u2a("读取总行数: ") & $effectRows & u2a(", 写入总行数:") & $(effectRows - skipFirstRows - skipLastRows)& u2a(", 跳过总行数:") & $(skipFirstRows + skipLastRows) & u2a("跳过前[") & $skipCols & u2a("]列")
    
    let et = getTime()
    echo u2a("耗时(秒): ") & $(et - st) 
 
    when defined(profile): 
        echo "DONE  @"  & format(getLocalTime(getTime()), "mm:ss") 
        echo "total/occupied/free: ", formatSize(getTotalMem()) , "\t", getOccupiedMem(),"(bytes)\t",getFreeMem()
       
    close(u2aConverter)
    close(a2uConveter)