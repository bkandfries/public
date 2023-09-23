Attribute VB_Name = "ActiveMacros"
Sub SelectTabsSameColor()

  Dim wsColor() As Integer
  Dim ws As Worksheet
  Dim ind As Integer
  Dim tabColor As Long
  
  
  ReDim wsNames(0)
  ReDim wsColor(0)
  tabColor = ActiveSheet.Tab.Color
  wsNames(0) = ActiveSheet.Name
  wsColor(0) = ActiveSheet.Tab.ColorIndex
  
  For Each ws In ActiveWorkbook.Sheets
    If ws.Visible = xlSheetVisible Then
    If ws.Tab.ColorIndex = wsColor(0) Then
      ReDim Preserve wsNames(UBound(wsNames) + 1)
      ReDim Preserve wsColor(UBound(wsColor) + 1)
      wsNames(UBound(wsNames)) = ws.Name
      wsColor(UBound(wsColor)) = ws.Tab.ColorIndex
    End If
    
    End If
  Next ws
  
  Sheets(wsNames).Select
  End Sub
  

Sub ColorTabsFromOtherWS()
Attribute ColorTabsFromOtherWS.VB_Description = "q"
Attribute ColorTabsFromOtherWS.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim Window As Windows
    Dim wbstring As String
    Dim OtherWBName As String
    Dim sheetname As String
    wbstring = InputBox("Enter matching String")
    ActiveWbName = ActiveWorkbook.Name
    Dim cws As Worksheet
    Set cws = ActiveSheet

    
    OtherWBName = "Not Fount"
    iCntr = 0
    Dim wbnameArray() As Variant
    For i = 1 To Workbooks.count
        iName = Workbooks(i).Name
       
        pos = InStr(1, iName, wbstring, 1)
        
        If iName <> ActiveWbName Then
            If pos > 0 Then
                iCntr = i
                OtherWBName = iName
                 Debug.Print (iName & " (i:" & i & ") vs " & ActiveWbName)
                
            End If
    End If
    Next
    
    If OtherWBName <> "Not Found" Then
        For Each ws In ActiveWindow.SelectedSheets
          '  ws.Select
            'ws.Activate
            sheetname = ws.Name
            
            ws.Tab.Color = GetTabColor(sheetname, OtherWBName)
            
        Next ws
            
    Else: MsgBox ("No Matching Sheet Found")
    End If

    

End Sub

Sub CreateTableOfContents()
    Dim i As Long

    On Error Resume Next
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    On Error GoTo 0

    ActiveCell.EntireColumn.NumberFormat = "@"

    For i = 1 To Sheets.count
      If Sheets(i).Visible < 2 Then 'Exculde super hidden sheets -1: visable, 1,:hidden 2:super hidden
        'Step 5: Add Hyperlink
         ActiveSheet.Hyperlinks.Add _
         Anchor:=ActiveSheet.Cells(i + 1, 1), _
             Address:="", _
            SubAddress:="'" & Sheets(i).Name & "'!A1", _
                TextToDisplay:="'" & (Sheets(i).Name)
      End If
    Next i

    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

Sub A_unhideAllNames()
'Unhide all names in the currently open Excel file
    For Each tempname In ActiveWorkbook.Names
        tempname.Visible = True
    Next
End Sub

Sub A_HideAllNames()
'Unhide all names in the currently open Excel file
    For Each tempname In ActiveWorkbook.Names
        tempname.Visible = False
    Next
End Sub

Sub AddLamdbdas()
    '
    'oLD VALUES:

    'ActiveWorkbook.Names.Add Name:="FILTERISNUM", RefersToR1C1:="=LAMBDA(List1,List2,UNIQUE(FILTER(List1,ISNUMBER(XMATCH(List1,List2)),"""")))"
    'ActiveWorkbook.Names.Add Name:="FILTERNA", RefersToR1C1:="=LAMBDA(List1,List2,UNIQUE(FILTER(List1,ISNA(XMATCH(List1,List2)),"""")))"
    'ActiveWorkbook.Names.Add Name:="SplitString", RefersToR1C1:="=LAMBDA(pcstring,LET(cell,pcstring,XMLStr,""<Array><Cell>""&SUBSTITUTE(cell,"","",""</Cell><Cell>"")&""</Cell></Array>"",Data,FILTERXML(XMLStr,""//Cell""),cleaned,SUBSTITUTE(SUBSTITUTE(TRIM(Data),""'"",""""),"""""""",""""),SplitXML,FILTER(cleaned,NOT(ISERROR(cleaned))),FormattTo5Char,TEXT(SplitXML,""00000""),UNIQUE(FormattTo5Char)))"
    'ActiveWorkbook.Names.Add Name:="JoinArray", RefersToR1C1:="=LAMBDA(array,""'""&TEXTJOIN("","",TRUE,SORT(array)))"
    'ActiveWorkbook.Names.Add Name:="ComparePCString", RefersToR1C1:="=LAMBDA(list1,list2,JoinArray(FILTERNA(SplitString(list1),SplitString(list2))))"
    ActiveWorkbook.Names.Add Name:="COUNTU", RefersToR1C1:="=LAMBDA(FilteredColumn,LET(array,FilteredColumn, start,INDEX(array,1), SUM(IF(FREQUENCY(IF(SUBTOTAL(3,OFFSET(start,ROW(array)-ROW(start),,1)), IF(array>"""",MATCH(""~""&array,array&"""",0))),ROW(array)-ROW(start)+1),1))))"
    'ActiveWorkbook.Names.Add Name:="PadStart", RefersToR1C1:="=LAMBDA(str,LET(val,str,IF(LEN(val)<5,REPT(""0"",5-LEN(val))&val,val)))"

    'ActiveWorkbook.Names.Add Name:="APPENDCOLS", RefersToR1C1:="=LAMBDA(array1,array2,     LET(         array1Rows, ROWS(array1),         array1Cols, COLUMNS(array1),         array2Rows, ROWS(array2),         array2Cols, COLUMNS(array2),         rowLen, MAX(array1Rows, array2Rows),                 colLen, array1Cols + array2Cols,         newArray, SEQUENCE(rowLen, colLen),         colIndex, MOD(newArray - 1, colLen) + 1,         rowIndex, 1 + ((newArray - colIndex) / colLen),                  resultArray, IF(             colIndex > array1Cols,             INDEX(array2, rowIndex, colIndex - array1Cols),             INDEX(array1, rowIndex, colIndex)         ),                  resultArray     ) )"
    'ActiveWorkbook.Names.Add Name:="APPENDROWS", RefersToR1C1:="=LAMBDA(array1,array2,     LET(         array1Rows, ROWS(array1),         colIndex, SEQUENCE(, MAX(COLUMNS(array1), COLUMNS(array2))),         rowIndex1, SEQUENCE(array1Rows + ROWS(array2)),         rowIndex2, rowIndex1 - array1Rows,         IF(             rowIndex2 >= 1,             INDEX(array2, rowIndex2, colIndex),             INDEX(array1, rowIndex1, colIndex)         )     ) )"
    'ActiveWorkbook.Names.Add Name:="DROPCOL", RefersToR1C1:="=LAMBDA(array,column,     MAKEARRAY(         ROWS(array),         COLUMNS(array) -1,         LAMBDA(i,j, INDEX(array, i, IF(j <column, j, j+1)))     ))"
'    ActiveWorkbook.Names.Add Name:="IFBLANK", RefersToR1C1:="= LAMBDA(value,value_if_blank,    IF(ISBLANK(value), value_if_blank, value))"
    'ActiveWorkbook.Names.Add Name:="JOINARRAY", RefersToR1C1:="= LAMBDA(array,    LET(array, array,        IF(AND(COUNTA(array)=1,LEN(array)=0),"""",         ""'"" & TEXTJOIN("","", FALSE, SORT(array)))))"
    'ActiveWorkbook.Names.Add Name:="LEFTANTI", RefersToR1C1:="= LAMBDA(LeftArray,RightArray,    JOINARRAY(FILTERNA(SPLITSTRING(LeftArray), SPLITSTRING(RightArray))))"
    ActiveWorkbook.Names.Add Name:="PADSTART", RefersToR1C1:="=LAMBDA(text, num_dig, LET( text, text, num_dig, num_dig, out, IF( LEN(text) < num_dig, REPT(""0"", num_dig - LEN(text)) & text, text ), iferror(VALUETOTEXT(out),out) ))"
    'ActiveWorkbook.Names.Add Name:="RIGHTANTI", RefersToR1C1:="= LAMBDA(LeftArray,RightArray,    JOINARRAY(FILTERNA(SPLITSTRING(RightArray), SPLITSTRING(RightArray))))"
    'ActiveWorkbook.Names.Add Name:="SPLITSTRINGv2", RefersToR1C1:="=LAMBDA(string, LET( cell, string, str, SUBSTITUTE(CLEAN(TRIM(cell)), ""'"", """"), XMLStr, ""<Array><Cell>"" & SUBSTITUTE(str, "","", ""</Cell><Cell>"") & ""</Cell></Array>"", SplitXML, FILTERXML(XMLStr, ""//Cell""), GetUnique, sort(Unique(SplitXML)), GetUnique ))"
    ActiveWorkbook.Names.Add Name:="CONCATX", RefersToR1C1:="=LAMBDA(lookup_value, lookup_range, return_range, LET( lookuparray, lookup_value, lookuprange, lookup_range, returnrange, return_range, uniquelist, sort(unique(FILTER(returnrange,ISNUMBER(XMATCH(lookuprange,lookuparray)),""""))), concat,TEXTJOIN("","",TRUE,uniquelist),concat))"
    ActiveWorkbook.Names.Add Name:="UNPIVOT", RefersToR1C1:="=LAMBDA(row_headers,column_headers,[data_set],[return_row_only],LET(row_headers, row_headers, column_headers, column_headers, data, data_set, IF(ISOMITTED(return_row_only), IF(ISOMITTED(data), HSTACK(TOCOL(IFNA(row_headers, column_headers)), TOCOL(IFNA(column_headers, row_headers))), HSTACK(TOCOL(IFNA(row_headers, column_headers)), TOCOL(IFNA(column_headers, row_headers)), TOCOL(data))), HSTACK(TOCOL(IFNA(row_headers, column_headers))))))"
    ActiveWorkbook.Names.Add Name:="StringSplit", RefersToR1C1:="=LAMBDA(string, LET( cell, string, str, SUBSTITUTE(CLEAN(TRIM(cell)), ""'"", """"), XMLStr, ""<Array><Cell>"" & SUBSTITUTE(str, "","", ""</Cell><Cell>"") & ""</Cell></Array>"", SplitXML, FILTERXML(XMLStr, ""//Cell""), FormatArray, TEXT(SplitXML, ""00000""), IFERROR(FILTER(FormatArray, NOT(ISERROR(FormatArray)), """"), """") ))"

    ActiveWorkbook.Names.Add Name:="COUNTAU", RefersToR1C1:="=LAMBDA(list1,list2, ""''"" & TEXTJOIN("","", TRUE, LET(List1, list2, List2, list2, UNIQUE(FILTER(TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(List1, ""'"", """"), """""""", """"), , "",""), ISNUMBER(XMATCH(TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(List1, ""'"", """"), """""""", """"), , "",""), TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(List2, ""'"", """"), """""""", """"), , "",""))), """")))))"
    ActiveWorkbook.Names.Add Name:="COMPAREPCSTRING", RefersToR1C1:="=LAMBDA(pc_str,in_pc_str,[reverse], LET( left, IF(reverse = 1, in_pc_str, pc_str), right, IF(reverse = 1, pc_str, in_pc_str), calculation, TEXTJOIN( "","", TRUE, FILTER( TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(left, ""'"", """"), """""""", """"), , "",""), ISNA( MATCH( TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(left, ""'"", """"), """""""", """"), , "",""), TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(right, ""'"", """"), """""""", """"), , "",""), 0 ) ), """" ) ), result, IFERROR(TRIM(calculation), """"), IF(LEN(result) > 0, ""''"" & result, """") ) )"
    ActiveWorkbook.Names.Add Name:="COMBINEPCLIST", RefersToR1C1:="=LAMBDA(pc_str_array, LET(array, pc_str_array, ""''"" & TEXTJOIN("","", FALSE, SORT(UNIQUE(TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(TEXTJOIN("","", TRUE, array), ""'"", """"), """""""", """"), , "","",FALSE))))))"
    ActiveWorkbook.Names.Add Name:="FILTERISNUM", RefersToR1C1:="=LAMBDA(List1,List2,UNIQUE(FILTER(List1,ISNUMBER(XMATCH(List1,List2)),"""")))"
    ActiveWorkbook.Names.Add Name:="FILTERNA", RefersToR1C1:="= LAMBDA(List1,List2, LET( List1, List1, List2, List2, UNIQUE(FILTER(List1, ISNA(XMATCH(List1, List2)), """")) ))"
    ActiveWorkbook.Names.Add Name:="REMOVEPC", RefersToR1C1:="=LAMBDA( remove_pc_str, from_pc_str,[reverse], LET( list1, from_pc_str, list2, remove_pc_str, reverse, 0, left, IF(reverse = 1, list2, list1), right, IF(reverse = 1, list1, list2), TEXTJOIN( "","", TRUE, FILTER( TEXTSPLIT(SUBSTITUTE(SUBSTITUTE(left, ""'"", """"), """""""", """"), , "",""), ISNA( MATCH( TEXTSPLIT( SUBSTITUTE(SUBSTITUTE(left, ""'"", """"), """""""", """"), , "","" ), TEXTSPLIT( SUBSTITUTE(SUBSTITUTE(right, ""'"", """"), """""""", """"), , "","" ), 0 ) ), """" ) ) ))"
    ActiveWorkbook.Names.Add Name:="FLOOKUP", RefersToR1C1:="=LAMBDA(lookup_value,lookup_array,return_array, LET(out, OFFSET(lookup_array, MATCH(lookup_value, OFFSET(lookup_array, 0, 0, ROWS(lookup_array), 1), 0) - 1, COLUMN(INDEX(return_array, 1, 1)) - COLUMN(INDEX(lookup_array, 1, 1)), 1, 1), out))"
    ActiveWorkbook.Names.Add Name:="DYNAMICRNG", RefersToR1C1:="=LAMBDA(firstCell,rowRange,[numberOfColumns], LET(rowOffset, ROW(firstCell), collOffset, IF(ISOMITTED(numberOfColumns), 0, numberOfColumns), INDEX(rowRange, rowOffset):OFFSET(INDEX(rowRange, MAX(ROW(rowRange) * (rowRange <> """"))), , collOffset)))"
End Sub

Sub ToggleGenerateGetPivotData()
 With Application
 .GenerateGetPivotData = Not .GenerateGetPivotData
 End With
End Sub


Sub FreezePanesAtSelection()

    Dim ws As Worksheet
    Application.ScreenUpdating = False

    'Loop through each selected worksheet
    For Each ws In ActiveWindow.SelectedSheets

        'Perform action.  E.g. hide selected worksheets
        ws.Activate
        On Error Resume Next
            Application.ActiveWindow.FreezePanes = False
        On Error GoTo 0

        With Application.ActiveWindow
            .FreezePanes = True
        End With

    Next ws


    Application.ScreenUpdating = True
End Sub

Sub UnFreezeWorksheet()

    Dim ws As Worksheet
    Application.ScreenUpdating = False

    'Loop through each selected worksheet
    For Each ws In ActiveWindow.SelectedSheets

        'Perform action.  E.g. hide selected worksheets
        ws.Activate
        On Error Resume Next
            Application.ActiveWindow.FreezePanes = False
        On Error GoTo 0
    Next ws

    Application.ScreenUpdating = True
End Sub

Sub FilterSelectedRange()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    For Each ws In ActiveWindow.SelectedSheets
        ws.Activate
        
        If ws.AutoFilterMode = True Then
        Range(ws.AutoFilter.Range.Address).AutoFilter
        End If
        If Selection.count > 1 Then
        Range(Selection.Address).AutoFilter
        Else
        Range("A11", ActiveCell.SpecialCells(xlLastCell)).AutoFilter
        End If
        
        Range("A1").Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    Next ws
    ActiveWindow.SelectedSheets.Item(1).Activate
End Sub


Sub Scroller()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ictr As Long
    Dim i As Long
    Dim ioffset As Long
    
    i = ActiveWindow.SelectedSheets.count
    ioffset = ActiveWindow.SelectedSheets(1).index
       
    For Each ws In ActiveWindow.SelectedSheets
      ws.Activate
      ActiveWindow.ScrollColumn = 1
      ActiveWindow.ScrollRow = 1
            
    If ws.index = ioffset + i - 1 Then Sheets(ActiveWindow.SelectedSheets(1).index).Activate
    Next ws
End Sub
 
Sub OpenSelectedWorkbook()
    On Error Resume Next
    
    Workbooks.Open (ActiveCell.Value2)
    If Err Then MsgBox ("No File")
    Exit Sub
    On Error GoTo 0
    
End Sub


Private Function GetTabColor(sheetname As String, OtherWB As String, Optional ClearMissing As Boolean) As Variant
    If IsWorkSheetExists2(sheetname, OtherWB) Then
        GetTabColor = Workbooks(OtherWB).Sheets(sheetname).Tab.Color
    ElseIf IsMissing(ClearMissing) = False Then
        If ClearMissing = True Then
            GetTabColor = False
        Else
            GetTabColor = Sheets(sheetname).Tab.Color
        End If
    Else
        GetTabColor = Sheets(sheetname).Tab.Color

    End If

End Function

Function IsWorkSheetExists2(sheetname As String, Optional workbookName As String) As Boolean
If IsMissing(workbookName) Then
        On Error Resume Next
        IsWorkSheetExists2 = Sheets(sheetname).Name <> ""
        On Error GoTo 0
        Else
            On Error Resume Next
            IsWorkbook = Workbooks(workbookName).Name <> ""
            On Error GoTo 0
                If IsWorkbook Then
                    On Error Resume Next
                    IsWorkSheetExists2 = Workbooks(workbookName).Sheets(sheetname).Name <> ""
                    On Error GoTo 0
                    End If
                End If
End Function

