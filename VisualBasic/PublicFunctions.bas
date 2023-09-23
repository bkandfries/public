Attribute VB_Name = "PublicFunctions"
 Public calcState, eventsState As Variant

Function FindStringInCellRange(Range_Obj_or_Str As Variant, search_string As Variant, Optional offset As Integer = 0) As Variant

    'FindStringInCellRange("$B$16:$B$20", "GSE", 7)

     
    'This UDF:
        '(1) Accepts 2 arguments: MyRange and MyString
        '(2) Finds a string passed as argument (MyString) in a cell range passed as argument (MyRange). The search is case-insensitive
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the cell range (MyRange) where the string (MyString) is found
    
    
    Dim myRangeObj As Range
    If TypeName(Range_Obj_or_Str) = "String" Then Set myRangeObj = Range(Range_Obj_or_Str) Else Set myRangeObj = Range_Obj_or_Str
    '    Debug.Print (myRangeObj.Address)
    
    With myRangeObj
                
        Debug.Print
        
        
    End With
     
     
     
End Function



 Function lastRow(sh As Worksheet)
    On Error Resume Next
    lastRow = sh.Cells.Find(what:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            searchdirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

 Function lastCol(sh As Worksheet)
    On Error Resume Next
    lastCol = sh.Cells.Find(what:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            searchdirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function


Function LastRange(Optional Worksheet As Variant, Optional SkipFirstRows As Integer = 0) As Range
        Dim ws As Worksheet
        If IsMissing(Worksheet) Then
            Set ws = ActiveSheet
        ElseIf VarType(Worksheet) = vbString Then
            Set ws = Worksheets(Worksheet)
            
        
        Else
            If IsObject(Worksheet) Then
                
                Set ws = Worksheet
            Else
              Set ws = ActiveWorkbook.Worksheets(Worksheet)
            End If
       End If
        Dim rng As Range
       
       
        With ws
            Set rng = .Range(.Cells(1, 1), .Cells(lastRow(ws), lastCol(ws)))
        End With
               
        If SkipFirstRows > 0 Then
        Set LastRange = rng.offset(SkipFirstRows).Resize(rng.Rows.count - SkipFirstRows, rng.Columns.count)
        Else
        Set LastRange = rng
        End If
        
End Function


Function FindStringAddress(Range_Obj_or_Str As Variant, search_string As Variant, Optional offset As Integer) As Variant

    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-find/
     
    'This UDF:
        '(1) Accepts 2 arguments: MyRange and MyString
        '(2) Finds a string passed as argument (MyString) in a cell range passed as argument (MyRange). The search is case-insensitive
        '(3) Returns the address (as an A1-style relative reference) of the first cell in the cell range (MyRange) where the string (MyString) is found
    
    
    Dim myRangeObj As Range
    If TypeName(Range_Obj_or_Str) = "String" Then Set myRangeObj = Range(Range_Obj_or_Str) Else Set myRangeObj = Range_Obj_or_Str
     Debug.Print (myRangeObj.Address)
    
    With myRangeObj
        FindStringAddress = .Find(what:=search_string, After:=.Cells(.Cells.count), LookIn:=xlValues, Lookat:=xlWhole, SearchOrder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).offset(0, offset).Address
   
        
        
    End With
End Function

Function RDB_Last(choice As Integer, rng As Range)
' By Ron de Bruin, 5 May 2008
' A choice of 1 = last row.
' A choice of 2 = last column.
' A choice of 3 = last cell.
   Dim lrw As Long
   Dim lCol As Integer

   Select Case choice

   Case 1:
      On Error Resume Next
      RDB_Last = rng.Find(what:="*", _
                          After:=rng.Cells(1), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          searchdirection:=xlPrevious, _
                          MatchCase:=False).Row
      On Error GoTo 0

   Case 2:
      On Error Resume Next
      RDB_Last = rng.Find(what:="*", _
                          After:=rng.Cells(1), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, _
                          searchdirection:=xlPrevious, _
                          MatchCase:=False).Column
      On Error GoTo 0

   Case 3:
      On Error Resume Next
      lrw = rng.Find(what:="*", _
                    After:=rng.Cells(1), _
                    Lookat:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    searchdirection:=xlPrevious, _
                    MatchCase:=False).Row
      On Error GoTo 0

      On Error Resume Next
      lCol = rng.Find(what:="*", _
                     After:=rng.Cells(1), _
                     Lookat:=xlPart, _
                     LookIn:=xlFormulas, _
                     SearchOrder:=xlByColumns, _
                     searchdirection:=xlPrevious, _
                     MatchCase:=False).Column
      On Error GoTo 0

      On Error Resume Next
      RDB_Last = rng.Parent.Cells(lrw, lCol).Address(False, False)
      If Err.Number > 0 Then
         RDB_Last = rng.Cells(1).Address(False, False)
         Err.Clear
      End If
      On Error GoTo 0

   End Select
End Function


Function Col_Letter(lngCol As Long) As String
    If lngCol = 0 Then lngCol = 1
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function



Public Function GetWorksheetObject(ByVal workbookObj As Variant, searchString As String) As Worksheet
'Searches wbstring and returns Workbook object
'
'DEPENDS ON:
    'Function GetWorkbookObject



'HOW TO USE'
' Set wb = GetWorkbookObject("tacker")

Dim iCntr%, wsName$, i%, pos%, iName$, matchCnt%
Dim wb As Workbook
wsName = "Not Fount"


'Resolve Workbook object 1. Direct wb object pass through, 2. search using GetWorkbookObject, 3. Activeworkbook
    On Error Resume Next
       If TypeName(workbookObj) = "Workbook" Then
            Set wb = workbookObj
        Else
            Set wb = GetWorkbookObject(workbookObj)
        If wb Is Nothing Then Set wb = ActiveWorkbook
        If Err Then Exit Function
        
        End If
        
    On Error GoTo 0
   
    iCntr = 0
    matchCnt = 0
    Dim wbnameArray() As Variant

'Search through worksheets in workbook "WB"
    For i = 1 To wb.Sheets.count
       iName = wb.Sheets(i).Name
       pos = InStr(1, iName, searchString, 1)
       If iName <> ActiveSheet.Name Then
            If InStr(1, iName, searchString, 1) > 0 Then
                matchCnt = matchCnt + 1
                iCntr = i
    
                wsName = iName
            End If
        End If
    Next
' check to see if more than 1 match => Exit

    If matchCnt > 1 Then
        Debug.Print "`n Multiple Sheets matches found"
        Exit Function
'If no match => Exit Function
    ElseIf wsName = "Not Found" Then
        Debug.Print "'nNot Found'n"
        Exit Function
        
    Else
        Set GetWorksheetObject = wb.Sheets(wsName)
        Debug.Print "Found: " & wsName
    End If
    
    
    
End Function






Private Function GetWorkbookObject(wbstring As Variant) As Workbook
'Searches wbstring and returns Workbook object
'  only if one match found
'  otherwise it returns empty


'DEPENDS ON:
'   N/A
    
'HOW TO USE'
' Set wb = GetWorkbookObject("tacker")




Dim iCntr%, OtherWBName$, i%, pos%, iName$, matchCnt%
    
    OtherWBName = "Not Fount"
    iCntr = 0
    matchCnt = 0
    Dim wbnameArray() As Variant
'Search through open Workbooks
    For i = 1 To Workbooks.count
       iName = Workbooks(i).Name
       pos = InStr(1, iName, wbstring, 1)
       If iName <> ActiveWorkbook.Name Then
            If InStr(1, iName, wbstring, 1) > 0 Then
                matchCnt = matchCnt + 1
                iCntr = i
    
                OtherWBName = iName
            End If
        End If
    Next
' check to see if more than 1 match => Exit

    If matchCnt > 1 Then
        Exit Function
'If no match => Exit Function
    ElseIf OtherWBName = "Not Found" Then
        Exit Function
    Else
         Set GetWorkbookObject = Workbooks(OtherWBName)
        
    End If
End Function





Sub TurnOffToBegin()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ActiveWorkbook

  ' Turn off Excel functionality to improve performance.
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
  ' Note: this is a sheet-level setting.
    ActiveSheet.DisplayPageBreaks = False
    Application.AutoRecover.Enabled = False
    For Each ws In wb.Worksheets
        ws.EnableCalculation = False
        Debug.Assert (IsWorkSheetExists(ws.Name))
        
      If Not ws.Visible = xlSheetVeryHidden Then
        ws.DisplayPageBreaks = False
      End If
    Next ws
    
    Set ws = Nothing
    Set wb = Nothing
    
    
End Sub


Sub TurnOnToEnd()
    
  ' Restore Excel settings to original state.
  Application.ScreenUpdating = True
  Application.DisplayStatusBar = True
  Application.AutoRecover.Enabled = True
  
 ' If Not IsEmpty(calcState) Then Application.Calculation = calcState
  
  
End Sub



Function IsWorkSheetExists(ByVal sheetname As String, Optional inWorkBook As Workbook) As Boolean
     On Error Resume Next
    Dim wb As Workbook
    If IsMissing(inWorkBook) Or IsEmpty(inWorkBook) Or (inWorkBook Is Nothing) Then
        Set wb = Workbooks(ActiveWorkbook.Name)
        Else
        Set wb = inWorkBook
    End If

  
  IsWorkSheetExists = wb.Worksheets(sheetname).Name <> ""
  On Error GoTo 0
End Function
  


