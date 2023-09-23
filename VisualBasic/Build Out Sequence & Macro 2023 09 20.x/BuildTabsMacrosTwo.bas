Attribute VB_Name = "BuildTabsMacrosTwo"
Option Explicit
Option Base 1
Option Compare Text
Public Enum fasdfsdf
cSheetName = 1
cParent = 2
CDescription = 4
cChildCOunt = 8
cLocation = 14
cSheetArray = 17
cFormula = 18
cPCString = 19
End Enum
Private Function InColumnA(nameToCheck As String, rngCol As Range) As Boolean
Dim c As Range
For Each c In rngCol.Rows
If c.Value = nameToCheck Then
InColumnA = True
Exit For
End If
Next c
End Function
Private Function ExistsInCollection(searchItem As Variant, col As Collection) As Boolean
Dim item As Variant
For Each item In col
If item = searchItem Then
ExistsInCollection = True
Exit For
End If
Next item
End Function
Private Function Exists(key As String, coll As Collection) As Boolean
On Error Resume Next
IsObject (coll.item(key))
Exists = True
On Error GoTo 0
End Function
Private Sub Main()
'  GetBuildOrderCollection
'   Call SUMonParentsSheets
'    Call UpdateCreteria
Application.Calculate
Call killLinks
'Call TurnOnToEnd
End Sub
Sub GetBuildOrderCollection()
Dim wb As Workbook
Set wb = ActiveWorkbook
Dim ws As Worksheet
Set ws = wb.Worksheets("TabOrder")
ws.Activate
Dim usedRange As Range
Set usedRange = LastRange(ws, 1)
Dim col As New Collection
Dim rowCount As Long, colCnt As Long
Dim cell As Range
For Each cell In usedRange.Columns(2).Rows
If cell.Value <> "" Then
If InColumnA(cell.Value, usedRange.Columns(1)) And Not Exists(cell.Value, col) Then
colCnt = colCnt + 1
col.Add cell.Value, cell.Value
End If
End If
Next cell
Dim i As Long, rCnt As Long, sheetName$, parentSheetName$
For i = usedRange.Rows.Count To 1 Step -1
'Check to see if parent exits, if so add after
If usedRange.Cells(i, 1).Value <> "" Then
sheetName = usedRange.Cells(i, 1).Value
parentSheetName = usedRange.Cells(i, 2).Value
If parentSheetName <> "" And ExistsInCollection(parentSheetName, col) And Not ExistsInCollection(sheetName, col) Then
col.Add sheetName, sheetName, After:=parentSheetName
Else
If Not ExistsInCollection(sheetName, col) Then
col.Add item:=sheetName, key:=sheetName
End If
End If
End If
Next i
Call builtTabsNewSheet(col)
End Sub
Function buillarray() As Variant
Dim wb As Workbook
Set wb = ActiveWorkbook
Dim ws As Worksheet
Set ws = wb.Worksheets("TabOrder")
ws.Activate
Dim usedRange As Range
Set usedRange = LastRange(ws, 1)
Dim col As New Collection
Dim rowCount As Long, colCnt As Long
Dim cell As Range
Dim arr As Variant
arr = usedRange
Dim DestArr As Variant
Dim iStart As Long, iEnd As Long
iStart = LBound(arr)
iEnd = UBound(arr)
ReDim DestArr(LBound(arr) To UBound(arr), 1 To 6)
Dim i As Long, j As Long
For i = iStart To iEnd
With usedRange.Rows(i)
DestArr(i, 1) = .Cells(i, cSheetName)
DestArr(i, 2) = .Cells(i, cParent)
DestArr(i, 3) = .Cells(i, CDescription)
DestArr(i, 4) = .Cells(i, cChildCOunt)
DestArr(i, 5) = .Cells(i, cLocation)
DestArr(i, 6) = .Cells(i, cPCString)
End With
Next i
buillarray = DestArr
End Function
Private Sub builtTabsNewSheet(tabs As Collection)
Dim rng As Range
Dim c As Range
Dim item As Variant
Dim forLoopWS As Worksheet, forLoopRange
Dim builtTabsArray() As String, ictr As Long
Dim answer As Integer
If RunThroughChecks = False Then
MsgBox "Failed Checks"
Exit Sub
End If
'Move sheets around
Dim wb As Workbook
Set wb = ActiveWorkbook
Dim summarySheet As Worksheet, tabOrderWorksheet As Worksheet
Set summarySheet = wb.Worksheets("Summary")
Set tabOrderWorksheet = wb.Worksheets("TabOrder")
summarySheet.Move Before:=Sheets(1)                     'Move to front
Worksheets("Parameters").Move Before:=Worksheets(Sheets.Count)  'Move to end
tabOrderWorksheet.Move Before:=Worksheets("Parameters") 'TabOrder must be last(ish)
ictr = 1
Dim lastCreatedWS$, trgtws As Worksheet
Dim smryRange As Range
Set smryRange = LastRange(summarySheet)
lastCreatedWS = Worksheets(tabOrderWorksheet.Index).Name
Dim arr As Variant
arr = buillarray
Dim iStrt As Long, iEnd As Long
Dim i As Long
Call TurnOffToBegin
For Each item In tabs
If Not IsWorkSheetExists(item) Then
ReDim Preserve builtTabsArray(1 To ictr)
wb.Worksheets.Add(Before:=Worksheets(lastCreatedWS), Type:=xlWBATWorksheet).Name = item
lastCreatedWS = item
builtTabsArray(ictr) = item
Debug.Print (Searcharray(arr, item, 5) & "  " & Searcharray(arr, item, 2))
ictr = ictr + 1
End If
Next item
Debug.Print UBound(builtTabsArray)
For item = LBound(builtTabsArray) To UBound(builtTabsArray)
With Worksheets(builtTabsArray(item))
.Activate
.Range("C12").Activate
summarySheet.Cells.Copy Destination:=.Cells
.Range(summarySheet.AutoFilter.Range.Address).AutoFilter
.Range("C12").Select
.Outline.ShowLevels RowLevels:=2
.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
With ActiveWindow
.FreezePanes = True
.Zoom = 85
.ScrollColumn = 1
.ScrollRow = 1
.DisplayOutline = False
End With
.Range("A1").Select
If Searcharray(arr, item, 5) <> "" Then
.Range("AS376").Value = Searcharray(arr, item, 5)
End If
End With
Next item 'end loop
Worksheets("Parameters").Range("B34").Value = arr(1, 6)
Call TurnOnToEnd
UpdateCreteria
On Error Resume Next
Worksheets("Total Net Income Summary").Range("C1").Value = 352
Worksheets("Adjusted EBITDA Summary").Range("C1").Value = 372
On Error GoTo 0
answer = MsgBox("Delete and continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Missing Summary")
Application.DisplayAlerts = False
If answer = vbYes Then summarySheet.Delete
Application.DisplayAlerts = True
End Sub
Private Function Searcharray(ByRef arr As Variant, ByVal searchStr As String, Optional offset As Long = 1) As String
Dim i As Long, j As Long
If IsArray(arr) Then
For i = LBound(arr, 1) To UBound(arr, 1)
If arr(i, 1) = searchStr Then
Searcharray = arr(i, offset)
End If
Next i
End If
End Function
Private Function RunThroughChecks() As Boolean
Dim answer As Variant
Select Case True
Case Not IsWorkSheetExists("TabOrder")
MsgBox "Missing taborder, exiting.."
Exit Function
Case Not IsWorkSheetExists("Parameters")
MsgBox "Missing Parementers, exiting.."
Exit Function
Case Not IsWorkSheetExists("Summary")
MsgBox "Missing Summary, exiting.."
Exit Function
Case Else
If IsWorkSheetExists("__Summary") Then
answer = MsgBox("Delete and continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Missing Summary")
If answer = vbYes Then
'Application.DisplayAlerts = False
Worksheets("__SUMMARY").Delete
'Application.DisplayAlerts = True
Else
Exit Function
End If
End If
End Select
RunThroughChecks = True
End Function
Private Sub delWS(wsName)
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Worksheets(wsName).Delete
On Error GoTo 0
Application.DisplayAlerts = True
End Sub
Sub killLinks()
Dim link As Variant
For Each link In ActiveWorkbook.LinkSources       'Loop through all the links in the workbook
ActiveWorkbook.BreakLink link, xlLinkTypeExcelLinks 'Break the link
Next link
End Sub
Sub FilterSelectionFunc(Optional filterRangeStr As String)
Dim ws As Worksheet
Set ws = ActiveSheet
For Each ws In ActiveWindow.SelectedSheets
ws.Activate
If ws.FilterMode = True Then ws.AutoFilterMode = False
If ws.AutoFilterMode = True Then ws.AutoFilter = False
On Error Resume Next
If Not IsEmpty(filterRangeStr) And filterRangeStr <> "" Then
Range(filterRangeStr).Cells(1, 1).AutoFilter
Range(filterRangeStr).Cells(1, 1).Select
Exit Sub
End If
If Selection.Rows.Count > 1 And Selection.Rows.Count < 1048570 Then
Range(Selection.Address).AutoFilter
Range(Selection.Address).Cells(1, 1).Select
Exit Sub
Else
Range("A11", ActiveCell.SpecialCells(xlLastCell).Resize(372, ActiveCell.SpecialCells(xlLastCell).Columns.Count)).AutoFilter
End If
On Error GoTo 0
Range("A1").Select
ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1
Next ws
ActiveWindow.SelectedSheets.item(1).Activate
End Sub
