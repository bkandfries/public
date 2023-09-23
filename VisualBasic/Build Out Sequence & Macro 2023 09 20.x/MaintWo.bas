Attribute VB_Name = "MaintWo"
Option Explicit
Dim srcWB As Workbook
Dim thisWB As Workbook
Sub MaintWo()
Set thisWB = ThisWorkbook
Dim trgtws As Worksheet
Dim trgtWB As Workbook
Dim rngHeaders As Range
Dim rngTabs As Range
Dim headerArray As Variant
Dim hLow As Long, hHigh As Long, iRow As Long, iCol As Long
Dim iLoc As String
Dim ipcstring As String
Dim ws As Worksheet
Set rngHeaders = ThisWorkbook.Sheets("SEQ Header").usedRange
Set rngHeaders = rngHeaders.offset(2, 0).Resize(rngHeaders.Rows.Count - 1, rngHeaders.Columns.Count)
Set rngTabs = ThisWorkbook.Sheets("SEQ TAB").usedRange
Application.Calculate
Application.Calculation = xlCalculationManual
Application.Calculate
'Get Headers for Loop
headerArray = ThisWorkbook.Worksheets("SEQ Header").Range("A1").CurrentRegion.offset(2)
For iRow = 1 To UBound(headerArray) - 1
If headerArray(iRow, 8) = "" Then
MsgBox "BlankRow, Exiting"
GoTo ExitMe:
End If
iLoc = headerArray(iRow, 1)
ipcstring = headerArray(iRow, 5)
If trgtWB Is Nothing Then
Set trgtWB = Workbooks.Open(headerArray(iRow, 8))
End If
'Set paremeters Page
With trgtWB.Worksheets("Parameters")
.Range("B33").Value2 = iLoc
.Range("B34").Value2 = ipcstring
End With
Debug.Print "Creating " & iLoc
'Call Other SubRoutines
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With
'Copy over relevent Taborder
Call CopyOverTabOrderTwo(ThisWorkbook.Worksheets("SEQ TAB"), trgtWB.Worksheets("TabOrder"), iLoc, 2)
'Delete where location not match from Evetns, Labor, and Salaried
Call DeleteSequence(trgtWB, iLoc)
trgtWB.Worksheets("TabOrder").Activate
trgtWB.Worksheets("TabOrder").Range("A1").Select
'Build Tabs out
BuildTab3
'Delete Summary Tab
Application.DisplayAlerts = False
trgtWB.Worksheets("Summary").Delete
Application.DisplayAlerts = True
'Set Location Sheet
Call SetLocationTabFormula(trgtWB)
'Calculate & Break Links
Application.Calculate
'Break Links
Dim link As Variant
For Each link In trgtWB.LinkSources       'Loop through all the links in the workbook
trgtWB.BreakLink link, xlLinkTypeExcelLinks 'Break the link
Next link
trgtWB.Worksheets("Parameters").Move After:=trgtWB.Worksheets(trgtWB.Worksheets.Count)
trgtWB.Close savechanges:=True
Set trgtWB = Nothing
Next iRow
GoTo ExitMe:
ExitMe:
With Application
.DisplayAlerts = True
.ScreenUpdating = True
.StatusBar = False
End With
Exit Sub
End Sub
Private Sub SetLocationTabFormula(trgtWB As Workbook)
Dim firstSheet$, lastSheet$
Dim lastCellCnt As Long
Dim ws As Worksheet
Dim locWS As Worksheet
Dim i As Long
Dim formulaStr As String
Dim ClosedSheetName As String
Set ws = trgtWB.Worksheets("TabOrder")
'Get First and last sheet and create base sum formula
ClosedSheetName = "None"
firstSheet = ws.Cells(3, 1).Value
lastSheet = firstSheet
For i = 2 To RDB_Last(1, trgtWB.Worksheets("TabOrder").Range("A:A"))
If Len(ws.Cells(i, 1)) > 0 Then
lastCellCnt = i
'Set LastSheet
lastSheet = ws.Cells(i, 1).Value
End If
If InStr(1, LCase$(ws.Cells(i, 1).Value), "closed") > 0 Then
ClosedSheetName = ws.Cells(i, 1).Value
End If
Next i
If lastCellCnt = 3 Then
formulaStr = "=SUM('" & firstSheet & "'!RC)"
ElseIf lastCellCnt > 3 Then
formulaStr = "=SUM('" & firstSheet & ":" & lastSheet & "'!RC)"
Else
formulaStr = "NA"
End If
'Update formula on Location sheet
Set locWS = trgtWB.Worksheets(ws.Cells(2, 1).Value)
Dim rng As Range
Set rng = locWS.Range("AS15:BD372")
locWS.Activate
ThisWorkbook.Worksheets("Static").Range("C15:N372").Copy
locWS.Range("AS15").PasteSpecial (xlPasteAll)
Application.CutCopyMode = False
Set rng = locWS.Range("AS15:BD372")
With rng.SpecialCells(xlCellTypeConstants, 23)
.Interior.Color = 16774131
.FormulaR1C1 = formulaStr
End With
With locWS.Tab
.ThemeColor = xlThemeColorAccent2
.TintAndShade = -0.249977111117893
End With
locWS.Range("A1").Activate
If ClosedSheetName <> "None" Then
trgtWB.Worksheets(ClosedSheetName).Tab.Color = 14277081
End If
locWS.Range("BJ15:BJ372").Clear
trgtWB.Worksheets("TabOrder").Move After:=trgtWB.Worksheets(trgtWB.Worksheets.Count)
trgtWB.Worksheets("TabOrder").Visible = xlHidden
trgtWB.Worksheets("Parameters").Move After:=Worksheets(lastSheet)
trgtWB.Worksheets(ws.Cells(2, 1).Value).Activate
End Sub
Private Sub CopyOverTabOrderTwo(SrcEventsWS As Worksheet, DestEventsWS As Worksheet, iLocationFilter As String, Optional rowOffset As Long = 2)
Dim fcol As Integer
Dim rng As Range
Dim srcRange As Range
Dim i As Long, row As Long
SrcEventsWS.AutoFilterMode = False
If DestEventsWS.Name = "TabOrder" Then
fcol = WorksheetFunction.Match("Location", SrcEventsWS.Rows(1), 0)
rowOffset = 2
Else
fcol = WorksheetFunction.Match("Loc", SrcEventsWS.Range("A6").CurrentRegion.Rows(1), 0)
rowOffset = 6
End If
Set rng = SrcEventsWS.Range("A" & rowOffset).CurrentRegion.offset(1)
DestEventsWS.Range("A" & rowOffset).CurrentRegion.offset(1).Clear
Application.ScreenUpdating = False
Application.EnableAnimations = False
Application.DisplayAlerts = False
For i = 2 To rng.Rows.Count
If rng.Cells(i, fcol).Value = iLocationFilter Then
rng.Rows(i).Copy
DestEventsWS.Range("A" & rowOffset).PasteSpecial xlPasteAll
rowOffset = rowOffset + 1
End If
Next i
ActiveWindow.ScrollRow = 1
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Private Sub buildTabSequence(trgtWB As Workbook)
trgtWB.Worksheets("TabOrder").Activate
trgtWB.Worksheets("TabOrder").Range("A1").Select
GetBuildOrderCollection
Application.Calculate
Dim link As Variant
For Each link In trgtWB.LinkSources       'Loop through all the links in the workbook
trgtWB.BreakLink link, xlLinkTypeExcelLinks 'Break the link
Next link
Application.DisplayAlerts = False
trgtWB.Worksheets("Summary").Delete
Application.DisplayAlerts = True
End Sub
Private Sub DeleteSequence(trgtWB As Workbook, iLocationFilter As String)
Dim ws As Worksheet
On Error Resume Next
'Events
Set ws = trgtWB.Worksheets("Events")
ws.Activate
ws.Range("A6").CurrentRegion.AutoFilter Field:=3, Criteria1:="<>" & iLocationFilter, Operator:=xlFilterValues
ws.Range("A6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.EntireRow.Delete
ActiveSheet.ShowAllData
ActiveSheet.Range("A6").Select
'Hourly Labor
Set ws = trgtWB.Worksheets("Hourly Labor")
ws.Activate
ws.Range("A6").CurrentRegion.AutoFilter Field:=3, Criteria1:="<>" & iLocationFilter, Operator:=xlFilterValues
ws.Range("A6").Select
ws.Range("A6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.EntireRow.Delete
ActiveSheet.ShowAllData
ActiveSheet.Range("A6").Select
'Salaried Labor
Set ws = trgtWB.Worksheets("Salaried Labor")
ws.Activate
ws.Range("A6").CurrentRegion.AutoFilter Field:=3, Criteria1:="<>" & iLocationFilter, Operator:=xlFilterValues
ws.Range("A6").Select
ws.Range("A6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.EntireRow.Delete
ActiveSheet.ShowAllData
ActiveSheet.Range("A6").Select
On Error GoTo 0
End Sub
