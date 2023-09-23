Attribute VB_Name = "Main"
Option Explicit
Dim srcWB As Workbook
Dim thisWB As Workbook
Sub Main()
Set thisWB = ThisWorkbook
Dim trgtws As Worksheet
Dim trgtWB As Workbook
Dim rngHeaders As Range
Dim rngTabs As Range
Dim headerArray As Variant
Dim hLow As Long, hHigh As Long, iRow As Long, iCol As Long
Dim iLoc As String
Dim ipcstring As String
Set rngHeaders = ThisWorkbook.Sheets("SEQ Header").usedRange
Set rngHeaders = rngHeaders.offset(2, 0).Resize(rngHeaders.Rows.Count - 1, rngHeaders.Columns.Count)
Set rngTabs = ThisWorkbook.Sheets("SEQ TAB").usedRange
If srcWB Is Nothing Then
Set srcWB = Workbooks.Open("R:\BUD\wip\base48MOTemplateFile.xlsm", False, True)
End If
Dim brakeI As Long
headerArray = ThisWorkbook.Worksheets("SEQ Header").Range("A1").CurrentRegion.offset(2)
For iRow = 1 To UBound(headerArray)
If iRow = RDB_Last(1, ThisWorkbook.Worksheets("SEQ Header").Range("A1").CurrentRegion) Then
Debug.Print "Exiting SUb"
Exit Sub
End If
If iRow = 15 Then
Debug.Print "Override failed"
Exit Sub
End If
iLoc = headerArray(iRow, 1)
ipcstring = headerArray(iRow, 5)
If trgtWB Is Nothing Then
Set trgtWB = Workbooks.Open(headerArray(iRow, 8))
End If
With trgtWB.Worksheets("Parameters")
.Range("B33").Value2 = iLoc
.Range("B34").Value2 = ipcstring
End With
Debug.Print "Creating " & iLoc
Call CopyOverTabOrder(ThisWorkbook.Worksheets("SEQ TAB"), trgtWB.Worksheets("TabOrder"), iLoc, 2)
Call CopyOverEvents(srcWB.Worksheets("Events"), trgtWB.Worksheets("Events"), iLoc, 6)
Call CopyOverEvents(srcWB.Worksheets("Hourly Labor"), trgtWB.Worksheets("Hourly Labor"), iLoc, 6)
Call CopyOverEvents(srcWB.Worksheets("Salaried Labor"), trgtWB.Worksheets("Salaried Labor"), iLoc, 6)
If iRow = 5 Then
Debug.Print iLoc
End If
trgtWB.Worksheets("TabOrder").Activate
trgtWB.Worksheets("TabOrder").Range("A1").Select
BuildTab3
Dim ws As Worksheet
For Each ws In trgtWB.Worksheets
ws.Range("A1:BE370").Calculate
Next ws
Dim link As Variant
For Each link In trgtWB.LinkSources       'Loop through all the links in the workbook
trgtWB.BreakLink link, xlLinkTypeExcelLinks 'Break the link
Next link
Application.DisplayAlerts = False
trgtWB.Worksheets("Summary").Delete
Application.DisplayAlerts = True
trgtWB.Close savechanges:=True
Set trgtWB = Nothing
Next iRow
End Sub
Sub CopyOverTabOrder(SrcEventsWS As Worksheet, DestEventsWS As Worksheet, iLocationFilter As String, Optional rowOffset As Long = 2)
Dim fcol As Integer
Dim rng As Range
Dim srcRange As Range
Dim i As Long, row As Long
SrcEventsWS.AutoFilterMode = False
Set rng = SrcEventsWS.usedRange
fcol = WorksheetFunction.Match("Location", rng.Rows(1), 0)
DestEventsWS.usedRange.offset(1).Clear
For i = 2 To rng.Rows.Count
If rng.Cells(i, fcol).Value = iLocationFilter Then
rng.Rows(i).Copy
DestEventsWS.Range("A" & rowOffset).PasteSpecial xlPasteAll
rowOffset = rowOffset + 1
End If
Next i
End Sub
Sub CopyOverTabs(TabOrderWS As Worksheet, iLocationFilter As String)
Dim fcol As Integer
Dim rng As Range
Set rng = thisWB.Worksheets("SEQ TAB").Range("A6").CurrentRegion
fcol = WorksheetFunction.Match("Location", rng.Range("1:1"), 0)
'    rng.AutoFilter field:=fCol, Criteria1:=iLocationFilter
'    rng.SpecialCells(xlCellTypeVisible).Copy Destination:=TabOrderWS.Range("A1")
End Sub
Sub CopyOverEvents(SrcEventsWS As Worksheet, DestEventsWS As Worksheet, iLocationFilter As String, Optional rowOffset As Long = 2)
Dim fcol As Integer
Dim rng As Range, srcRange As Range, destRange As Range
Dim i As Long, row As Long, colCnt As Long, j As Long
Dim addr As String
Dim formulaColumns() As Integer
Dim forumlaColumnsCount As Integer
SrcEventsWS.AutoFilterMode = False
Set rng = SrcEventsWS.Range("A6").CurrentRegion.offset(1)
fcol = WorksheetFunction.Match("Loc", DestEventsWS.Range("A6").CurrentRegion.Rows(1), 0)
row = rowOffset
colCnt = rng.Columns.Count
forumlaColumnsCount = 0
For i = 1 To rng.Columns.Count
If rng.Rows(1).Columns(i).HasFormula Then
forumlaColumnsCount = forumlaColumnsCount + 1
ReDim Preserve formulaColumns(forumlaColumnsCount)
formulaColumns(forumlaColumnsCount) = i
End If
Next i
i = 0
For i = 1 To rng.Rows.Count
If rng.Cells(i, fcol).Value = iLocationFilter Then
DestEventsWS.Range("A" & rowOffset).Resize(1, rng.Columns.Count).Formula2 = rng.Rows(i).Formula2
rowOffset = rowOffset + 1
End If
Next i
Debug.Print DestEventsWS.Range("A" & row - 1).CurrentRegion.offset.Address
End Sub
