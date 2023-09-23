Attribute VB_Name = "NewBuiilder"
Option Explicit
Private Sub buildOrderCol() 'As Collection
Dim col As New Collection
Dim ws As Worksheet
Dim rng As Range
Dim lastRow As Long, lastCol As Long
Dim arr As Variant, outArray As Variant
Set ws = Worksheets("taborder")
Set rng = LastRange(ws)
Dim i As Long, j As Long
arr = rng.offset(1).Resize(rng.Rows.Count - 1, rng.Columns.Count)
For i = LBound(arr, 1) To UBound(arr, 1)
If arr(i, 1) <> "" Then
If arr(i, 2) <> "" And existsInArray(arr, arr(i, 2)) And Not IsWorkSheetExists(arr(i, 2)) Then
If Not ExistsInCollection(arr(i, 2), col) Then col.Add arr(i, 2), arr(i, 2)
End If
End If
Next i
Dim item As Variant
For i = 1 To col.Count
Debug.Print col.item(1)
Next i
End Sub
Private Function existsInArray(ByRef arr As Variant, searchString As Variant, Optional col As Long = 1) As Boolean
Dim i As Long
On Error Resume Next
For i = LBound(arr, 1) To UBound(arr, 1)
If arr(i, col) = searchString Then
existsInArray = True
Exit For
End If
Next i
On Error GoTo 0
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
Sub testCreateMultipleSHeets(col As Collection)
Dim wb As Workbook, ws As Worksheet
Dim rng As Range, i As Long, iCntr As Long
Dim sheetsCreated() As String
Set wb = Workbooks("03MO SEI Template - With 910 & 920 Depts 2023 09 10a.xlsm")
wb.Activate
'Quick Checks for TaborOrder Page & Summary Page
If IsWorkSheetExists("TabOrder") = False Then
MsgBox ("Error: 'Summary' tab not found!")
Exit Sub
ElseIf IsWorkSheetExists("Parameters") = False Then
MsgBox ("No parameters tab")
Exit Sub
ElseIf IsWorkSheetExists("Summary") = False Then
MsgBox ("No Summary tab")
Exit Sub
End If
Dim wsSrc As Worksheet, rngSrc As Range, arrSrc As Variant
Set wsSrc = wb.Worksheets("SUMMARY")
Set rngSrc = LastRange(wsSrc)
arrSrc = rngSrc.Value2
Dim wsTab As Worksheet, rngTab As Range, arrTab As Variant
Set wsTab = wb.Worksheets("Taborder")
Set rngTab = LastRange(wsTab, 1)
Dim c As Variant
i = 1
Dim wsTrg As Worksheet
Dim forSheetNames() As String, lastCreatedSheet As String
Dim forSheetName As String
Dim arrA As Variant
i = 1
lastCreatedSheet = Worksheets(wsTab.Index - 2)
For Each arrA In col
If Not arrA = "" And Not IsWorkSheetExists(arrA) Then
ReDim Preserve forSheetNames(1 To i)
forSheetNames(i) = arrA
'DEBUG
delWS (arrA)
Set wsTrg = wb.Worksheets.Add(After:=Worksheets(lastCreatedSheet), _
Type:=xlWBATWorksheet)
lastCreatedSheet = arrA
With wsTrg
.Name = forSheetNames(i)
.Activate
.Range("C11").Activate
With ActiveWindow
.FreezePanes = True
.Zoom = 85
.ScrollColumn = 1
.ScrollRow = 1
End With
'.Range(wsSrc.UsedRange.Address).Cells.Interior = wsSrc.UsedRange.Cells.Interior
End With
i = i + 1
End If
Next arrA
forSheetNames(0) = wsSrc.Name
Dim arrc() As String
For i = 1 To UBound(forSheetNames)
ReDim Preserve arrc(i - 1)
arrc(i - 1) = forSheetNames(i)
Next i
Dim lCol As Long, lcoll As String
lCol = lastCol(wsSrc)
lcoll = Col_Letter(lCol)
wsSrc.Activate
Columns("A:" & lcoll).Copy
Worksheets(arrc).Select
Worksheets(arrc(1)).Activate
Columns("A:A").Select
ActiveSheet.Paste
For Each ws In wb.Worksheets
If ws.Visible = xlSheetVisible Then
ws.Activate
Range("a1").Select
End If
Next ws
End Sub
Private Sub cpData()
End Sub
Private Sub DeleteSheetsInTabOrder()
Dim tabOrder As Variant
Dim i As Long
Dim DisplayAlertsState As Variant
DisplayAlertsState = Application.DisplayAlerts
tabOrder = LastRange(ThisWorkbook.Worksheets("Taborder")).Columns(1)
On Error Resume Next
Dim ws As Worksheet
Application.DisplayAlerts = False
For Each ws In Worksheets
If ws.Visible <> xlSheetVeryHidden Then
For i = UBound(tabOrder, 1) To LBound(tabOrder, 1) Step -1
If ws.Name = tabOrder(i, 1) Then
Debug.Print "Deleting ws :" & ws.Name
ws.Cells.Clear
ws.Delete
End If
Next i
End If
Next ws
Application.DisplayAlerts = DisplayAlertsState
On Error GoTo 0
End Sub
Private Sub delWS(wsName)
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Worksheets(wsName).Delete
On Error GoTo 0
Application.DisplayAlerts = True
End Sub

