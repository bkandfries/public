Attribute VB_Name = "OtherSubs"
Option Explicit
Option Base 1
Option Compare Text
Sub SUMonParentsSheets()
Dim wb As Workbook
Dim ws As Worksheet
Dim trgtws As Worksheet, tgrtRng As Range, trgtBigRange As Range
Dim sheetName$, formulaString$
Dim rng As Range, row As Range, rngData As Range
Dim i As Long, j As Long, breakFlag As Long
Dim formulacol&, childrencol&
Dim AdjEBITDARow As Integer
Dim childSheetNames() As String, childStr As String, iChild As Variant
Dim addressString$, addrStrArray() As String
Set wb = ActiveWorkbook
Set ws = wb.Worksheets("Taborder")
Set rng = LastRange(ws)
addressString = ThisWorkbook.Worksheets("StaticRaw").usedRange.SpecialCells(xlCellTypeConstants).Address
For i = 2 To rng.Rows.Count
'Debug.Print rng.Cells(i, 1) & ":" & rng.Cells(i, 18)
If rng.Cells(i, 8) > 0 And rng.Cells(i, 18) <> "" And IsWorkSheetExists(rng.Cells(i, 1).Value) Then
childStr = rng.Cells(i, 17).Text
childStr = Trim(Replace(childStr, "'", ""))
childSheetNames = Split(childStr, ",")
breakFlag = 0
For Each iChild In childSheetNames
If Not IsWorkSheetExists(iChild) Then
'       Debug.Print "Worksheet missing: " & iChild
breakFlag = breakFlag + 1
End If
Next iChild
If breakFlag = 0 Then
sheetName = rng.Cells(i, 1)
Set trgtws = wb.Worksheets(sheetName)
Set trgtBigRange = Union(trgtws.Range(Split(addressString, ",")(0)), trgtws.Range(Split(addressString, ",")(1)))
For j = 2 To UBound(Split(addressString, ","))
Set trgtBigRange = Union(trgtBigRange, trgtws.Range(Split(addressString, ",")(j)))
Next j
trgtws.Activate
trgtBigRange.Interior.Color = 15923173
formulaString = Replace(rng.Cells(i, 18), "'=", "")
'Debug.Print (formulaString)
With trgtBigRange
.FormulaR1C1 = formulaString
.Parent.Tab.Color = RGB(102, 102, 153)
End With
trgtws.Range("BJ15:BJ372").Clear
trgtws.Range("A1").Select
End If
End If
' Debug.Assert rng.Cells(i, 1).Value <> "BWI Consol"
If IsWorkSheetExists(rng.Cells(i, 1).Value) And InStr(1, rng.Cells(i, 1).Value, "Closed") Then
With wb.Worksheets(rng.Cells(i, 1).Value).Range(addressString)
.Interior.Color = xlNone
.Value = 0
.Parent.Tab.Color = RGB(217, 217, 217)
'.Range("A1").Select
End With
Else
'Debug.Print "No Closed on:" & rng.Cells(i, 1).Value
End If
Next i
For Each ws In wb.Worksheets
ws.Range("A2").Calculate
Next ws
End Sub
Sub updateMEC()
Dim wb As Workbook, ws As Worksheet, rng As Range
Set wb = ActiveWorkbook
Set ws = wb.Worksheets("Taborder")
Set rng = LastRange(ws)
Dim i As Long, childSheetNames() As String, childStr$, breakFlag%, iChild As Variant
For i = 2 To rng.Rows.Count
'Debug.Print rng.Cells(i, 1) & ":" & rng.Cells(i, 18)
Dim trng As Range
Dim formulaRev$, formulaNI$, formulaAEBITDA$, baseFormula$
Dim trgtws As Worksheet
If rng.Cells(i, 8) > 0 And rng.Cells(i, 18) <> "" And IsWorkSheetExists(rng.Cells(i, 1).Value) Then
childStr = rng.Cells(i, 17).Text
childStr = Trim(Replace(childStr, "'", ""))
childSheetNames = Split(childStr, ",")
breakFlag = 0
For Each iChild In childSheetNames
If Not IsWorkSheetExists(iChild) Then
'       Debug.Print "Worksheet missing: " & iChild
breakFlag = breakFlag + 1
End If
'Debug.Assert breakFlag = 0
Next iChild
If breakFlag = 0 Then
Set trgtws = wb.Worksheets(rng.Cells(i, 1).Value)
trgtws.Activate
Debug.Assert rng.Cells(i, 18).Value <> ""
baseFormula = Replace(Replace(Replace(rng.Cells(i, 18), "'=", ""), "[", ""), "]", "")
formulaRev = Replace(baseFormula, "??", "52")
formulaNI = Replace(baseFormula, "??", "352")
formulaAEBITDA = Replace(baseFormula, "??", "372")
Set trng = Range("C376")
trgtws.Range("$C$375,$E$375,$G$375,$I$375,$L$375,$N$375,$Q$375").Formula2R1C1 = "=" & formulaRev
trgtws.Range("$C$376,$E$376, $G$376,$I$376,$L$376,$N$376,$Q$376").Formula2R1C1 = "=" & formulaNI
trgtws.Range("$C$377,$E$377, $G$377,$I$377,$L$377,$N$377,$Q$377").Formula2R1C1 = "=" & formulaAEBITDA
ActiveWindow.ScrollRow = 1
ActiveWindow.ScrollColumn = 1
End If
End If
Next i
End Sub
