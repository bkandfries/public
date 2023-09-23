Attribute VB_Name = "BuildTabsMacros"
Option Explicit
Private Function GetChildSheetsOffset(CurrentRow) As Integer
Dim nRowCounter As Integer
Dim nTabOrderRows As Integer
Dim ws As Worksheet
Dim rowsheet As String, rowparent As String, forSheetName As String, Forparent As String, ssheetname As String
Set ws = Sheets("TabOrder")
rowsheet = ws.Cells(CurrentRow, 1).Value
rowparent = ws.Cells(CurrentRow, 2).Value
'Return Last Sheet Index by Default
GetChildSheetsOffset = Sheets("TabOrder").Index - 1
'If Parent Exists, then Index will be after Parent Sheet Index + any Children that already exits
If IsWorkSheetExists(rowparent) Then
GetChildSheetsOffset = Sheets(rowparent).Index + 1
For nRowCounter = 1 To (CurrentRow - 1) 'Loop through all Previous Rows in TabORder
If Len(Trim(ws.Cells(nRowCounter, 1).Value)) > 0 Then 'Skip Rows with Blank PCs
forSheetName = ws.Cells(nRowCounter, 1).Value
Forparent = ws.Cells(nRowCounter, 2).Value
If (Forparent = rowparent) And (IsWorkSheetExists(forSheetName)) And (forSheetName <> ssheetname) Then
GetChildSheetsOffset = GetChildSheetsOffset + 1
End If
End If
Next nRowCounter
End If
End Function
Sub BuildTab3()
Dim nRowCounter As Integer
Dim nTabOrderRows As Integer
Dim ws As Worksheet
Dim ChildOffset As Integer
Dim newWS As Worksheet
Set ws = Sheets("TabOrder")
nTabOrderRows = ws.usedRange.Cells.Rows.Count
Dim SUMMARY As String
SUMMARY = "SUMMARY"
Dim summaryWS As Worksheet
Dim rowPCstr$, rowlocstr$, rowsitestr$, rowbustr$, rowdescription$, cPCList$, cLOCList$, cSITEList$, cBUSUNITList$, CDescription$
'Quick Checks for TaborOrder Page & Summary Page
If IsWorkSheetExists("TabOrder") = False Then
MsgBox ("Error: 'Summary' tab not found!")
Exit Sub
ElseIf IsWorkSheetExists("Parameters") = False Then
MsgBox ("No parameters tab")
Exit Sub
ElseIf IsWorkSheetExists(SUMMARY) = False Then
MsgBox ("No Summary tab")
Exit Sub
End If
Set summaryWS = Sheets(SUMMARY)
'Move TabOrder & Summary to Front
'Sheets(Sheets.Count).Visible = True
summaryWS.Move Before:=Sheets(1)
ws.Move Before:=Sheets(Sheets.Count)
Application.EnableAnimations = False
Application.ScreenUpdating = True
Application.Calculation = xlCalculationManual
summaryWS.EnableCalculation = False
summaryWS.DisplayPageBreaks = False
Dim sheetName$, sheetparent$, buildindex&, BuildAfterSheetName$
For nRowCounter = 2 To nTabOrderRows
sheetName = ws.Cells(nRowCounter, 1).Value
sheetparent = ws.Cells(nRowCounter, 2).Value
'1.Check to see if either Row is blank or already exits
If Len(Trim(sheetName)) > 0 And IsWorkSheetExists(sheetName) = False Then
buildindex = GetChildSheetsOffset(nRowCounter)
BuildAfterSheetName = Sheets(buildindex).Name
Application.ScreenUpdating = True
Application.StatusBar = "Building Sheet " & sheetName & ", Please Wait..."
Application.ScreenUpdating = False
'Debug.Print ("Building Sheet " & sheetName & ", After " & BuildAfterSheetName)
Dim cpws As Worksheet
summaryWS.Copy After:=Sheets(buildindex)
Set newWS = Sheets(buildindex + 1)
newWS.Name = Left(sheetName, 31)
'Grab Row Values
rowPCstr = ws.Cells(nRowCounter, 19).Value          ' column q
rowlocstr = ws.Cells(nRowCounter, 17).Value         ' column O
rowsitestr = ws.Cells(nRowCounter, 5).Value         ' column E
rowbustr = ws.Cells(nRowCounter, 6).Value           ' column F
rowdescription = ws.Cells(nRowCounter, 4).Value     ' column D
'Defaults
cPCList = "*"
cLOCList = "*"
cSITEList = "*"
cBUSUNITList = "*"
CDescription = "IGNORE"
If Len(Trim(rowPCstr)) > 0 Then
cPCList = "'" & Replace(Replace(rowPCstr, """", ""), "'", "")
End If
If Len(Trim(rowlocstr)) > 0 Then cLOCList = rowlocstr
If Len(Trim(rowsitestr)) > 0 Then cSITEList = rowsitestr
If Len(Trim(rowbustr)) > 0 Then cBUSUNITList = rowbustr
If Len(Trim(rowdescription)) > 0 Then CDescription = rowdescription
newWS.Range("C5:C7").NumberFormat = "@" ' ConvertTo Text
'newWS.Cells(4, 3).Value = cSITEList
'newWS.Cells(5, 3).Value = cBUSUNITList
'newWS.Cells(7, 3).Value = cLocList
newWS.Cells(4, 3).Value = "*"
newWS.Cells(5, 3).Value = "*"
newWS.Cells(7, 3).Value = "*"
newWS.Cells(6, 3).Value = cPCList
If Not CDescription = "IGNORE" And CDescription <> "" Then
newWS.Range("B9").NumberFormat = "@"
newWS.Cells(9, 2).Value = CDescription
End If
Application.StatusBar = False
End If '1. End Blank Row check
Next nRowCounter
Application.StatusBar = False
Application.ScreenUpdating = True
Application.ScreenUpdating = True
ExitMe:
Application.StatusBar = False
Application.ScreenUpdating = True
Application.ScreenUpdating = True
Exit Sub
End Sub
Private Function ColLetter(colNumber As Integer) As String
ColLetter = Left(Cells(1, colNumber).Address(False, False), Len(Cells(1, colNumber).Address(False, False)) - 1)
End Function
Public Sub UpdateCreteria()
Dim nRow As Integer
Dim nRowCounter As Integer
Dim ws As Worksheet
Dim tabname As String
Dim cPCList As String
Dim cLOCList As String
Dim cSITEList As String
Dim cBUSUNITList As String
Dim CDescriptionList As String
Dim wsDoesExist As Boolean
Set ws = Worksheets("TABORDER")
ws.Activate
Range("A1").Activate
nRow = ws.usedRange.Cells.Rows.Count
For nRowCounter = 2 To nRow
If Len(Trim(ws.Cells(nRowCounter, 1).Value)) > 0 Then
tabname = Trim(ws.Cells(nRowCounter, 1).Value)
cPCList = "*"
cLOCList = "*"
cSITEList = "*"
cBUSUNITList = "*"
CDescriptionList = "ZZZZZ"
' Get Profit Center value
If Len(ws.Cells(nRowCounter, 19).Value) > 0 Then ' column q
cPCList = ws.Cells(nRowCounter, 19).Value
Else
If Len(tabname) = 5 And tabname <> "Sharp" And (tabname = "00000" Or Val(tabname) > 0) Then
cPCList = "'" & tabname
End If
End If
If Len(ws.Cells(nRowCounter, 4).Value) > 0 Then CDescriptionList = ws.Cells(nRowCounter, 4).Value
If Len(ws.Cells(nRowCounter, 17).Value) > 0 Then cLOCList = ws.Cells(nRowCounter, 17).Value
If Len(ws.Cells(nRowCounter, 5).Value) > 0 Then cSITEList = ws.Cells(nRowCounter, 5).Value
If Len(ws.Cells(nRowCounter, 6).Value) > 0 Then cBUSUNITList = ws.Cells(nRowCounter, 6).Value
wsDoesExist = IsWorkSheetExists(tabname)
If wsDoesExist Then
Sheets(tabname).Range("C4:C7").NumberFormat = "@"
Sheets(tabname).Range("c6").Value = cPCList
If Len(ws.Cells(nRowCounter, cLocation).Value) = 3 Then
Sheets(tabname).Range("A1").Value = ws.Cells(nRowCounter, cLocation).Value
End If
If Not CDescriptionList = "ZZZZZ" Then
Sheets(tabname).Range("B9").Value = CDescriptionList
If Len(cPCList) > 10 Then
Sheets(tabname).Range("B10").Value = "Multiple"
End If
End If
End If
End If
Next nRowCounter
End Sub
Private Sub FilterSelection()
Dim ws As Worksheet
Set ws = ActiveSheet
For Each ws In ActiveWindow.SelectedSheets
ws.Activate
If ws.AutoFilterMode = True Then
Range(ws.AutoFilter.Range.Address).AutoFilter
End If
If Selection.Count > 1 Then
Range(Selection.Address).AutoFilter
Else
Range("A11", ActiveCell.SpecialCells(xlLastCell)).AutoFilter
End If
Next ws
ActiveWindow.SelectedSheets.item(1).Activate
End Sub
Private Sub removepagebreaks()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.DisplayPageBreaks = False
Next ws
End Sub
