Attribute VB_Name = "Scractvh"
Option Explicit

Sub ResetArrayFormulas()
    Dim ws As Worksheet
    Dim DataRange As Range
    Dim iRow As Long
    Dim icol As Interior
    Set ws = ActiveSheet
    Set DataRange = ws.UsedRange

    DataRange.Formula2 = DataRange.Formula


End Sub

Sub InsertSEIFormulas()
    Dim srcWB As Workbook
    Dim targetWB As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    Dim sourceWBSearchString$, targetWBSearchString$
    sourceWBSearchString = "03MO Tem"
    targetWBSearchString = "48MO"




    For i = 1 To Workbooks.count
        If InStr(1, Workbooks(i).Name, sourceWBSearchString) Then Set srcWB = Workbooks(i)
        If InStr(1, Workbooks(i).Name, targetWBSearchString) Then Set targetWB = Workbooks(i)
    Next i

    If srcWB Is Nothing Or targetWB Is Nothing Then
        MsgBox "Src or target workbook not found!"
        Exit Sub
    End If


    If CheckWSIfExists("NectariAddinForExcelProperties", srcWB.Name) Then
        If CheckWSIfExists("NectariAddinForExcelProperties", targetWB.Name) Then
            Application.DisplayAlerts = False
            Set ws = targetWB.Worksheets("NectariAddinForExcelProperties")
            ws.Visible = xlSheetVisible
            ws.Delete
            targetWB.Worksheets(1).Visible = xlSheetVeryHidden
            Application.DisplayAlerts = True
        End If
        srcWB.Sheets("NectariAddinForExcelProperties").Copy before:=targetWB.Sheets(1)
    End If
    Set ws = Nothing

    Dim srcWS As Workbook
    If CheckWSIfExists("BLANK", srcWB.Name) Then
    Set srcWS = srcWB.Worksheets("BLANK")


    End If


End Sub
Private Sub fillacrossheet()
Dim strNames() As Variant
Dim i, nSelectedSheets As Integer
Dim ws As Worksheet

   
    Application.EnableAnimations = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Sheets("Summary").EnableCalculation = False
    Sheets("Summary").DisplayPageBreaks = False


nSelectedSheets = ActiveWindow.SelectedSheets.count

ReDim strNames(0 To nSelectedSheets)
strNames(0) = "Summary"

i = 1
For Each ws In ActiveWindow.SelectedSheets
    strNames(i) = ws.Name
    i = i + 1

Next ws

Debug.Print Join(strNames, ", ")
Sheets("Summary").Activate

Sheets(strNames).FillAcrossSheets _
 Worksheets("Summary").UsedRange



    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationManual
    Sheets("Summary").EnableCalculation = True
End Sub

Private Sub FilterSelection()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    For Each ws In ActiveWindow.SelectedSheets

        If ws.AutoFilterMode = True Then
            Range(ws.AutoFilter.Range.Address).AutoFilter
        End If


        Range("A11").Resize(358, ActiveCell.SpecialCells(xlLastCell).Column).AutoFilter
        Range("C11").Select
        ActiveWindow.Zoom = 85
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    Next ws

End Sub

Private Function CheckWSIfExists(sheetname, Optional workbookName As String) As Boolean

  On Error Resume Next
    Dim wb As Workbook
    If workbookName = "" Then
        Set wb = ActiveWorkbook
        Else
        Set wb = Workbooks(workbookName)
    End If

  
  CheckWSIfExists = wb.Worksheets(sheetname).Name <> ""
  On Error GoTo 0
End Function


Public Sub HelloWorld()
    MsgBox "Hello World"
End Sub

Sub UpdateCreteria2()
    Dim rg As Range
    Set rg = LastRange("TabOrder", 1)
    Dim arr As Variant
    arr = rg
    
    Dim rowCount As Long, i As Integer, j As Integer
    rowCount = rg.Rows.count
    Dim sheetname As String, pcString As String
    For i = LBound(arr) To UBound(arr)
        sheetname = arr(i, 1)
        'Debug.Assert sheetName <> "Region Davis Consol"
        pcString = arr(i, 19)
        If CheckWSIfExists(sheetname) Then

            Worksheets(sheetname).Range("C6").Value2 = pcString
        
        End If
        
    Next i
    

    Debug.Print rg.Address
End Sub

Sub UPDATEME()
    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        Debug.Print ws.Name
        ws.Calculate
    Next ws
End Sub
