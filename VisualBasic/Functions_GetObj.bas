Attribute VB_Name = "Functions_GetObj"
Option Explicit
'
'TOC:
'   GetRangeObject
'   GetWorksheetObject
'   GetWorkbookObject
'
'
'
'
Private Function getRangeObject(wbNameOrIndex As String, wsNameOrIndex As String, Optional cellString As String = "a1", Optional ignoreHeadersRow As Boolean = True) As Range
 'RETURNS range Object based on workbook\worksheet object or search string
 '
 'DEPENDS ON:
    'Function GetWorkbookObject
    'Function GetWorkSheetObject
 
 
 
 'WBNameorIndex can be: workbook object, workbook search string, defaults to activeworkbook
 'WSNameOrIndex can be: worksheet Object, worksheet search string
 
 ' Dim rng As Range
 ' Set rng = getRangeObject("Updater", "deice", , False)
    
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim returnRange As Range
    Dim errorCode$
    errorCode = "PASS"
    
    
    On Error Resume Next
        
        Set wb = GetWorkbookObject(wbNameOrIndex)
        If Err Or wb Is Nothing Then errorCode = "Invalid Error Code; "
        Set ws = GetWorksheetObject(wbNameOrIndex, wsNameOrIndex)
        If Err Or ws Is Nothing Then errorCode = errorCode & "`n Invalid SheetName; `n"
                 
        If IsMissing(cellString) Or UCase$(cellString) = "A1" Then
            If ignoreHeadersRow = True Then
                Set returnRange = ws.UsedRange.offset(1, 0).Resize(ws.UsedRange.Rows.count - 1, ws.UsedRange.Columns.count)
                If Err Then errorCode = errorCode & "`nInvalid destination cell address; "
                Else
                Set returnRange = ws.UsedRange
                    If Err Then errorCode = errorCode & "`nInvalid destination cell address 2; "
                End If
                    
            Else
                Set returnRange = ws.Range(cellString)
                    If Err Then errorCode = errorCode & "`nInvalid destination cell address 3; `n"
            End If
        


       ' Debug.Print (errorCode)
       Set getRangeObject = returnRange
         On Error GoTo 0
         
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



