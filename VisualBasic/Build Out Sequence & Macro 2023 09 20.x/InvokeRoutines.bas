Attribute VB_Name = "InvokeRoutines"
Option Explicit
Sub InvokeBuildTabs()
Dim wb As Workbook
On Error Resume Next
Set wb = Workbooks(ThisWorkbook.Names("wbName").RefersToRange.Text)
If Err Then
MsgBox "Invalid Name"
Err.Clear
Exit Sub
End If
On Error GoTo 0
wb.Activate
GetBuildOrderCollection
' Call SUMonParentsSheets
Call UpdateCreteria
Set wb = Nothing
End Sub
Sub InvokeUpdateParentSheets()
Dim wb As Workbook
On Error Resume Next
Set wb = Workbooks(ThisWorkbook.Names("wbName").RefersToRange.Text)
If Err Then
MsgBox "Invalid Name"
Err.Clear
Exit Sub
End If
On Error GoTo 0
wb.Activate
SUMonParentsSheets
Set wb = Nothing
End Sub
Sub ActivateTabOrder()
Dim wb As Workbook
On Error Resume Next
Set wb = Workbooks(ThisWorkbook.Names("wbName").RefersToRange.Text)
If Err Then
MsgBox "Invalid Name"
Err.Clear
Exit Sub
End If
On Error GoTo 0
wb.Worksheets("TabOrder").Activate
Set wb = Nothing
End Sub
Sub InvokeUpdateMEC()
Dim wb As Workbook
On Error Resume Next
Set wb = Workbooks(ThisWorkbook.Names("wbName").RefersToRange.Text)
If Err Then
MsgBox "Invalid Name"
Err.Clear
Exit Sub
End If
On Error GoTo 0
wb.Activate
updateMEC
Set wb = Nothing
End Sub
