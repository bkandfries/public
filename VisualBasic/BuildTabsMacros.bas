Attribute VB_Name = "BuildTabsMacros"

Option Explicit

Private Function GetChildSheetsOffset(CurrentRow) As Integer
    Dim nRowCounter As Integer
    Dim nTabOrderRows As Integer
    Dim ws As Worksheet
    Dim rowsheet As Long, rowparent As Long
    Dim forSheetName As String, forParent As String, sSheetName As String
    
        
    
    Set ws = Sheets("TabOrder")
    rowsheet = ws.Cells(CurrentRow, 1).Value
    rowparent = ws.Cells(CurrentRow, 2).Value
    
    'Return Last Sheet Index by Default
    GetChildSheetsOffset = Sheets("TabOrder").index - 1
    
    'If Parent Exists, then Index will be after Parent Sheet Index + any Children that already exits
    If IsWorkSheetExists(rowparent) Then
    GetChildSheetsOffset = Sheets(rowparent).index
    
    For nRowCounter = 1 To (CurrentRow - 1) 'Loop through all Previous Rows in TabORder
        If Len(Trim(ws.Cells(nRowCounter, 1).Value)) > 0 Then 'Skip Rows with Blank PCs
            forSheetName = ws.Cells(nRowCounter, 1).Value
            forParent = ws.Cells(nRowCounter, 2).Value
            
            
            If (forParent = rowparent) And (IsWorkSheetExists(forSheetName)) And (forSheetName <> sSheetName) Then
            GetChildSheetsOffset = GetChildSheetsOffset + 1
        End If
        
    End If
    Next nRowCounter
    
    End If
    End Function


Sub BuildTab3_Personal()
    Dim nRowCounter As Integer
    Dim nTabOrderRows As Integer
    Dim ws As Worksheet
    Dim ChildOffset As Integer
    Dim newWS As Worksheet
    Set ws = Sheets("TabOrder")
    nTabOrderRows = ws.UsedRange.Cells.Rows.count
    Dim SUMMARY As String
    SUMMARY = "SUMMARY"
    Dim summaryWS As Worksheet
    
    
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
    summaryWS.Move before:=Sheets(1)
    summaryWS.Move before:=Sheets(2)
        
        
        Application.EnableAnimations = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        summaryWS.EnableCalculation = False
        summaryWS.DisplayPageBreaks = False
        
        For nRowCounter = 2 To nTabOrderRows
    
    
    Dim sheetname As String, sheetparent As String, BuildAfterSheetName As String
    Dim buildindex As Long
    Dim rowPCStr As String, rowLocStr As String, rowSiteStr As String, rowBUStr As String, rowDescription As String, CDescription As String
    Dim cPCList As String, cLocList As String, cSITEList As String, cBUSUNITList As String
    Dim cleanedPcStr As String
    
    
    sheetname = ws.Cells(nRowCounter, 1).Value
    sheetparent = ws.Cells(nRowCounter, 2).Value
    
        '1.Check to see if either Row is blank or already exits
        If Len(Trim(sheetname)) > 0 And IsWorkSheetExists(sheetname) = False Then '1.
            buildindex = GetChildSheetsOffset(nRowCounter)
            BuildAfterSheetName = Sheets(buildindex).Name
            
            Application.ScreenUpdating = True
            Application.StatusBar = "Building Sheet " & sheetname & ", Please Wait..."
            Application.ScreenUpdating = False
            
            'Debug.Print ("Building Sheet " & sheetName & ", After " & BuildAfterSheetName)
            
            summaryWS.Copy After:=Sheets(buildindex)
            Set newWS = Sheets(buildindex + 1)
            
            newWS.Name = Left(sheetname, 31)
            
            'Grab Row Values
            rowPCStr = ws.Cells(nRowCounter, 19).Value          ' column q
            rowLocStr = ws.Cells(nRowCounter, 17).Value         ' column O
            rowSiteStr = ws.Cells(nRowCounter, 5).Value         ' column E
            rowBUStr = ws.Cells(nRowCounter, 6).Value           ' column F
            rowDescription = ws.Cells(nRowCounter, 4).Value     ' column D
            
            'Defaults
            cPCList = "*"
            cLocList = "*"
            cSITEList = "*"
            cBUSUNITList = "*"
            CDescription = "IGNORE"
            
            
            If Len(Trim(rowPCStr)) > 0 Then
                'Cleanup on the PC first
                cleanedPcStr = Replace(Replace(rowPCStr, """", ""), "'", "")
            cPCList = "'" & cleanedPcStr
            End If
            If Len(Trim(rowLocStr)) > 0 Then cLocList = rowLocStr
            If Len(Trim(rowSiteStr)) > 0 Then cSITEList = rowSiteStr
            If Len(Trim(rowBUStr)) > 0 Then cBUSUNITList = rowBUStr
            If Len(Trim(rowDescription)) > 0 Then CDescription = rowDescription
            
            newWS.Range("C5:C7").NumberFormat = "@" ' ConvertTo Text
            newWS.Cells(4, 3).Value = cSITEList
            newWS.Cells(5, 3).Value = cBUSUNITList
            newWS.Cells(6, 3).Value = cPCList
            newWS.Cells(7, 3).Value = cLocList
            If Not CDescription = "IGNORE" Then
                newWS.Range("B9").NumberFormat = "@"
            newWS.Cells(9, 2).Value = CDescription
            End If
            
            
            
        Application.StatusBar = False
        
        
        End If '1. End Blank Row check
        Next nRowCounter
        
        
        Application.StatusBar = False
        Application.ScreenUpdating = True
        
    Application.ScreenUpdating = True
exitMe:
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Application.ScreenUpdating = True
    Exit Sub
    
        
                    
End Sub

                    
Private Function ColLetter(ColNumber As Integer) As String
    ColLetter = Left(Cells(1, ColNumber).Address(False, False), Len(Cells(1, ColNumber).Address(False, False)) - 1)
End Function

Private Sub UpdateCreteria()

  Dim nRow As Integer
  Dim nRowCounter As Integer
  Dim ws As Worksheet
  Dim tabname As String
  Dim cPCList As String
  Dim cLocList As String
  Dim cSITEList As String
  Dim cBUSUNITList As String
  Dim CDescription As String
  Dim wsDoesExist As Boolean




  Set ws = Sheets("TABORDER")

  ws.Activate
  Range("A1").Activate
  nRow = ws.UsedRange.Cells.Rows.count


  For nRowCounter = 2 To nRow
    If Len(Trim(ws.Cells(nRowCounter, 1).Value)) > 0 Then

      tabname = Trim(ws.Cells(nRowCounter, 1).Value)


      cPCList = "*"
      cLocList = "*"
      cSITEList = "*"
      cBUSUNITList = "*"
      CDescription = "ZZZZZ"


      ' Get Profit Center value
      If Len(ws.Cells(nRowCounter, 19).Value) > 0 Then ' column q
        cPCList = ws.Cells(nRowCounter, 19).Value
      Else
        If Len(tabname) = 5 And tabname <> "Sharp" And (tabname = "00000" Or Val(tabname) > 0) Then
          cPCList = "'" & tabname
        End If
      End If

      ' Get SheetDescription
      If Len(ws.Cells(nRowCounter, 4).Value) > 0 Then ' column d
        CDescription = ws.Cells(nRowCounter, 4).Value
      End If

      ' Get Location
      If Len(ws.Cells(nRowCounter, 17).Value) > 0 Then ' column O
        cLocList = ws.Cells(nRowCounter, 17).Value
      End If

      '
      '               Get SITE value
      If Len(ws.Cells(nRowCounter, 5).Value) > 0 Then ' column q
        cSITEList = ws.Cells(nRowCounter, 5).Value
      End If
      '               Get BUS UNIT value
      If Len(ws.Cells(nRowCounter, 6).Value) > 0 Then ' column q
        cBUSUNITList = ws.Cells(nRowCounter, 6).Value
      End If

      

      If IsWorkSheetExists(tabname) Then

        Sheets(tabname).Range("C4:C7").NumberFormat = "@"
        Sheets(tabname).Range("c6").Value = cPCList


        If Not CDescription = "ZZZZZ" Then
          Sheets(tabname).Range("B9").Value = CDescription
          If Len(cPCList) > 10 Then Sheets(tabname).Range("B10").Value = "Multiple"
          End If

        End If

      End If

    Next nRowCounter

End Sub



