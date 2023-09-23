Attribute VB_Name = "Functions_CopyClear"
Option Explicit



Private Function ClearAndCopyRange(srcRange As Variant, destRange As Variant, Optional ignoreSourceHeaders As Boolean = True, Optional ignoreDestHeaders As Boolean = True, Optional destSheet As Variant)
    
    'Destwb.Sheets("tabname").Range("address")
    'Destwb.Sheets("tabname").usedrange or Destwb.Sheets("tabname").currentRange
    'ActiveSheet.UsedRange
    '[a1:a2
    Dim src As Range
    Dim dest As Range
    
    If srcRange Is Nothing Or destRange Is Nothing Then
        
        If srcRange Is Nothing Then Debug.Print ("No SrcRange")
        If destRange Is Nothing Then Debug.Print ("No destRange")
    
    Exit Function
    End If
    
    
    If ignoreSourceHeaders = True Then
        Set src = srcRange.offset(1, 0).Resize(srcRange.Rows.count - 1, srcRange.Columns.count)
    Else
        Set src = srcRange
    End If

    ' ignoreDestHeaders when clearing, true by default
    If ignoreDestHeaders = True Then
        destRange.offset(1, 0).Resize(destRange.Rows.count - 1, destRange.Columns.count).Clear
        Set dest = destRange.offset(1, 0).Resize(src.Rows.count, src.Columns.count)
        Else
        destRange.Clear
        Set dest = destRange.offset(0, 0).Resize(src.Rows.count, src.Columns.count)
        
    End If
         
    
    dest.Value2 = src.Value2


End Function
