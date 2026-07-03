    ' Utility Helper

Public Function LastRowSafe(ByVal ws As Worksheet, ByVal col As String) As Long

    ValidateWorksheet ws, "LastRowSafe.ws"

    LastRowSafe = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    If LastRowSafe < 1 Then LastRowSafe = 1

End Function


    '================================================================================================


    ' Utility Helper
    
Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = ""
    Else
        SafeText = Trim(CStr(v))
    End If
End Function


    '================================================================================================


    ' Utility Helper

Private Function SafeCDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        SafeCDbl = 0
    ElseIf Trim(CStr(v)) = "" Then
        SafeCDbl = 0
    ElseIf IsNumeric(v) Then
        SafeCDbl = CDbl(v)
    Else
        SafeCDbl = 0
    End If
End Function