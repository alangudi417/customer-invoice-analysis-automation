
    ' SAFE OPEN WORKBOOK
Public Function OpenWBSafe(ByVal fullPath As String) As Workbook

    On Error GoTo ErrHandler

    If Len(Dir(fullPath)) = 0 Then
        Err.Raise 9999, "OpenWBSafe", "File not found: " & fullPath
    End If

    Set OpenWBSafe = Workbooks.Open(fullPath)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "OpenWBSafe", "Cannot open file: " & fullPath & " | " & Err.Description

End Function

    '=============================================================================================

    ' SAFE GET SHEET
Public Function GetSheetSafe(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet

    On Error GoTo ErrHandler

    If wb Is Nothing Then
        Err.Raise 9999, "GetSheetSafe", "Workbook is Nothing while trying to get sheet '" & sheetName & "'"
    End If

    On Error Resume Next
    Set GetSheetSafe = wb.Worksheets(sheetName)
    On Error GoTo ErrHandler

    If GetSheetSafe Is Nothing Then
        Err.Raise 9999, "GetSheetSafe", "Missing sheet '" & sheetName & "' in workbook: " & wb.Name
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GetSheetSafe", Err.Description

End Function