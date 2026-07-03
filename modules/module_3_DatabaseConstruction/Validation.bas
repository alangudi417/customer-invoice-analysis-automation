    ' Validation Helper

Private Sub ValidateWorksheet(ByVal ws As Worksheet, ByVal varName As String)
    If ws Is Nothing Then
        Err.Raise 9999, "ValidateWorksheet", "Worksheet '" & varName & "' is Nothing"
    End If
End Sub


    '==============================================================================================


Private Sub ValidateDictionary(ByVal dict As Object, ByVal varName As String)
    If dict Is Nothing Then
        Err.Raise 9999, "ValidateDictionary", "Dictionary '" & varName & "' is Nothing"
    End If
End Sub