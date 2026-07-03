    ' DICTIONARY BUILDERS
    ' Here is also where the magic comes!

Public Function BuildUDMDictionarySafe(ByVal ws As Worksheet) As Object

    On Error GoTo ErrHandler

    ValidateWorksheet ws, "BuildUDMDictionarySafe.ws"

    Dim dict As Object
    Dim lr As Long, i As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    lr = LastRowSafe(ws, "L")
    If lr < 2 Then
        Set BuildUDMDictionarySafe = dict
        Exit Function
    End If

    For i = 2 To lr
        key = SafeText(ws.Cells(i, "L").Value)   ' key used to match invoice/case title

        If key <> "" Then
            dict(key) = Array( _
                ws.Cells(i, "H").Value, _
                ws.Cells(i, "I").Value, _
                ws.Cells(i, "G").Value, _
                ws.Cells(i, "N").Value, _
                ws.Cells(i, "O").Value, _
                ws.Cells(i, "P").Value, _
                ws.Cells(i, "J").Value, _
                ws.Cells(i, "K").Value, _
                ws.Cells(i, "M").Value, _
                ws.Cells(i, "Q").Value, _
                ws.Cells(i, "AD").Value, _
                ws.Cells(i, "T").Value, _
                ws.Cells(i, "Y").Value, _
                ws.Cells(i, "Z").Value _
            )
        End If
    Next i

    Set BuildUDMDictionarySafe = dict
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "BuildUDMDictionarySafe", Err.Description

End Function


    '==================================================================================


    ' Safe dictionary building

Public Function BuildNotesDictionarySafe(ByVal ws As Worksheet) As Object

    On Error GoTo ErrHandler

    ValidateWorksheet ws, "BuildNotesDictionarySafe.ws"

    Dim dict As Object
    Dim lr As Long, i As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    lr = LastRowSafe(ws, "C")
    If lr < 2 Then
        Set BuildNotesDictionarySafe = dict
        Exit Function
    End If

    For i = 2 To lr
        key = SafeText(ws.Cells(i, "C").Value)
        If key <> "" Then
            dict(key) = ws.Cells(i, "AW").Value
        End If
    Next i

    Set BuildNotesDictionarySafe = dict
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "BuildNotesDictionarySafe", Err.Description

End Function


    '==================================================================================


    ' Safe dictionary building

Public Function BuildSimpleDictionarySafe(ByVal ws As Worksheet, ByVal keyCol As String, ByVal valCol As String) As Object

    On Error GoTo ErrHandler

    ValidateWorksheet ws, "BuildSimpleDictionarySafe.ws"

    Dim dict As Object
    Dim lr As Long, i As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    lr = LastRowSafe(ws, keyCol)
    If lr < 2 Then
        Set BuildSimpleDictionarySafe = dict
        Exit Function
    End If

    For i = 2 To lr
        key = SafeText(ws.Cells(i, keyCol).Value)
        If key <> "" Then
            dict(key) = ws.Cells(i, valCol).Value
        End If
    Next i

    Set BuildSimpleDictionarySafe = dict
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "BuildSimpleDictionarySafe", Err.Description

End Function