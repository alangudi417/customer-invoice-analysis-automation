Private Sub ApplyWarboardFormulas(ByVal ws As Worksheet, ByVal fDate As String)

    On Error GoTo ErrHandler

    ValidateWorksheet ws, "ApplyWarboardFormulas.ws"

    Dim lr As Long

    lr = LastRowSafe(ws, "J")
    If lr < 2 Then Exit Sub

    ' Clear requested columns (rows only, keep headers)
    ws.Range("P2:P" & lr).ClearContents   ' Taxes
    ws.Range("Q2:Q" & lr).ClearContents   ' Tariff
    ws.Range("R2:R" & lr).ClearContents   ' Total Billed
    ws.Range("S2:S" & lr).ClearContents   ' Total ZREPRINT
    ws.Range("T2:T" & lr).ClearContents   ' Dif
    ws.Range("W2:W" & lr).ClearContents   ' Paid?
    ws.Range("AD2:AD" & lr).ClearContents ' Filter

        ' Insert formulas:
    ' Taxes
    ws.Range("P2").Formula = _
        "=IF(COUNTIF($J$2:J2,J2)=1," & _
        "IFERROR(INDEX('[VBRP " & fDate & ".xlsx]Pivot'!$B:$B," & _
        "MATCH(J2,'[VBRP " & fDate & ".xlsx]Pivot'!$A:$A,0)),""""),"""")"

    ' Tariff
    ws.Range("Q2").Formula = _
        "=IF(COUNTIF($J$2:J2,J2)=1," & _
        "IFERROR(INDEX('[TAR001 " & fDate & ".xlsx]Pivot'!$B:$B," & _
        "MATCH(J2,'[TAR001 " & fDate & ".xlsx]Pivot'!$A:$A,0)),""""),"""")"

    ' Total Billed
    ws.Range("R2").Formula = _
        "=IF(COUNTIF($J$2:J2,J2)=1," & _
        "SUMIFS($O:$O,$J:$J,J2)+SUMIFS($P:$P,$J:$J,J2)+SUMIFS($Q:$Q,$J:$J,J2),"""")"

    ' Total ZREPRINT
    ws.Range("S2").Formula = _
        "=IF(COUNTIF($J$2:J2,J2)=1," & _
        "IFERROR(INDEX('[ZREPRINT " & fDate & ".xlsx]Pivot'!$C:$C," & _
        "MATCH(J2,'[ZREPRINT " & fDate & ".xlsx]Pivot'!$A:$A,0)),""""),"""")"

    ' Filter
    ws.Range("AD2").Formula = _
        "=IF(AND(R2="""",S2=""""),""N/A"",IF(R2=S2,""Ok"",""Check""))"

    ' Lookup Filter
    ws.Range("T2").Formula = _
        "=IFERROR(INDEX(AD:AD,MATCH(J2,J:J,0)),"""")"

    ' Paid?
    ws.Range("W2").Formula = _
        "=IF(OR(S2="""",V2=""""),""N/A"",IF(S2<>V2,""Short Paid"",""Not Paid""))"

    ' Fill formulas down
    If lr > 2 Then
        ws.Range("P2:P" & lr).FillDown
        ws.Range("Q2:Q" & lr).FillDown
        ws.Range("R2:R" & lr).FillDown
        ws.Range("S2:S" & lr).FillDown
        ws.Range("AD2:AD" & lr).FillDown
        ws.Range("T2:T" & lr).FillDown
        ws.Range("W2:W" & lr).FillDown
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "ApplyWarboardFormulas", _
              "Error applying formulas in sheet '" & ws.Name & "': " & Err.Description

End Sub