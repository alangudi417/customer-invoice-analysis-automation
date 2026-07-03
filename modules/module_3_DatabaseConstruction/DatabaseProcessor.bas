    ' PROCESS MAIN LOGIC
    ' Here is where the magic comes!
Private Sub ProcessSheetSafe( _
    ByVal wsDB As Worksheet, _
    ByVal wsSrc As Worksheet, _
    ByVal dictUDM As Object, _
    ByVal dictNotes As Object, _
    ByVal dictVBRP As Object, _
    ByVal dictTAR As Object, _
    ByVal dictZREP As Object, _
    ByVal isTLA As Boolean)

    On Error GoTo ErrHandler

    ' Validate inputs
    ValidateWorksheet wsDB, "wsDB"
    ValidateWorksheet wsSrc, "wsSRC"
    ValidateDictionary dictUDM, "dictUDM"
    ValidateDictionary dictNotes, "dictNotes"
    ValidateDictionary dictVBRP, "dictVBRP"
    ValidateDictionary dictTAR, "dictTAR"
    ValidateDictionary dictZREP, "dictZREP"

    Dim arr As Variant
    Dim i As Long, r As Long, lr As Long
    Dim dictStatus As Object, dictSeen As Object
    Dim inv As String, cat As String

    Dim qty As Double, price As Double
    Dim subtotal As Double, tax As Double, tariff As Double, zrep As Double
    Dim total As Double

    Set dictStatus = CreateObject("Scripting.Dictionary")
    Set dictSeen = CreateObject("Scripting.Dictionary")

    ' Clear old database data
    wsDB.Rows("2:" & wsDB.Rows.Count).ClearContents


    '===========================
    ' PHASE 1 - Load source data
    '===========================
    
    lr = LastRowSafe(wsSrc, "A")
    If lr < 2 Then Exit Sub   ' no source data

    arr = wsSrc.Range("A2:I" & lr).Value
    r = 2

    For i = 1 To UBound(arr, 1)

        If isTLA Then
            cat = SafeText(arr(i, 7))
            If Not (cat = "ZCCC" Or cat = "ZDEF" Or cat = "ZTLA") Then
                GoTo SkipRow
            End If
        End If

        wsDB.Cells(r, "J").Value = arr(i, 1) ' Invoice
        wsDB.Cells(r, "G").Value = arr(i, 2)
        wsDB.Cells(r, "H").Value = arr(i, 3)
        wsDB.Cells(r, "I").Value = arr(i, 4)
        wsDB.Cells(r, "L").Value = arr(i, 5)
        wsDB.Cells(r, "K").Value = arr(i, 6)
        wsDB.Cells(r, "M").Value = arr(i, 8)
        wsDB.Cells(r, "N").Value = arr(i, 9)

        r = r + 1

SkipRow:
    Next i

    lr = LastRowSafe(wsDB, "J")
    If lr < 2 Then Exit Sub

    '===========================
    ' PHASE 2 - Fill DB data
    '===========================
    
    For i = 2 To lr

        inv = SafeText(wsDB.Cells(i, "J").Value)
        If inv = "" Then GoTo NextI


        ' UDM data
        Dim udmData As Variant


        If dictUDM.Exists(inv) Then
            udmData = dictUDM(inv)

            If IsArray(udmData) Then
                If UBound(udmData) >= 13 Then
                    wsDB.Cells(i, "A").Value = udmData(0)   ' Days
                    wsDB.Cells(i, "B").Value = udmData(1)   ' Dispute
                    wsDB.Cells(i, "C").Value = udmData(2)   ' AOR
                    wsDB.Cells(i, "D").Value = udmData(3)   ' Customer Number
                    wsDB.Cells(i, "E").Value = udmData(4)   ' Customer Name
                    wsDB.Cells(i, "F").Value = udmData(5)   ' Cause Desc
                    wsDB.Cells(i, "G").Value = udmData(6)   ' PO
                    wsDB.Cells(i, "H").Value = udmData(7)   ' Sales Order

                    wsDB.Cells(i, "U").Value = udmData(8)   ' Inv Date
                    wsDB.Cells(i, "V").Value = udmData(9)   ' Disputed Amount
                    wsDB.Cells(i, "X").Value = udmData(10)  ' Manager Name
                    wsDB.Cells(i, "Y").Value = udmData(11)  ' Owner
                    wsDB.Cells(i, "Z").Value = udmData(12)  ' Processor
                    wsDB.Cells(i, "AA").Value = udmData(13) ' Processor Name
                Else
                    Err.Raise 9999, "ProcessSheetSafe", _
                        "UDM array for invoice '" & inv & "' does not contain all expected fields."
                End If
            Else
                Err.Raise 9999, "ProcessSheetSafe", _
                    "UDM data for invoice '" & inv & "' is not an array."
            End If
        End If



        ' Notes (AB)
        If SafeText(wsDB.Cells(i, "B").Value) <> "" Then
            If dictNotes.Exists(SafeText(wsDB.Cells(i, "B").Value)) Then
                wsDB.Cells(i, "AB").Value = dictNotes(SafeText(wsDB.Cells(i, "B").Value))
            End If
        End If

        ' Formulas
        qty = SafeCDbl(wsDB.Cells(i, "M").Value)
        price = SafeCDbl(wsDB.Cells(i, "N").Value)

        subtotal = qty * price

        If dictVBRP.Exists(inv) Then
            tax = SafeCDbl(dictVBRP(inv))
        Else
            tax = 0
        End If

        If dictTAR.Exists(inv) Then
            tariff = SafeCDbl(dictTAR(inv))
        Else
            tariff = 0
        End If

        If dictZREP.Exists(inv) Then
            zrep = SafeCDbl(dictZREP(inv))
        Else
            zrep = 0
        End If

        total = subtotal + tax + tariff

        wsDB.Cells(i, "O").Value = subtotal
        wsDB.Cells(i, "P").Value = tax
        wsDB.Cells(i, "Q").Value = tariff
        wsDB.Cells(i, "R").Value = total
        wsDB.Cells(i, "S").Value = zrep

        
        ' Status by invoice
        If Not dictSeen.Exists(inv) Then
            If Round(zrep, 2) = Round(total, 2) Then
                dictStatus(inv) = "Ok"
            Else
                dictStatus(inv) = "Check"
            End If
            dictSeen.Add inv, True
        End If

        wsDB.Cells(i, "T").Value = dictStatus(inv)

        
        ' Payment status
        If SafeText(wsDB.Cells(i, "S").Value) = "" Or SafeText(wsDB.Cells(i, "V").Value) = "" Then
            wsDB.Cells(i, "W").Value = "N/A"
        ElseIf Round(SafeCDbl(wsDB.Cells(i, "S").Value), 2) <> Round(SafeCDbl(wsDB.Cells(i, "V").Value), 2) Then
            wsDB.Cells(i, "W").Value = "Short Paid"
        Else
            wsDB.Cells(i, "W").Value = "Not Paid"
        End If

        wsDB.Cells(i, "AC").Value = Application.WorksheetFunction.CountIf(wsDB.Range("J:J"), inv)

NextI:
    Next i

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "ProcessSheetSafe", _
              "Error in sheet '" & wsDB.Name & "' using source '" & wsSrc.Name & "': " & Err.Description

End Sub