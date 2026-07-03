Option Explicit

' This first macro will help to analyze the data from ZSO_SEARCH, and separate the invoices data in 3 categories:
' one-line invoices, multi-lines invoices and TLAs

Sub invoice_categories()

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsData As Worksheet
    Dim wsOneLine As Worksheet
    Dim wsTLA As Worksheet
    Dim wsMulti As Worksheet
    Dim wsPivot As Worksheet
    Dim ptSheet As Worksheet

    Dim LastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim i As Long

    Dim PivotCache As PivotCache
    Dim PivotTable As PivotTable
    Dim PivotRange As Range
    Dim DataRange As Range

    Dim FilePath As String
    Dim fileName As String
    Dim key As String
    Dim cat As String

    Dim dict As Object
    Dim ans As VbMsgBoxResult

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

        ' Ask before proceeding
    ans = MsgBox("Do you have the ZSO_SEARCH file ready?" & vbCrLf & vbCrLf & _
                 "Select Yes to continue.", vbYesNo + vbQuestion, "Confirmation")
    If ans = vbNo Then GoTo ErrorHandler

    'OPEN ZSO FILE
    FilePath = "/sample_data/input/ZSO_Search/"
    fileName = "ZSO_SEARCH " & Format(Date, "mm-dd-yy") & ".xlsx"

    Set wb = Workbooks.Open(FilePath & fileName)
    Set wsSource = wb.Sheets("Sheet1")

    'CREATE TABS
    wsSource.Copy After:=wsSource
    Set wsData = ActiveSheet
    wsData.Name = "Data ZSO"

    Set wsOneLine = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsOneLine.Name = "One Line Invoice"

    Set wsTLA = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsTLA.Name = "TLA"

    Set wsMulti = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsMulti.Name = "Mult Lines Invoice"

    Set wsPivot = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPivot.Name = "Pivot"

    'TEXT TO COLUMNS AV
    With wsData.Columns("AV")
        .TextToColumns Destination:=.Cells(1, 1), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False
    End With

    'CREATE CONCA COLUMN
    LastRow = wsData.Cells(wsData.Rows.Count, "AV").End(xlUp).Row
    wsData.Range("BT1").Value = "Conca"
    For r = 2 To LastRow
        wsData.Cells(r, "BT").Value = _
            wsData.Cells(r, "AV").Value & _
            wsData.Cells(r, "AW").Value & _
            wsData.Cells(r, "AX").Value
    Next r

    'REMOVE DUPLICATES
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(LastRow, lastCol)).RemoveDuplicates _
        Columns:=72, Header:=xlYes

    LastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    'CREATE PIVOT TABLE
    Set DataRange = wsData.Range("A1").CurrentRegion
    Set ptSheet = wb.Sheets.Add
    
    Set PivotCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=DataRange)

    Set PivotTable = PivotCache.CreatePivotTable( _
        TableDestination:=ptSheet.Range("A3"), _
        TableName:="ZSO_Pivot")

    With PivotTable

        .PivotFields("Billing document").Orientation = xlRowField
        .PivotFields("PO No.").Orientation = xlRowField
        .PivotFields("SO Order number").Orientation = xlRowField
        .PivotFields("Delivery document number").Orientation = xlRowField
        .PivotFields("Manufacturer part number").Orientation = xlRowField
        .PivotFields("Proc Strategry").Orientation = xlRowField
        .PivotFields("Item category").Orientation = xlRowField

        .AddDataField .PivotFields("Billing quantity"), _
            "Sum of Billing quantity", xlSum

        .AddDataField .PivotFields("Unit Price"), _
            "Average of Unit Price", xlAverage

        .RowAxisLayout xlTabularRow

        .RepeatAllLabels xlRepeatLabels

        .ColumnGrand = False
        .RowGrand = False

    End With

    Dim pf As PivotField

    For Each pf In PivotTable.RowFields
        pf.Subtotals(1) = False
    Next pf

    'COPY PIVOT TO PIVOT TAB
    Set PivotRange = ptSheet.UsedRange
    PivotRange.Copy
    
    wsPivot.Range("A1").PasteSpecial xlPasteValues
    wsPivot.Columns.AutoFit
    ptSheet.Delete
    
    LastRow = wsPivot.Cells(wsPivot.Rows.Count, "A").End(xlUp).Row
    
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To LastRow
        key = wsPivot.Cells(i, "A").Value

        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
    Next i
    
    wsPivot.Range("J1").Value = "TLA?"
    For i = 2 To LastRow

        cat = wsPivot.Cells(i, "G").Value

        If cat = "ZCCC" Or _
           cat = "ZDEF" Or _
           cat = "ZSUB" Or _
           cat = "ZTAE" Or _
           cat = "ZTLA" Or _
           cat = "ZTAL" Or _
           cat = "Z002" Then

            wsPivot.Cells(i, "J").Value = "YES"
        Else
            wsPivot.Cells(i, "J").Value = "NO"
        End If

    Next i


    ' Here's where the magic comes:
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
        'FAST ARRAY-BASED SPLIT
    Dim data As Variant
    Dim oneArr() As Variant
    Dim tlaArr() As Variant
    Dim multiArr() As Variant

    Dim oneCount As Long, tlaCount As Long, multiCount As Long
    Dim c As Long
    Dim colCount As Long

    colCount = 10
    data = wsPivot.Range("A2:J" & LastRow).Value

    ReDim oneArr(1 To UBound(data), 1 To colCount)
    ReDim tlaArr(1 To UBound(data), 1 To colCount)
    ReDim multiArr(1 To UBound(data), 1 To colCount)

    For i = 1 To UBound(data)

        Dim cnt As Long
        Dim flag As String
        Dim invoiceKey As String
        
        invoiceKey = data(i, 1)

        If dict.Exists(invoiceKey) Then
            cnt = dict(invoiceKey)
        Else
            cnt = 0
        End If
        
        flag = data(i, 10)

        If flag = "YES" Then
            tlaCount = tlaCount + 1
            For c = 1 To colCount
                tlaArr(tlaCount, c) = data(i, c)
            Next c

        ElseIf cnt = 1 Then

            oneCount = oneCount + 1
            For c = 1 To colCount
                oneArr(oneCount, c) = data(i, c)
            Next c
            
        Else

            multiCount = multiCount + 1
            For c = 1 To colCount
                multiArr(multiCount, c) = data(i, c)
            Next c

        End If
    
    Next i

            ' 1. ONE LINE INVOICE
    If oneCount > 0 Then
        wsOneLine.Range("A1").Resize(1, colCount).Value = wsPivot.Range("A1:J1").Value
        wsOneLine.Range("A2").Resize(oneCount, colCount).Value = oneArr
    End If


            ' 2. TLA

    If tlaCount > 0 Then
        wsTLA.Range("A1").Resize(1, colCount).Value = wsPivot.Range("A1:J1").Value
        wsTLA.Range("A2").Resize(tlaCount, colCount).Value = tlaArr
    End If


    Dim dictTLA As Object
    Set dictTLA = CreateObject("Scripting.Dictionary")

    Dim lastTLA As Long

            ' In the TLA tab, categorize all invoices related to the top TLA (ZDEF, ZTLA, ZCCC)
    lastTLA = wsTLA.Cells(wsTLA.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastTLA

        key = wsTLA.Cells(i, "A").Value

        If Not dictTLA.Exists(key) Then
            dictTLA.Add key, True
        End If

    Next i


            ' 3. MULT LINES INVOICE
    If multiCount > 0 Then
        wsMulti.Range("A1").Resize(1, colCount).Value = wsPivot.Range("A1:J1").Value
        wsMulti.Range("A2").Resize(multiCount, colCount).Value = multiArr
    End If

    Dim lastMulti As Long
    lastMulti = wsMulti.Cells(wsMulti.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    For i = lastMulti To 2 Step -1
        key = wsMulti.Cells(i, "A").Value

        If dictTLA.Exists(key) Then
            wsMulti.Rows(i).Delete
        End If
    Next i

    Application.ScreenUpdating = True

    wsOneLine.Columns.AutoFit
    wsTLA.Columns.AutoFit
    wsMulti.Columns.AutoFit
    
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
    Application.CutCopyMode = False
        
    'SAVE AND CLOSE
    wb.Save
    wb.Close

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "The ZSO_SEARCH info has been filtered", vbInformation

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

MsgBox "Error: " & Err.Description, vbExclamation

End Sub