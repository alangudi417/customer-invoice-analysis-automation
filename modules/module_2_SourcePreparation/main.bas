Option Explicit


Sub sources_filter()

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsPivot As Worksheet
    Dim ptSheet As Worksheet

    Dim PivotCache As PivotCache
    Dim PivotTable As PivotTable
    Dim PivotRange As Range
    Dim DataRange As Range

    Dim FilePath As String
    Dim fileName As String

    Dim TodayDate As String
    Dim ans As VbMsgBoxResult

    TodayDate = Format(Date, "mm-dd-yy")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

        ' Ask before proceeding
    ans = MsgBox("Do you have the files (VBRP, TAR001, UDM_Notes and ZREPRINT)ready?" & vbCrLf & vbCrLf & _
                 "Select Yes to continue.", vbYesNo + vbQuestion, "Confirmation")
    If ans = vbNo Then GoTo ErrorHandler


    '====================================================
' Phase 1: UDM NOTES
    '====================================================

    'Open UDM_Notes file
    FilePath = "/sample_data/input/ZSO_Search/UDM Comments/"
    fileName = "UDM_Notes " & Format(Date, "mm-dd-yy") & ".xlsx"

    Set wb = Workbooks.Open(FilePath & fileName)

    'Change text to column
    With wb.Sheets(1).Columns("C")
        .TextToColumns Destination:=.Cells(1, 1), _
                       DataType:=xlDelimited
    End With

    'Save changes and close
    wb.Save
    wb.Close



    '====================================================
' Phase 2: TAR001
    '====================================================

    'Open TARIFF001 file
    FilePath = "/sample_data/input/ZSO_Search/TAR001/"
    fileName = "TAR001 " & TodayDate & ".xlsx"

    Set wb = Workbooks.Open(FilePath & fileName)
    Set wsSource = wb.Sheets("Customer Billings")

    On Error Resume Next
    wb.Sheets("Pivot").Delete
    On Error GoTo ErrorHandler

    ' Create new tab "Pivot"
    Set wsPivot = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPivot.Name = "Pivot"

    'Change text to column
    With wsSource.Columns("Z")
        .TextToColumns Destination:=.Cells(1, 1), _
                       DataType:=xlDelimited
    End With


    Set DataRange = wsSource.UsedRange

    'Prepare the pivot
    Set ptSheet = wb.Sheets.Add
    ptSheet.Name = "PivotTemp"

    Set PivotCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=DataRange)

    Set PivotTable = PivotCache.CreatePivotTable( _
        TableDestination:=ptSheet.Range("A3"), _
        TableName:="TariffPivot")

    With PivotTable
        .PivotFields("Billing Document").Orientation = xlRowField
    
        With .PivotFields("Total stacked tariff charge")
            .Orientation = xlDataField
            .Function = xlSum
        End With

        .RowGrand = False
        .ColumnGrand = False
    End With

    ptSheet.UsedRange.Copy

    'Paste values from the pivot table in the tab (Pivot)
    wsPivot.Range("A1").PasteSpecial xlPasteValues
    wsPivot.Columns("B").NumberFormat = "$#,##0.00"
    wsPivot.Columns.AutoFit

    Application.CutCopyMode = False

    ptSheet.Delete

    ' Save changes and close
    wb.Save
    wb.Close



    '====================================================
' Phase 3: VBRP
    '====================================================

    'Open VBRP file
    FilePath = "/sample_data/input/ZSO_Search/VBRP/"
    fileName = "VBRP " & TodayDate & ".xlsx"

    Set wb = Workbooks.Open(FilePath & fileName)
    Set wsSource = wb.Sheets("Sheet1")

    On Error Resume Next
    wb.Sheets("Pivot").Delete
    On Error GoTo ErrorHandler

    ' Create new tab "Pivot"
    Set wsPivot = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPivot.Name = "Pivot"

    'Change text to column
    With wsSource.Columns("A")
        .TextToColumns Destination:=.Cells(1, 1), _
                       DataType:=xlDelimited
    End With


    Set DataRange = wsSource.UsedRange

    'Prepare the pivot
    Set ptSheet = wb.Sheets.Add
    ptSheet.Name = "PivotTemp"

    Set PivotCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=DataRange)

    Set PivotTable = PivotCache.CreatePivotTable( _
        TableDestination:=ptSheet.Range("A3"), _
        TableName:="TaxPivot")

    With PivotTable
        .PivotFields("Billing Document").Orientation = xlRowField

        With .PivotFields("Tax amount")
            .Orientation = xlDataField
            .Function = xlSum
        End With

        .RowGrand = False
        .ColumnGrand = False

    End With

    ptSheet.UsedRange.Copy

    'Paste values from the pivot table in the tab (Pivot)
    wsPivot.Range("A1").PasteSpecial xlPasteValues
    wsPivot.Columns("B").NumberFormat = "$#,##0.00"
    wsPivot.Columns.AutoFit

    Application.CutCopyMode = False

    ptSheet.Delete

    'Save changes and close
    wb.Save
    wb.Close


    '====================================================
' Phase 4: ZREPRINT
    '====================================================

    'Open ZREPRINT file
    FilePath = "/sample_data/input/ZSO_Search/ZREPRINT/"
    fileName = "ZREPRINT " & Format(Date, "mm-dd-yy") & ".xlsx"

    Set wb = Workbooks.Open(FilePath & fileName)
    Set wsSource = wb.Sheets("Sheet1")

    On Error Resume Next
    wb.Sheets("Pivot").Delete
    On Error GoTo ErrorHandler

    ' Create new tab "Pivot"
    Set wsPivot = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPivot.Name = "Pivot"

    'Change text to column
    With wsSource.Columns("C")
        .TextToColumns Destination:=.Cells(1, 1), _
                       DataType:=xlDelimited
    End With

    Set DataRange = wsSource.UsedRange

    'Prepare the pivot
    Set ptSheet = wb.Sheets.Add
    ptSheet.Name = "PivotTemp"

    Set PivotCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=DataRange)

    Set PivotTable = PivotCache.CreatePivotTable( _
        TableDestination:=ptSheet.Range("A3"), _
        TableName:="ReprintPivot")

    With PivotTable

        .PivotFields("Document Number").Orientation = xlRowField
        .PivotFields("Doc. Typ Descr.").Orientation = xlRowField
        .PivotFields("Amount").Orientation = xlRowField

        .RowAxisLayout xlTabularRow
        
        .RowGrand = False
        .ColumnGrand = False
    
    End With

    Dim pField As PivotField

    For Each pField In PivotTable.RowFields
        pField.Subtotals = Array(False, False, False, False, False, False, _
                                 False, False, False, False, False, False)
    Next pField

    ptSheet.UsedRange.Copy
    
    'Paste values from the pivot table in the tab (Pivot)
    wsPivot.Range("A1").PasteSpecial xlPasteValues
    wsPivot.Columns("C").NumberFormat = "$#,##0.00"
    wsPivot.Columns.AutoFit

    Application.CutCopyMode = False
    ptSheet.Delete

    'Save changes and close
    wb.Save
    wb.Close



' Closing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "The files have been successfully filtered", vbInformation

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

MsgBox "Error: " & Err.Description, vbExclamation

End Sub