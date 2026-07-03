' DATABASE STRUCTURE MACRO
Public Sub database_integration_board()

    On Error GoTo ErrHandler

    '-----------------------------
    ' App state backup
    '-----------------------------
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '-----------------------------
    ' Variables
    '-----------------------------
    Dim fDate As String
    Dim basePath As String

    Dim wbDB As Workbook
    Dim wbZSO As Workbook
    Dim wbUDM As Workbook
    Dim wbNotes As Workbook
    Dim wbVBRP As Workbook
    Dim wbTAR As Workbook
    Dim wbZREP As Workbook

    Dim wsDB1 As Worksheet, wsDB2 As Worksheet, wsDB3 As Worksheet
    Dim wsZSO1 As Worksheet, wsZSO2 As Worksheet, wsZSO3 As Worksheet
    Dim wsUDM As Worksheet, wsNotes As Worksheet
    Dim wsVBRP As Worksheet, wsTAR As Worksheet, wsZREP As Worksheet

    Dim dictUDM As Object
    Dim dictNotes As Object
    Dim dictVBRP As Object
    Dim dictTAR As Object
    Dim dictZREP As Object


    ' Date and path
    fDate = Format(Date, "mm-dd-yy")

    If Right$("/sample_data/input/ZSO_Search/", 1) = "/" Then
        basePath = "/sample_data/input/ZSO_Search/"
    Else
        basePath = "/sample_data/input/ZSO_Search/" & "/"
    End If


    ' OPEN WORKBOOKS
    Set wbDB = OpenWBSafe("/sample_data/output/Database/Database " & fDate & ".xlsm")
    
    Set wbZSO = OpenWBSafe(basePath & "ZSO_Search/ZSO_SEARCH " & fDate & ".xlsx")
    Set wbUDM = OpenWBSafe(basePath & "Boards/Board/UDM_Dispute " & fDate & ".xlsm")
    Set wbNotes = OpenWBSafe(basePath & "UDM Comments/UDM_Notes " & fDate & ".xlsx")
    Set wbVBRP = OpenWBSafe(basePath & "VBRP/VBRP " & fDate & ".xlsx")
    Set wbTAR = OpenWBSafe(basePath & "TAR001/TAR001 " & fDate & ".xlsx")
    Set wbZREP = OpenWBSafe(basePath & "ZREPRINT/ZREPRINT " & fDate & ".xlsx")


    ' VALIDATE SHEETS
    Set wsDB1 = GetSheetSafe(wbDB, "One Line")
    Set wsDB2 = GetSheetSafe(wbDB, "Mult Line")
    Set wsDB3 = GetSheetSafe(wbDB, "TLA")

    Set wsZSO1 = GetSheetSafe(wbZSO, "One Line Invoice")
    Set wsZSO2 = GetSheetSafe(wbZSO, "Mult Lines Invoice")
    Set wsZSO3 = GetSheetSafe(wbZSO, "TLA")

    Set wsUDM = GetSheetSafe(wbUDM, "Open")
    Set wsNotes = GetSheetSafe(wbNotes, "Sheet1")

    Set wsVBRP = GetSheetSafe(wbVBRP, "Pivot")
    Set wsTAR = GetSheetSafe(wbTAR, "Pivot")
    Set wsZREP = GetSheetSafe(wbZREP, "Pivot")


    ' BUILD DICTIONARIES
    Set dictUDM = BuildUDMDictionarySafe(wsUDM)
    Set dictNotes = BuildNotesDictionarySafe(wsNotes)
    Set dictVBRP = BuildSimpleDictionarySafe(wsVBRP, "A", "B")
    Set dictTAR = BuildSimpleDictionarySafe(wsTAR, "A", "B")
    Set dictZREP = BuildSimpleDictionarySafe(wsZREP, "A", "C")

    ValidateDictionary dictUDM, "dictUDM"
    ValidateDictionary dictNotes, "dictNotes"
    ValidateDictionary dictVBRP, "dictVBRP"
    ValidateDictionary dictTAR, "dictTAR"
    ValidateDictionary dictZREP, "dictZREP"


    ' PROCESS SHEETS
    ProcessSheetSafe wsDB1, wsZSO1, dictUDM, dictNotes, dictVBRP, dictTAR, dictZREP, False
    ProcessSheetSafe wsDB2, wsZSO2, dictUDM, dictNotes, dictVBRP, dictTAR, dictZREP, False
    ProcessSheetSafe wsDB3, wsZSO3, dictUDM, dictNotes, dictVBRP, dictTAR, dictZREP, True


''Recently added''
    ' APPLY ACTIVE FORMULAS TO MULT LINE AND TLA
    ApplyWarboardFormulas wsDB2, fDate   ' Mult Line
    ApplyWarboardFormulas wsDB3, fDate   ' TLA

    wsDB2.Calculate
    wsDB3.Calculate


    ' FORMAT COLUMN AB - REMOVE WRAP TEXT
    wsDB1.Columns("AB").WrapText = False
    wsDB2.Columns("AB").WrapText = False
    wsDB3.Columns("AB").WrapText = False


    ' SAVE & CLEANUP
    wbDB.Save

CleanExit:
    On Error Resume Next

    If Not wbZSO Is Nothing Then wbZSO.Close False
    If Not wbUDM Is Nothing Then wbUDM.Close False
    If Not wbNotes Is Nothing Then wbNotes.Close False
    If Not wbVBRP Is Nothing Then wbVBRP.Close False
    If Not wbTAR Is Nothing Then wbTAR.Close False
    If Not wbZREP Is Nothing Then wbZREP.Close False
    If Not wbDB Is Nothing Then wbDB.Close False

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation

    On Error GoTo 0

    MsgBox "Macro completed successfully.", vbInformation
    Exit Sub

ErrHandler:
    Dim errMsg As String
    errMsg = "ERROR " & Err.Number & vbCrLf & _
             "Description: " & Err.Description & vbCrLf & _
             "Source: " & Err.Source
    MsgBox errMsg, vbCritical, "database_structure failed"

    Resume CleanExit

End Sub