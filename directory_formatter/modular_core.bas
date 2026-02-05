Option Explicit

'=========================================================
' DIRECTORY PRINTER - CORE MODULE
'=========================================================
' Main entry point: BuildPrintableDirectory
' Orchestrates data import, processing, and output
'=========================================================

'=========================
' USER CONFIG
'=========================
Private Const INPUT_SHEET As String = "PASTE-HERE"
Private Const OUT_BY_NAME As String = "PRINT-BY-NAME"
Private Const OUT_BY_UNIT As String = "PRINT-BY-UNIT"
Private Const OUT_BY_UNIT_TOC As String = "PRINT-BY-UNIT-TOC"

' Page number prefixes
Private Const PAGE_PREFIX_BY_NAME As String = "A"
Private Const PAGE_PREFIX_BY_UNIT As String = "B"
Private Const PAGE_PREFIX_TOC As String = "B"

' Set to True if you want each unit to start on a new page
Private Const START_EACH_UNIT_ON_NEW_PAGE As Boolean = True

Public Sub BuildPrintableDirectory()
    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim wsIn As Worksheet
    Set wsIn = GetSheetOrNothing(wb, INPUT_SHEET)
    If wsIn Is Nothing Then
        MsgBox "Can't find the input sheet named '" & INPUT_SHEET & "'." & vbCrLf & _
               "Create it (or rename your paste sheet) and paste the website export starting at cell A1.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    Dim lastRow As Long, lastCol As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    lastCol = wsIn.Cells(1, wsIn.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 1 Then
        MsgBox "No data found. Paste the website export into '" & INPUT_SHEET & "' starting at A1.", vbExclamation
        GoTo CleanExit
    End If

    Dim data As Variant
    data = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(lastRow, lastCol)).Value2

    ' Required columns by header
    Dim cName As Long, cPhone As Long, cNumber As Long, cStreet As Long
    Dim cStreetUnit As Long, cHoaUnit As Long, cDistrict As Long, cIsMember As Long

    cName = FindHeaderColumn(data, "Directory Names")
    cPhone = FindHeaderColumn(data, "Directory Phone Numbers")
    cNumber = FindHeaderColumn(data, "Number")
    cStreet = FindHeaderColumn(data, "Street")
    cStreetUnit = FindHeaderColumn(data, "Unit")
    cHoaUnit = FindHeaderColumn(data, "HOA Unit")
    cDistrict = FindHeaderColumn(data, "District")
    cIsMember = FindHeaderColumn(data, "Is Member")

    If cName = 0 Or cPhone = 0 Or cNumber = 0 Or cStreet = 0 Or cIsMember = 0 Then
        MsgBox "Missing one or more required headers in row 1." & vbCrLf & _
               "Expected: Directory Names, Directory Phone Numbers, Number, Street, Unit, HOA Unit, District, Is Member", vbExclamation
        GoTo CleanExit
    End If

    ' Unit grouping key: prefer HOA Unit; if missing, fall back to District; else blank
    Dim cUnitGroup As Long
    If cHoaUnit <> 0 Then
        cUnitGroup = cHoaUnit
    ElseIf cDistrict <> 0 Then
        cUnitGroup = cDistrict
    Else
        cUnitGroup = 0
    End If

    ' Output sheets (cleared each run)
    Dim wsByName As Worksheet, wsByUnit As Worksheet, wsTOC As Worksheet
    Set wsByName = GetOrCreateSheet(wb, OUT_BY_NAME, True)
    Set wsByUnit = GetOrCreateSheet(wb, OUT_BY_UNIT, True)
    Set wsTOC = GetOrCreateSheet(wb, OUT_BY_UNIT_TOC, True)

    ' Output buffers
    ' PRINT-BY-NAME: Last, First, Phone, StreetNo, StreetName(+Unit X), Member?
    ' PRINT-BY-UNIT staging:
    '   UnitKey, UnitAlpha, UnitNum, StreetNo, StreetName(+Unit X), Last, First, Phone, Member?, IsResident
    Dim outName() As Variant, outUnit() As Variant
    Dim cap As Long: cap = (lastRow - 1) * 5 + 200
    ReDim outName(1 To cap, 1 To 6)
    ReDim outUnit(1 To cap, 1 To 10)

    Dim rn As Long: rn = 1
    Dim ru As Long: ru = 1

    outName(1, 1) = "Last Name"
    outName(1, 2) = "First Name"
    outName(1, 3) = "Phone #"
    outName(1, 4) = "Street number"
    outName(1, 5) = "Street Name"
    outName(1, 6) = "Member?"

    outUnit(1, 1) = "UnitKey"
    outUnit(1, 2) = "UnitAlpha"
    outUnit(1, 3) = "UnitNum"
    outUnit(1, 4) = "Number"
    outUnit(1, 5) = "Street Name"
    outUnit(1, 6) = "Last Name"
    outUnit(1, 7) = "First Name"
    outUnit(1, 8) = "Phone"
    outUnit(1, 9) = "Member?"
    outUnit(1, 10) = "IsResident"

    Dim r As Long
    For r = 2 To lastRow
        Dim rawNames As String, rawPhones As String
        rawNames = NzStr(data(r, cName))
        rawPhones = NzStr(data(r, cPhone))

        Dim isMemberText As String
        isMemberText = IIf(IsMemberTrue(data(r, cIsMember)), "Yes", "No")

        ' Expand people + phones together; preserve "Resident" flag
        Dim firsts() As String, lasts() As String, phones() As String, isRes() As Boolean
        Dim peopleCount As Long
        peopleCount = ExpandPeoplePhonesWithResidentFlag(rawNames, rawPhones, firsts, lasts, phones, isRes)
        If peopleCount = 0 Then GoTo NextRow

        ' Address fields
        Dim streetNo As String, streetName As String, streetUnit As String, streetNameOut As String
        streetNo = CleanCellTextSingleLine(NzStr(data(r, cNumber)))
        streetName = CleanCellTextSingleLine(NzStr(data(r, cStreet)))
        
        ' Skip rows with no address
        If Len(streetNo) = 0 And Len(streetName) = 0 Then GoTo NextRow

        streetUnit = ""
        If cStreetUnit <> 0 Then streetUnit = CleanCellTextSingleLine(NzStr(data(r, cStreetUnit)))

        streetNameOut = streetName
        If Len(streetUnit) > 0 Then
            streetNameOut = streetNameOut & "  Unit " & streetUnit
        End If

        ' Unit group key + sort keys
        Dim unitKey As String, unitAlpha As String
        Dim unitNum As Long

        If cUnitGroup <> 0 Then
            unitKey = CleanCellTextSingleLine(NzStr(data(r, cUnitGroup)))
        Else
            unitKey = ""
        End If
        unitAlpha = UnitAlphaKey(unitKey)
        unitNum = UnitNumericKey(unitKey)

        Dim iPerson As Long
        For iPerson = 1 To peopleCount
            ' Grow buffers if needed
            If rn + 1 > UBound(outName, 1) Then ReDim Preserve outName(1 To UBound(outName, 1) + 5000, 1 To 6)
            If ru + 1 > UBound(outUnit, 1) Then ReDim Preserve outUnit(1 To UBound(outUnit, 1) + 5000, 1 To 10)

            ' PRINT-BY-NAME: omit Resident
            If Not isRes(iPerson) Then
                rn = rn + 1
                outName(rn, 1) = lasts(iPerson)
                outName(rn, 2) = firsts(iPerson)
                outName(rn, 3) = phones(iPerson)
                outName(rn, 4) = streetNo
                outName(rn, 5) = streetNameOut
                outName(rn, 6) = isMemberText
            End If

            ' PRINT-BY-UNIT: include everyone, including Resident
            ru = ru + 1
            outUnit(ru, 1) = unitKey
            outUnit(ru, 2) = unitAlpha
            outUnit(ru, 3) = unitNum
            outUnit(ru, 4) = streetNo
            outUnit(ru, 5) = streetNameOut
            outUnit(ru, 6) = lasts(iPerson)
            outUnit(ru, 7) = firsts(iPerson)
            outUnit(ru, 8) = phones(iPerson)
            outUnit(ru, 9) = isMemberText
            outUnit(ru, 10) = IIf(isRes(iPerson), "1", "0")
        Next iPerson

NextRow:
    Next r

    '=========================
    ' WRITE PRINT-BY-NAME
    '=========================
    If rn < 2 Then
        wsByName.Range("A1").Value2 = "No entries were produced for PRINT-BY-NAME."
    Else
        wsByName.Range("A1").Resize(rn, 6).Value2 = outName
        SortByName wsByName, rn
        FormatPrintSheet wsByName, "A:F", PAGE_PREFIX_BY_NAME
    End If

    '=========================
    ' WRITE PRINT-BY-UNIT
    '=========================
    If ru < 2 Then
        wsByUnit.Range("A1").Value2 = "No entries were produced for PRINT-BY-UNIT."
        wsTOC.Range("A1").Value2 = "No TOC produced (no units)."
    Else
        wsByUnit.Range("A1").Resize(ru, 10).Value2 = outUnit
        SortUnitStagingByAlphaNumeric wsByUnit, ru
        BuildUnitPrintLayoutAndTOC wsByUnit, ru, wsTOC, PAGE_PREFIX_BY_UNIT, START_EACH_UNIT_ON_NEW_PAGE
        FormatPrintSheet wsByUnit, "A:F", PAGE_PREFIX_BY_UNIT
        FormatTOCSheet wsTOC, PAGE_PREFIX_TOC
    End If

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "BuildPrintableDirectory failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
