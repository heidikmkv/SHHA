Option Explicit


'==========================================================
' MODULE: modular_core
'==========================================================


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
Private Const OUT_BY_NAME_2COL As String = "PRINT-BY-NAME-2COL"
Private Const OUT_BY_UNIT_2COL As String = "PRINT-BY-UNIT-2COL"

' Page number prefixes
Private Const PAGE_PREFIX_BY_NAME As String = "A"
Private Const PAGE_PREFIX_BY_UNIT As String = "B"
Private Const PAGE_PREFIX_TOC As String = "B"
Private Const PAGE_PREFIX_BY_NAME_2COL As String = "C"
Private Const PAGE_PREFIX_BY_UNIT_2COL As String = "D"

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

    '=========================
    ' WRITE TWO-COLUMN SHEETS
    '=========================
    Dim wsByName2 As Worksheet, wsByUnit2 As Worksheet
    Set wsByName2 = GetOrCreateSheet(wb, OUT_BY_NAME_2COL, True)
    Set wsByUnit2 = GetOrCreateSheet(wb, OUT_BY_UNIT_2COL, True)

    If rn < 2 Then
        wsByName2.Range("A1").Value2 = "No entries (PRINT-BY-NAME is empty)."
    Else
        BuildTwoColumnByName wsByName, wsByName2, PAGE_PREFIX_BY_NAME_2COL
    End If

    If ru < 2 Then
        wsByUnit2.Range("A1").Value2 = "No entries (PRINT-BY-UNIT is empty)."
    Else
        BuildTwoColumnByUnit wsByUnit, wsByUnit2, PAGE_PREFIX_BY_UNIT_2COL
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

'==========================================================
' MODULE: modular_parsing
'==========================================================


'=========================================================
' DIRECTORY PRINTER - PARSING MODULE
'=========================================================
' Name and phone expansion with Resident flag handling
'=========================================================

'=========================================================
' EXPANSION: PEOPLE + PHONES (WITH RESIDENT FLAG)
' Returns count; arrays are 1..count
'=========================================================
Public Function ExpandPeoplePhonesWithResidentFlag( _
    ByVal rawNames As String, _
    ByVal rawPhones As String, _
    ByRef firstArr() As String, _
    ByRef lastArr() As String, _
    ByRef phoneArr() As String, _
    ByRef isResidentArr() As Boolean _
) As Long

    rawNames = CleanCellTextPreserveLF(rawNames)
    rawPhones = CleanCellTextPreserveLF(rawPhones)

    Dim nameLines() As String, phoneLines() As String
    nameLines = SplitOnNewlinesPreserve(rawNames)
    phoneLines = SplitOnNewlinesPreserve(rawPhones)

    ' Phones: collect nonblank
    Dim phones() As String
    Dim phoneCount As Long: phoneCount = 0
    Dim pl As Variant, ph As String
    For Each pl In phoneLines
        ph = Trim$(CStr(pl))
        If Len(ph) > 0 Then
            phoneCount = phoneCount + 1
            If phoneCount = 1 Then
                ReDim phones(1 To 1)
            Else
                ReDim Preserve phones(1 To phoneCount)
            End If
            phones(phoneCount) = ph
        End If
    Next pl

    ' People: from name lines; expand "First1 & First2 Last"; preserve Resident as a person
    Dim peopleFirst() As String, peopleLast() As String
    Dim peopleIsRes() As Boolean
    Dim peopleCount As Long: peopleCount = 0

    Dim nl As Variant, line As String
    For Each nl In nameLines
        line = Application.WorksheetFunction.Trim(Trim$(CStr(nl)))
        If Len(line) = 0 Then GoTo NextNameLine

        If StrComp(line, "Resident", vbTextCompare) = 0 Then
            AddPersonWithResident peopleFirst, peopleLast, peopleIsRes, peopleCount, "Resident", "", True
            GoTo NextNameLine
        End If

        Dim tokens() As String
        tokens = Split(line, " ")

        If UBound(tokens) >= 1 Then
            Dim lastN As String, firstChunk As String
            lastN = tokens(UBound(tokens))
            firstChunk = JoinTokens(tokens, 0, UBound(tokens) - 1)

            If InStr(1, firstChunk, "&", vbTextCompare) > 0 Then
                Dim parts() As String, fp As Variant, oneFirst As String
                parts = Split(firstChunk, "&")
                For Each fp In parts
                    oneFirst = Application.WorksheetFunction.Trim(CStr(fp))
                    If Len(oneFirst) > 0 Then
                        AddPersonWithResident peopleFirst, peopleLast, peopleIsRes, peopleCount, oneFirst, lastN, False
                    End If
                Next fp
                GoTo NextNameLine
            End If
        End If

        Dim f As String, l As String
        ParseFirstLast line, f, l
        AddPersonWithResident peopleFirst, peopleLast, peopleIsRes, peopleCount, f, l, False

NextNameLine:
    Next nl

    If peopleCount = 0 Then
        ExpandPeoplePhonesWithResidentFlag = 0
        Exit Function
    End If

    ' Assign phones
    Dim assignedPhones() As String
    ReDim assignedPhones(1 To peopleCount)

    Dim i As Long
    If phoneCount = 0 Then
        ' blank
    ElseIf phoneCount = 1 Then
        For i = 1 To peopleCount
            assignedPhones(i) = phones(1)
        Next i
    ElseIf phoneCount >= peopleCount Then
        For i = 1 To peopleCount
            assignedPhones(i) = phones(i)
        Next i
    Else
        For i = 1 To peopleCount
            If i <= phoneCount Then
                assignedPhones(i) = phones(i)
            Else
                assignedPhones(i) = phones(phoneCount)
            End If
        Next i
    End If

    ReDim firstArr(1 To peopleCount)
    ReDim lastArr(1 To peopleCount)
    ReDim phoneArr(1 To peopleCount)
    ReDim isResidentArr(1 To peopleCount)

    For i = 1 To peopleCount
        firstArr(i) = peopleFirst(i)
        lastArr(i) = peopleLast(i)
        phoneArr(i) = assignedPhones(i)
        isResidentArr(i) = peopleIsRes(i)
    Next i

    ExpandPeoplePhonesWithResidentFlag = peopleCount
End Function

Private Sub AddPersonWithResident(ByRef peopleFirst() As String, ByRef peopleLast() As String, ByRef peopleIsRes() As Boolean, ByRef peopleCount As Long, ByVal firstName As String, ByVal lastName As String, ByVal isRes As Boolean)
    peopleCount = peopleCount + 1
    If peopleCount = 1 Then
        ReDim peopleFirst(1 To 1)
        ReDim peopleLast(1 To 1)
        ReDim peopleIsRes(1 To 1)
    Else
        ReDim Preserve peopleFirst(1 To peopleCount)
        ReDim Preserve peopleLast(1 To peopleCount)
        ReDim Preserve peopleIsRes(1 To peopleCount)
    End If
    peopleFirst(peopleCount) = firstName
    peopleLast(peopleCount) = lastName
    peopleIsRes(peopleCount) = isRes
End Sub

'=========================================================
' NAME PARSING
'=========================================================
Private Sub ParseFirstLast(fullName As String, ByRef firstName As String, ByRef lastName As String)
    Dim t As String
    t = Application.WorksheetFunction.Trim(Trim$(fullName))

    If Len(t) = 0 Then
        firstName = ""
        lastName = ""
        Exit Sub
    End If

    Dim tokens() As String
    tokens = Split(t, " ")

    If UBound(tokens) = 0 Then
        firstName = ""
        lastName = tokens(0)
        Exit Sub
    End If

    lastName = tokens(UBound(tokens))
    firstName = JoinTokens(tokens, 0, UBound(tokens) - 1)
End Sub

Private Function JoinTokens(tokens() As String, startIdx As Long, endIdx As Long) As String
    Dim i As Long, s As String
    s = ""
    If endIdx < startIdx Then
        JoinTokens = ""
        Exit Function
    End If
    For i = startIdx To endIdx
        If i = startIdx Then
            s = tokens(i)
        Else
            s = s & " " & tokens(i)
        End If
    Next i
    JoinTokens = Application.WorksheetFunction.Trim(s)
End Function

'==========================================================
' MODULE: modular_sorting
'==========================================================


'=========================================================
' DIRECTORY PRINTER - SORTING MODULE
'=========================================================
' Sort routines for BY-NAME and BY-UNIT layouts
'=========================================================

Public Sub SortByName(ws As Worksheet, lastRow As Long)
    ' A Last, B First, C Phone, D Street#, E Street(+Unit), F Member?
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("D2:D" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("E2:E" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:F" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Public Sub SortUnitStagingByAlphaNumeric(ws As Worksheet, lastRow As Long)
    ' Staging columns:
    ' A UnitKey, B UnitAlpha, C UnitNum, D StreetNo, E StreetName(+Unit), F Last, G First, H Phone, I Member?, J IsResident
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("C2:C" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("D2:D" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("E2:E" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("F2:F" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("G2:G" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:J" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

'==========================================================
' MODULE: modular_layout
'==========================================================


'=========================================================
' DIRECTORY PRINTER - LAYOUT & FORMATTING MODULE
'=========================================================
' Sheet building and formatting functions
'=========================================================

Public Sub BuildUnitPrintLayoutAndTOC(wsUnit As Worksheet, lastDataRow As Long, wsTOC As Worksheet, pagePrefix As String, startNewPageEachUnit As Boolean)
    ' Input staging (wsUnit):
    ' A UnitKey | B UnitAlpha | C UnitNum | D StreetNo | E StreetName(+Unit) | F Last | G First | H Phone | I Member? | J IsResident
    '
    ' Output wsUnit:
    ' Number | Street Name | Last Name | First Name | Phone | Member?
    ' with unit headers merged + shaded
    '
    ' Output wsTOC:
    ' Unit | Page  (e.g., "South Unit 10" | "B-3")

    Dim src As Variant
    src = wsUnit.Range("A1:J" & lastDataRow).Value2

    ' Reset both sheets completely each run
    wsUnit.Cells.Clear
    wsUnit.ResetAllPageBreaks
    wsTOC.Cells.Clear

    ' TOC headers
    wsTOC.Range("A1").Value2 = "Unit"
    wsTOC.Range("B1").Value2 = "Page"
    With wsTOC.Range("A1:B1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Unit sheet column headers
    wsUnit.Range("A1").Value2 = "Number"
    wsUnit.Range("B1").Value2 = "Street Name"
    wsUnit.Range("C1").Value2 = "Last Name"
    wsUnit.Range("D1").Value2 = "First Name"
    wsUnit.Range("E1").Value2 = "Phone"
    wsUnit.Range("F1").Value2 = "Member?"
    With wsUnit.Range("A1:F1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim outR As Long: outR = 1
    Dim tocR As Long: tocR = 1
    Dim pageNum As Long: pageNum = 0

    Dim curUnit As String: curUnit = ""

    Dim r As Long
    For r = 2 To UBound(src, 1)
        Dim unitKey As String
        unitKey = Trim$(NzStr(src(r, 1)))

        If StrComp(unitKey, curUnit, vbTextCompare) <> 0 Then
            curUnit = unitKey
            outR = outR + 1

            ' Optional page break so each unit starts on a new page
            If startNewPageEachUnit Then
                If pageNum > 0 Then
                    wsUnit.HPageBreaks.Add Before:=wsUnit.Rows(outR)
                End If
            End If

            ' Compute "logical" page number for TOC:
            ' If startNewPageEachUnit is true, this will be 1,2,3... per unit reliably.
            pageNum = pageNum + 1

            ' Write TOC row
            tocR = tocR + 1
            wsTOC.Cells(tocR, 1).Value2 = IIf(Len(curUnit) > 0, curUnit, "(No Unit)")
            wsTOC.Cells(tocR, 2).Value2 = pagePrefix & "-" & CStr(pageNum)

            ' Pretty merged unit header row
            With wsUnit.Range(wsUnit.Cells(outR, 1), wsUnit.Cells(outR, 6))
                .Merge
                .Value2 = IIf(Len(curUnit) > 0, curUnit, "(No Unit)")
                .Interior.Color = RGB(235, 235, 235)
                .Font.Bold = True
                .Font.Size = 14
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsUnit.Rows(outR).RowHeight = 22
        End If

        outR = outR + 1
        wsUnit.Cells(outR, 1).Value2 = src(r, 4) ' StreetNo
        wsUnit.Cells(outR, 2).Value2 = src(r, 5) ' StreetName(+Unit X)
        wsUnit.Cells(outR, 3).Value2 = src(r, 6) ' Last
        wsUnit.Cells(outR, 4).Value2 = src(r, 7) ' First
        wsUnit.Cells(outR, 5).Value2 = src(r, 8) ' Phone
        wsUnit.Cells(outR, 6).Value2 = src(r, 9) ' Member?

        With wsUnit.Range(wsUnit.Cells(outR, 1), wsUnit.Cells(outR, 6))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
    Next r

    ' Basic TOC formatting
    With wsTOC.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    wsTOC.Columns("A:B").AutoFit

    ' Repeat header row on prints
    wsUnit.PageSetup.PrintTitleRows = "$1:$1"
End Sub

Public Sub FormatPrintSheet(ws As Worksheet, colRange As String, pagePrefix As String)
    Dim ur As Range
    Set ur = ws.UsedRange

    ur.HorizontalAlignment = xlCenter
    ur.VerticalAlignment = xlCenter
    ur.WrapText = False

    ws.Rows(1).Font.Bold = True
    ws.Columns(colRange).AutoFit
    
    ' Column alignment (directory style)
    ' PRINT-BY-NAME: Last/First left, Phone/Member center, Street# center, Street left
    ' PRINT-BY-UNIT: Number center, Street left, Last/First left, Phone center, Member center
    ws.Columns("A").HorizontalAlignment = xlCenter
    ws.Columns("B").HorizontalAlignment = xlLeft
    ws.Columns("C").HorizontalAlignment = xlLeft
    ws.Columns("D").HorizontalAlignment = xlLeft
    ws.Columns("E").HorizontalAlignment = xlCenter
    ws.Columns("F").HorizontalAlignment = xlCenter
    
    FreezeTopRow ws

    With ws.PageSetup
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .PrintTitleRows = "$1:$1"
        .FirstPageNumber = 1
        .CenterFooter = pagePrefix & "-&P"
    End With
End Sub

Public Sub FormatTOCSheet(ws As Worksheet, pagePrefix As String)
    ws.Columns("A:B").AutoFit
    FreezeTopRow ws

    With ws.PageSetup
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .PrintTitleRows = "$1:$1"
        .FirstPageNumber = 1
        .CenterFooter = pagePrefix & "-&P"
    End With
End Sub

Private Sub FreezeTopRow(ws As Worksheet)
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

'==========================================================
' MODULE: modular_two_column
'==========================================================


'=========================================================
' DIRECTORY PRINTER - TWO-COLUMN LAYOUT MODULE
'=========================================================
' Senior-friendly two-column formatting for compact booklet printing
' Optimized for readability with generous spacing and larger fonts
'=========================================================

' Layout: Column A (rows 1-40) | Gap Col | Column B (rows 1-40)
' Then repeats downward, alternating left/right columns
' Headers every 40 rows with page break logic

Private Const TWO_COL_ROWS_PER_COLUMN As Long = 38   ' Data rows per column before break (38 + 2 header spacing)
Private Const TWO_COL_LEFT_COL_START As Long = 1
Private Const TWO_COL_GAP_COL As Long = 8             ' Column H is spacer
Private Const TWO_COL_RIGHT_COL_START As Long = 9

' Senior-friendly fonts
Private Const FONT_NAME_BODY As String = "Calibri"
Private Const FONT_SIZE_BODY As Double = 12           ' 12pt for readability
Private Const FONT_SIZE_HEADER As Double = 13        ' 13pt for headers
Private Const FONT_SIZE_SECTION As Double = 13       ' Section headers

Public Sub BuildTwoColumnByName(wsNameSrc As Worksheet, wsNameTwoCol As Worksheet, pagePrefix As String)
    ' Convert PRINT-BY-NAME to TWO-COLUMN layout
    ' Source: A Last | B First | C Phone | D StreetNo | E Street | F Member?
    ' Output: Two columns side-by-side, 38 rows per column

    Dim src As Variant
    Dim lastRow As Long
    lastRow = wsNameSrc.Cells(wsNameSrc.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        wsNameTwoCol.Range("A1").Value2 = "No entries (source sheet empty)."
        Exit Sub
    End If
    
    src = wsNameSrc.Range("A1:F" & lastRow).Value2
    
    wsNameTwoCol.Cells.Clear
    wsNameTwoCol.ResetAllPageBreaks
    
    ' Initialize column headers for both columns
    WriteTwoColHeaders wsNameTwoCol
    
    Dim outRow As Long: outRow = 2          ' Start below main headers
    Dim srcRow As Long: srcRow = 2
    Dim colLeft As Long: colLeft = 1        ' Column A
    Dim colRight As Long: colRight = 9      ' Column I
    Dim dataRowInCol As Long: dataRowInCol = 0
    Dim isLeftColumn As Boolean: isLeftColumn = True
    
    ' Populate data rows
    Do While srcRow <= UBound(src, 1)
        If dataRowInCol >= TWO_COL_ROWS_PER_COLUMN Then
            ' Transition to right column or next page set
            If isLeftColumn Then
                isLeftColumn = False
                outRow = 2  ' Reset to row 2 for right column
                dataRowInCol = 0
            Else
                ' Both columns filled, insert page break and move to next set
                wsNameTwoCol.HPageBreaks.Add Before:=wsNameTwoCol.Rows(outRow)
                isLeftColumn = True
                outRow = 2
                dataRowInCol = 0
                ' Re-write headers for new page
                WriteTwoColHeaders wsNameTwoCol
            End If
        End If
        
        Dim startCol As Long
        If isLeftColumn Then
            startCol = colLeft
        Else
            startCol = colRight
        End If
        
        ' Write data row
        wsNameTwoCol.Cells(outRow, startCol + 0).Value2 = src(srcRow, 1)      ' Last
        wsNameTwoCol.Cells(outRow, startCol + 1).Value2 = src(srcRow, 2)      ' First
        wsNameTwoCol.Cells(outRow, startCol + 2).Value2 = src(srcRow, 3)      ' Phone
        wsNameTwoCol.Cells(outRow, startCol + 3).Value2 = src(srcRow, 4)      ' StreetNo
        wsNameTwoCol.Cells(outRow, startCol + 4).Value2 = src(srcRow, 5)      ' Street
        wsNameTwoCol.Cells(outRow, startCol + 5).Value2 = src(srcRow, 6)      ' Member?
        
        ' Format data row
        With wsNameTwoCol.Range(wsNameTwoCol.Cells(outRow, startCol), wsNameTwoCol.Cells(outRow, startCol + 5))
            .Font.Name = FONT_NAME_BODY
            .Font.Size = FONT_SIZE_BODY
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
        
        ' Alignment for each column: Last/First left, Phone/StreetNo center, Street left, Member center
        wsNameTwoCol.Cells(outRow, startCol + 0).HorizontalAlignment = xlLeft
        wsNameTwoCol.Cells(outRow, startCol + 1).HorizontalAlignment = xlLeft
        wsNameTwoCol.Cells(outRow, startCol + 2).HorizontalAlignment = xlCenter
        wsNameTwoCol.Cells(outRow, startCol + 3).HorizontalAlignment = xlCenter
        wsNameTwoCol.Cells(outRow, startCol + 4).HorizontalAlignment = xlLeft
        wsNameTwoCol.Cells(outRow, startCol + 5).HorizontalAlignment = xlCenter
        
        outRow = outRow + 1
        dataRowInCol = dataRowInCol + 1
        srcRow = srcRow + 1
    Loop
    
    ' Finalize formatting
    FormatTwoColSheet wsNameTwoCol, pagePrefix
End Sub

Public Sub BuildTwoColumnByUnit(wsUnitSrc As Worksheet, wsUnitTwoCol As Worksheet, pagePrefix As String)
    ' Convert PRINT-BY-UNIT to TWO-COLUMN layout
    ' Preserves unit section headers (merged rows)
    ' Source: A Number | B Street | C Last | D First | E Phone | F Member? (with unit headers)
    ' Output: Two columns, honoring section breaks

    Dim src As Variant
    Dim lastRow As Long
    lastRow = wsUnitSrc.Cells(wsUnitSrc.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        wsUnitTwoCol.Range("A1").Value2 = "No entries (source sheet empty)."
        Exit Sub
    End If
    
    src = wsUnitSrc.Range("A1:F" & lastRow).Value2
    
    wsUnitTwoCol.Cells.Clear
    wsUnitTwoCol.ResetAllPageBreaks
    
    ' Initialize column headers
    WriteTwoColHeaders wsUnitTwoCol, True ' IsUnitLayout = True
    
    Dim outRow As Long: outRow = 2
    Dim srcRow As Long: srcRow = 2
    Dim colLeft As Long: colLeft = 1
    Dim colRight As Long: colRight = 9
    Dim dataRowInCol As Long: dataRowInCol = 0
    Dim isLeftColumn As Boolean: isLeftColumn = True
    
    ' Populate with unit sections
    Do While srcRow <= UBound(src, 1)
        ' Check if this row is a unit header (merged in source, or has only col A filled)
        Dim isUnitHeader As Boolean
        isUnitHeader = (Len(NzStr(src(srcRow, 1))) > 0 And _
                       Len(NzStr(src(srcRow, 2))) = 0 And _
                       Len(NzStr(src(srcRow, 3))) = 0)
        
        If isUnitHeader Then
            ' Unit header - might need page break
            ' Check if we have room in current column
            If dataRowInCol > 0 Then
                ' Switch columns or pages to keep unit headers readable
                If isLeftColumn Then
                    isLeftColumn = False
                    outRow = 2
                    dataRowInCol = 0
                Else
                    wsUnitTwoCol.HPageBreaks.Add Before:=wsUnitTwoCol.Rows(outRow)
                    isLeftColumn = True
                    outRow = 2
                    dataRowInCol = 0
                    WriteTwoColHeaders wsUnitTwoCol, True
                End If
            End If
            
            Dim startCol As Long
            If isLeftColumn Then
                startCol = colLeft
            Else
                startCol = colRight
            End If
            
            ' Write unit header (merged across 6 columns)
            With wsUnitTwoCol.Range(wsUnitTwoCol.Cells(outRow, startCol), wsUnitTwoCol.Cells(outRow, startCol + 5))
                .Merge
                .Value2 = NzStr(src(srcRow, 1))
                .Interior.Color = RGB(220, 220, 220)
                .Font.Name = FONT_NAME_BODY
                .Font.Size = FONT_SIZE_SECTION
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsUnitTwoCol.Rows(outRow).RowHeight = 18
            
            outRow = outRow + 1
            dataRowInCol = dataRowInCol + 1
        Else
            ' Regular data row
            If dataRowInCol >= TWO_COL_ROWS_PER_COLUMN Then
                If isLeftColumn Then
                    isLeftColumn = False
                    outRow = 2
                    dataRowInCol = 0
                Else
                    wsUnitTwoCol.HPageBreaks.Add Before:=wsUnitTwoCol.Rows(outRow)
                    isLeftColumn = True
                    outRow = 2
                    dataRowInCol = 0
                    WriteTwoColHeaders wsUnitTwoCol, True
                End If
            End If
            
            startCol = IIf(isLeftColumn, colLeft, colRight)
            
            ' Write data
            wsUnitTwoCol.Cells(outRow, startCol + 0).Value2 = src(srcRow, 1)      ' Number
            wsUnitTwoCol.Cells(outRow, startCol + 1).Value2 = src(srcRow, 2)      ' Street
            wsUnitTwoCol.Cells(outRow, startCol + 2).Value2 = src(srcRow, 3)      ' Last
            wsUnitTwoCol.Cells(outRow, startCol + 3).Value2 = src(srcRow, 4)      ' First
            wsUnitTwoCol.Cells(outRow, startCol + 4).Value2 = src(srcRow, 5)      ' Phone
            wsUnitTwoCol.Cells(outRow, startCol + 5).Value2 = src(srcRow, 6)      ' Member?
            
            ' Format
            With wsUnitTwoCol.Range(wsUnitTwoCol.Cells(outRow, startCol), wsUnitTwoCol.Cells(outRow, startCol + 5))
                .Font.Name = FONT_NAME_BODY
                .Font.Size = FONT_SIZE_BODY
                .VerticalAlignment = xlCenter
                .WrapText = False
            End With
            
            wsUnitTwoCol.Cells(outRow, startCol + 0).HorizontalAlignment = xlCenter     ' Number
            wsUnitTwoCol.Cells(outRow, startCol + 1).HorizontalAlignment = xlLeft       ' Street
            wsUnitTwoCol.Cells(outRow, startCol + 2).HorizontalAlignment = xlLeft       ' Last
            wsUnitTwoCol.Cells(outRow, startCol + 3).HorizontalAlignment = xlLeft       ' First
            wsUnitTwoCol.Cells(outRow, startCol + 4).HorizontalAlignment = xlCenter     ' Phone
            wsUnitTwoCol.Cells(outRow, startCol + 5).HorizontalAlignment = xlCenter     ' Member?
            
            outRow = outRow + 1
            dataRowInCol = dataRowInCol + 1
        End If
        
        srcRow = srcRow + 1
    Loop
    
    FormatTwoColSheet wsUnitTwoCol, pagePrefix
End Sub

'=========================================================
' TWO-COLUMN HELPERS
'=========================================================

Public Sub WriteTwoColHeaders(ws As Worksheet, Optional isUnitLayout As Boolean = False)
    ' Write headers for both columns (left and right)
    ' Row 1: Headers for both left and right columns
    ' Gap column (H) left blank
    
    Dim leftStart As Long: leftStart = 1
    Dim rightStart As Long: rightStart = 9
    Dim gapCol As Long: gapCol = 8
    
    If Not isUnitLayout Then
        ' BY-NAME headers
        ' Left column: Last | First | Phone | StreetNo | Street | Member?
        ws.Cells(1, leftStart + 0).Value2 = "Last Name"
        ws.Cells(1, leftStart + 1).Value2 = "First Name"
        ws.Cells(1, leftStart + 2).Value2 = "Phone #"
        ws.Cells(1, leftStart + 3).Value2 = "St. #"
        ws.Cells(1, leftStart + 4).Value2 = "Street"
        ws.Cells(1, leftStart + 5).Value2 = "Mbr?"
        
        ' Right column: same
        ws.Cells(1, rightStart + 0).Value2 = "Last Name"
        ws.Cells(1, rightStart + 1).Value2 = "First Name"
        ws.Cells(1, rightStart + 2).Value2 = "Phone #"
        ws.Cells(1, rightStart + 3).Value2 = "St. #"
        ws.Cells(1, rightStart + 4).Value2 = "Street"
        ws.Cells(1, rightStart + 5).Value2 = "Mbr?"
    Else
        ' BY-UNIT headers
        ' Left column: Number | Street | Last | First | Phone | Member?
        ws.Cells(1, leftStart + 0).Value2 = "Number"
        ws.Cells(1, leftStart + 1).Value2 = "Street"
        ws.Cells(1, leftStart + 2).Value2 = "Last Name"
        ws.Cells(1, leftStart + 3).Value2 = "First Name"
        ws.Cells(1, leftStart + 4).Value2 = "Phone"
        ws.Cells(1, leftStart + 5).Value2 = "Mbr?"
        
        ' Right column: same
        ws.Cells(1, rightStart + 0).Value2 = "Number"
        ws.Cells(1, rightStart + 1).Value2 = "Street"
        ws.Cells(1, rightStart + 2).Value2 = "Last Name"
        ws.Cells(1, rightStart + 3).Value2 = "First Name"
        ws.Cells(1, rightStart + 4).Value2 = "Phone"
        ws.Cells(1, rightStart + 5).Value2 = "Mbr?"
    End If
    
    ' Format headers
    With ws.Range(ws.Cells(1, leftStart), ws.Cells(1, leftStart + 5))
        .Font.Name = FONT_NAME_BODY
        .Font.Size = FONT_SIZE_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    With ws.Range(ws.Cells(1, rightStart), ws.Cells(1, rightStart + 5))
        .Font.Name = FONT_NAME_BODY
        .Font.Size = FONT_SIZE_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ws.Rows(1).RowHeight = 20
    
    ' Set row heights for data area to be generous
    Dim r As Long
    For r = 2 To 200
        ws.Rows(r).RowHeight = 16.5  ' Generous spacing for readability
    Next r
End Sub

Public Sub FormatTwoColSheet(ws As Worksheet, pagePrefix As String)
    ' Column widths for readability
    Dim colWidths() As Variant
    ReDim colWidths(1 To 16)
    
    ' Left column widths
    colWidths(1) = 16  ' Last Name
    colWidths(2) = 16  ' First Name
    colWidths(3) = 13  ' Phone #
    colWidths(4) = 7   ' St. #
    colWidths(5) = 20  ' Street
    colWidths(6) = 5   ' Member?
    colWidths(7) = 1.5 ' Gap
    
    ' Right column widths (same as left)
    colWidths(8) = 16
    colWidths(9) = 16
    colWidths(10) = 13
    colWidths(11) = 7
    colWidths(12) = 20
    colWidths(13) = 5
    colWidths(14) = 1.5
    
    Dim c As Long
    For c = 1 To 14
        ws.Columns(c).ColumnWidth = colWidths(c)
    Next c
    
    ' Page setup for 8.5x11" double-sided booklet
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter       ' 8.5x11"
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        
        .CenterFooter = pagePrefix & "-&P"
        .PrintTitleRows = "$1:$1"
    End With
    
    ' Freeze header row
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

Public Function NzStr(v As Variant) As String
    If IsError(v) Then
        NzStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

'==========================================================
' MODULE: modular_helpers
'==========================================================


'=========================================================
' DIRECTORY PRINTER - HELPERS MODULE
'=========================================================
' Text processing, unit sorting, and sheet utilities
'=========================================================

'=========================================================
' MEMBER FLAG -> Yes/No
'=========================================================
Public Function IsMemberTrue(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        IsMemberTrue = False
        Exit Function
    End If

    If VarType(v) = vbBoolean Then
        IsMemberTrue = (v = True)
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(v))
    If s = "1" Then
        IsMemberTrue = True
        Exit Function
    End If

    On Error Resume Next
    IsMemberTrue = (CDbl(s) = 1)
    On Error GoTo 0
End Function

'=========================================================
' UNIT SORT KEYS (NO ACTIVEX / NO REGEX)
'=========================================================
Public Function UnitNumericKey(unitKey As String) As Long
    Dim n As Variant
    n = RightmostNumber_NoRegex(unitKey)
    If IsNull(n) Then
        UnitNumericKey = 999999
    Else
        UnitNumericKey = CLng(n)
    End If
End Function

Public Function UnitAlphaKey(unitKey As String) As String
    Dim s As String
    s = CleanCellTextSingleLine(unitKey)
    s = RemoveDigits_NoRegex(s)
    UnitAlphaKey = NormalizeSpaces(s)
End Function

Private Function RightmostNumber_NoRegex(ByVal s As String) As Variant
    s = CleanCellTextSingleLine(s)

    Dim i As Long
    Dim digits As String
    digits = ""

    ' Find last digit
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) Like "#" Then Exit For
    Next i

    If i = 0 Then
        RightmostNumber_NoRegex = Null
        Exit Function
    End If

    ' Collect contiguous digit run
    For i = i To 1 Step -1
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            digits = ch & digits
        Else
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then
        RightmostNumber_NoRegex = Null
    Else
        RightmostNumber_NoRegex = CLng(digits)
    End If
End Function

Private Function RemoveDigits_NoRegex(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            out = out & " "
        Else
            out = out & ch
        End If
    Next i
    RemoveDigits_NoRegex = out
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    NormalizeSpaces = Application.WorksheetFunction.Trim(s)
End Function

'=========================================================
' TEXT HELPERS
'=========================================================
Public Function SplitOnNewlinesPreserve(s As String) As String()
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    SplitOnNewlinesPreserve = Split(t, vbLf)
End Function

Public Function CleanCellTextPreserveLF(s As String) As String
    Dim t As String
    t = Replace(s, Chr$(34), "")
    t = Replace(t, ChrW$(160), " ")
    t = Replace(t, vbTab, " ")

    Dim i As Long, ch As String, code As Long, out As String
    out = ""
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        code = AscW(ch)
        If ch = vbLf Or ch = vbCr Then
            out = out & ch
        ElseIf code >= 32 Then
            out = out & ch
        End If
    Next i
    CleanCellTextPreserveLF = Trim$(out)
End Function

Public Function CleanCellTextSingleLine(s As String) As String
    Dim t As String
    t = Replace(s, Chr$(34), "")
    t = Replace(t, ChrW$(160), " ")
    t = Replace(t, vbTab, " ")
    t = Application.WorksheetFunction.Clean(t)
    CleanCellTextSingleLine = Trim$(t)
End Function

Public Function NzStr(v As Variant) As String
    If IsError(v) Then
        NzStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

'=========================================================
' HEADER MATCHING
'=========================================================
Public Function FindHeaderColumn(data As Variant, headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = UBound(data, 2)
    For c = 1 To lastCol
        If NormalizeHeader(NzStr(data(1, c))) = NormalizeHeader(headerName) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

Private Function NormalizeHeader(s As String) As String
    NormalizeHeader = LCase$(Trim$(CleanCellTextSingleLine(s)))
End Function

'=========================================================
' SHEET HELPERS
'=========================================================
Public Function GetSheetOrNothing(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function GetOrCreateSheet(wb As Workbook, sheetName As String, Optional wipe As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    ElseIf wipe Then
        ws.Cells.Clear
    End If

    Set GetOrCreateSheet = ws
End Function
