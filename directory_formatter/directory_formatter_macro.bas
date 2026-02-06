Option Explicit

'=========================================================
' DIRECTORY PRINTER MACRO (COPY/PASTE MODULE)
'=========================================================
' Input sheet:  PASTE-HERE
' Output sheets:
'   1) PRINT-BY-NAME          (omits people listed as "Resident")
'   2) PRINT-BY-UNIT          (includes "Resident")
'
' Page number prefixes (printed in Center Footer):
'   PRINT-BY-NAME : A-1, A-2, ...
'   PRINT-BY-UNIT : B-1, B-2, ...
'
' Sorting of units:
'   - Primary: unit alpha text with digits removed (e.g. "South Unit", "North")
'   - Secondary: rightmost number in unit string (e.g. "South Unit 10" -> 10)
'   - Then: Street number / street / name
'
' Street unit formatting:
'   If Unit column has a value, append to street as "Unit X" (NOT "Unit: X")
'
' Names supported:
'   - "First Last"
'   - "First1 & First2 Last"  -> becomes 2 rows (First1 Last, First2 Last)
'   - newline-separated names -> become multiple rows, duplicating address
' Phones:
'   - newline-separated phones align to people when counts match
'
' Membership:
'   - prints BOTH members and non-members
'   - outputs Member? as Yes / No
'
' Formatting:
'   - button press overwrites prior content/formatting by clearing target sheets
'   - center & middle justification for entries
'   - PRINT-BY-UNIT unit headers are merged, light background, larger centered text
'   - optional: each unit starts on a new page (manual page breaks)
'=========================================================

'=========================
' USER CONFIG
'=========================
Private Const INPUT_SHEET As String = "PASTE-HERE"
Private Const OUT_BY_NAME As String = "PRINT-BY-NAME"
Private Const OUT_BY_UNIT As String = "PRINT-BY-UNIT"
Private Const OUT_BY_UNIT_TOC As String = "PRINT-BY-UNIT-TOC"

Private Const SETTINGS_SHEET As String = "Instructions"
Private Const SETTINGS_FONT_SIZE_CELL As String = "C30"
Private Const SETTINGS_START_NEW_PAGE_CELL As String = "C31"
Private Const SETTINGS_ZEBRA_CELL As String = "C32"
Private Const SETTINGS_PREFIX_BY_NAME_CELL As String = "C33"
Private Const SETTINGS_PREFIX_BY_UNIT_CELL As String = "C34"

' Page number prefixes
Private Const PAGE_PREFIX_BY_NAME As String = "A"
Private Const PAGE_PREFIX_BY_UNIT As String = "B"
Private Const PAGE_PREFIX_TOC As String = "B"

' Set to True if you want each unit to start on a new page in PRINT-BY-UNIT
Private Const START_EACH_UNIT_ON_NEW_PAGE As Boolean = True

Public Sub BuildPrintableDirectory()
    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim wsSettings As Worksheet
    Set wsSettings = GetSheetOrNothing(wb, SETTINGS_SHEET)

    Dim settingFontSize As Double
    settingFontSize = ReadSettingNumber(wsSettings, SETTINGS_FONT_SIZE_CELL, 10)
    Dim settingHeaderSize As Double
    settingHeaderSize = settingFontSize + 1
    Dim settingStartNewPage As Boolean
    settingStartNewPage = ReadSettingYesNo(wsSettings, SETTINGS_START_NEW_PAGE_CELL, START_EACH_UNIT_ON_NEW_PAGE)
    Dim settingZebra As Boolean
    settingZebra = ReadSettingYesNo(wsSettings, SETTINGS_ZEBRA_CELL, False)
    Dim settingPrefixByName As String
    settingPrefixByName = ReadSettingText(wsSettings, SETTINGS_PREFIX_BY_NAME_CELL, PAGE_PREFIX_BY_NAME)
    Dim settingPrefixByUnit As String
    settingPrefixByUnit = ReadSettingText(wsSettings, SETTINGS_PREFIX_BY_UNIT_CELL, PAGE_PREFIX_BY_UNIT)
    Dim settingPrefixTOC As String
    settingPrefixTOC = settingPrefixByUnit

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

    ' Required columns by header (new PASTE-HERE format)
    Dim cFirst As Long, cLast As Long, cNumber As Long, cStreet As Long
    Dim cStreetUnit As Long, cHoaUnit As Long, cDistrict As Long, cPhone As Long
    Dim cIsMember As Long, cListName As Long, cListPhone As Long, cNameSort As Long

    cFirst = FindHeaderColumn(data, "First Name")
    cLast = FindHeaderColumn(data, "Last Name")
    cNumber = FindHeaderColumn(data, "Street Number")
    cStreet = FindHeaderColumn(data, "Street Name")
    cStreetUnit = FindHeaderColumn(data, "Street Unit")
    cDistrict = FindHeaderColumn(data, "District")
    cHoaUnit = FindHeaderColumn(data, "HOA Unit")
    cPhone = FindHeaderColumn(data, "Phone")
    cIsMember = FindHeaderColumn(data, "Is Member")
    cListName = FindHeaderColumn(data, "List Name in Directory")
    cListPhone = FindHeaderColumn(data, "List Phone in Directory")
    cNameSort = FindHeaderColumn(data, "Name Sort Order")

    If cFirst = 0 Or cLast = 0 Or cNumber = 0 Or cStreet = 0 Or cIsMember = 0 Or cListName = 0 Or cListPhone = 0 Then
        MsgBox "Missing one or more required headers in row 1." & vbCrLf & _
               "Expected: First Name, Last Name, Street Number, Street Name, Street Unit, HOA Unit, District, Phone, Is Member, List Name in Directory, List Phone in Directory, Name Sort Order", vbExclamation
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
    ' PRINT-BY-NAME: Last, First, Phone, StreetNo, StreetName(+Unit X), Member?, NameSort
    ' PRINT-BY-UNIT staging:
    '   UnitKey, UnitAlpha, UnitNum, StreetNo, StreetName(+Unit X), Last, First, Phone, Member?, IsResident
    Dim outName() As Variant, outUnit() As Variant
    Dim cap As Long: cap = (lastRow - 1) * 5 + 200
    ReDim outName(1 To cap, 1 To 7)
    ReDim outUnit(1 To cap, 1 To 10)

    Dim rn As Long: rn = 1
    Dim ru As Long: ru = 1

    outName(1, 1) = "Last Name"
    outName(1, 2) = "First Name"
    outName(1, 3) = "Phone #"
    outName(1, 4) = "Street number"
    outName(1, 5) = "Street Name"
    outName(1, 6) = "Member?"
    outName(1, 7) = "NameSort"

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
        Dim firstName As String, lastName As String, phoneVal As String
        firstName = NzStr(data(r, cFirst))
        lastName = NzStr(data(r, cLast))
        phoneVal = NzStr(data(r, cPhone))

        Dim listName As Boolean, listPhone As Boolean
        listName = IsMemberTrue(data(r, cListName))
        listPhone = IsMemberTrue(data(r, cListPhone))

        Dim isMemberText As String
        isMemberText = IIf(IsMemberTrue(data(r, cIsMember)), "Yes", "No")

        Dim displayFirst As String, displayLast As String, displayPhone As String
        If listName Then
            displayFirst = firstName
            displayLast = lastName
        Else
            displayFirst = "Resident"
            displayLast = ""
        End If

        If listName And listPhone Then
            displayPhone = phoneVal
        Else
            displayPhone = ""
        End If

        ' Address fields
        Dim streetNo As String, streetName As String, streetUnit As String, streetNameOut As String
        streetNo = CleanCellTextSingleLine(NzStr(data(r, cNumber)))
        streetName = CleanCellTextSingleLine(NzStr(data(r, cStreet)))
        
        ' Skip rows with no address (bad export rows)
        If Len(streetNo) = 0 And Len(streetName) = 0 Then GoTo NextRow

        streetUnit = ""
        If cStreetUnit <> 0 Then streetUnit = CleanCellTextSingleLine(NzStr(data(r, cStreetUnit)))

        streetNameOut = streetName
        If Len(streetUnit) > 0 Then
            streetNo = streetNo & "-" & streetUnit
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

        Dim nameSortVal As Long
        If cNameSort <> 0 Then
            If Len(Trim$(NzStr(data(r, cNameSort)))) = 0 Then
                nameSortVal = 999999
            Else
                nameSortVal = CLng(Val(NzStr(data(r, cNameSort))))
            End If
        Else
            nameSortVal = 999999
        End If

        ' Grow buffers if needed
        If rn + 1 > UBound(outName, 1) Then ReDim Preserve outName(1 To UBound(outName, 1) + 5000, 1 To 7)
        If ru + 1 > UBound(outUnit, 1) Then ReDim Preserve outUnit(1 To UBound(outUnit, 1) + 5000, 1 To 10)

        ' PRINT-BY-NAME: omit Resident
        If listName Then
            rn = rn + 1
            outName(rn, 1) = displayLast
            outName(rn, 2) = displayFirst
            outName(rn, 3) = displayPhone
            outName(rn, 4) = streetNo
            outName(rn, 5) = streetNameOut
            outName(rn, 6) = isMemberText
            outName(rn, 7) = nameSortVal
        End If

        ' PRINT-BY-UNIT: include everyone, including Resident
        ru = ru + 1
        outUnit(ru, 1) = unitKey
        outUnit(ru, 2) = unitAlpha
        outUnit(ru, 3) = unitNum
        outUnit(ru, 4) = streetNo
        outUnit(ru, 5) = streetNameOut
        outUnit(ru, 6) = displayLast
        outUnit(ru, 7) = displayFirst
        outUnit(ru, 8) = displayPhone
        outUnit(ru, 9) = isMemberText
        outUnit(ru, 10) = IIf(listName, "0", "1")

NextRow:
    Next r

    '=========================
    ' WRITE PRINT-BY-NAME
    '=========================
    If rn < 2 Then
        wsByName.Range("A1").Value2 = "No entries were produced for PRINT-BY-NAME."
    Else
        wsByName.Range("A1").Resize(rn, 7).Value2 = outName
        SortByName wsByName, rn
        wsByName.Columns("G").Hidden = True
        FormatPrintSheet wsByName, "A:F", settingPrefixByName, settingFontSize, settingHeaderSize, settingZebra
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
        BuildSimpleTOC wsByUnit, ru, wsTOC, settingFontSize, settingHeaderSize
        BuildUnitPrintLayout wsByUnit, ru, settingPrefixByUnit, settingStartNewPage, settingFontSize, settingHeaderSize, settingZebra
        FormatPrintSheet wsByUnit, "A:F", settingPrefixByUnit, settingFontSize, settingHeaderSize, settingZebra
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

'=========================================================
' MEMBER FLAG -> Yes/No
'=========================================================
Private Function IsMemberTrue(v As Variant) As Boolean
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
' SORTING
'=========================================================
Private Sub SortByName(ws As Worksheet, lastRow As Long)
    ' A Last, B First, C Phone, D Street#, E Street(+Unit), F Member?
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("G2:G" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("D2:D" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("E2:E" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub SortUnitStagingByAlphaNumeric(ws As Worksheet, lastRow As Long)
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

'=========================================================
' SIMPLE TOC (UNIT LIST ONLY)
'=========================================================
Private Sub BuildSimpleTOC(wsUnit As Worksheet, lastDataRow As Long, wsTOC As Worksheet, baseFontSize As Double, headerFontSize As Double)
    ' Extract unique units from sorted staging data and create a simple list
    ' Page numbers can be added manually by the user
    
    Dim src As Variant
    src = wsUnit.Range("A1:J" & lastDataRow).Value2
    
    wsTOC.Cells.Clear
    
    ' Headers
    wsTOC.Range("A1").Value2 = "Unit"
    wsTOC.Range("B1").Value2 = "Page"
    With wsTOC.Range("A1:B1")
        .Font.Bold = True
        .Font.Size = headerFontSize
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Dim tocR As Long: tocR = 1
    Dim curUnit As String: curUnit = ""
    
    Dim r As Long
    For r = 2 To UBound(src, 1)
        Dim unitKey As String
        unitKey = Trim$(NzStr(src(r, 1)))
        
        If StrComp(unitKey, curUnit, vbTextCompare) <> 0 Then
            curUnit = unitKey
            tocR = tocR + 1
            wsTOC.Cells(tocR, 1).Value2 = IIf(Len(curUnit) > 0, curUnit, "(No Unit)")
            wsTOC.Cells(tocR, 2).Value2 = ""  ' Empty - for manual entry
        End If
    Next r
    
    ' Format TOC
    With wsTOC.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Size = baseFontSize
    End With
    wsTOC.Rows(1).Font.Size = headerFontSize
    wsTOC.Columns("A:B").AutoFit
    
    ' Freeze top row
    wsTOC.Activate
    wsTOC.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

'=========================================================
' UNIT PRINT LAYOUT
'=========================================================
Private Sub BuildUnitPrintLayout(wsUnit As Worksheet, lastDataRow As Long, pagePrefix As String, startNewPageEachUnit As Boolean, baseFontSize As Double, headerFontSize As Double, useZebra As Boolean)
    ' Input staging (wsUnit):
    ' A UnitKey | B UnitAlpha | C UnitNum | D StreetNo | E StreetName(+Unit) | F Last | G First | H Phone | I Member? | J IsResident
    '
    ' Output wsUnit:
    ' Number | Street Name | Last Name | First Name | Phone | Member?
    ' with unit headers merged + shaded

    Dim src As Variant
    src = wsUnit.Range("A1:J" & lastDataRow).Value2

    ' Reset sheet completely each run
    wsUnit.Cells.Clear
    wsUnit.ResetAllPageBreaks

    ' Unit sheet column headers
    wsUnit.Range("A1").Value2 = "Number"
    wsUnit.Range("B1").Value2 = "Street Name"
    wsUnit.Range("C1").Value2 = "Last Name"
    wsUnit.Range("D1").Value2 = "First Name"
    wsUnit.Range("E1").Value2 = "Phone"
    wsUnit.Range("F1").Value2 = "Member?"
    With wsUnit.Range("A1:F1")
        .Font.Bold = True
        .Font.Size = headerFontSize
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim outR As Long: outR = 1
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
                If outR > 2 Then
                    wsUnit.HPageBreaks.Add Before:=wsUnit.Rows(outR)
                End If
            End If

            ' Pretty merged unit header row
            With wsUnit.Range(wsUnit.Cells(outR, 1), wsUnit.Cells(outR, 6))
                .Merge
                .Value2 = IIf(Len(curUnit) > 0, curUnit, "(No Unit)")
                .Interior.Color = RGB(235, 235, 235)
                .Font.Bold = True
                .Font.Size = headerFontSize
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsUnit.Rows(outR).RowHeight = RowHeightForFont(headerFontSize)
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
            .Font.Size = baseFontSize
        End With
        wsUnit.Rows(outR).RowHeight = RowHeightForFont(baseFontSize)
    Next r

    If useZebra Then
        ApplyZebraStripes wsUnit, 2, wsUnit.UsedRange.Rows.Count, 1, 6
    End If

    ' Repeat header row on prints
    wsUnit.PageSetup.PrintTitleRows = "$1:$1"
End Sub

'=========================================================
' PRINT FORMATTING (page prefix A-/B-)
'=========================================================
Private Sub FormatPrintSheet(ws As Worksheet, colRange As String, pagePrefix As String, baseFontSize As Double, headerFontSize As Double, useZebra As Boolean)
    Dim ur As Range
    Set ur = ws.UsedRange

    ur.HorizontalAlignment = xlCenter
    ur.VerticalAlignment = xlCenter
    ur.WrapText = False
    ur.Font.Size = baseFontSize

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Size = headerFontSize
    ws.Columns(colRange).AutoFit
    ' Column alignment (directory style)
    If ws.Name = OUT_BY_NAME Then
        ws.Columns("A").HorizontalAlignment = xlLeft
        ws.Columns("B").HorizontalAlignment = xlLeft
        ws.Columns("C").HorizontalAlignment = xlCenter
        ws.Columns("D").HorizontalAlignment = xlCenter
        ws.Columns("E").HorizontalAlignment = xlLeft
        ws.Columns("F").HorizontalAlignment = xlCenter
    Else
        ' PRINT-BY-UNIT
        ws.Columns("A").HorizontalAlignment = xlCenter
        ws.Columns("B").HorizontalAlignment = xlLeft
        ws.Columns("C").HorizontalAlignment = xlLeft
        ws.Columns("D").HorizontalAlignment = xlLeft
        ws.Columns("E").HorizontalAlignment = xlCenter
        ws.Columns("F").HorizontalAlignment = xlCenter
    End If
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

    If ws.Name = OUT_BY_NAME Then
        ws.Rows("2:" & ws.UsedRange.Rows.Count).RowHeight = RowHeightForFont(baseFontSize)
    End If

    If useZebra Then
        ApplyZebraStripes ws, 2, ws.UsedRange.Rows.Count, 1, 6
    End If
End Sub

Private Sub FreezeTopRow(ws As Worksheet)
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

'=========================================================
' EXPANSION: PEOPLE + PHONES (WITH RESIDENT FLAG)
' Returns count; arrays are 1..count
'=========================================================
Private Function ExpandPeoplePhonesWithResidentFlag( _
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
' UNIT SORT KEYS (NO ACTIVEX / NO REGEX)
'=========================================================
Private Function UnitNumericKey(unitKey As String) As Long
    Dim n As Variant
    n = FirstNumber_NoRegex(unitKey)
    If IsNull(n) Then
        UnitNumericKey = 999999
    Else
        UnitNumericKey = CLng(n)
    End If
End Function

Private Function UnitAlphaKey(unitKey As String) As String
    Dim s As String
    Dim firstDigitPos As Long
    s = CleanCellTextSingleLine(unitKey)
    firstDigitPos = FirstDigitPosition(s)
    If firstDigitPos > 1 Then
        UnitAlphaKey = NormalizeSpaces(Left$(s, firstDigitPos - 1))
    ElseIf firstDigitPos = 1 Then
        UnitAlphaKey = ""
    Else
        UnitAlphaKey = NormalizeSpaces(s)
    End If
End Function

Private Function FirstDigitPosition(ByVal s As String) As Long
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            FirstDigitPosition = i
            Exit Function
        End If
    Next i
    FirstDigitPosition = 0
End Function

Private Function FirstNumber_NoRegex(ByVal s As String) As Variant
    s = CleanCellTextSingleLine(s)

    Dim i As Long
    Dim digits As String
    digits = ""

    i = FirstDigitPosition(s)
    If i = 0 Then
        FirstNumber_NoRegex = Null
        Exit Function
    End If

    Dim ch As String
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            digits = digits & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    If Len(digits) = 0 Then
        FirstNumber_NoRegex = Null
    Else
        FirstNumber_NoRegex = CLng(digits)
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

'=========================================================
' TEXT HELPERS
'=========================================================
Private Function SplitOnNewlinesPreserve(s As String) As String()
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    SplitOnNewlinesPreserve = Split(t, vbLf)
End Function

Private Function CleanCellTextPreserveLF(s As String) As String
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

Private Function CleanCellTextSingleLine(s As String) As String
    Dim t As String
    t = Replace(s, Chr$(34), "")
    t = Replace(t, ChrW$(160), " ")
    t = Replace(t, vbTab, " ")
    t = Application.WorksheetFunction.Clean(t)
    CleanCellTextSingleLine = Trim$(t)
End Function

Private Function NzStr(v As Variant) As String
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
Private Function FindHeaderColumn(data As Variant, headerName As String) As Long
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
Private Function GetSheetOrNothing(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(wb As Workbook, sheetName As String, Optional wipe As Boolean = False) As Worksheet
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

'=========================================================
' SETTINGS + FORMATTING HELPERS
'=========================================================
Private Function ReadSettingText(ws As Worksheet, cellAddress As String, defaultValue As String) As String
    If ws Is Nothing Then
        ReadSettingText = defaultValue
        Exit Function
    End If
    Dim v As String
    v = Trim$(NzStr(ws.Range(cellAddress).Value2))
    If Len(v) = 0 Then
        ReadSettingText = defaultValue
    Else
        ReadSettingText = v
    End If
End Function

Private Function ReadSettingNumber(ws As Worksheet, cellAddress As String, defaultValue As Double) As Double
    If ws Is Nothing Then
        ReadSettingNumber = defaultValue
        Exit Function
    End If
    Dim v As String
    v = Trim$(NzStr(ws.Range(cellAddress).Value2))
    If Len(v) = 0 Then
        ReadSettingNumber = defaultValue
    Else
        ReadSettingNumber = CDbl(Val(v))
    End If
End Function

Private Function ReadSettingYesNo(ws As Worksheet, cellAddress As String, defaultValue As Boolean) As Boolean
    If ws Is Nothing Then
        ReadSettingYesNo = defaultValue
        Exit Function
    End If
    Dim v As String
    v = LCase$(Trim$(NzStr(ws.Range(cellAddress).Value2)))
    If v = "yes" Or v = "y" Or v = "true" Or v = "1" Then
        ReadSettingYesNo = True
    ElseIf v = "no" Or v = "n" Or v = "false" Or v = "0" Then
        ReadSettingYesNo = False
    Else
        ReadSettingYesNo = defaultValue
    End If
End Function

Private Sub ApplyZebraStripes(ws As Worksheet, startRow As Long, endRow As Long, startCol As Long, endCol As Long)
    Dim r As Long
    Dim shade As Boolean
    shade = False

    For r = startRow To endRow
        If ws.Cells(r, startCol).MergeCells Then
            shade = False
        Else
            shade = Not shade
            If shade Then
                ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Interior.Color = RGB(245, 245, 245)
            Else
                ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Interior.ColorIndex = xlColorIndexNone
            End If
        End If
    Next r
End Sub

Private Function RowHeightForFont(baseFontSize As Double) As Double
    RowHeightForFont = baseFontSize + 4
End Function