Option Explicit

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
