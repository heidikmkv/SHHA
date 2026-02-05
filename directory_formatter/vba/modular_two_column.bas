Option Explicit

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

Private Sub WriteTwoColHeaders(ws As Worksheet, Optional isUnitLayout As Boolean = False)
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

Private Sub FormatTwoColSheet(ws As Worksheet, pagePrefix As String)
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

Private Function NzStr(v As Variant) As String
    If IsError(v) Then
        NzStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function
