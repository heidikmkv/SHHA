Option Explicit

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
