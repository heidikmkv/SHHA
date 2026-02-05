Option Explicit

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
