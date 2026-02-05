Option Explicit

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
