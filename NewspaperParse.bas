Attribute VB_Name = "NewspaperParse"
Option Explicit

'Change Log:
'112903 rjk: Created to extract intelligence from the newspaper

Enum enumNewspaperSections
    NP_UNKNOWN
    NP_HEADER
    NP_TELECOMMUNICATIONS
End Enum

Public ActivityCounts() As Integer ' 2d array  first index specifies hour that activity occurred, second specifies country relations doing the activity

Public Sub ParseNewsPaper(strLine As String)
Static npSection As enumNewspaperSections
Dim strOrigin As String
Dim strDestination As String
Dim strTemp As String
Dim Origin As Integer
Dim Destination As Integer

If Len(strLine) = 0 Then
    Exit Sub
End If

If InStr(strLine, "-=[  EMPIRE NEWS  ]=-") > 0 Then
    npSection = NP_HEADER
    ReDim ActivityCounts(0 To 24, 0 To Nations.Count)
    Nations.ClearTelegramCount
    Exit Sub
ElseIf InStr(strLine, "===  Telecommunications  ===") > 0 Then
    npSection = NP_TELECOMMUNICATIONS
    Exit Sub
End If

If InStr(strLine, "::") > 0 Then
    Exit Sub
End If

If InStr(strLine, ":") = 14 Then
    Dim hh As Integer
    Dim Index As Integer
    
    hh = Val(Mid$(strLine, 12, 2))
    strTemp = Right$(strLine, Len(strLine) - 18)
    If (InStr(strTemp, "times")) > 0 Then
        'Remove the number for now, may want to count in the future.
        If Len(strTemp) = InStr(strTemp, "times") + 4 Then
            strTemp = Left$(strTemp, InStr(strTemp, " times") - 1)
            strTemp = Left$(strTemp, InStrRev(strTemp, " ") - 1)
        Else
            strTemp = Right$(strTemp, Len(strTemp) - InStr(strTemp, "times") - 5)
        End If
    End If
    For Index = 0 To Nations.Count
        If Len(Nations.NationName(Index)) > 0 Then
            If InStr(strTemp, Nations.NationName(Index)) = 1 Then
                ActivityCounts(hh, Index) = ActivityCounts(hh, Index) + 1
                Exit For
            End If
        End If
    Next Index
End If

Select Case npSection
Case NP_TELECOMMUNICATIONS
    '     ===  Telecommunications  ===
    'Sat Nov 29 09:01  two times 9 telexes 1 2 3
    'Sat Nov 29 08:56  9 telexes 1 2 3
    'Sat Nov 29 08:54  two times ron5 telexes 1 2 3
    'Sun Nov 23 12:43  1 2 3 sends a telegram to ron5
    'Remove the date
    strLine = Right$(strLine, Len(strLine) - 18)
    If (InStr(strLine, "times")) > 0 Then
        'Remove the number for now, may want to count in the future.
        If Len(strLine) = InStr(strLine, "times") + 4 Then
            strLine = Left$(strLine, InStr(strLine, " times") - 1)
            strLine = Left$(strLine, InStrRev(strLine, " ") - 1)
        Else
            strLine = Right$(strLine, Len(strLine) - InStr(strLine, "times") - 5)
        End If
    End If
    If InStr(strLine, "telexes") > 0 Then
        strOrigin = Left$(strLine, InStr(strLine, " telexes") - 1)
        strDestination = Right$(strLine, Len(strLine) - InStr(strLine, "telexes") - 7)
    ElseIf InStr(strLine, "sends a telegram to") > 0 Then
        strOrigin = Left$(strLine, InStr(strLine, " sends a telegram to") - 1)
        strDestination = Right$(strLine, Len(strLine) - InStr(strLine, "sends a telegram to") - 19)
    End If
    Origin = Nations.NationNumber(strOrigin)
    Destination = Nations.NationNumber(strDestination)
    If Origin <> -1 And Destination <> -1 Then
        Nations.telegramCount(Origin, Destination) = Nations.telegramCount(Origin, Destination) + 1
    End If
End Select
End Sub

