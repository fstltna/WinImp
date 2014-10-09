Attribute VB_Name = "PlayersParse"
Option Explicit

'Change Log:
'120103 rjk: Created to determine and display the current players.

Public currentPlayersNumber() As Integer
Public currentPlayersName() As String

'
'Mon Dec 01 08:26:37 2003
'            #                                   time  idle last command
'10         10                   ronk@localhost  0:02    4s
'1 player
Public Sub ParsePlayers(strLine As String)

If Len(strLine) = 0 Then
    Exit Sub
End If

If InStr(strLine, ":") > 0 Then
    If InStr(InStr(strLine, ":") + 1, strLine, ":") > 0 Then
        Exit Sub 'skip date line
    End If
End If

If Mid$(strLine, 13, 1) = "#" Then 'Header line
    ReDim currentPlayersNumber(0)
    ReDim currentPlayersName(0)
    Exit Sub
End If

If InStr(strLine, "player") Then
    Exit Sub
End If

If Val(Mid$(strLine, 12, 2)) <> CountryNumber Then
    currentPlayersNumber(UBound(currentPlayersNumber)) = Val(Mid$(strLine, 12, 2))
    currentPlayersName(UBound(currentPlayersName)) = Trim$(Left$(strLine, 10))
    ReDim Preserve currentPlayersNumber(0 To (UBound(currentPlayersNumber) + 1))
    ReDim Preserve currentPlayersName(0 To (UBound(currentPlayersName) + 1))
End If
End Sub

Public Sub InitializePlayers()
ReDim currentPlayersNumber(0)
ReDim currentPlayersName(0)
End Sub



