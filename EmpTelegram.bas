Attribute VB_Name = "EmpTelegram"
Option Explicit

'Change Log:
'08xx03 efj: removed dead variables
'101503 rjk: Fixed telegram so does not delete blank lines.
'101603 rjk: Fixed telegram to break message into lines and remove LF or CRLF combinations.
'103003 rjk: Updated SectorLost to make the timestamp format consistent
'            Added Import Intelligence
'110403 rjk: Added the automatic functionality to convert a series of telegrams into one IntelligenceReport
'070204 rjk: Simplify EventMarkNukes
'070304 rjk: Switched Enemy timestamp to UTC
'080304 rjk: Added Check for St@r W@rs space travel (flying through space)
'070604 rjk: Added Plague Marking
'210704 rjk: Removed Check for St@r W@rs space travel (flying through space)
'100804 rjk: Added deity parsing to EventStarvationNotice function.
'110804 rjk: Added the motd to telegram window for deity mode.
'290804 rjk: Always save telegrams, unless it is a duplicate.
'210705 rjk: Added option to filter Unit Damage when shelled.
'100606 rjk: Fixed a bug with sailing order results in production report.
'            The second part of the production report was being lost.
'050706 rjk: Do not erase minefields when flying over them or from bulletins.
'170906 rjk: Added TTS for telegrams and bulletins.
'020907 rjk: Switched to database field for TimeZoneAdj.
'171007 rjk: Fix the parsing of bulletins for split country names
'311007 rjk: Mark sectors that are now fully ours in yellow

Public Const TELEGRAM_CLASS_INBOX = 1
Public Const TELEGRAM_CLASS_BULLETIN = 2
Public Const TELEGRAM_CLASS_PLAYER = 3
Public Const TELEGRAM_CLASS_PRODUCTION = 4
Public Const TELEGRAM_CLASS_REPORTS = 5
Public Const TELEGRAM_CLASS_INTELLIGENCE = 6 '103003 rjk: Added Import Intelligence
Public Const TELEGRAM_CLASS_GENERAL = 7
Public Const TELEGRAM_CLASS_COMBAT = 8

Public Const NUMBER_OF_TELEGRAM_CLASSES = 7 '103003 rjk: Added Import Intelligence

Public Const BULLETIN_COASTWATCH = 1
Public Const BULLETIN_SHIP_BOARDED = 2
Public Const BULLETIN_THREAT = 3
Public Const BULLETIN_SONAR_PING = 4
Public Const BULLETIN_PLANE_OVERFLIGHT = 5
Public Const BULLETIN_SHELLED = 6
Public Const BULLETIN_BOMBED = 7
Public Const BULLETIN_AIR_ASSAULT = 8
Public Const BULLETIN_SPY_CAUGHT = 9
Public Const BULLETIN_MISSILE_ATTACK = 10
Public Const BULLETIN_NAVEL_ASSAULT = 11
Public Const BULLETIN_NUKE_BLAST = 12
Public Const BULLETIN_LOST_TO_ATTACK = 13
Public Const BULLETIN_DIPLOMACY = 14
Public Const BULLETIN_WENT_HOSTILE = 15
Public Const BULLETIN_ATTACK_REPULSED = 15
Public Const BULLETIN_CASULTIES = 16
Public Const BULLETIN_UNIT_REACTS = 17
Public Const BULLETIN_UNIT_RETREATS = 18
Public Const BULLETIN_UNIT_DIES = 19
Public Const BULLETIN_SHIP_SUNK = 20
Public Const BULLETIN_SCOUT_REPORT = 21
Public Const BULLETIN_SHIP_TORPED = 22
Public Const BULLETIN_SUB_SPOTTED = 23
Public Const BULLETIN_FLEW_OVER = 24
Public Const BULLETIN_TORP_SIGHTED = 25

Public RecievedNoticeOfMapChange As Boolean
Public TelegramMessage As String
Public LastShip As Integer
Public EventMarkers As New EmpSectorMarker
Public EventX As Integer
Public EventY As Integer
Public strLastProductionReport As String

Dim CurrentTelegramLine As Integer
Dim NewTelegramID As Integer
Dim CurrentTelegramClass As Integer

Public Sub IncomingTelegramLine(strLine As String, NewTele As Boolean)
On Error GoTo Empire_Error
Static bSaveTelegram As Boolean

TelegramMessage = strLine
If NewTele Then
    CurrentTelegramClass = ProcessTelegramHeader(strLine)
    'set short message
    Select Case CurrentTelegramClass
        Case TELEGRAM_CLASS_BULLETIN
            TelegramMessage = "BULLETIN"
        Case TELEGRAM_CLASS_PRODUCTION
            TelegramMessage = "Production Report"
        Case TELEGRAM_CLASS_PLAYER
            TelegramMessage = Left$(strLine, InStr(strLine, "(") - 1)
    End Select

    bSaveTelegram = StartNewTelegram(strLine)
Else
    If bSaveTelegram Then
        AddTelegramLine strLine
    End If
    If CurrentTelegramClass = TELEGRAM_CLASS_BULLETIN Then
        IncomingBulletin strLine
    'ships moved by "orders" post the coastwatch bulletins in
    'the production telegram.
    ElseIf CurrentTelegramClass = TELEGRAM_CLASS_PRODUCTION And _
        InStr(strLine, "sighted at") > 0 Then
        CurrentTelegramClass = TELEGRAM_CLASS_BULLETIN
        IncomingBulletin strLine
        CurrentTelegramClass = TELEGRAM_CLASS_PRODUCTION
    End If
End If

Select Case CurrentTelegramClass
    Case TELEGRAM_CLASS_BULLETIN
        If frmOptions.bTTSBulletins And Len(strLine) > 0 Then
'            frmReport.ttsReport.Tag = "Normal"
'            frmReport.ttsReport.Speak strLine
        End If
    Case TELEGRAM_CLASS_PLAYER
        If frmOptions.bTTSTelegrams And Len(strLine) > 0 Then
'            frmReport.ttsReport.Tag = "Normal"
'            frmReport.ttsReport.Speak strLine
        End If
End Select

Exit Sub
Empire_Error:
EmpireError "IncomingTelegramLine", vbNullString, strLine
End Sub

'new telegram
Private Function StartNewTelegram(strHeader As String) As Boolean
Dim n As Integer
Dim n2 As Integer
Dim hldIndex As String

StartNewTelegram = True
On Error GoTo StartNewTelegram_Error

rsTeleHead.Index = "Title"
rsTeleHead.Seek "=", strHeader
If Not rsTeleHead.NoMatch Then
    rsTeleBody.Index = "PrimaryKey"
'    StartNewTelegram = False
    If strLastProductionReport = strHeader Then
        hldIndex = rsTeleBody.Index
        NewTelegramID = rsTeleHead.Fields("id")
        CurrentTelegramLine = 1
        rsTeleBody.Index = "ID"
        rsTeleBody.Seek "=", NewTelegramID
        If Not rsTeleBody.NoMatch Then
            While rsTeleBody.Fields("id") = NewTelegramID
                CurrentTelegramLine = CurrentTelegramLine + 1
                rsTeleBody.MoveNext
            Wend
        End If
        rsTeleBody.Index = hldIndex
        StartNewTelegram = True
    Else
        StartNewTelegram = False
    End If
    Exit Function
End If
'Get the last telegram number
rsTeleHead.Index = "PrimaryKey"
If rsTeleHead.BOF And rsTeleHead.EOF Then
    NewTelegramID = 1    'First Telegram
Else
    rsTeleHead.MoveLast
    NewTelegramID = rsTeleHead.Fields("ID") + 1
End If

'Create New Telegram header
rsTeleHead.AddNew
rsTeleHead.Fields("ID") = NewTelegramID
rsTeleHead.Fields("Title") = strHeader
rsTeleHead.Fields("Read") = "N"
rsTeleHead.Fields("Exported") = "N"
n = InStr(strHeader, "dated")
If n > 0 Then
    rsTeleHead.Fields("Received") = EmpireTimeString(Trim$(Mid$(strHeader, n + 6)))
Else
    rsTeleHead.Fields("Received") = DateAdd("n", rsVersion.Fields("Time Zone Adj"), Now)
End If

'Figure out telegram class
Select Case CurrentTelegramClass
    Case TELEGRAM_CLASS_PRODUCTION
        rsTeleHead.Fields("Class") = TELEGRAM_CLASS_PRODUCTION
        rsTeleHead.Fields("From") = 0
        rsTeleHead.Fields("Subject") = rsTeleHead.Fields("Received")
        strLastProductionReport = strHeader

'Deity Bulletin
    Case TELEGRAM_CLASS_BULLETIN
        rsTeleHead.Fields("Class") = TELEGRAM_CLASS_BULLETIN
        rsTeleHead.Fields("From") = 0
        rsTeleHead.Fields("Subject") = "Bulletin"

'Saved Report
    Case TELEGRAM_CLASS_REPORTS
        rsTeleHead.Fields("Class") = TELEGRAM_CLASS_REPORTS
        rsTeleHead.Fields("From") = 0
        rsTeleHead.Fields("Subject") = strHeader
        rsTeleHead.Fields("Read") = "Y"

    Case TELEGRAM_CLASS_PLAYER
        'Saved telegram
        rsTeleHead.Fields("Class") = TELEGRAM_CLASS_PLAYER
        rsTeleHead.Fields("Subject") = "..."
        n = InStr(strHeader, "(#")
        If n > 0 Then
            n2 = InStr(n + 2, strHeader, ")")
            rsTeleHead.Fields("From") = CInt(Mid$(strHeader, n + 2, n2 - n - 2))
        Else
            rsTeleHead.Fields("From") = 0
        End If
        
        If InStr(strHeader, "Telegram To ") > 0 Then
            rsTeleHead.Fields("Read") = "Y"
        End If
        
    Case Else
        rsTeleHead.Fields("Subject") = "..."
        n = InStr(strHeader, "(#")
        If n > 0 Then
            n2 = InStr(n + 2, strHeader, ")")
            rsTeleHead.Fields("From") = CInt(Mid$(strHeader, n + 2, n2 - n - 2))
        Else
            rsTeleHead.Fields("From") = 0
        End If
End Select
CurrentTelegramClass = rsTeleHead.Fields("Class")
rsTeleHead.Update
CurrentTelegramLine = 1
frmTelegram.TreeDirty = True
Exit Function

StartNewTelegram_Error:
MsgBox "ERROR - Your telegram database is corrupted.  WinACE was not able" + _
" to save the latested telegram.  You should take Export, then Clear, then Import off" + _
" the File Menu to repair the database", vbOKOnly + vbCritical, "Database Error"
End Function

Private Sub AddTelegramLine(strLine As String)
Dim pos As Integer
Dim posLF As Integer '101603 rjk: Used to break message into lines  and remove LFs.
Dim posCRLF As Integer '101603 rjk: Used to break message into lines and remove CRLF combinations
If NewTelegramID = 0 Or CurrentTelegramLine = 0 Then
    'MsgBox strLine, vbExclamation + vbOKOnly, "Error - Tried to add Telegram Line before starting Telegram"
    Exit Sub
End If

'091803 rjk: strLine now contains the whole telegram, so must break apart for insertion
'            into the database.
'            May or May not contain the whole telegram.  Use Line to tell if it is the end
'101503 rjk: Fixed to store blank lines
'101603 rjk: Break message into lines with vbLF found
pos = 0
    
'110403 rjk: Remove any subject lines from the middle of the telegram, the server
'            joins telegrams together if they are send close enough together
If CurrentTelegramClass = TELEGRAM_CLASS_INTELLIGENCE Then
    If InStr(strLine, "+++ Intelligence Report") > 0 Then
        Exit Sub
    End If
End If

On Error GoTo AddTelegramLine_Error
Do
    rsTeleBody.AddNew
    rsTeleBody.Fields("ID") = NewTelegramID
    rsTeleBody.Fields("Line") = CurrentTelegramLine
    posLF = InStr(pos + 1, strLine, vbLf)
    posCRLF = InStr(pos + 1, strLine, vbCrLf)
    
    If posLF = 0 Then
        posLF = Len(strLine)
    End If
    'Check for max length
    'If Len(strLine) - pos > 255 Then
    If (posLF - pos) >= 255 Then
        rsTeleBody.Fields("Body") = Mid$(strLine, pos + 1, 255)
        pos = pos + 255
        If pos = Len(strLine) And Right$(strLine, 1) <> vbLf Then '101503 rjk: If the length is exact, then the add blank to force a blank on insertion into textbox.
            rsTeleBody.Update
            CurrentTelegramLine = CurrentTelegramLine + 1
            rsTeleBody.AddNew
            rsTeleBody.Fields("ID") = NewTelegramID
            rsTeleBody.Fields("Line") = CurrentTelegramLine
            rsTeleBody.Fields("Body") = ""
        End If
    ElseIf posLF = 0 Then
        rsTeleBody.Fields("Body") = ""
    Else
        If (posLF = (posCRLF + 1)) And posCRLF <> 0 Then
            rsTeleBody.Fields("Body") = Mid$(strLine, pos + 1, posCRLF - pos - 1)
        ElseIf Mid$(strLine, posLF, 1) = vbLf Then
            rsTeleBody.Fields("Body") = Mid$(strLine, pos + 1, posLF - pos - 1)
        Else
            rsTeleBody.Fields("Body") = Mid$(strLine, pos + 1)
        End If
        pos = posLF
    End If
    
    rsTeleBody.Update
    CurrentTelegramLine = CurrentTelegramLine + 1
Loop While pos < Len(strLine) '101503 rjk: Moved the check to the bottom of the loop to write blank lines.

If Len(Trim$(strLine)) = 0 Then
    Exit Sub
End If

'Skip the rest if you can
If CurrentTelegramClass = TELEGRAM_CLASS_PLAYER Or CurrentTelegramClass = TELEGRAM_CLASS_INTELLIGENCE Then
    If CurrentTelegramLine > 2 Then
        Exit Sub
    Else
        rsTeleHead.Seek "=", NewTelegramID
        If Not rsTeleHead.NoMatch Then
            '103003 rjk: Added Intelligence Report (special Player telegram)
            If InStr(strLine, "+++ Intelligence Report") > 0 Then
                ConvertToIntelligenceTelegrams strLine '110403 rjk: Added a new function to combined telegrams to together
            Else
                rsTeleHead.Edit
                rsTeleHead.Fields("Subject") = Left$(strLine, 50)
                rsTeleHead.Update
            End If
        End If
    End If
End If

If CurrentTelegramClass <> TELEGRAM_CLASS_BULLETIN Then
    Exit Sub
End If
        
rsTeleHead.Seek "=", NewTelegramID
If Not rsTeleHead.NoMatch Then
   
    rsTeleHead.Edit
    If InStr(strLine, "sighted") > 0 Then
        rsTeleHead.Fields("Subject") = "Coastwatch"
    ElseIf InStr(strLine, "boarded") > 0 Then
        rsTeleHead.Fields("Subject") = "Ship Boarded"
    ElseIf InStr(strLine, "considered") > 0 Then
        rsTeleHead.Fields("Subject") = "Threat"
    ElseIf InStr(strLine, "onar ping") > 0 Then
        rsTeleHead.Fields("Subject") = "Sonar Ping"
    ElseIf InStr(strLine, "spotted") > 0 Then
        rsTeleHead.Fields("Subject") = "Plane Overflight"
    ElseIf InStr(strLine, "shelled") > 0 Then
        rsTeleHead.Fields("Subject") = "Shelled"
    ElseIf InStr(strLine, "bombing") > 0 Then
        rsTeleHead.Fields("Subject") = "Bombing Raid"
    ElseIf InStr(strLine, "air-assault") > 0 Then
        rsTeleHead.Fields("Subject") = "Air Assault"
    ElseIf InStr(strLine, "spy ") > 0 Then
        rsTeleHead.Fields("Subject") = "Spy Caught"
    ElseIf InStr(strLine, "missile attack ") > 0 Then
        rsTeleHead.Fields("Subject") = "Missile Attack"
    ElseIf InStr(strLine, " assault") > 0 Then
        rsTeleHead.Fields("Subject") = "Naval Assualt"
    ElseIf InStr(strLine, " nuclear device did") > 0 Then
        rsTeleHead.Fields("Subject") = "Nuke Blast"
    ElseIf rsTeleHead.Fields("Subject") = "Bulletin" Then
            If InStr(strLine, "defending") > 0 Then
                rsTeleHead.Fields("Subject") = "Ground Attack"
            ElseIf InStr(strLine, "relation") > 0 Then
                rsTeleHead.Fields("Subject") = "Diplomacy"
            End If
    End If
    
    rsTeleHead.Update
End If
Exit Sub

AddTelegramLine_Error:
MsgBox "ERROR - Your telegram database is corrupted.  WinACE was not able" + _
" to save the latest telegram.  You should take Export, then Clear, then Import off" + _
" the File Menu to repair the database", vbOKOnly + vbCritical, "Database Error"

End Sub

Private Function ProcessTelegramHeader(strHeader As String) As Integer
On Error GoTo Empire_Error

'Figure out telegram class
If InStr(strHeader, "Production Report") > 0 Then
    ProcessTelegramHeader = TELEGRAM_CLASS_PRODUCTION
'Deity Bulletin
ElseIf InStr(strHeader, "BULLETIN") > 0 Then
    ProcessTelegramHeader = TELEGRAM_CLASS_BULLETIN
'Saved Report
ElseIf InStr(strHeader, "> Saved ") > 0 Then
    ProcessTelegramHeader = TELEGRAM_CLASS_REPORTS
'Saved telegram
ElseIf InStr(strHeader, "Telegram To ") > 0 Then
    ProcessTelegramHeader = TELEGRAM_CLASS_PLAYER
'Telegram from players
ElseIf InStr(strHeader, "Telegram from") > 0 Then
    ProcessTelegramHeader = TELEGRAM_CLASS_PLAYER
Else
    ProcessTelegramHeader = 0
End If
Exit Function
Empire_Error:
EmpireError "ProcessTelegramHeader", vbNullString, strHeader
End Function

'110403 rjk: New function added to convert a series of telegrams into one IntelligenceReport
Private Sub ConvertToIntelligenceTelegrams(strSubject As String)
Dim strError As String
Dim hldIndex As String
Dim LastLineNumber As Integer

strError = "Database Error"

On Error GoTo Empire_Error

rsTeleHead.Seek "=", NewTelegramID - 1
If Not rsTeleHead.NoMatch Then
    If rsTeleHead.Fields("Subject") = Mid$(strSubject, 25) And _
       rsTeleHead.Fields("Class") = TELEGRAM_CLASS_INTELLIGENCE Then
        'delete the header
        rsTeleHead.Seek "=", NewTelegramID
        If Not rsTeleHead.NoMatch Then
            rsTeleHead.Delete
            NewTelegramID = NewTelegramID - 1
        Else
            strError = "Did not find header to delete"
            GoTo Empire_Error
        End If
        rsTeleHead.Seek "=", NewTelegramID
        If rsTeleHead.NoMatch Then
            strError = "Did not find first header"
            GoTo Empire_Error
        End If
        'remove the body - should delete the intelligence subject line
        hldIndex = rsTeleBody.Index
        rsTeleBody.Index = "PrimaryKey"
        rsTeleBody.Seek "=", NewTelegramID + 1, 1
        If Not rsTeleBody.NoMatch Then
            rsTeleBody.Delete
        Else
            rsTeleBody.Index = hldIndex
            strError = "Did not find the new telegram body to delete"
            GoTo Empire_Error
        End If
        LastLineNumber = 0
        Do
            LastLineNumber = LastLineNumber + 1
            rsTeleBody.Seek "=", NewTelegramID, LastLineNumber
        Loop While Not rsTeleBody.NoMatch
        If LastLineNumber = 1 Then
            rsTeleBody.Index = hldIndex
            strError = "Did not find first telegram body"
            GoTo Empire_Error
        End If
        rsTeleBody.Seek "=", NewTelegramID, LastLineNumber - 1
        If Not rsTeleBody.NoMatch Then
            CurrentTelegramLine = LastLineNumber
            CurrentTelegramClass = TELEGRAM_CLASS_INTELLIGENCE
        Else
            rsTeleBody.Index = hldIndex
            strError = "Did not find last telegram body"
            GoTo Empire_Error
        End If
        rsTeleBody.Index = hldIndex
        Exit Sub
    End If
End If
rsTeleHead.Seek "=", NewTelegramID
If rsTeleHead.NoMatch Then
    strError = "Did not find original header"
    GoTo Empire_Error
End If
rsTeleHead.Edit
rsTeleHead.Fields("Subject") = Mid$(strSubject, 25)
rsTeleHead.Fields("Class") = TELEGRAM_CLASS_INTELLIGENCE
rsTeleHead.Update
CurrentTelegramClass = TELEGRAM_CLASS_INTELLIGENCE
Exit Sub

Empire_Error:
EmpireError "ConvertToIntelligenceTelegrams", strError, strSubject
End Sub

Private Function BulletinType(strLine As String) As Integer
On Error GoTo Empire_Error

If InStr(strLine, "Torpedo sight") > 0 Then
    BulletinType = BULLETIN_TORP_SIGHTED
ElseIf InStr(strLine, " sighted ") > 0 Then
    BulletinType = BULLETIN_COASTWATCH
ElseIf InStr(strLine, " boarded ") > 0 Then
    BulletinType = BULLETIN_SHIP_BOARDED
ElseIf InStr(strLine, "considered attacking you") > 0 Then
    BulletinType = BULLETIN_THREAT
ElseIf InStr(strLine, "onar ping") > 0 Then
    BulletinType = BULLETIN_SONAR_PING
ElseIf InStr(strLine, " spotted ") > 0 Then
    BulletinType = BULLETIN_PLANE_OVERFLIGHT
ElseIf InStr(strLine, "shelled") > 0 Then
    BulletinType = BULLETIN_SHELLED
ElseIf InStr(strLine, "bombing") > 0 Then
    BulletinType = BULLETIN_BOMBED
ElseIf InStr(strLine, "bombs did") > 0 Then
    BulletinType = BULLETIN_BOMBED
ElseIf InStr(strLine, "air-assault") > 0 Then
    BulletinType = BULLETIN_AIR_ASSAULT
ElseIf InStr(strLine, " spy ") > 0 Then
    BulletinType = BULLETIN_SPY_CAUGHT
ElseIf InStr(strLine, "flying over ") > 0 Then
    BulletinType = BULLETIN_FLEW_OVER
ElseIf InStr(strLine, " tracks ") > 0 And InStr(strLine, " subs at ") > 0 Then
    BulletinType = BULLETIN_SUB_SPOTTED
ElseIf InStr(strLine, "surfacing noises") > 0 Then
    BulletinType = BULLETIN_SUB_SPOTTED
ElseIf InStr(strLine, "eriscope spotted") > 0 Then
    BulletinType = BULLETIN_SUB_SPOTTED
ElseIf InStr(strLine, " missile attack ") > 0 Then
    BulletinType = BULLETIN_MISSILE_ATTACK
ElseIf InStr(strLine, " assault") > 0 Then
    BulletinType = BULLETIN_NAVEL_ASSAULT
ElseIf InStr(strLine, " nuclear device did") > 0 Then
    BulletinType = BULLETIN_NUKE_BLAST
ElseIf InStr(strLine, " troops taking") > 0 Then
    BulletinType = BULLETIN_LOST_TO_ATTACK
ElseIf InStr(strLine, " troops attacking") > 0 Then
    BulletinType = BULLETIN_ATTACK_REPULSED
ElseIf InStr(strLine, " troops defending") > 0 Then
    BulletinType = BULLETIN_CASULTIES
ElseIf InStr(strLine, " reacts to") > 0 Then
    BulletinType = BULLETIN_UNIT_REACTS
ElseIf InStr(strLine, " retreats at") > 0 Then
    BulletinType = BULLETIN_UNIT_RETREATS
ElseIf InStr(strLine, " dies defending") > 0 Or _
    InStr(strLine, "retreat, and dies") > 0 Then
    BulletinType = BULLETIN_UNIT_DIES
ElseIf InStr(strLine, "Scouts report") > 0 Then
    BulletinType = BULLETIN_SCOUT_REPORT
ElseIf InStr(strLine, " torpedoed") > 0 Then
    BulletinType = BULLETIN_SHIP_TORPED
ElseIf InStr(strLine, " sunk!") > 0 Then
    BulletinType = BULLETIN_SHIP_SUNK
ElseIf InStr(strLine, " their relations with you") > 0 Then
    BulletinType = BULLETIN_DIPLOMACY
ElseIf InStr(strLine, " downgraded to") > 0 And _
    InStr(strLine, "Hostile") > 0 Then
    BulletinType = BULLETIN_WENT_HOSTILE
Else

End If
'metathemus missile attack did 36 damage in 4,20
Exit Function
Empire_Error:
EmpireError "BulletinType", vbNullString, strLine
End Function


Private Sub IncomingBulletin(strLine As String)
On Error GoTo Empire_Error
' Dim ntype As Integer    8/2003 efj  removed

If CurrentTelegramClass <> TELEGRAM_CLASS_BULLETIN Then
    Exit Sub
End If

Select Case BulletinType(strLine)
    Case BULLETIN_COASTWATCH
        BulletinCoastwatch strLine
    Case BULLETIN_LOST_TO_ATTACK
        BulletinSectorLost strLine
    Case BULLETIN_AIR_ASSAULT
        BulletinAirAssault strLine
    Case BULLETIN_NAVEL_ASSAULT
        BulletinNavelAssault strLine
    Case BULLETIN_SHELLED
        BulletinShelled strLine
    Case BULLETIN_THREAT
        BulletinThreat strLine
    Case BULLETIN_PLANE_OVERFLIGHT
        BulletinPlaneOverflight strLine
    Case BULLETIN_SPY_CAUGHT
        BulletinSpyCaught strLine
    Case BULLETIN_SONAR_PING
        BulletinSonarPing strLine
    Case BULLETIN_SHIP_TORPED
        BulletinShipTorped strLine
    Case BULLETIN_SHIP_SUNK
        BulletinShipSunk strLine
    Case BULLETIN_BOMBED
        BulletinBombed strLine
    Case BULLETIN_SUB_SPOTTED
        BulletinSubDetected strLine
    Case BULLETIN_FLEW_OVER
        BulletinFlewOver strLine
    Case BULLETIN_TORP_SIGHTED
        BulletinTorpedoSighted strLine
    Case BULLETIN_MISSILE_ATTACK
        BulletinMissileAttack strLine
        
        
End Select
Exit Sub
Empire_Error:
EmpireError "IncomingBulletin", vbNullString, strLine
End Sub
Private Sub BulletinCoastwatch(strLine As String)
On Error GoTo Empire_Error
'From coastwatch
'WinACE2 (#  8) hc   heavy cruiser (#2) @ 22,4

'From coastwatch bulletin
'SNAFU frg  frigate (#12) sighted at -27,5
'Big Rock frg  frigate (#1) sighted at 2,4

'The idea here is to run this thru the normal coastwatch
'parsing routine, which requires a few modification to the string
Dim strLook As String
Dim strOwner As String
Dim n As Integer
Dim n2 As Integer
Dim NatNum As Integer

'first replace "sighted at" with "@"
strLook = strLine
n = InStr(strLine, "sighted at")
If n < 2 Then Exit Sub
strLook = Left$(strLine, n - 1) + "@" + Mid$(strLine, n + 10)
TelegramMessage = strLook
'Then find the owner name.  If it is a mulipart name, like
'"Park Place" then make sure you try with all parts
n = 0
strOwner = vbNullString
NatNum = -1
While NatNum < 0
    n = InStr(n + 1, strLook, " ")
    If n = 0 Then Exit Sub
    strOwner = Left$(strLook, n - 1)
    NatNum = Nations.NationNumber(strOwner)
Wend

For n = 1 To Len(Nations.NationName(NatNum))
    If Left$(strLine, n) <> Left$(Nations.NationName(NatNum), n) Then
        Exit For
    End If
Next

'for telegram message, take out he ship type, -- desription is sufficent
n2 = InStr(n + 1, strLook, " ")
TelegramMessage = Left$(strLook, n) + Mid$(strLook, n2 + 1)

strLook = Left$(strLook, n) + "(#" + CStr(NatNum) + ") " + Mid$(strLook, n + 1)

'Now Process it
Dim VARsec As Variant
VARsec = ParseCoastWatch(strLook)
If IsArray(VARsec) Then
    AddEnemyUnitInfo VARsec
    LastShip = CInt(VARsec(EU_ID))
    RecievedNoticeOfMapChange = True
End If
Exit Sub
Empire_Error:
EmpireError "BulletinCoastwatch", vbNullString, strLine

End Sub
Private Sub BulletinSectorLost(strLine As String)
On Error GoTo Empire_Error
'SiliconSorcerer (#1) lost 118 troops taking -5,11

Dim i As Integer
Dim n As Integer
Dim n2 As Integer
Dim owner As Integer

i = InStr(strLine, " troops taking")
If i > 0 Then
    
    'get the owner
    n = InStr(strLine, "(#")
    If n > 0 Then
        n2 = InStr(n + 2, strLine, ")")
        owner = CInt(Trim$(Mid$(strLine, n + 2, n2 - n - 2)))
    End If

    SectorLost Mid$(strLine, i + 14), owner
    
End If
Exit Sub
Empire_Error:
EmpireError "BulletinSectorLost", vbNullString, strLine
End Sub
Private Sub BulletinAirAssault(strLine As String)
On Error GoTo Empire_Error
'Guitest (#8) lost 0 troops assaulting and taking 3,1
'Guitest (#8) lost 36 troops trying to air-assault 2,0

Dim i As Integer
Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim owner As Integer

'get the owner
n = InStr(strLine, "(#")
If n > 0 Then
    n2 = InStr(n + 2, strLine, ")")
    owner = CInt(Trim$(Mid$(strLine, n + 2, n2 - n - 2)))
End If

i = InStr(strLine, " and taking")
If i > 0 Then
    'Assault was Successfull
    SectorLost Mid$(strLine, i + 11), owner
Else
    'assault failed - mark sector
    i = InStr(strLine, " air-assault")
    If i > 0 Then
        'get sector number
        strSect = Trim$(Mid$(strLine, i + 12))
        
        MarkSector strSect, owner, "Para-drop by"
    End If
End If

Exit Sub
Empire_Error:
EmpireError "BulletinAirAssault", vbNullString, strLine

End Sub
Private Sub BulletinNavelAssault(strLine As String)
On Error GoTo Empire_Error
'Guitest (#8) lost 0 troops assaulting and taking 3,1
'Guitest (#8) lost 50 troops trying to assault 8,0

Dim i As Integer
Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim owner As Integer

'get the owner
n = InStr(strLine, "(#")
If n > 0 Then
    n2 = InStr(n + 2, strLine, ")")
    owner = CInt(Trim$(Mid$(strLine, n + 2, n2 - n - 2)))
End If

i = InStr(strLine, " and taking")
If i > 0 Then
    SectorLost Mid$(strLine, i + 11), owner
Else
    'assault failed - mark sector
    i = InStr(strLine, " assault")
    If i > 0 Then
        'get sector number
        strSect = Trim$(Mid$(strLine, i + 8))
        
        MarkSector strSect, owner, "Assault by"
    End If
End If
Exit Sub
Empire_Error:
EmpireError "BulletinNavelAssault", vbNullString, strLine
End Sub
Private Sub BulletinShelled(strLine As String)
On Error GoTo Empire_Error
'Country #8 shelled sector 3,3 for 39 damage.
'Country #8 shelled dd   destroyer (#9) in 0,10 for 34 damage.
'Country #8 shelled sb   submarine (#1) in -6,8 for 31 damage.
Dim i As Integer
' Dim n As Integer    8/2003 efj  removed
' Dim n2 As Integer    8/2003 efj  removed
Dim strSect As String
Dim owner As Integer

On Error Resume Next
'get the owner
owner = CInt(Trim$(StringBetween(strLine, " #", " shelled ")))

i = InStr(strLine, " sector ")
If i > 0 Then
    'get sector number
    strSect = StringBetween(strLine, " sector ", " for ")
Else
    strSect = StringBetween(strLine, " in ", " for ")
End If
    
MarkSector strSect, owner, "Shelled by"
Exit Sub
Empire_Error:
EmpireError "BulletinShelled", vbNullString, strLine
End Sub

Private Sub BulletinThreat(strLine As String)
On Error GoTo Empire_Error
'disneyland (#12) considered attacking you @-8,0

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim owner As Integer

'get the owner
n = InStr(strLine, "(#")
If n > 0 Then
    n2 = InStr(n + 2, strLine, ")")
    owner = CInt(Trim$(Mid$(strLine, n + 2, n2 - n - 2)))
End If

'get sector number
n = InStr(strLine, "@")
strSect = Trim$(Mid$(strLine, n + 1))

MarkSector strSect, owner, "Threated by"
Exit Sub
Empire_Error:
EmpireError "BulletinThreat", vbNullString, strLine
    
End Sub
Private Sub BulletinSpyCaught(strLine As String)
On Error GoTo Empire_Error
'Lothlorien (#3) spy caught in -16,-6
'10 (#10) spy deported from 4,2


' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim owner As Integer

'get the owner
n = InStr(strLine, "(#")
If n > 0 Then
    n2 = InStr(n + 2, strLine, ")")
    owner = CInt(Trim$(Mid$(strLine, n + 2, n2 - n - 2)))
End If

'get sector number
n = InStr(strLine, "caught in")
If n > 0 Then
    strSect = Trim$(Mid$(strLine, n + 10))
    MarkSector strSect, owner, "Spy Caught"
Else
    n = InStr(strLine, "deported From")
    If n > 0 Then
        strSect = Trim$(Mid$(strLine, n + 14))
        MarkSector strSect, owner, "Spy Caught"
    End If
End If
Exit Sub
Empire_Error:
EmpireError "BulletinSpyCaught", vbNullString, strLine
End Sub
Private Sub BulletinPlaneOverflight(strLine As String)
On Error GoTo Empire_Error
'Lothlorien planes spotted over -16,-6
'nuke_em planes spotted over ships in 9,15

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
' Dim n2 As Integer    8/2003 efj  removed
Dim strSect As String
Dim owner As Integer
Dim strOwner As String

'get the owner
n = InStr(strLine, "planes spotted")
If n > 0 Then
    strOwner = Trim$(Left$(strLine, n - 1))
    owner = Nations.NationNumber(strOwner)
End If

'get sector number
n = InStr(strLine, "over ships in")
If n = 0 Then
    n = InStr(strLine, " spotted over")
End If
strSect = Trim$(Mid$(strLine, n + 13))

MarkSector strSect, owner, "Planes spotted"
Exit Sub
Empire_Error:
EmpireError "BulletinPlaneOverflight", vbNullString, strLine
End Sub
Private Sub BulletinMissileAttack(strLine As String)
On Error GoTo Empire_Error
'Lothlorien planes spotted over -16,-6
'nuke_em planes spotted over ships in 9,15

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
' Dim n2 As Integer    8/2003 efj  removed
Dim strSect As String
Dim owner As Integer
Dim strOwner As String

'get the owner
n = InStr(strLine, "missile attack")
If n > 0 Then
    strOwner = Trim$(Left$(strLine, n - 1))
    owner = Nations.NationNumber(strOwner)
End If

'get sector number
n = InStr(strLine, "damage in ")
strSect = Trim$(Mid$(strLine, n + 10))

MarkSector strSect, owner, "Missile Attack"
Exit Sub
Empire_Error:
EmpireError "BulletinMissileAttack", vbNullString, strLine
End Sub
Private Sub BulletinSonarPing(strLine As String)
On Error GoTo Empire_Error
'Sonar ping from 5,-11 detected by dd   destroyer (#37)!

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
Dim strSect As String

'get sector number
n = InStr(strLine, "from ")
n2 = InStr(strLine, "detected ")
strSect = Trim$(Mid$(strLine, n + 4, n2 - n - 5))

MarkSector strSect, -2, "Sonar Ping"
Exit Sub
Empire_Error:
EmpireError "BulletinSonarPing", vbNullString, strLine
End Sub
Private Sub BulletinShipTorped(strLine As String)
On Error GoTo Empire_Error
'sub in -36,-48 torpedoed pt   patrol boat (#378) for 82 damage.

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
Dim strSect As String

'get sector number
n = InStr(strLine, "sub in ")
n2 = InStr(strLine, "torpedoed ")
strSect = Trim$(Mid$(strLine, n + 6, n2 - n - 7))

MarkSector strSect, -2, "Sub Fired"
Exit Sub
Empire_Error:
EmpireError "BulletinShipTorped", vbNullString, strLine
End Sub
Private Sub BulletinTorpedoSighted(strLine As String)
On Error GoTo Empire_Error
'Torpedo sighted @ 10,-4 by frg  frigate (#5)

Dim strSect As String

'get sector number
strSect = Trim$(StringBetween(strLine, "@", " by "))
MarkSector strSect, -2, "Torp Spotted"
Exit Sub
Empire_Error:
EmpireError "BulletinTorpedoSighted", vbNullString, strLine
End Sub
Public Sub BulletinShipSunk(strLine As String)
On Error GoTo Empire_Error
'dd   destroyer (#88) sunk!

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
' Dim strSect As String    8/2003 efj  removed
Dim shipid As Integer
Dim hldIndex

hldIndex = rsShip.Index
rsShip.Index = "PrimaryKey"

On Error GoTo BulletinShipSunk_Exit

'get the ship nember
n = InStr(strLine, "(#")
If n > 0 Then
    n2 = InStr(n + 2, strLine, ")")
    shipid = CInt(Trim$(Mid$(strLine, n + 2, n2 - n - 2)))

    rsShip.Seek "=", shipid
    If Not rsShip.NoMatch Then
        rsShip.Delete
    Else
        'check enemy database to see if its in there
        rsEnemyUnit.MoveFirst
        While Not rsEnemyUnit.EOF
            If rsEnemyUnit.Fields("id") = shipid And _
               rsEnemyUnit.Fields("class") = "s" Then
               rsEnemyUnit.Delete
            End If
            rsEnemyUnit.MoveNext
        Wend
    End If
    
End If
BulletinShipSunk_Exit:
rsShip.Index = hldIndex
Exit Sub
Empire_Error:
EmpireError "BulletinShipSunk", vbNullString, strLine
End Sub
Private Sub BulletinBombed(strLine As String)
On Error GoTo Empire_Error
'nuke_em bombs did 68 damage to sb   submarine (#98) at 17,9
'nuke_em bombs did 68 damage to dd   destroyer (#88) at 9,1
'WinACE2 bombing raid did 53 damage in 0,4

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim owner As Integer
Dim strOwner As String

'get the owner
n = InStr(strLine, "bombs did ")
If n = 0 Then
    n = InStr(strLine, "bombing raid ")
    n2 = InStr(strLine, "e in ")
Else
    n2 = InStr(strLine, ") at")
End If

If n > 0 Then
    strOwner = Trim$(Left$(strLine, n - 1))
    owner = Nations.NationNumber(strOwner)
End If

'get sector number
strSect = Trim$(Mid$(strLine, n2 + 4))
MarkSector strSect, owner, "Bombed"
Exit Sub
Empire_Error:
EmpireError "BulletinBombed", vbNullString, strLine
End Sub
Private Sub BulletinSubDetected(strLine As String)
On Error GoTo Empire_Error
'as   anti-sub plane #494 tracks Coke subs at 2,18
'as   anti-sub plane #494 detects sub movement in 1,19

'sb   submarine (#1642) detected surfacing noises in -26,-26.
'Periscope spotted in -30,-20 by pt patrol boat (#1475)
'dd   destroyer (#245) detected surfacing noises in -35,-11.

' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim owner As Integer
Dim strOwner As String

'get the owner
n = InStr(strLine, "tracks ")
If n > 0 Then
    n2 = InStr(strLine, "subs at")
    strOwner = Trim$(Mid$(strLine, n + 5, n2 - n - 6))
    owner = Nations.NationNumber(strOwner)
    strSect = Trim$(Mid$(strLine, n2 + 7))
Else
    'get sector number
    n = InStr(strLine, "ment in")
    If n > 0 Then
        strSect = Trim$(Mid$(strLine, n + 7))
        owner = -2
    Else
        'get sector number
        n = InStr(strLine, "spotted in")
        If n > 0 Then
            n = n + 10
            n2 = InStr(strLine, " by ")
            strSect = Trim$(Mid$(strLine, n, n2 - n + 1))
            owner = -2
        Else
            n = InStr(strLine, "noises in")
            If n > 0 Then
                strSect = Trim$(Mid$(strLine, n + 9))
                owner = -2
            End If
        End If
    End If
End If

'get sector number
MarkSector strSect, owner, "Sub Spotted"
Exit Sub
Empire_Error:
EmpireError "BulletinSubDetected", vbNullString, strLine
End Sub
Public Sub BulletinFlewOver(strLine As String)
On Error GoTo Empire_Error
'flying over gold mine at 4,-14

Dim n As Integer
Dim strSect As String
Dim strChar As String
Dim strDes As String
Dim sx As Integer
Dim sy As Integer

'get the sector type
strDes = StringBetween(strLine, "over ", " at ")
If Len(strDes) = 0 Then
    Exit Sub
End If
    
'080304 rjk: Removed custom cases as should not be need any more
n = 1
Do
    If Mid$(SectDesString(n), 4, Len(strDes)) = strDes Then
        strChar = Left$(SectDesString(n), 1)
        Exit Do
    End If
    n = n + 1
Loop Until Len(SectDesString(n)) = 0

If Len(strChar) > 0 Then
    'get sector number
    strSect = StringBetween(strLine, " at ")
    If ParseSectors(sx, sy, strSect) Then
        'update bmap
        rsBmap.Seek "=", sx, sy
        If rsBmap.NoMatch Then
            rsBmap.AddNew
            rsBmap.Fields("x") = sx
            rsBmap.Fields("y") = sy
        Else
            rsBmap.Edit
            '050706 rjk: Do not erase minefields when flying over them.
            If rsBmap.Fields("des") = "X" And strChar = "." Then
                strChar = "X"
            End If
        End If
        rsBmap.Fields("des") = strChar
        rsBmap.Update
        
        'update sector
        rsEnemySect.Seek "=", sx, sy
        If Not rsEnemySect.NoMatch Then
            rsEnemySect.Edit
            rsEnemySect.Fields("des") = strChar
            rsEnemySect.Update
        End If
    End If
End If
Exit Sub
Empire_Error:
EmpireError "BulletinFlewOver", vbNullString, strLine

End Sub


'WinACE2 bombing raid did 53 damage in 0,4

'        support Values
'        forts   ships   units   planes
'attacker    0.67    0.00    0.00    0.00
'defender    0.23    0.00    0.00    0.00

'2 fighters are intercepting Cathouse planes!
'Escher     Cathouse    strength int odds  damage
' f2  #400   f2  #226      4/3    53  0.43   24/29  cleared
' f2  #401   f2  #226      4/2    29  0.33   12/17  cleared


Private Sub SectorLost(ByVal strSector As String, owner As Integer)
On Error GoTo Empire_Error

'Dim n As Integer       removed 8/03 ej
Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, strSector) Then

    'Get enemy Sector Information
    rsEnemySect.Seek "=", sx, sy
    If rsEnemySect.NoMatch Then
        rsEnemySect.AddNew
        rsEnemySect.Fields("x") = sx
        rsEnemySect.Fields("y") = sy
        rsEnemySect.Fields("eff") = -1
        rsEnemySect.Fields("road") = -1
        rsEnemySect.Fields("rail") = -1
        rsEnemySect.Fields("defense") = -1
        rsEnemySect.Fields("civ") = -1
        rsEnemySect.Fields("shell") = -1
        rsEnemySect.Fields("gun") = -1
        rsEnemySect.Fields("pet") = -1
        rsEnemySect.Fields("food") = -1
        rsEnemySect.Fields("iron") = -1
        rsEnemySect.Fields("bar") = -1
    Else
        rsEnemySect.Edit
    End If
    
    rsEnemySect.Fields("owner") = owner
    rsEnemySect.Fields("oldowner") = CountryNumber
    rsEnemySect.Fields("mil") = -1
    
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        'copy sector info to enemy info database
        rsEnemySect.Fields("des") = rsSectors.Fields("des")
        rsEnemySect.Fields("eff") = rsSectors.Fields("eff")
        rsEnemySect.Fields("road") = rsSectors.Fields("road")
        rsEnemySect.Fields("rail") = rsSectors.Fields("rail")
        rsEnemySect.Fields("defense") = rsSectors.Fields("defense")
        rsEnemySect.Fields("civ") = rsSectors.Fields("civ")
        rsEnemySect.Fields("shell") = rsSectors.Fields("shell")
        rsEnemySect.Fields("gun") = rsSectors.Fields("gun")
        rsEnemySect.Fields("pet") = rsSectors.Fields("pet")
        rsEnemySect.Fields("food") = rsSectors.Fields("food")
        rsEnemySect.Fields("iron") = rsSectors.Fields("iron")
        rsEnemySect.Fields("bar") = rsSectors.Fields("bar")
        
        'now delete sector record
        rsSectors.Delete
        RecievedNoticeOfMapChange = True
    End If
    
    'add/update enemy info record
    rsEnemySect.Fields("timestamp") = GetWinACEUTC '070304 rjk: Switched Enemy timestamp to UTC
    rsEnemySect.Update

End If
Exit Sub
Empire_Error:
EmpireError "SectorLost", vbNullString, strSector
End Sub

Public Sub MarkSector(strSector As String, Nation As Integer, Msg As String)
On Error GoTo Empire_Error

' Dim n As Integer    8/2003 efj  removed
Dim sx As Integer
Dim sy As Integer
Dim strNation As String

If ParseSectors(sx, sy, strSector) Then
    EventX = sx
    EventY = sy
    If Nation >= 0 Then
        strNation = " " + Nations.NationName(Nation)
        EventMarkers.Add sx, sy, Nation, Msg + strNation
    Else
        EventMarkers.Add sx, sy, Nation, Msg
    End If
    RecievedNoticeOfMapChange = True
End If

Exit Sub
Empire_Error:
EmpireError "MarkSector", strSector, Msg
End Sub

Public Sub EventMarkNukes()
On Error GoTo Empire_Error

'Mark all sectors with Nukes
If Not (rsNuke.EOF And rsNuke.BOF) Then
    EventMarkers.Clear
    rsNuke.MoveFirst
    While Not rsNuke.EOF
        MarkSector SectorString(rsNuke.Fields("x"), rsNuke.Fields("y")), -1, vbNullString
        rsNuke.MoveNext
    Wend
    frmDrawMap.DrawHexes
Else
    MsgBox "No Nukes found", vbOKOnly + vbExclamation, "Mark Nuke Request"
End If

Exit Sub
Empire_Error:
EmpireError "EventMarkNukes", vbNullString, vbNullString
End Sub

Public Sub EventStarvationNotice(strLine As String)
On Error GoTo Empire_Error
' -20,14   h *  100% will starve 507 people. 28 more food needed

Dim n As Integer
Dim n2 As Integer
Dim strSect As String
Dim strPeople As String
Dim strFood As String

'get sector number
If bDeity Then
    strLine = Trim$(strLine)
    n = InStr(strLine, " ")
    strLine = Trim$(Mid$(strLine, n))
    n = InStr(strLine, " ")
    strSect = Trim$(Left$(strLine, n))
Else
    strSect = Trim$(Left$(strLine, 10))
End If

'get number of people
n = InStr(strLine, "starve")
n = n + 7
n2 = InStr(n, strLine, " ")
strPeople = Mid$(strLine, n, n2 - n)

'get number of people
n = InStr(n2, strLine, "people")
n = n + 8
n2 = InStr(n, strLine, " ")
strFood = Mid$(strLine, n, n2 - n)

MarkSector strSect, -1, "Starvation " + strPeople + "p/" + strFood + "f"
Exit Sub
Empire_Error:
EmpireError "EventStarvationNotice", vbNullString, strLine
End Sub

Public Sub EventChe(strLine As String)
On Error GoTo Empire_Error
'Production minimally disrupted by terrorists in 2,-2
'Guerrilla warfare in -2,0
'  body count: troops: 1, rebels: 0
'Revolutionary subversion reported in 3,3!

' Dim n As Integer    8/2003 efj  removed
' Dim n2 As Integer    8/2003 efj  removed
Static strSect As String
Dim StrMsg As String
' Dim strFood As String    8/2003 efj  removed

'handle terrorist
If InStr(strLine, "terrorists") > 0 Then
    'get sector number
    strSect = StringBetween(strLine, "terrorists in ")
    StrMsg = Trim$(StringBetween(strLine, "Production ", " disrupted"))
    If Right$(StrMsg, 2) = "ly" Then
        StrMsg = Left$(StrMsg, Len(StrMsg) - 2)
    End If
    MarkSector strSect, -1, "Che: " + StrMsg
    strSect = vbNullString
End If

'handle revolutionaries
If InStr(strLine, "Revolutionary") > 0 Then
    'get sector number
    strSect = StringBetween(strLine, " in ", "!")
    StrMsg = Trim$(StringBetween(strLine, "Revolutionary ", " reported"))
    MarkSector strSect, -1, "Che: " + StrMsg
    strSect = vbNullString
End If

'Save Sector for Guerrilla warfare
If InStr(strLine, "Guerrilla warfare") > 0 Then
    'get sector number
    strSect = StringBetween(strLine, " in ")
End If

'mark after getting body count
If InStr(strLine, "body count") > 0 Then
    'get sector number
    If Len(strSect) > 0 Then
        StrMsg = Trim$(StringBetween(strLine, "troops: ", ","))
        StrMsg = "mil:" + StrMsg + " che:" + Trim$(StringBetween(strLine, "rebels:"))
        MarkSector strSect, -1, "Che: " + StrMsg
        strSect = vbNullString
    End If
End If

'mark sectors that are now fully ours in yellow
'Sector -28,-12 is now fully yours
If InStr(strLine, " is now fully yours") > 0 Then
    'get sector number
    strSect = Trim$(StringBetween(strLine, "Sector ", " is now fully yours"))
    MarkSector strSect, -3, "Now fully yours"
    strSect = vbNullString
End If

Exit Sub
Empire_Error:
EmpireError "EventChe", vbNullString, strLine
End Sub

Public Sub EventPlague(strLine As String)
On Error GoTo Empire_Error
'0,0 battling PLAGUE
'Outbreak of PLAGUE in -28,0!
'PLAGUE deaths reported in -24,0.
Dim strSect As String

'handle terrorist
If InStr(strLine, " battling") > 0 Then
    'get sector number
    strSect = Left$(strLine, InStr(strLine, " battling"))
    MarkSector strSect, -1, "Plague: Battling"
End If

'handle revolutionaries
If InStr(strLine, "Outbreak") > 0 Then
    'get sector number
    strSect = StringBetween(strLine, " in ", "!")
    MarkSector strSect, -1, "Plague: Outbreak"
End If

'mark after getting body count
If InStr(strLine, "deaths") > 0 Then
    strSect = Trim$(StringBetween(strLine, " in ", "."))
    MarkSector strSect, -1, "Plague: Deaths"
End If

Exit Sub
Empire_Error:
EmpireError "EventPlague", vbNullString, strLine
End Sub

Public Sub ExportTelegrams(ExportAll As Boolean)
On Error GoTo Empire_Error
Dim filenumber As Integer
Dim teleID As Integer
Dim ntele As Integer
Dim FileName As String
Dim n As Integer

If VerifySubDirectory("Exports", True) Then
    FileName = App.path + "\Exports\" + GameName + " Telegrams"
Else
    FileName = App.path + "\" + GameName + " Telegrams"
End If

'Partial extracts go to unique file names
n = 1
If Not ExportAll Then
    While Len(Dir(FileName + str$(n) + ".txt")) > 0
        n = n + 1
    Wend
    FileName = FileName + str$(n)
End If

FileName = FileName + ".txt"
filenumber = FreeFile   ' Get unused file number.
Open FileName For Output As #filenumber

rsTeleHead.Index = "PrimaryKey"
If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
    rsTeleHead.MoveFirst
    While Not rsTeleHead.EOF
        teleID = rsTeleHead.Fields("ID")
        If ExportAll Or rsTeleHead.Fields("Exported") = "N" Then
            ntele = ntele + 1
            Print #filenumber, rsTeleHead.Fields("Title")
            
            'Write out telegram body into file
            rsTeleBody.Seek "=", teleID, 1
            If Not rsTeleBody.NoMatch Then
                Do While (Not rsTeleBody.EOF)
                    If rsTeleBody.Fields("ID") <> teleID Then
                        Exit Do
                    End If
                    Print #filenumber, rsTeleBody.Fields("Body")
                    rsTeleBody.MoveNext
                Loop
            End If
            rsTeleHead.Edit
            rsTeleHead.Fields("Exported") = "Y"
            rsTeleHead.Update
        End If
        rsTeleHead.MoveNext
    Wend
    rsTeleHead.MoveLast
End If

Close #filenumber
MsgBox CStr(ntele) + " Telegrams exported to file " + FileName
Exit Sub
Empire_Error:
EmpireError "ExportTelegrams", vbNullString, vbNullString
End Sub


Public Sub ImportTelegrams()
On Error GoTo Empire_Error

Dim filenumber As Integer
' Dim teleID As Integer    8/2003 efj  removed
Dim FirstTeleID As Integer
Dim LastTeleID As Integer
Dim FileName As String
Dim strLine As String
Dim ntele As Integer
Dim NewTele As Boolean
Dim found As Boolean
Dim nFound As Integer
Dim X As Integer

On Error GoTo ImportTelegrams_Exit

ChDir App.path

Load frmCommonDialog
' Set CancelError is True
frmCommonDialog.cdb.CancelError = True
' Set flags
frmCommonDialog.cdb.Flags = cdlOFNHideReadOnly
' Set filters
frmCommonDialog.cdb.Filter = "All Files (*.*)|*.*|Text Files" & _
"(*.txt)|*.txt"
' Specify default filter
frmCommonDialog.cdb.FilterIndex = 2
' Display name of selected file
frmCommonDialog.cdb.FileName = vbNullString

If VerifySubDirectory("Exports", True) Then
    frmCommonDialog.cdb.InitDir = App.path + "\Exports"
End If

' Display the open dialog box
frmCommonDialog.cdb.ShowOpen

'frmCommonDialog.cdb.FileTitle

FileName = frmCommonDialog.cdb.FileName
filenumber = FreeFile   ' Get unused file number.
Open FileName For Input As #filenumber
frmReport.Caption = "Imported Telegrams"
frmReport.ClearReport

'Set which was the last telegram in the database
If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
    rsTeleHead.MoveLast
    FirstTeleID = rsTeleHead.Fields("ID") + 1
Else
    FirstTeleID = 1
End If

While Not EOF(filenumber)
    Line Input #filenumber, strLine
    NewTele = (Left$(strLine, 1) = ">") And ((InStr(strLine, " dated ") > 0) _
                Or (InStr(strLine, ">Telegram To ") > 0) Or (InStr(strLine, "> Saved") > 0))
    'Make sure canidate is a valid type
    If NewTele Then
        If ProcessTelegramHeader(strLine) = 0 Then
            NewTele = False
        End If
    End If
    
    If NewTele Then
        ntele = ntele + 1
        
        'search to see if this tele is already in our database
        found = False
        If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
            rsTeleHead.MoveFirst
            While Not rsTeleHead.EOF And Not found
'                teleID = rsTeleHead.Fields("ID")
                If Trim$(strLine) = Trim$(rsTeleHead.Fields("Title")) Then
                    found = True
                    nFound = nFound + 1
                End If
                rsTeleHead.MoveNext
            Wend
        End If
    End If
        
    'Check for air combat records
'    If Not AirCombatResults Then
'        If InStr(strLine, "strength int odds") > 0 Then
'            AirCombatResults = True
'            ParseAirCombat strLine
'        End If
'    Else
'        If InStr(mid$(strLine, 50, 15), "abort") > 0 Or _
'            InStr(mid$(strLine, 50, 15), "clear") > 0 Or _
'            InStr(mid$(strLine, 50, 15), "shot") > 0 Then
'            ParseAirCombat strLine
'        Else
'            AirCombatResults = False
'        End If
'    End If
    
    If Not found Then
        IncomingTelegramLine strLine, NewTele
        frmReport.AddReportLine strLine
    End If
Wend
Close #filenumber
frmReport.AddReportLine "Telegram Import"
frmReport.AddReportLine "Import File Name:   " + FileName
frmReport.AddReportLine "Telegrams read:     " + CStr(ntele)
frmReport.AddReportLine "Telegrams imported: " + CStr(ntele - nFound)

frmReport.Show vbModeless

'Set which was the last telegram in the database
If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
    rsTeleHead.MoveLast
    LastTeleID = rsTeleHead.Fields("ID")
Else
    LastTeleID = -1
End If

'Mark all the imported telegrams as already been exported.
For X = FirstTeleID To LastTeleID
    rsTeleHead.Seek "=", X
    If Not rsTeleHead.NoMatch Then
        rsTeleHead.Edit
        rsTeleHead.Fields("Exported") = "Y"
        rsTeleHead.Update
    End If
Next X

ImportTelegrams_Exit:
Unload frmCommonDialog

Exit Sub
Empire_Error:
EmpireError "ImportTelegrams", vbNullString, vbNullString
End Sub

Public Sub SendTelegram1(strHeader As String, strBody As String)
Dim strTarget() As String
Dim iTelegrams As Integer
Dim n As Integer
Dim X As Integer
Dim iCurrentLength As Integer
Dim iLineLength As Integer

ReDim strTarget(1 To 1) As String

iTelegrams = 1
iCurrentLength = 0

Dim strSource As String
Dim strTemp As String
Dim strTel As String

strSource = strBody
'remove carrage returns
n = InStr(strBody, vbCr)
While n > 0
    strBody = Left$(strSource, n - 1) + Mid$(strSource, n + 1)
    n = InStr(strSource, vbCr)
Wend

Do
    n = InStr(strSource, vbLf)
    If n = 0 Then n = Len(strSource) + 1
    strTemp = Left$(strSource, n - 1)
    iLineLength = UTF8Length(strTemp)
    strSource = Mid$(strSource, n + 1)
    
'    If chkSave.Value = vbChecked And Option1.Value Then
'         IncomingTelegramLine strTemp, False, True
'    End If
    
    'just in case we get a real long line, split it up
    While iLineLength > 1021
        n = 1021 - iCurrentLength
        strTarget(iTelegrams) = strTarget(iTelegrams) + MaxUTF8Length(strTemp, n)
        strTemp = Mid$(strTemp, n + 1)
        iCurrentLength = 0
        iTelegrams = iTelegrams + 1 '110303 rjk: Fixed to support lines longer than 2044
        ReDim Preserve strTarget(1 To iTelegrams)
    Wend
    
    If (iCurrentLength + iLineLength) >= 1022 Then
        iCurrentLength = 0
        iTelegrams = iTelegrams + 1
        ReDim Preserve strTarget(1 To iTelegrams)
    End If
    
    iCurrentLength = iCurrentLength + iLineLength + 1
    strTarget(iTelegrams) = strTarget(iTelegrams) + strTemp + vbLf

Loop While Len(strSource) > 0

'run thru multiple telegrams if necessary
For n = 1 To iTelegrams
    frmEmpCmd.SubmitEmpireCommand "te1 " + strTarget(n), False
    
    If frmTelegram.Option1.Value Then
        'Build recipient string
        strTel = vbNullString
        For X = 0 To frmTelegram.lbNation.ListCount - 1
            If frmTelegram.lbNation.Selected(X) Then
                strTel = strTel + CStr(frmTelegram.lbNation.ItemData(X)) + " "
            End If
        Next X
        frmEmpCmd.SubmitEmpireCommand "telegram " + strTel, False
    ElseIf frmTelegram.Option2.Value Then
        frmEmpCmd.SubmitEmpireCommand "announce ", False
    ElseIf frmTelegram.Option3.Value Then
        'Build recipient string
        strTel = vbNullString
        For X = 0 To frmTelegram.lbNation.ListCount - 1
            If frmTelegram.lbNation.Selected(X) Then
                strTel = CStr(frmTelegram.lbNation.ItemData(X))
            End If
        Next X
        frmEmpCmd.SubmitEmpireCommand "flash " + strTel, False
    ElseIf frmTelegram.optMotd.Value Then
        frmEmpCmd.SubmitEmpireCommand "turn mess ", False
    End If
Next n
End Sub

Public Function UnitDamage(strResponse As String) As Boolean
If Left$(strResponse, 1) <> Chr$(9) Or InStr(strResponse, " #") = 0 Or _
   InStr(strResponse, " takes ") = 0 Then
    UnitDamage = False
Else
    UnitDamage = True
End If
End Function
