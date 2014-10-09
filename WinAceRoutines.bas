Attribute VB_Name = "WinAceRoutines"
Option Explicit

'Change Log:
'010903 rjk: Added for additional exploration modes
'010903 rjk: Corrected OffsetSectorLocation Map wraparound
'010903 rjk: Corrected ExploreOut to proper account for the units being explored with
'280903 rjk: Switched Display const to enum and moved to frmDrawMap.
'            Added timestamp to ndump.
'            Fixed ndump parsing to skip country code in deity mode.
'051003 rjk: Removed timestamp ndump
'071003 rjk: Moved Nation list logic into relation report, removes the need for READ_STATE_REQ_REPORT,
'            and "report *"
'071003 rjk: In case the country name is different do nation first before relation.
'141003 rjk: Moved "show sect s" here from FormLoad frmDrawMap to be inside the database refresh
'            Added update to grid when processing lost items (scrap, scuttle)
'            Adding timestamp for ndump, processing lost nuke items.
'151003 rjk: Changed AddEnemyUnitInfo to preserve information from other sources.
'            Modified AddEnemySectInfo to preserve information from different intelligence sources
'171003 rjk: Added Strength fields to Sector database
'181003 rjk: Removed the check on EU_X and EU_Y in AddEnemyUnitInfo and AddEnemySectInfo as should always be known.
'            Removed update the EU_X and EU_Y in AddEnemySectInfo as it should not change.
'221003 rjk: Modified CopyGameInfo/CopySector: added road/rail/defense/mines, fixed stage 1 so headquarters and airport do
'            not overlap the harbor, modified stage 2 to use the sector offset, added owner to build to solve a problem
'            deity building units (uses owner 1 as fake owner for filler units) and unit dumps for stage 1 to 3.
'271003 rjk: Modified the ExportNation/edit command to work with syntax supported next version of the server.
'            Added detection in ProcessDump for when grid update is required and only update when required.
'281003 rjk: Always fill in the country field for rsSectors, rsShip, rsLand, rsPlane, rsNuke.
'281003 rjk: Added bridge span and tower to ExportNation - not complete.
'291003 rjk: Do not add records for ourselves, when importing intelligence
'301003 rjk: Updated SetOwner to make the timestamp format consistent
'            Moved ExportIntelligence to its own form
'            Moved upgrade for rsEnemySect and rsEnemyUnit timestamp to UpdateEnemyTimestamp
'            Added ImportIntelligenceTelegram for processing a intelligence report via a telegram.
'            Modified ImportIntelligence to ignore the header insert for the telegrams.
'311003 rjk: Added String Definition for hldIndex.
'011103 rjk: Added Friendly Plane GIF.
'031103 rjk: Added Neutral Plane and Land Unit GIF.
'041103 rjk: Changed the pos Variable to a Long to accomodate large Intelligence Reports in ImportIntelligenceTelegram
'141103 rjk: When we get a new unit, ensure there is not an old enemy intelligence record with the same id, if present
'            delete it.
'181103 rjk: Import Intelligence, changed processing to ignore headers.
'            Switched to a Yes/No question for Clearing Enemy Intelligence
'201103 rjk: Update the sector panel if a new strength report comes in for the sector being displayed
'            Added option to control strength updates
'            Added protection to not overflow the sec_def and tot_def fields in the database (LOTR II)
'241103 rjk: In ParseStrength, moved the sector panel until after the record is updated.
'281103 rjk: Moved Unit picture loading EmpUnitView
'291103 rjk: Added READ_STATE_NEWSPAPER to gather intelligence from the newspaper.
'291103 rjK: Remove the question on clearing the intelligence as there is a separate form now.
'            Can now delete individual items when clear the intelligence.
'011203 rjk: Added READ_STATE_PLAYERS for displaying current players
'031203 rjk: Added detection of enemy mines in ParseStrength
'071203 rjk: Changed hldindex to hldIndex.
'121203 rjk: Added Items, a collection of item, use to store information and characteristics
'            about the current items in this game.
'291203 rjk: Need to ensure there are entries in the Nuke table before ProcessLostDump entries
'250104 rjk: Added y (St@r W@rs) to the production report.
'010204 rjk: Switched to "PrimaryKey" (id) index for processing nuke dumps and losses.
'040204 rjk: Switched to ProcessDump to use the header for processing.
'070204 rjk: Changed ProcessLostDump to process the nuke id with type.
'200204 rjk: Check for a quoted string at the end of the list - ship name, combined if across multiple tokens
'210204 rjk: Fixed a bug in the string processing for ProcessDump.
'260204 rjk: Added show sect c/b to ensure the database is fill, as it starts empty
'070304 rjk: Switch Enemy timestamp to UTC
'070304 rjk: Switch to CSng Val for regional settings
'120304 rjk: Skip the invalid nations (-1) in AirCombatReport
'270304 rjk: Switched over to DeleteAllRecords for clearing a table
'030404 rjk: Switched to use the avail from the server instead of calculating it,
'080404 rjk: Added Null check to lmines.
'240404 rjk: Fixed ImportIntelligence to skip our sectors.
'260604 rjk: Fixed error when a satellite report detects changing owner for an enemy unit.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'020704 rjk: Clear any enemy sect intelligence if a sector is turned into a wasteland
'300704 rjk: Added Days Range to ClearEnemyInfo.
'220904 rjk: Modified GetCurrentStrength to not request strength of sea sectors when the dump suppresses sea sectors.
'011204 rjk: Restricted ImportIntelligence to existing files.
'011204 rjk: Added sdes and fixed lmines for CopySector function.
'220305 rjk: Added 'Version' to Version table.
'220305 rjk: Added SubmitTelegramRead to single function for requesting telegrams.
'            Added version check to deal syntax change in 4.2.21.
'240405 rjk: Added Three custom script buttons.
'100505 rjk: Changed Three Three custom script buttons to Ten.
'250505 rjk: Added Telegram Warning, Soft Limit and Hard Limit.
'160705 rjk: Added http (XmlRpc) proxy.  Added Get Oil Content.
'120905 rjk: Changed GetOilContent to only do oil content for oe/od ships.
'180206 rjk: Added GetShips, GetNukes, GetLandUnits, GetPlanes, GetLost, GetNation,
'            GetCountryList, IsArrayEmpty to support xdump for 4.3.0.  Added xdump
'            support to SendFullDump.  Added meta table store for 4.3.0.
'            Replace ndump, sdump, pdump, ldump, nation, relation and list with
'            GetNukes, GetShips, GetPlanes, GetLandUnits, GetNation, GetNation and
'            GetCountryList.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'210306 rjk: Moved the RequestMetaTables to DrawMap form loading.
'220406 rjk: Enhanced ClearEnemyInfo to clear AirCombat records by date.
'230406 rjk: Added Flash Logging.
'140506 rjk: Use shells for lcms for South Pacific Atlantis.
'220506 rjk: Switched NSC_STRINGY to be value extracted from meta-type table.
'            Incorporate 4.3.4 server changes for xdump nat and coun.
'            SP: Atlantis also uses changes for xdump nat and coun.
'270506 rjk: Added relations for 4.3.0 to 4.3.3 servers and for SP: Atlantis.
'280506 rjk: Fixed tech production for SP: Atlantis for ProdSummaryReport.
'110606 rjk: Enhanced script buttons to support jump points.
'050706 rjk: Added a check for sea mines when exploring.  Do not attempt to
'            explore into sea with mines.
'090806 rjk: Fixed ProdSummaryReport, it was reporting shortages of lcms for building
'            unit.  There was a mistake in the bSPAtlantis code.
'230806 rjk: Added Import/Export Sea Information
'181006 rjk: Added Import Offset to ImportIntelligence.
'181006 rjk: Made CheckForOldEnemyUnit public for use XDumpParse.
'090607 rjk: In ExportNation(), added Land Unit counts and Plane counts
'            for 4.3.6 servers. Fill in zeros for lmines when they
'            are not known.
'090907 rjk: Added code for using xdump game and xdump update to get
'            next update for server 4.3.10 and newer.  New functions
'            GetUpdate, UpdateCommands() and StartUpdateTimer().
'010108 rjk: Add GetOrders function to allow the use of xdump for ship orders
'            for the newer servers.
'120108 rjk: Removed unused local variable tsDump.
'260408 rjk: Added 'q' for Heavy Metal II biofuel plant to Production Summary
'120508 rjk: Added '~' for Heavy Metal II to Production Summary
'240508 rjk: Fixed Production Summary Report to use build lcm and hcm for city production.
'010608 rjk: Fixed Production Summary Report to properly calculate the exceed civs for city
'            and other sectors.
'140608 rjk: Added packing_type to sector chr tofix the packing factor for mobility calculations.
'            Added additional debugging for production summary overflow. Fixed 'q' to draw
'            production ratio from sector chr and corrected the default ratios.
'020708 rjk: Fixed ApplySectorDamage when the enemy efficiency is not unknown, fixes
'            Invalid use of Null errors when shelling.
'120708 rjk: Fixed Excess Civ bug for Big Cities.
'250409 rjk: Do not include stopped units in the production summary.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

'EmpCmd Constants for seting expected empire results
Public Const READ_STATE_FIRSTSERVERINIT = 1
Public Const READ_STATE_USERCMDOK = 2
Public Const READ_STATE_COUNCMDOK = 3
Public Const READ_STATE_PASSCMDOK = 4
Public Const READ_STATE_KILLPROC = 5
Public Const READ_STATE_PLAYINIT = 6
Public Const READ_STATE_SECTOR_DUMP = 7
Public Const READ_STATE_PLANE_DUMP = 8
Public Const READ_STATE_LAND_DUMP = 9
Public Const READ_STATE_SHIP_DUMP = 10
Public Const READ_STATE_FULL_BMAP = 11
Public Const READ_STATE_REPORT = 12
'Public Const READ_STATE_REQ_REPORT = 13 100703 rjk: Not required as Nation list is done in relations now.
Public Const READ_STATE_REQ_RELATIONS = 14
Public Const READ_STATE_EXPLORE_CIRCLE = 15
Public Const READ_STATE_UNIT_MOVE = 16
Public Const READ_STATE_NEXT_UPDATE = 17
Public Const READ_STATE_FIRE_SEQ = 18
Public Const READ_STATE_LAUNCH = 19
Public Const READ_STATE_BOMB = 20
Public Const READ_STATE_TELEGRAM_READ = 21
Public Const READ_STATE_TELEGRAM_WRITE = 22
Public Const READ_STATE_LOOK = 23
Public Const READ_STATE_SPY = 24
Public Const READ_STATE_RECON = 25
Public Const READ_STATE_COASTWATCH = 26
Public Const READ_STATE_ATTACK = 27
Public Const READ_STATE_VERSION = 28
Public Const READ_STATE_LOST_DUMP = 29
Public Const READ_STATE_ORIGIN = 30
Public Const READ_STATE_LISTSANC = 31
Public Const READ_STATE_STARVATION = 32
Public Const READ_STATE_SONAR = 33
Public Const READ_STATE_SHOW = 34
Public Const READ_STATE_PARA = 35
Public Const READ_STATE_BREAK = 36
Public Const READ_STATE_NATION = 37
Public Const READ_STATE_NUKE_DUMP = 38
Public Const READ_STATE_BOMBCL = 39
Public Const READ_STATE_ORDER = 40
Public Const READ_STATE_SHOWORDER = 41
Public Const READ_STATE_LANDMISSION = 42
Public Const READ_STATE_SHIPMISSION = 43
Public Const READ_STATE_PLANEMISSION = 44
Public Const READ_STATE_EDIT = 45
Public Const READ_STATE_TELEGRAM_READ_CL = 46
Public Const READ_STATE_WIRE = 47
Public Const READ_STATE_STRENGTH = 48
Public Const READ_STATE_NEWSPAPER = 49 '112903 rjk: Added to gather intelligence from the newspaper.
Public Const READ_STATE_PLAYERS = 50 '120103 rjk: Added for displaying current players
Public Const READ_STATE_SCUTTLE = 51 '040404 rjk: Added for automatic answer to scuttle and scrap
Public Const READ_STATE_XDUMP = 52 '280804 rjk: Added xdump commands
Public Const READ_STATE_OPTIONS = 53 '280505 rjk: Added for login options (UTF-8).
Public Const READ_STATE_FLASH_WRITE = 54

'Public Const TYPE_ALL = 0 'moved to frmDrawMap and made into an enum
'Public Const TYPE_CAPITAL = 1
'Public Const TYPE_ESCORT = 2
'Public Const TYPE_SUPPORTSHIP = 3
'Public Const TYPE_SUB = 4
'Public Const TYPE_CARRIER = 5
'Public Const TYPE_TROOPSHIPS = 6
'Public Const TYPE_MERCHANTS = 7
'Public Const TYPE_FIGHTERS = 8
'Public Const TYPE_TACTICAL = 8
'Public Const TYPE_BOMBERS = 9
'Public Const TYPE_ESCORTPLANES = 10
'Public Const TYPE_TRANSPORTS = 11
'Public Const TYPE_RECON = 12
'Public Const TYPE_MISSILES = 13
'Public Const TYPE_SATELLITES = 14
'Public Const TYPE_CHOPPER = 15
'Public Const TYPE_ANTISUB = 16
'Public Const TYPE_COMBAT = 17
'Public Const TYPE_ARTILLERY = 18
'Public Const TYPE_SUPPORT = 19
'Public Const TYPE_ENGINEERS = 20
''Public Const TYPE_OTHER = 21
'Public Const TYPE_INTERCEPTORS = 21
'Public Const TYPE_CARGO = 22
'Public Const TYPE_SWEEP = 23
'Public Const TYPE_MINE = 24
'Public Const TYPE_PARATROOPS = 25
'Public Const TYPE_ALL_ESCORTS = 26
'Public Const TYPE_ALL_BOMBERS = 27
'Public Const TYPE_MISSILES_AND_SATELLITES = 28
'Public Const TYPE_RESTORE = 99

Public Const BF_NEVERABORT = 0

Public Const BF_ABORTONRETURNFIRE = 1
Public Const BF_ABORTREMAINING = 2

'Public Const DSP_SHIP = 1
'Public Const DSP_LAND = 2
'Public Const DSP_PLANE = 3
'Public Const DSP_ENEMY = 4
'Public Const DSP_LIST = 5

'Public Const DSP_ALL = 1 'moved to frmDrawMap and made into an enum
'Public Const DSP_FLEET = 2
'Public Const DSP_SECT = 3
'Public Const DSP_PLANE_RANGE = 4
'Public Const DSP_MISSILE_RANGE = 5
'Public Const DSP_ATTACK_RANGE = 6

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public Declare Function GetTimeZoneInformation Lib "kernel32" _
        (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
Private Const TIME_ZONE_ID_STANDARD& = 1
Private Const TIME_ZONE_ID_UNKNOWN& = 0
Private Const TIME_ZONE_ID_DAYLIGHT& = 2

Public tz As TIME_ZONE_INFORMATION

'Public Const MaxExploreRadius = 10 'why so small? drk@unxsoft.com 7/15/02 fixme
Public Const MaxExploreRadius = 20
Enum enumExploreMode  '09/01/03 rjk: Added for additional exploration modes
    EXP_COMPLETE = 1
    EXP_BORDER = 2
    EXP_WHEEL = 3
End Enum

Enum enumVersion  '21/03/05 rjk: Added for Version Check Function
    VER_EXACT = -1
    VER_LESS = 0
    VER_MORE = 1
End Enum

Public MapSizeX As Integer
Public MapSizeY As Integer
Public bDeity As Boolean
Public bUTF8 As Boolean
Public bXMLRPC As Boolean
Public xmlrpcServer As EmpXMLRPC

Public TimeIdle As Long

Public picGreenLight As Picture
Public picRedLight As Picture
Public picBothLights As Picture

Public Nations As New EmpNationList
Public Items As EmpItems
Public etsTable As EmpTables
Public estsTable As EmpSymbolTables
Public CountryNumber As Integer

Public BatchScript1File As String
Public BatchScript1Time As Long
Public BatchScriptUpdate As String

Public FlashLogFileNumber As Integer

Enum enumTelegramSound
    tsNoSound = 0
    ts1Beep = 1
    ts2Beeps = 2
    ts3Beeps = 3
End Enum

Enum enumTelegramOptionType
    totTelegram = 0
    totAnnouncement = 1
    totFlash = 2
    totMOTD = 3
End Enum

Enum enumTelegramLevel
    tlWarning = 0
    tlSoftLimit = 1
    tlHardLimit = 2
End Enum

Public Type tTelegram
    bEnabled As Boolean
    iColumn As Integer
    eTelegramSound As enumTelegramSound
End Type
    
Public tTelegramOptions(totTelegram To totMOTD, tlWarning To tlHardLimit) As tTelegram

Public Type tScriptButton
    strFileName As String
    bSendOutputReportWindow As Boolean
    strImageFileName As String
    bJumpPoint As Boolean
    iJumpPoint As Integer
End Type
Public Const NUMBER_SCRIPT_BUTTONS = 10

Public tScriptButtons(1 To NUMBER_SCRIPT_BUTTONS) As tScriptButton

Public Sub LoadBmap(buf As String, StartX As Integer, linelen As Integer)
Dim CurY As Integer
Dim CurX As Integer
Dim X As Integer
Dim fieldname As String
Dim BufferLen As Integer
Dim hldIndex As String

On Error GoTo LoadBmap_Exit

hldIndex = rsEnemySect.Index
rsEnemySect.Index = "PrimaryKey"

CurY = Val(Left$(buf, 4))

'Figure out how far bmap goes
BufferLen = Len(buf)
If linelen < BufferLen Then
    BufferLen = linelen
End If

If Left$(buf, 4) = "    " Then
    Exit Sub
End If

For X = 5 To BufferLen
    If Mid$(buf, X, 1) <> " " Then
        CurX = StartX + X - 6
        rsBmap.Seek "=", CurX, CurY
        If rsBmap.NoMatch Then
            rsBmap.AddNew
            fieldname = "x"
            rsBmap.Fields(fieldname) = CurX
            fieldname = "y"
            rsBmap.Fields(fieldname) = CurY
            fieldname = "des"
            rsBmap.Fields(fieldname) = Mid$(buf, X, 1)
            rsBmap.Update
        ElseIf rsBmap.Fields("des") <> Mid$(buf, X, 1) Then
            rsBmap.Edit
            fieldname = "des"
            rsBmap.Fields(fieldname) = Mid$(buf, X, 1)
            rsBmap.Update
        End If
        If Mid$(buf, X, 1) = "\" Then '020704 rjk: Clear any enemy sect intelligence if a sector is turned into a wasteland
            rsEnemySect.Seek "=", CurX, CurY
            If Not rsEnemySect.NoMatch Then
                rsEnemySect.Delete
            End If
        End If
    End If
Next X
rsEnemySect.Index = hldIndex
Exit Sub

'Error handling routine
LoadBmap_Exit:
On Error Resume Next
If Len(hldIndex) > 0 Then
    rsEnemySect.Index = hldIndex
End If
EmpireError "LoadBmap", fieldname, buf
End Sub

Public Sub ImportBmap()
On Error GoTo Empire_Error
Dim CurY As Integer
Dim CurX As Integer
Dim X As Integer
Dim BufferLen As Integer
Dim filenumber As Integer
Dim FileName As String
Dim bdes As String


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
On Error GoTo ImportBmap_Exit
frmCommonDialog.cdb.ShowOpen
On Error GoTo Empire_Error

'frmCommonDialog.cdb.FileTitle

FileName = frmCommonDialog.cdb.FileName
filenumber = FreeFile   ' Get unused file number.
Open FileName For Input As #filenumber

Dim Line As Integer
Dim bmaphead1 As String
Dim bmaphead2 As String
Dim bmaphead3 As String
Dim StartX As Integer
Dim AdjX As Integer
Dim AdjY As Integer
Dim temp As Integer
Dim linelen As Integer
Dim overwrite As Boolean
Dim strLine As String
Dim strSect As String

start:
strSect = InputBox("Enter your x,y co-ordinates for the imported bmaps origin")
If Not ParseSectors(AdjX, AdjY, strSect) Then
    If MsgBox("Sectors not valid", vbRetryCancel + vbCritical, "Import Bmap Error") = vbRetry Then
        GoTo start
    End If
    Exit Sub
End If

overwrite = (MsgBox("Do you want to overwrite your existing bmap", vbYesNo + vbQuestion, "Import Bmap") = vbYes)
    
While Not EOF(filenumber)
    Line Input #filenumber, strLine
    Line = Line + 1
    If Line = 1 Then
        bmaphead1 = strLine
        linelen = Len(RTrim$(bmaphead1))
    ElseIf Line = 2 Then
        bmaphead2 = strLine
        'Calculate the starting x
        StartX = Val(Mid$(bmaphead1, 6, 1)) * 10 + Val(Mid$(bmaphead2, 6, 1))
        temp = Val(Mid$(bmaphead1, 7, 1)) * 10 + Val(Mid$(bmaphead2, 7, 1))
        If temp < StartX Then
            StartX = -1 * StartX
        End If
    ElseIf Line = 3 And Left$(strLine, 4) = "    " Then
        'Handle three line bmap - recompute start x
        bmaphead3 = strLine
        StartX = Val(Mid$(bmaphead1, 6, 1)) * 100 + Val(Mid$(bmaphead2, 6, 1)) * 10 + Val(Mid$(bmaphead3, 6, 1))
        temp = Val(Mid$(bmaphead1, 7, 1)) * 100 + Val(Mid$(bmaphead2, 7, 1)) * 10 + Val(Mid$(bmaphead3, 7, 1))
        If temp < StartX Then
            StartX = -1 * StartX
        End If
    Else
        CurY = Val(Left$(strLine, 4)) + AdjY
        
        'Figure out how far bmap goes
        BufferLen = Len(strLine)
        If linelen < BufferLen Then
            BufferLen = linelen
        End If
        
        If Left$(strLine, 4) <> "    " Then
            For X = 5 To BufferLen
                If Mid$(strLine, X, 1) <> " " Then
                    CurX = StartX + X - 6 + AdjX
                    
                    'Adjust the x and y to compesate for
                    'going over the boundry
                    OffsetSectorLocation CurX, CurY, 0, 0
                    
                    'process sector by sector
                    bdes = Mid$(strLine, X, 1)
                    If bdes = "?" Then
                        bdes = "0"
                    End If
                    
                    rsBmap.Seek "=", CurX, CurY
                    If rsBmap.NoMatch Then
                        frmEmpCmd.SubmitEmpireCommand "bdes " + SectorString(CurX, CurY) _
                             + " " + bdes, False
                    ElseIf rsBmap.Fields("des") <> bdes Then
                        'check and make sure land and water line up with
                        'current bmap
                        If (InStr(".=X", rsBmap.Fields("des")) > 0 And _
                            InStr(".=X", bdes) = 0) Or _
                           (InStr(".=X", rsBmap.Fields("des")) = 0 And _
                            InStr(".=X", bdes) > 0) Then
                            
                            If MsgBox("Warning.  Possible error in bmap lineup. " + _
                             "Current Bmap shows '" + rsBmap.Fields("des") + "', imported bmap shows '" + _
                             bdes + "' Do you want to continue?", vbExclamation + vbYesNo, "Import Bmap") _
                              = vbNo Then Exit Sub
                        End If
                        If overwrite Then
                            frmEmpCmd.SubmitEmpireCommand "bdes " + SectorString(CurX, _
                                CurY) + " " + bdes, False
                        'alway mark map if minefield
                        ElseIf rsBmap.Fields("des") = "." And bdes = "X" Then
                            frmEmpCmd.SubmitEmpireCommand "bdes " + SectorString(CurX, _
                                CurY) + " " + bdes, False
                        End If
                    End If
                End If
            Next X
        End If
    End If
Wend
Exit Sub

'Error handling routine
ImportBmap_Exit:
Exit Sub
Empire_Error:
EmpireError "ImportBmap", vbNullString, vbNullString

End Sub

Public Sub RefreshDataBase()
On Error GoTo Empire_Error
frmEmpCmd.ForceUpdates = True

'Send Code to start database update
frmEmpCmd.SubmitEmpireCommand "db1", False

'If version information needs filling, request version
If rsVersion.BOF And rsVersion.EOF Then
    frmEmpCmd.SubmitEmpireCommand "version", False
End If

If VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
    GetNation
    GetCountryList
Else
    GetCountryList
    GetNation
End If
If Not VersionCheck(4, 3, 0) = VER_LESS Then
    RequestConfigurationTables
Else
    'used to calculate the max safe civ at connection when autoupdate is on
    frmEmpCmd.SubmitEmpireCommand "show sect s", False
    '022604 rjk: Added show sect c/b to ensure the database is fill, as it starts empty
    frmEmpCmd.SubmitEmpireCommand "show sect c", False
    frmEmpCmd.SubmitEmpireCommand "show sect b", False
End If

'Request an update of Sector information
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
GetOContent '110605 rjk: Added ability to get OContent for Sea Sectors

'Request an update of Plane information
GetPlanes
'Request an update of land information
GetLandUnits
'Request an update of ship information
GetShips

'Request an update of nuke information
GetNukes

'Request update lost information
GetLost

'Request an update of ship information
frmEmpCmd.SubmitEmpireCommand "bmap *", False

'send Code to end database update
frmEmpCmd.SubmitEmpireCommand "db2", False
Exit Sub
Empire_Error:
EmpireError "RefreshDataBase", vbNullString, vbNullString
End Sub
Public Function ProcessSectorDump(strLine As String, ByRef CurX As Integer, ByRef CurY As Integer) As String

Dim strTokens() As String
' Dim top As Integer    8/2003 efj  removed
Dim numTokens As Integer
' Dim i As Integer    8/2003 efj  removed
Dim n As Integer
Dim hldIndex As String
Dim var As String
On Error GoTo ProcessSectorDump_Exit

'check for an announcement or tele message in the middle of a dump
If InStr(strLine, "You have") > 0 Then
    ProcessSectorDump = vbNullString
    Exit Function
End If

'Save and restore the index on entry and exit
hldIndex = rsSectors.Index
rsSectors.Index = "PrimaryKey"

ParseDumpString strLine, strTokens(), bDeity

numTokens = UBound(strTokens)
For n = LBound(strTokens) To numTokens
    If n > 3 Or n = 2 Then
        var = rsSectors.Fields(n).Name
        rsSectors.Fields(n) = strTokens(n)
    ElseIf n = 3 Then
        var = rsSectors.Fields(n).Name
        If strTokens(n) = "_" Then
            strTokens(n) = " "
        End If
        rsSectors.Fields(n) = strTokens(n)
    ElseIf n = 0 Then
        CurX = CInt(strTokens(n))
    Else
        CurY = CInt(strTokens(n))
        rsSectors.Seek "=", CurX, CurY
        If rsSectors.NoMatch Then
            rsSectors.AddNew
        Else
            rsSectors.Edit
        End If
        var = rsSectors.Fields(0).Name
        rsSectors.Fields(0) = CurX
        rsSectors.Fields(1) = CurY
    End If
Next n
If Not bDeity Then '102803 rjk: Always fill in the country field
    rsSectors.Fields("country") = CountryNumber
End If
rsSectors.Update
hldIndex = rsSectors.Index
ProcessSectorDump = strTokens(0) + "," + strTokens(1)
Exit Function

'Error handling routine
ProcessSectorDump_Exit:
EmpireError "ProcessSectorDump", var, strLine
ProcessSectorDump = vbNullString
End Function

Public Sub ParseStrength(strLine As String)
Dim sx As Integer
Dim sy As Integer
Dim hldIndex As String
Dim var As String
On Error GoTo ParseStrength_Exit

'[##:##] str #1
'
'
'Mon Oct 23 22:44:16 1995
'DEFENSE STRENGTH               land  sect   sector  reacting    total
'  sect       eff  mil  units  mines  mult  defense     units  defense
'   5,-5   * 100%   95                1.00       95        48      143
'   4,-4   c 100%                     4.00        0       192      192
'   6,-4   f   0%   78     48         2.00      252                252
'  -1,-3   c 100%                     4.00        0        80       80
'   1,-3   c 100%                     4.00        0        80       80
'   3,-3   c 100%                  ?  4.00        0        80       80
'   7,-3   ! 100%   53            90  2.80      148       134      282
'7 sectors
'1234567890123456789012345678901234567890123456789012345678901234567890

'check for an announcement or tele message in the middle of a dump
If InStr(strLine, "You have") > 0 Then
    Exit Sub
End If

If bDeity Then 'the first four characters are used for the country
    strLine = Mid$(strLine, 5)
End If

If InStr(strLine, ",") = 0 Then 'Ensure it is a sector line
    Exit Sub
End If

'Save and restore the index on entry and exit
hldIndex = rsSectors.Index
rsSectors.Index = "PrimaryKey"

If Not ParseSectors(sx, sy, Left$(strLine, 9)) Then
    EmpireError "ParseStrength", "Not valid sector", strLine
    Exit Sub
End If

rsSectors.Seek "=", sx, sy
If rsSectors.NoMatch Then
    EmpireError "ParseStrength", "Sector not found", strLine
    Exit Sub
End If

rsSectors.Edit
var = "units"
rsSectors.Fields("units") = Val(Mid$(strLine, 23, 6))
var = "lmines"
If Mid$(strLine, 35, 1) = "?" Then '120303 rjk: display enemy mines in red "?"
    rsSectors.Fields("lmines") = -1
Else
    rsSectors.Fields("lmines") = Val(Mid$(strLine, 30, 6))
End If
var = "sec_mult"
rsSectors.Fields("sec_mult") = Val(Mid$(strLine, 37, 5))
var = "sec_def"
If Val(Mid$(strLine, 43, 8)) > 32767 Then '112003 rjk: Added protection to not overflow the db variable (LOTR II)
    rsSectors.Fields("sec_def") = 32767
Else
    rsSectors.Fields("sec_def") = Val(Mid$(strLine, 43, 8))
End If
var = "runits"
rsSectors.Fields("runits") = Val(Mid$(strLine, 52, 9))
var = "tot_def"
If Val(Mid$(strLine, 62, 8)) > 32767 Then '112003 rjk: Added protection to not overflow the db variable (LOTR II)
    rsSectors.Fields("tot_def") = 32767
Else
    rsSectors.Fields("tot_def") = Val(Mid$(strLine, 62, 8))
End If

rsSectors.Update
hldIndex = rsSectors.Index

'112003 rjk: Update the sector panel if a new strength report comes in for the sector being displayed
'112403 rjk: Moved to record is actually up to date.
If sx = frmDrawMap.SelX And sy = frmDrawMap.SelY Then
    frmDrawMap.FillSectorBox sx, sy
End If

Exit Sub

'Error handling routine
ParseStrength_Exit:
EmpireError "ParseStrength", var, strLine
End Sub

Public Sub ProcessLostDump(strLine As String)

Dim strTokens() As String
' Dim top As Integer    8/2003 efj  removed
' Dim numTokens As Integer    8/2003 efj  removed
' Dim CurX As Integer    8/2003 efj  removed
' Dim CurY As Integer    8/2003 efj  removed
' Dim i As Integer    8/2003 efj  removed
' Dim n As Integer    8/2003 efj  removed
On Error GoTo ProcessLostDump_Exit

'Since Deity gets all information, he doesn't process a lost dump.
'Simplest way around problem of items lost to player A were deleting
'entries for the same item that were rebuilt by player B
If bDeity Then
    Exit Sub
End If

'check for an announcement or tele message in the middle of a dump
If InStr(strLine, "You have") > 0 Then
    Exit Sub
End If

ParseDumpString strLine, strTokens(), bDeity
' numTokens = UBound(strTokens)    8/2003 efj  removed

Select Case strTokens(LS_TYPE)
    Case LT_SECTOR
        rsSectors.Seek "=", CInt(strTokens(LS_X)), CInt(strTokens(LS_Y))
        If Not rsSectors.NoMatch Then
            rsSectors.Delete
        End If
    Case LT_SHIP
        rsShip.Seek "=", CInt(strTokens(LS_ID))
        If Not rsShip.NoMatch Then
            rsShip.Delete
        End If
    Case LT_PLANE
        rsPlane.Seek "=", CInt(strTokens(LS_ID))
        If Not rsPlane.NoMatch Then
            rsPlane.Delete
        End If
    Case LT_LAND
        rsLand.Seek "=", CInt(strTokens(LS_ID))
        If Not rsLand.NoMatch Then
            rsLand.Delete
        End If
    Case LT_NUKE '101403 rjk: add to process lost nukes, needed for timestamp
        rsNuke.Seek "=", CInt(strTokens(LS_ID))
        While Not rsNuke.NoMatch
            rsNuke.Delete
            rsNuke.Seek "=", CInt(strTokens(LS_ID))
        Wend
End Select

'101403 rjk: added to update grid when doing scrap/scuttle operations.
If frmDrawMap.SelX = CInt(strTokens(LS_X)) And frmDrawMap.SelY = CInt(strTokens(LS_Y)) Then
    If strTokens(LS_TYPE) = LT_SHIP Or strTokens(LS_TYPE) = LT_PLANE Or _
       strTokens(LS_TYPE) = LT_LAND Then
        frmDrawMap.FillGrid
        frmDrawMap.Refresh
    Else
        frmDrawMap.FillSectorBox frmDrawMap.SelX, frmDrawMap.SelY
    End If
End If

Exit Sub

'Error handling routine
ProcessLostDump_Exit:
EmpireError "ProcessLostDump", vbNullString, strLine

End Sub

'102703 rjk: Modified to detect if a fill grid is required.
Public Function ProcessDump(strLine As String, ByRef rs As Recordset, strFields() As String) As Boolean
Dim strTokens() As String
Dim strLineFiltered As String
Dim n As Integer
Dim id As Long
Dim hldIndex As String
Dim var As String
Dim bQuote As Boolean

ProcessDump = False '102703 rjk: Added to detect when grid is required.

On Error GoTo ProcessDump_Exit

'check for an announcement or tele message in the middle of a dump
If InStr(strLine, "You have") > 0 Then
    Exit Function
End If

'Save and restore the index on entry and exit
hldIndex = rs.Index
rs.Index = "PrimaryKey"

'ParseDumpString strLine, strTokens(), bDeity
'Remove extra spaces, 210204 rjk: Fix a bug
bQuote = False
strLine = Trim$(strLine)
strLineFiltered = Left$(strLine, 1)
For n = 2 To Len(strLine)
    If Mid$(strLine, n, 1) = Chr$(34) Then
        bQuote = Not bQuote
        strLineFiltered = strLineFiltered + Chr$(34)
    Else
        If Mid$(strLine, n - 1, 1) <> " " Or Mid$(strLine, n, 1) <> " " Or bQuote Then
            strLineFiltered = strLineFiltered + Mid$(strLine, n, 1)
        End If
    End If
Next n
strTokens = Split(strLineFiltered)

'200204 rjk: Check for a quoted string at the end of the list (ship name), combined if across multiple tokens
If InStr(strTokens(UBound(strTokens)), Chr$(34)) > 0 Then
    Dim strBuild As String
    n = UBound(strTokens)
    While Left$(strTokens(n), 1) <> Chr$(34)
        strBuild = " " + strTokens(n) + strBuild
        n = n - 1
    Wend
    strTokens(n) = strTokens(n) + strBuild
    ReDim Preserve strTokens(n)
End If

If UBound(strTokens) <> UBound(strFields) Then
    EmpireError "ProcessDump", "Different number of fields than the header", strLine
    Exit Function
End If

If strFields(0) = "own" Then
    n = 1 'Process the country at the end, do not have a record yet
Else
    n = LBound(strTokens)
End If
Do
    var = strFields(n)
    If strFields(n) = "id" Then
        id = CLng(strTokens(n)) '111403 rjk: id is long so switch from CInt to CLng
        rs.Seek "=", id
        If rs.NoMatch Then
            rs.AddNew
            CheckForOldEnemyUnit id, Left$(rs.Name, 1) '111403 rjk: Ensure there is not an old enemy intelligence for this ship.
        Else
            rs.Edit
        End If
    End If
    Select Case rs.Fields(strFields(n)).Type
    Case dbInteger
        rs.Fields(strFields(n)).Value = CInt(strTokens(n))
    Case dbLong
        rs.Fields(strFields(n)).Value = CLng(strTokens(n))
    Case dbSingle
        rs.Fields(strFields(n)).Value = Val(strTokens(n)) '070304 rjk: Switch to Val for regional settings
    Case dbText
        If strFields(n) = "name" Then
            rs.Fields(strFields(n)).Value = Mid$(strTokens(n), 2, Len(strTokens(n)) - 2)
        Else
            rs.Fields(strFields(n)).Value = strTokens(n)
        End If
    End Select
    n = n + 1
Loop Until n > UBound(strTokens)

'102703 rjk: Added to detect when grid is required.
If (rs.Fields("x").Value = frmDrawMap.SelX) And (rs.Fields("y").Value = frmDrawMap.SelY) Then
    ProcessDump = True
End If

If strFields(0) = "own" Then '102803 rjk: Always fill in the country field
    rs.Fields("country") = strTokens(0)
Else
    rs.Fields("country") = CountryNumber
End If
rs.Update
rs.Index = hldIndex

Exit Function

'Error handling routine
ProcessDump_Exit:
EmpireError "ProcessDump", var, strLine

End Function

'111403 rjk: Ensure there is not an old enemy intelligence for this ship.
Public Sub CheckForOldEnemyUnit(id As Long, strClass As String)
Dim hldIndex As String
hldIndex = rsEnemyUnit.Index
rsEnemyUnit.Index = "ByID"
rsEnemyUnit.Seek "=", strClass, id
If Not rsEnemyUnit.NoMatch Then
    rsEnemyUnit.Delete
End If
rsEnemyUnit.Index = hldIndex
End Sub

'101403 rjk: Changed ProcessNukeDump to accept partial dumps
Public Sub ProcessNukeDump(strLine As String, ByRef rs As Recordset)
'id x y num type
Dim n As Integer
Dim numTokens As Integer
Dim strTokens() As String
Dim hldIndex As String
Dim var As String

On Error GoTo ProcessNukeDump_Exit
hldIndex = rs.Index
rs.Index = "PrimaryKey"

ParseDumpString strLine, strTokens(), bDeity

numTokens = UBound(strTokens)

rs.Seek "=", strTokens(LBound(strTokens)), strTokens(4)
If Not rs.NoMatch Then
    rs.Edit
Else
    rs.AddNew
    rs.Fields("id").Value = strTokens(LBound(strTokens))
End If

For n = LBound(strTokens) + 1 To numTokens
    var = rs.Fields(n).Name
    rs.Fields(n) = strTokens(n)
Next n

If Not bDeity Then '102803 rjk: Always fill in the country field
    rsNuke.Fields("country") = CountryNumber
End If

rs.Update

'101403 rjk: added to update sector when doing disarm operations.
If frmDrawMap.SelX = CInt(strTokens(LBound(strTokens) + 1)) And frmDrawMap.SelY = CInt(strTokens(LBound(strTokens) + 2)) Then
    frmDrawMap.FillSectorBox frmDrawMap.SelX, frmDrawMap.SelY
End If

rs.Index = hldIndex
Exit Sub

'Error handling routine
ProcessNukeDump_Exit:

rs.Index = hldIndex
EmpireError "ProcessNukeDump", var, strLine
End Sub

Public Sub BlitzSetup(Initial As Boolean, sx1 As Integer, sy1 As Integer, sx2 As Integer, sy2 As Integer)
On Error GoTo Empire_Error

Dim strSect As String
Dim Cap As String
Dim Harbour As String
Dim Oil As String
' Dim x As Integer    8/2003 efj  removed
Dim mines As Integer
Dim hldIndex As String

'set the index
hldIndex = rsSectors.Index
rsSectors.Index = "UpdateOrder"

If Initial Then
    mines = 1
Else
    mines = 0
End If

Harbour = vbNullString

rsSectors.MoveFirst
While Not rsSectors.EOF

    If Initial Or _
        (rsSectors.Fields("x") >= sx1 And rsSectors.Fields("x") <= sx2 And _
        rsSectors.Fields("y") >= sy1 And rsSectors.Fields("y") <= sy2) Then
           
        rsSectors.Edit
        
        '0,0 is a gold mine
        If Initial And rsSectors.Fields("x") = 0 And rsSectors.Fields("y") = 0 Then
            rsSectors.Fields("des") = "g"
            frmEmpCmd.SubmitEmpireCommand "des 0,0 g", False
            frmEmpCmd.SubmitEmpireCommand "thr d 0,0 1", False
        
        '2,0 is a mine
        ElseIf Initial And rsSectors.Fields("x") = 2 And rsSectors.Fields("y") = 0 Then
            rsSectors.Fields("des") = "m"
            frmEmpCmd.SubmitEmpireCommand "des 2,0 m", False
            frmEmpCmd.SubmitEmpireCommand "thr i 2,0 1", False
        
        ElseIf Not Initial And rsSectors.Fields("des") = "h" Then
            Harbour = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            
        ElseIf Not Initial And rsSectors.Fields("des") = "=" Then
            
        'Dim oil well
        ElseIf rsSectors.Fields("ocontent") > 10 Then
            rsSectors.Fields("des") = "o"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " o", False
            frmEmpCmd.SubmitEmpireCommand "thr o " + strSect + " 1", False
            If rsSectors.Fields("ocontent") > 90 And Oil = vbNullString Then
                Oil = strSect
            End If
        
        'Dim up to 3 mines
        ElseIf rsSectors.Fields("min") > 90 And mines < 3 Then
            rsSectors.Fields("des") = "m"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " m", False
            frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 1", False
            mines = mines + 1
            
        'Dim up to several golds
        ElseIf rsSectors.Fields("gold") > 20 Then
            rsSectors.Fields("des") = "g"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " g", False
            frmEmpCmd.SubmitEmpireCommand "thr d " + strSect + " 1", False
        
        'Finally - make it a highway
        Else
            rsSectors.Fields("des") = "+"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " +", False
        End If
        
        rsSectors.Update
    End If
    
    rsSectors.MoveNext
Wend

'capital
Cap = SetOpenHex("c", 0, Initial, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + Cap + " c", False

If Initial Then
    'Reset Capital
    frmEmpCmd.SubmitEmpireCommand "cap " + Cap, False

    'move people to oil well
    frmEmpCmd.SubmitEmpireCommand "move c 0,0 250 " + Oil, False
    
    'build library and move people
    strSect = SetOpenHex("l", 2, Initial, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " l", False
    frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 150", False
    frmEmpCmd.SubmitEmpireCommand "move c 2,0 150 " + strSect, False
    
    'build tech center and move people
    strSect = SetOpenHex("t", 2, Initial, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " t", False
    frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 120", False
    frmEmpCmd.SubmitEmpireCommand "thr o " + strSect + " 60", False
    frmEmpCmd.SubmitEmpireCommand "thr d " + strSect + " 999", False
    frmEmpCmd.SubmitEmpireCommand "move c 2,0 100 " + strSect, False
End If

'lcm factory
strSect = SetOpenHex("j", 2, Initial, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " j", False
frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 999", False
frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 1", False
frmEmpCmd.SubmitEmpireCommand "move c 0,0 250 " + strSect, False

'harbour
If Len(Harbour) = 0 Then
    Harbour = SetOpenHex("h", 1, Initial, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + Harbour + " h", False
    If Initial Then
        frmEmpCmd.SubmitEmpireCommand "move c 2,0 300 " + Harbour, False
    End If
End If

'bank
strSect = SetOpenHex("b", 0, Initial, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " b", False
frmEmpCmd.SubmitEmpireCommand "thr d " + strSect + " 300", False

'hcm
strSect = SetOpenHex("k", 2, Initial, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " k", False
frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 999", False
frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 1", False

If Not Initial Then
    'enlistment center
    strSect = SetOpenHex("e", 2, Initial, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " e", False
    frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 200", False
    
    'another hcm
    strSect = SetOpenHex("k", 2, Initial, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " k", False
    frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 999", False
    frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 1", False

End If

'distribution
If Initial Then
    strSect = "* "
Else
    strSect = CStr(sx1) + ":" + CStr(sx2) + "," + CStr(sy1) + ":" + CStr(sy2) + " "
End If

frmEmpCmd.SubmitEmpireCommand "dist " + strSect + Harbour, False
frmEmpCmd.SubmitEmpireCommand "thr c " + strSect + "768", False
frmEmpCmd.SubmitEmpireCommand "thr u " + strSect + "868", False

frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors True
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength
frmEmpCmd.SubmitEmpireCommand "db2", False

'restore Index
rsSectors.Index = hldIndex
Exit Sub
Empire_Error:
EmpireError "BlitzSetup", vbNullString, vbNullString

End Sub


Private Function SetOpenHex(strDes As String, Coast As Integer, _
    Initial As Boolean, sx1 As Integer, sy1 As Integer, sx2 As Integer, sy2 As Integer) As String

'Coast Values
'0 = not on coast
'1 = on coast
'2 = don't care
'3 = don't care - non essential

On Error GoTo Empire_Error
Dim SectString As String
Dim X As Integer

If Coast <= 2 Then
    SectString = "-+gom"
Else
    SectString = "-+"
End If

For X = 1 To Len(SectString)
    If Len(SetOpenHex) = 0 Then
        rsSectors.MoveFirst
        Do
            If Not ((rsSectors.Fields("x") = 0 And rsSectors.Fields("y") = 0) Or _
                (rsSectors.Fields("x") = 2 And rsSectors.Fields("y") = 0)) And _
                rsSectors.Fields("des") = Mid$(SectString, X, 1) Then
                
                If Initial Or _
                    (rsSectors.Fields("x") >= sx1 And rsSectors.Fields("x") <= sx2 And _
                    rsSectors.Fields("y") >= sy1 And rsSectors.Fields("y") <= sy2) Then
                    
                    Select Case Coast
                        Case 0
                            If rsSectors.Fields("coast") = 0 Then
                                SetOpenHex = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
                                rsSectors.Edit
                                rsSectors.Fields("des") = strDes
                                rsSectors.Update
                                Exit Do
                            End If
                        Case 1
                            If rsSectors.Fields("coast") = 1 Then
                                SetOpenHex = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
                                rsSectors.Edit
                                rsSectors.Fields("des") = strDes
                                rsSectors.Update
                                Exit Do
                            End If
                        Case 2, 3
                            SetOpenHex = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
                            rsSectors.Edit
                            rsSectors.Fields("des") = strDes
                            rsSectors.Update
                            Exit Do
                    End Select
                End If
            End If
            
            rsSectors.MoveNext
        Loop Until rsSectors.EOF
    End If
Next X
Exit Function
Empire_Error:
EmpireError "SetOpenHex", vbNullString, vbNullString

End Function
Public Sub IslandDevelop(sx1 As Integer, sy1 As Integer, sx2 As Integer, sy2 As Integer)
On Error GoTo Empire_Error

Dim strSect As String
' Dim Cap As String    8/2003 efj  removed
' Dim Harbour As String    8/2003 efj  removed
' Dim Oil As String    8/2003 efj  removed
' Dim x As Integer    8/2003 efj  removed
Dim mines As Integer
Dim hldIndex As String

'set the index
hldIndex = rsSectors.Index
rsSectors.Index = "UpdateOrder"

mines = 0
' Harbour = vbnullstring    8/2003 efj  removed

'Pass one - redesignate old mines to highways
rsSectors.MoveFirst
While Not rsSectors.EOF

    If (rsSectors.Fields("x") >= sx1 And rsSectors.Fields("x") <= sx2 And _
        rsSectors.Fields("y") >= sy1 And rsSectors.Fields("y") <= sy2) Then
           
        rsSectors.Edit
        
        Select Case rsSectors.Fields("des")
            Case "g"
                If rsSectors.Fields("gold") = 0 Then
                    rsSectors.Fields("des") = "+"
                    strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
                    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " +", False
                End If
            Case "o"
                If rsSectors.Fields("ocontent") < 5 Then
                    rsSectors.Fields("des") = "+"
                    strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
                    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " +", False
                End If
            Case "m"
                If rsSectors.Fields("gold") > 0 Then
                    rsSectors.Fields("des") = "g"
                    strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
                    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " g", False
                    frmEmpCmd.SubmitEmpireCommand "thr d " + strSect + " 1", False
                Else
                    mines = mines + 1
                End If
            End Select
        
        rsSectors.Update
    End If
    
    rsSectors.MoveNext
Wend

'Pass two - designate new oilwells and mines
rsSectors.MoveFirst
While Not rsSectors.EOF

    If (rsSectors.Fields("x") >= sx1 And rsSectors.Fields("x") <= sx2 And _
        rsSectors.Fields("y") >= sy1 And rsSectors.Fields("y") <= sy2) And _
        rsSectors.Fields("des") = "+" Then
           
        rsSectors.Edit
        
        If rsSectors.Fields("gold") > 0 Then
            rsSectors.Fields("des") = "g"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " g", False
            frmEmpCmd.SubmitEmpireCommand "thr d " + strSect + " 1", False
        ElseIf rsSectors.Fields("oil") > 5 Then
            rsSectors.Fields("des") = "o"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " o", False
            frmEmpCmd.SubmitEmpireCommand "thr o " + strSect + " 1", False
        ElseIf rsSectors.Fields("min") > 80 And mines < 4 Then
            mines = mines + 1
            rsSectors.Fields("des") = "m"
            strSect = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            frmEmpCmd.SubmitEmpireCommand "des " + strSect + " m", False
            frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 1", False
        End If
        
        rsSectors.Update
    End If
    
    rsSectors.MoveNext
Wend

'build airport
strSect = SetOpenHex("*", 0, False, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " *", False
frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 300", False
frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 150", False
frmEmpCmd.SubmitEmpireCommand "thr s " + strSect + " 100", False
frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 200", False
frmEmpCmd.SubmitEmpireCommand "thr p " + strSect + " 800", False
    
'build fortress
strSect = SetOpenHex("f", 1, False, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " f", False
frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 200", False
frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 100", False
frmEmpCmd.SubmitEmpireCommand "thr s " + strSect + " 50", False
frmEmpCmd.SubmitEmpireCommand "thr g " + strSect + " 10", False
    
'build shell factory
strSect = SetOpenHex("i", 2, False, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " i", False
frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 200", False
frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 100", False
frmEmpCmd.SubmitEmpireCommand "thr s " + strSect + " 1", False
    
'build defense plant
strSect = SetOpenHex("d", 2, False, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " d", False
frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 100", False
frmEmpCmd.SubmitEmpireCommand "thr o " + strSect + " 20", False
frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 200", False
frmEmpCmd.SubmitEmpireCommand "thr g " + strSect + " 1", False

'headquarters
strSect = SetOpenHex("!", 2, False, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " !", False
frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 300", False
frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 150", False
frmEmpCmd.SubmitEmpireCommand "thr s " + strSect + " 100", False
frmEmpCmd.SubmitEmpireCommand "thr g " + strSect + " 20", False
frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 200", False

'refinery
strSect = SetOpenHex("%", 2, False, sx1, sy1, sx2, sy2)
frmEmpCmd.SubmitEmpireCommand "des " + strSect + " %", False
frmEmpCmd.SubmitEmpireCommand "thr p " + strSect + " 1", False
frmEmpCmd.SubmitEmpireCommand "thr o " + strSect + " 20", False

'hcm
strSect = SetOpenHex("k", 3, False, sx1, sy1, sx2, sy2)
If Len(strSect) > 0 Then
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " k", False
    frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 999", False
    frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 1", False
End If

'lcm factory
If Len(strSect) > 0 Then
    strSect = SetOpenHex("j", 3, False, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " j", False
    frmEmpCmd.SubmitEmpireCommand "thr i " + strSect + " 999", False
    frmEmpCmd.SubmitEmpireCommand "thr l " + strSect + " 1", False
End If

'build fortress
If Len(strSect) > 0 Then
    strSect = SetOpenHex("f", 3, False, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " f", False
    frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 200", False
    frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 100", False
    frmEmpCmd.SubmitEmpireCommand "thr s " + strSect + " 50", False
    frmEmpCmd.SubmitEmpireCommand "thr g " + strSect + " 10", False
End If
    
'enlistment center
If Len(strSect) > 0 Then
    strSect = SetOpenHex("e", 3, False, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " e", False
    frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 118", False
End If
    
'build fortress
If Len(strSect) > 0 Then
    strSect = SetOpenHex("f", 3, False, sx1, sy1, sx2, sy2)
    frmEmpCmd.SubmitEmpireCommand "des " + strSect + " f", False
    frmEmpCmd.SubmitEmpireCommand "thr m " + strSect + " 200", False
    frmEmpCmd.SubmitEmpireCommand "thr h " + strSect + " 100", False
    frmEmpCmd.SubmitEmpireCommand "thr s " + strSect + " 50", False
    frmEmpCmd.SubmitEmpireCommand "thr g " + strSect + " 10", False
End If
    
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors True
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength
frmEmpCmd.SubmitEmpireCommand "db2", False

'restore Index
rsSectors.Index = hldIndex
Exit Sub
Empire_Error:
EmpireError "IslandDevelop", vbNullString, vbNullString

End Sub
       
'101503 rjk: Modified to preserve information from different intelligence sources
Public Sub AddEnemySectInfo(varInfo As Variant)
'Information is in the variant designation
Dim n As Integer
Dim sx As Integer
Dim sy As Integer

'if you are not getting an array, then quit
If Not IsArray(varInfo) Then
    Exit Sub
End If

On Error GoTo AddEnemySectInfo_Exit

n = ES_X
sx = varInfo(ES_X)
n = ES_Y
sy = varInfo(ES_Y)

'Find out if we already have the record
rsEnemySect.Seek "=", sx, sy
If rsEnemySect.NoMatch Then
    rsEnemySect.AddNew
Else
    rsEnemySect.Edit
End If
    
'Setup the bmap record also
rsBmap.Seek "=", sx, sy
If rsBmap.NoMatch Then
    rsBmap.AddNew
    rsBmap.Fields("x") = sx
    rsBmap.Fields("y") = sy
Else
    rsBmap.Edit
End If
    
For n = 1 To UBound(varInfo)

    If varInfo(n) = "???" Then
        varInfo(n) = "-1"
    End If
    
    'Put things in the right fields
    '101503 rjk: Modified to preserve information from different intelligence sources
    '101803 rjk: Removed updating EU_X and EU_Y as it should not change.
    Select Case n
    Case ES_DES
        If varInfo(n) <> "???" Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("des") = Left$(varInfo(n), 1)
            rsBmap.Fields("des").Value = Left$(varInfo(n), 1)
            rsBmap.Update
        End If
    Case ES_X
        If rsEnemySect.NoMatch Then
            rsEnemySect.Fields("x") = CInt(varInfo(n))
        End If
    Case ES_Y
        If rsEnemySect.NoMatch Then
            rsEnemySect.Fields("y") = CInt(varInfo(n))
        End If
    Case ES_OWNER
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("owner") = CInt(varInfo(n))
        End If
    Case ES_OLDOWNER
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("oldowner") = CInt(varInfo(n))
        End If
    Case ES_EFFICIENCY
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("eff") = CInt(varInfo(n))
        End If
    Case ES_ROAD
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("road") = CInt(varInfo(n))
        End If
    Case ES_RAIL
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("rail") = CInt(varInfo(n))
        End If
    Case ES_DEFENSE
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("defense") = CInt(varInfo(n))
        End If
    Case ES_CIV
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("civ") = CInt(varInfo(n))
        End If
    Case ES_MIL
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("mil") = CInt(varInfo(n))
        End If
    Case ES_SHELL
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("shell") = CInt(varInfo(n))
        End If
    Case ES_GUN
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("gun") = CInt(varInfo(n))
        End If
    Case ES_PETROL
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("pet") = CInt(varInfo(n))
        End If
    Case ES_FOOD
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("food") = CInt(varInfo(n))
        End If
    Case ES_BARS
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("bar") = CInt(varInfo(n))
        End If
    Case ES_IRON
        If CInt(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
            rsEnemySect.Fields("iron") = CInt(varInfo(n))
        End If
    End Select
Next n
rsEnemySect.Fields("timestamp") = GetWinACEUTC
rsEnemySect.Update


Exit Sub

'Error handling routine
AddEnemySectInfo_Exit:
Dim var As String
Dim strLine As String
var = CStr(n)
For n = 1 To UBound(varInfo)
    strLine = strLine + CStr(varInfo(n)) + ","
Next n
EmpireError "AddEnemySectInfo", var, strLine
End Sub

'101503 rjk: Modified to preserve information from different intelligence sources
'            Assumed that the ID and unit type is sufficient to identify record
'            Do not use owner as it is always present.
Public Sub AddEnemyUnitInfo(varInfo As Variant)

'Information is in the variant designation

Dim n As Integer
Dim sowner As Integer
Dim sclass As String
Dim sid As Integer
Dim swake As String
Dim Index As Integer

'if you are not getting an array, then quit
If Not IsArray(varInfo) Then
    Exit Sub
End If

On Error GoTo AddEnemyUnitInfo_Exit

n = EU_OWNER
sowner = CInt(varInfo(EU_OWNER))
n = EU_ID
sid = CInt(varInfo(EU_ID))
n = EU_TYPE

'get the class
sclass = CStr(varInfo(0))

'Find out if we already have the record
'rsEnemyUnit.Index = "PrimaryKey"
'rsEnemyUnit.Seek "=", sowner, sclass, sid
'101503 rjk: Switch to only sclass and sid as owner is not always present
rsEnemyUnit.Index = "ByID"
rsEnemyUnit.Seek "=", sclass, sid
If rsEnemyUnit.NoMatch Then
    rsEnemyUnit.AddNew
    '101503 rjk: move sclass and sid as they can not change, so only do once
    'record class
    rsEnemyUnit.Fields("class") = sclass
    rsEnemyUnit.Fields("id") = sid
    swake = vbNullString
ElseIf Val(rsEnemyUnit.Fields("owner")) <> Val(varInfo(EU_OWNER)) Then
    rsEnemyUnit.Delete '260604 rjk: Added to deal with satellite changing owner.
    rsEnemyUnit.AddNew
    rsEnemyUnit.Fields("class") = sclass
    rsEnemyUnit.Fields("id") = sid
    swake = vbNullString
Else
    rsEnemyUnit.Edit
    'for ships, keep a "wake" that shows where they have been
    swake = rsEnemyUnit.Fields("wake")
    If sclass = "s" Then
        If Val(varInfo(EU_X)) <> Val(rsEnemyUnit.Fields("x")) Or _
           Val(varInfo(EU_Y)) <> Val(rsEnemyUnit.Fields("y")) Then
            If ((rsEnemyUnit.Fields("x") + rsEnemyUnit.Fields("y")) Mod 2) = 0 Then
                If Right$(swake, 1) <> ";" Then  'clear wake from earlier versions
                    swake = vbNullString
                End If
                swake = swake + SectorString(rsEnemyUnit.Fields("x"), _
                    rsEnemyUnit.Fields("y")) + ";"
                While Len(swake) > 50
                    n = InStr(swake, ";")
                    swake = Mid$(swake, n + 1)
                Wend
            End If
        End If
    End If
End If
    
'record class move to add step
'rsEnemyUnit.Fields("class") = sclass

For n = 1 To UBound(varInfo)

    If varInfo(n) = "???" Then
        varInfo(n) = "-1"
    End If
    
    'Put things in the right fields
    '101503 rjk: do not erase something we already know.
    Select Case n
'        Case EU_ID moved to add step
'            rsEnemyUnit.Fields("id") = sid
    '101803 rjk: Removed the check on EU_X and EU_Y as should always be known.
    Case EU_X
        rsEnemyUnit.Fields("x") = CInt(varInfo(n))
    Case EU_Y
        rsEnemyUnit.Fields("y") = CInt(varInfo(n))
    Case EU_OWNER
        If sowner <> -1 Or rsEnemyUnit.NoMatch Then
            rsEnemyUnit.Fields("owner") = sowner
        End If
    Case EU_TYPE
        If varInfo(n) <> "???" Or rsEnemyUnit.NoMatch Then
            rsEnemyUnit.Fields("type") = varInfo(n)
        End If
    Case EU_MIL
        If CInt(varInfo(n)) <> -1 Or rsEnemyUnit.NoMatch Then
            rsEnemyUnit.Fields("mil") = CInt(varInfo(n))
        End If
    Case EU_TECH
        If CInt(varInfo(n)) <> -1 Or rsEnemyUnit.NoMatch Then
            rsEnemyUnit.Fields("tech") = CInt(varInfo(n))
        End If
    Case EU_EFF
        If CInt(varInfo(n)) <> -1 Or rsEnemyUnit.NoMatch Then
            rsEnemyUnit.Fields("eff") = CInt(varInfo(n))
        End If
    End Select
    
Next n
rsEnemyUnit.Fields("wake") = swake
rsEnemyUnit.Fields("timestamp") = GetWinACEUTC

rsEnemyUnit.Update
Exit Sub

'Error handling routine
AddEnemyUnitInfo_Exit:
Dim var As String
Dim strLine As String
var = CStr(n)
For n = 1 To UBound(varInfo)
    strLine = strLine + CStr(varInfo(n)) + ","
Next n

EmpireError "AddEnemyUnitInfo", var, strLine

End Sub

'112903 rjk: Remove the question on clearing the intelligence as there is a separate form now.
'            Can now delete individual items.
Public Sub ClearEnemyInfo(bSectors As Boolean, bShips As Boolean, bLands As Boolean, bPlanes As Boolean, bAirCombat As Boolean, _
                          bAllied As Boolean, bNeutral As Boolean, bEnemy As Boolean, iDays As Integer)
Dim rel As Integer
Dim dExpiry As Date
Dim dRecord As Date

dExpiry = DateAdd("d", -iDays, Now)
                

On Error GoTo Empire_Error
'if there is no enemy information, then exit
If (rsEnemyUnit.BOF And rsEnemyUnit.EOF) And (rsEnemySect.BOF And rsEnemySect.EOF) And (rsAirCombat.BOF And rsAirCombat.EOF) Then
    Exit Sub
End If

If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
    rsEnemyUnit.MoveFirst
    While Not (rsEnemyUnit.EOF)
        dRecord = ConvertToLocalTimeFromWinACEUTC(rsEnemyUnit.Fields("timestamp").Value)
        If DateDiff("d", dExpiry, dRecord) <= 0 Then
            If (rsEnemyUnit.Fields("class") = "s" And bShips) Or _
               (rsEnemyUnit.Fields("class") = "l" And bLands) Or _
               (rsEnemyUnit.Fields("class") = "p" And bPlanes) Then
                rel = Nations.YourRelation(rsEnemyUnit.Fields("owner"))
                If ((rel = iREL_AT_WAR Or rel = iREL_HOSTILE Or rel = iREL_MOBILIZING Or rel = iREL_SITZKRIEG) And bEnemy) Or _
                   ((rel = iREL_ALLIED Or (rel = iREL_FRIENDLY And frmOptions.friendlyUnit = FRIEND_ALLIED)) And bAllied) Or _
                   ((rel = iREL_NEUTRAL Or (rel = iREL_FRIENDLY And frmOptions.friendlyUnit = FRIEND_NEUTRAL)) And bNeutral) Or _
                   rel = 0 Then
                    rsEnemyUnit.Delete
                    rsEnemyUnit.MoveFirst
                Else
                    rsEnemyUnit.MoveNext
                End If
            Else
                rsEnemyUnit.MoveNext
            End If
        Else
            rsEnemyUnit.MoveNext
        End If
    Wend
End If

If bSectors Then
    DeleteAllRecordsOlderThen rsEnemySect, dExpiry
End If

If bAirCombat Then
    DeleteAllAirCombatRecordsOlderThen dExpiry
End If

Exit Sub
Empire_Error:
EmpireError "ClearEnemyInfo", vbNullString, vbNullString
End Sub

Public Sub UpdateLights()
On Error GoTo Empire_Error
'set the green and red lights in the main window status bar
If frmEmpCmd.CmdinProgress Or Not frmEmpCmd.ConnectedtoHost Then
    Set frmDrawMap.sb1.Panels.Item("Lights").Picture = picRedLight
    frmDrawMap.sb1.Panels.Item("Lights").ToolTipText = "Empire Server Unavailable"
    Set frmEmpCmd.imgLights.Picture = picRedLight
Else
    Set frmDrawMap.sb1.Panels.Item("Lights").Picture = picGreenLight
    frmDrawMap.sb1.Panels.Item("Lights").ToolTipText = "Empire Server Available"
    Set frmEmpCmd.imgLights.Picture = picGreenLight
End If
Exit Sub
Empire_Error:
EmpireError "UpdateLights", vbNullString, vbNullString
End Sub

Public Sub SetOwner(sx1 As Integer, sy1 As Integer, sx2 As Integer, sy2 As Integer, Nation As Integer)
On Error GoTo Empire_Error

Dim hldIndex As String

'set the index
hldIndex = rsBmap.Index
rsBmap.Index = "PrimaryKey"

rsBmap.MoveFirst
While Not rsBmap.EOF

    If (rsBmap.Fields("x") >= sx1 And rsBmap.Fields("x") <= sx2 And _
        rsBmap.Fields("y") >= sy1 And rsBmap.Fields("y") <= sy2) Then
        rsEnemySect.Seek "=", rsBmap.Fields("x"), rsBmap.Fields("y")
        If rsBmap.Fields("des") <> "." And rsBmap.Fields("des") <> "X" Then
            If rsEnemySect.NoMatch Then
                rsEnemySect.AddNew
                rsEnemySect.Fields("x") = rsBmap.Fields("x")
                rsEnemySect.Fields("y") = rsBmap.Fields("y")
                rsEnemySect.Fields("oldowner") = 0
                rsEnemySect.Fields("eff") = 0
                rsEnemySect.Fields("road") = 0
                rsEnemySect.Fields("rail") = 0
                rsEnemySect.Fields("defense") = 0
                rsEnemySect.Fields("civ") = 0
                rsEnemySect.Fields("mil") = 0
                rsEnemySect.Fields("shell") = 0
                rsEnemySect.Fields("gun") = 0
                rsEnemySect.Fields("pet") = 0
                rsEnemySect.Fields("food") = 0
                rsEnemySect.Fields("bar") = 0
                rsEnemySect.Fields("iron") = 0
                rsEnemySect.Fields("timestamp") = GetWinACEUTC '103003 rjk: Updated to make the timestamp format consistent
            Else
                rsEnemySect.Edit
            End If
            rsEnemySect.Fields("owner") = Nation
            rsEnemySect.Fields("des") = rsBmap.Fields("des")
            
            rsEnemySect.Update
        End If
    End If
    
    rsBmap.MoveNext
Wend


'restore Index
rsBmap.Index = hldIndex
frmDrawMap.DrawHexes
Exit Sub
Empire_Error:
EmpireError "SetOwner", vbNullString, vbNullString
End Sub

'102203 rjk: Added as temporary, need to combine with CopyGameInfo
'            still needs a lot of work.
'102703 rjk: Modifiy the edit command to work with syntax supported next version of the server.
Public Sub ExportNation()
On Error GoTo Empire_Error

Dim filenumber As Integer
Dim nLocations As Integer
Dim nSect As Integer
Dim nUnit As Integer
Dim FileName As String
Dim strCmd As String
Dim nID As Integer
Dim hldIndex As String

If VerifySubDirectory("Exports", True) Then
    FileName = App.path + "\Exports\" + GameName + " NationScript.txt"
Else
    FileName = App.path + "\" + GameName + " NationScript.txt"
End If

filenumber = FreeFile   ' Get unused file number.
Open FileName For Output As #filenumber

Print #filenumber, "edit country 0 T 1000"

hldIndex = rsShip.Index
rsShip.Index = "PrimaryKey"

If Not (rsShip.BOF And rsShip.EOF) Then
    If VersionCheck(4, 3, 0) <> VER_LESS Then
        ComputeUnitCountsForShips
    End If
    Print #filenumber, "designate 0,0 h"
    Print #filenumber, "setsector eff 0,0 100"
    'need to add index to ensure we start at the lowest number first
    rsShip.MoveFirst
    nID = 0
    While Not rsShip.EOF
        If nID = rsShip.Fields("id") Then
            rsBuild.Seek "=", "s", rsShip.Fields("type")
            If Not rsBuild.NoMatch Then
                'Need to supply the initial build supplies 20%
                Print #filenumber, "give lcms 0,0 " + CStr(rsBuild.Fields("lcm") \ 5#)
                Print #filenumber, "give hcms 0,0 " + CStr(rsBuild.Fields("hcm") \ 5#)
                Print #filenumber, "setsector avail 0,0 " + CStr(rsBuild.Fields("avail") \ 5#)
                'May need money as well
                Print #filenumber, "build ship 0,0 " + rsShip.Fields("type") + " 1 " + CStr(rsShip.Fields("tech"))
                strCmd = " U " + CStr(rsShip.Fields("id"))
                strCmd = strCmd + " O " + CStr(rsShip.Fields("country"))
                strCmd = strCmd + " L " + CStr(rsShip.Fields("x")) + "," + CStr(rsShip.Fields("y"))
                strCmd = strCmd + " E " + CStr(rsShip.Fields("eff"))
                strCmd = strCmd + " M " + CStr(rsShip.Fields("mob"))
                'strCmd= strCmd + " T " + CStr(rsShip.Fields("tech")) already in the build command
                strCmd = strCmd + " F " + CStr(rsShip.Fields("flt"))
                strCmd = strCmd + " H " + CStr(rsShip.Fields("he"))
                strCmd = strCmd + " P " + CStr(rsShip.Fields("pln"))
                strCmd = strCmd + " B " + CStr(rsShip.Fields("fuel"))
                strCmd = strCmd + " X " + CStr(rsShip.Fields("xl"))
                strCmd = strCmd + " Y " + CStr(rsShip.Fields("land"))
                'strCmd= strCmd + " R " + rsShip.Fields("retreat_path")
                'strCmd= strCmd + " W " + CStr(rsShip.Fields("retreat_flags"))
                'strCmd= strCmd + " a " + CStr(rsShip.Fields("plague_stage"))
                'strCmd= strCmd + " b " + rsShip.Fields("plague_time")
                'cargo
                strCmd = strCmd + " c " + CStr(rsShip.Fields("civ"))
                strCmd = strCmd + " m " + CStr(rsShip.Fields("mil"))
                strCmd = strCmd + " u " + CStr(rsShip.Fields("uw"))
                strCmd = strCmd + " f " + CStr(rsShip.Fields("food"))
                strCmd = strCmd + " s " + CStr(rsShip.Fields("shell"))
                strCmd = strCmd + " g " + CStr(rsShip.Fields("gun"))
                strCmd = strCmd + " p " + CStr(rsShip.Fields("petrol"))
                strCmd = strCmd + " i " + CStr(rsShip.Fields("iron"))
                strCmd = strCmd + " d " + CStr(rsShip.Fields("dust"))
                strCmd = strCmd + " o " + CStr(rsShip.Fields("oil"))
                'strCmd= strCmd + " ? " + CStr(rsShip.Fields("bar"))
                strCmd = strCmd + " l " + CStr(rsShip.Fields("lcm"))
                strCmd = strCmd + " h " + CStr(rsShip.Fields("hcm"))
                strCmd = strCmd + " r " + CStr(rsShip.Fields("rad"))
                Print #filenumber, "edit ship 0" + strCmd
                If Len(rsShip.Fields("name")) > 0 Then
                    Print #filenumber, "name " + CStr(rsShip.Fields("id")) + " " + Chr$(34) + rsShip.Fields("name") + Chr$(34)
                End If
                nUnit = nUnit + 1
                rsShip.MoveNext
            Else
                MsgBox ("Ship not found in the build information")
                Exit Sub
            End If
        Else
            rsBuild.MoveFirst
            rsBuild.Seek "=", "s", "fb" 'needs to be more general - will not always been there
            Print #filenumber, "give lcms 0,0 " + CStr(rsBuild.Fields("lcm") \ 5#)
            Print #filenumber, "give hcms 0,0 " + CStr(rsBuild.Fields("hcm") \ 5#)
            Print #filenumber, "setsector avail 0,0 " + CStr(rsBuild.Fields("avail") \ 5#)
            'May need money as well
            Print #filenumber, "build ship 0,0 fb 1 "
            Print #filenumber, "edit ship 0 O 1"
        End If
        nID = nID + 1
    Wend
End If
rsShip.Index = hldIndex

hldIndex = rsPlane.Index
rsPlane.Index = "PrimaryKey"

If Not (rsPlane.BOF And rsPlane.EOF) Then
    Print #filenumber, "designate 0,0 *"
    Print #filenumber, "setsector eff 0,0 100"
    'need to add index to ensure we start at the lowest number first
    rsPlane.MoveFirst
    nID = 0
    While Not rsPlane.EOF
        If nID = rsPlane.Fields("id") Then
            rsBuild.Seek "=", "p", rsPlane.Fields("type")
            If Not rsBuild.NoMatch Then
                'Need to supply the initial build supplies 10%
                Print #filenumber, "give lcms 0,0 " + CStr(rsBuild.Fields("lcm") \ 10#)
                Print #filenumber, "give hcms 0,0 " + CStr(rsBuild.Fields("hcm") \ 10#)
                Print #filenumber, "give mil 0,0 " + CStr(rsBuild.Fields("mil") \ 10#)
                Print #filenumber, "setsector avail 0,0 " + CStr(rsBuild.Fields("avail") \ 10#)
                'May need money as well
                Print #filenumber, "build plane 0,0 " + rsPlane.Fields("type") + " 1 " + CStr(rsPlane.Fields("tech"))
                strCmd = " U " + CStr(rsPlane.Fields("id"))
                strCmd = strCmd + " O " + CStr(rsPlane.Fields("country"))
                strCmd = strCmd + " L " + CStr(rsPlane.Fields("x")) + "," + CStr(rsPlane.Fields("y"))
                strCmd = strCmd + " e " + CStr(rsPlane.Fields("eff"))
                strCmd = strCmd + " M " + CStr(rsPlane.Fields("mob"))
                'strCmd = strCmd + " e " + CStr(rsPlane.Fields("tech")) alreay in the build command
                strCmd = strCmd + " F " + CStr(rsPlane.Fields("wing"))
                strCmd = strCmd + " B " + CStr(rsPlane.Fields("fuel"))
                Print #filenumber, "edit plane 0" + strCmd
                nUnit = nUnit + 1
                rsPlane.MoveNext
            Else
                MsgBox ("Plane not found in the build information")
                Exit Sub
            End If
        Else
            rsBuild.MoveFirst
            rsBuild.Seek "=", "p", "f1" 'needs to be more general - will not always been there
            Print #filenumber, "give lcms 0,0 " + CStr(rsBuild.Fields("lcm") \ 10#)
            Print #filenumber, "give hcms 0,0 " + CStr(rsBuild.Fields("hcm") \ 10#)
            Print #filenumber, "give mil 0,0 " + CStr(rsBuild.Fields("mil") \ 10#)
            Print #filenumber, "setsector avail 0,0 " + CStr(rsBuild.Fields("avail") \ 10#)
            'May need money as well
            Print #filenumber, "build plane 0,0 f1 1 "
            Print #filenumber, "edit plane 0 O 1"
        End If
        nID = nID + 1
    Wend
End If
rsPlane.Index = hldIndex

hldIndex = rsLand.Index
rsLand.Index = "PrimaryKey"

If Not (rsLand.BOF And rsLand.EOF) Then
    If VersionCheck(4, 3, 0) <> VER_LESS Then
        ComputeUnitCountsForLandUnits
    End If
    Print #filenumber, "designate 0,0 !"
    Print #filenumber, "setsector eff 0,0 100"
    'need to add index to ensure we start at the lowest number first
    rsLand.MoveFirst
    nID = 0
    While Not rsLand.EOF
        If nID = rsLand.Fields("id") Then
            rsBuild.Seek "=", "l", rsLand.Fields("type")
            If Not rsBuild.NoMatch Then
                'Need to supply the initial build supplies 10%
                Print #filenumber, "give lcms 0,0 " + CStr(rsBuild.Fields("lcm") \ 10#)
                Print #filenumber, "give hcms 0,0 " + CStr(rsBuild.Fields("hcm") \ 10#)
                Print #filenumber, "give mil 0,0 " + CStr(rsBuild.Fields("mil") \ 10#)
                Print #filenumber, "setsector avail 0,0 " + CStr(rsBuild.Fields("avail") \ 10#)
                'May need money as well
                Print #filenumber, "build land 0,0 " + rsLand.Fields("type") + " 1 " + CStr(rsLand.Fields("tech"))
                strCmd = " U " + CStr(rsLand.Fields("id"))
                strCmd = strCmd + " O " + CStr(rsLand.Fields("country"))
                strCmd = strCmd + " L " + CStr(rsLand.Fields("x")) + "," + CStr(rsLand.Fields("y"))
                strCmd = strCmd + " e " + CStr(rsLand.Fields("eff"))
                strCmd = strCmd + " M " + CStr(rsLand.Fields("mob"))
                'strCmd = strCmd + " t " + CStr(rsLand.Fields("tech")) alreay in the build command
                strCmd = strCmd + " a " + CStr(rsLand.Fields("army"))
                strCmd = strCmd + " B " + CStr(rsLand.Fields("fuel"))
                strCmd = strCmd + " X " + CStr(rsLand.Fields("xl"))
                strCmd = strCmd + " Y " + CStr(rsLand.Fields("land"))
                'strCmd = strCmd + " P " + rsLand.Fields("radius")
                'strCmd= strCmd + " Z " + CStr(rsLand.Fields("retreat_percentage"))
                'strCmd= strCmd + " R " + rsShip.Fields("retreat")
                'strCmd= strCmd + " W " + CStr(rsShip.Fields("retreat_flags"))
                'cargo
                'strCmd = strCmd + " c " + CStr(rsLand.Fields("civ"))
                strCmd = strCmd + " m " + CStr(rsLand.Fields("mil"))
                'strCmd = strCmd + " u " + CStr(rsLand.Fields("uw"))
                strCmd = strCmd + " f " + CStr(rsLand.Fields("food"))
                strCmd = strCmd + " s " + CStr(rsLand.Fields("shell"))
                strCmd = strCmd + " g " + CStr(rsLand.Fields("gun"))
                strCmd = strCmd + " p " + CStr(rsLand.Fields("petrol"))
                strCmd = strCmd + " i " + CStr(rsLand.Fields("iron"))
                strCmd = strCmd + " d " + CStr(rsLand.Fields("dust"))
                strCmd = strCmd + " o " + CStr(rsLand.Fields("oil"))
                'strCmd= strCmd + " ? " + CStr(rsLand.Fields("bar"))
                strCmd = strCmd + " l " + CStr(rsLand.Fields("lcm"))
                strCmd = strCmd + " h " + CStr(rsLand.Fields("hcm"))
                strCmd = strCmd + " r " + CStr(rsLand.Fields("rad"))
                Print #filenumber, "edit land 0" + strCmd
                nUnit = nUnit + 1
                rsLand.MoveNext
            Else
                MsgBox ("Land Unit not found in the build information")
                Exit Sub
            End If
        Else
            rsBuild.MoveFirst
            rsBuild.Seek "=", "l", "cav" 'needs to be more general - will not always been there
            Print #filenumber, "give lcms 0,0 " + CStr(rsBuild.Fields("lcm") \ 10#)
            Print #filenumber, "give hcms 0,0 " + CStr(rsBuild.Fields("hcm") \ 10#)
            Print #filenumber, "give mil 0,0 " + CStr(rsBuild.Fields("mil") \ 10#)
            Print #filenumber, "setsector avail 0,0 " + CStr(rsBuild.Fields("avail") \ 10#)
            'May need money as well
            Print #filenumber, "build land 0,0 cav 1 "
            Print #filenumber, "edit unit 0 O 1"
        End If
        nID = nID + 1
    Wend
End If
rsLand.Index = hldIndex

'need to add an index to control the order of reading
If Not (rsBmap.BOF And rsBmap.EOF) Then
    rsBmap.MoveFirst
    While Not rsBmap.EOF
        If Not IsNumeric(rsBmap.Fields("des")) Then 'need to deal sharebmap data
            nLocations = nLocations + 1
            '102803 rjk: Added bridge span and tower - needs more yet.
            If rsBmap.Fields("des") = "=" Then
                Print #filenumber, "designate " + CStr(rsBmap.Fields("x")) + "," + CStr(rsBmap.Fields("y")) + " ."
                Print #filenumber, "designate " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " #"
                Print #filenumber, "give hcms " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " 100" 'needs to be read from the sector info.
                Print #filenumber, "setsector avail " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " 40" 'needs to be read from the sector info.
                Print #filenumber, "build b " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " g"
                'need to restore sector
            ElseIf rsBmap.Fields("des") = "@" Then
                Print #filenumber, "designate " + CStr(rsBmap.Fields("x")) + "," + CStr(rsBmap.Fields("y")) + " ."
                Print #filenumber, "designate " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " #"
                Print #filenumber, "give hcms " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " 400" 'needs to be read from the sector info.
                Print #filenumber, "setsector avail " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " 160" 'needs to be read from the sector info.
                Print #filenumber, "build t " + CStr(rsBmap.Fields("x") + 2) + "," + CStr(rsBmap.Fields("y")) + " g"
                'need to restore sector
            Else
                Print #filenumber, "designate " + CStr(rsBmap.Fields("x")) + "," + CStr(rsBmap.Fields("y")) + " " + rsBmap.Fields("des")
            End If
        Else 'set to a plain for now
            Print #filenumber, "designate " + CStr(rsBmap.Fields("x")) + "," + CStr(rsBmap.Fields("y")) + " ~"
        End If
        rsBmap.MoveNext
    Wend
End If

If Not (rsSectors.BOF And rsSectors.EOF) Then
    rsSectors.MoveFirst
    While Not rsSectors.EOF
        nSect = nSect + 1
        'SetResource Set What (iron, gold, oil, uranium, fertility)?  Set What (, , , , )?
        Print #filenumber, "setresource iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("iron"))
        Print #filenumber, "setresource gold " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("gold"))
        Print #filenumber, "setresource oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("ocontent"))
        Print #filenumber, "setresource uranium " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("uran"))
        Print #filenumber, "setresource fertility " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("fert"))
        
        Print #filenumber, "edit land " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " R " + CStr(rsSectors.Fields("road")) + _
                            " r " + CStr(rsSectors.Fields("rail")) + " d " + CStr(rsSectors.Fields("defense"))
        
        'SetSector Give What (iron, gold, oil, uranium, fertility, owner, eff., Mob., Work, Avail., oldown, mines)?  , , , , , ., ., , ., , )?
        Print #filenumber, "setsector owner " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("country"))
        'Print #filenumber, "setsector eff" + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -100"
        Print #filenumber, "setsector eff " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("eff"))
        Print #filenumber, "setsector mob " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("mob"))
        'Print #filenumber, "setsector work " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -100"
        Print #filenumber, "setsector work " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("work"))
        'Print #filenumber, "setsector avail " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        Print #filenumber, "setsector avail " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("avail"))
        'Print #filenumber, "setsector mines " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -65536 " 'Just a guess
        If IsNull(rsSectors.Fields("lmines")) Then
            Print #filenumber, "setsector mines " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " 0"
        Else
            Print #filenumber, "setsector mines " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("lmines"))
        End If
       
        'Extract from rsEnemySector table when "*" is set
        If rsSectors.Fields("*") = "*" Then
            'lookup sector in rsEnemySector table
            'Print #filenumber, "setsector " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " oldown " + )
        End If
        
        'Commodity
        'Print #filenumber, "give bar " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("bar") > 0 Then
            Print #filenumber, "give bar " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("bar"))
        End If
        'Print #filenumber, "give civ " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("civ") > 0 Then
            Print #filenumber, "give civ " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("civ"))
        End If
        'Print #filenumber, "give dust " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("dust") > 0 Then
            Print #filenumber, "give dust " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("dust"))
        End If
        'Print #filenumber, "give food " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("food") > 0 Then
            Print #filenumber, "give food " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("food"))
        End If
        'Print #filenumber, "give gun " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("gun") > 0 Then
            Print #filenumber, "give gun " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("gun"))
        End If
        'Print #filenumber, "give hcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("hcm") > 0 Then
            Print #filenumber, "give hcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("hcm"))
        End If
        'Print #filenumber, "give iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("iron") > 0 Then
            Print #filenumber, "give iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("iron"))
        End If
        'Print #filenumber, "give lcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("lcm") > 0 Then
            Print #filenumber, "give lcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("lcm"))
        End If
        'Print #filenumber, "give mil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("mil") > 0 Then
            Print #filenumber, "give mil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("mil"))
        End If
        'Print #filenumber, "give oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("oil") > 0 Then
            Print #filenumber, "give oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("oil"))
        End If
        'Print #filenumber, "give pet " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("pet") > 0 Then
            Print #filenumber, "give pet " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("pet"))
        End If
        'Print #filenumber, "give rad " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("rad") > 0 Then
            Print #filenumber, "give rad " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("rad"))
        End If
        'Print #filenumber, "give shell " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("shell") > 0 Then
            Print #filenumber, "give shell " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("shell"))
        End If
        'Print #filenumber, "give uran " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
        If rsSectors.Fields("uran") > 0 Then
            Print #filenumber, "give uran " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("uran"))
        End If

        'Set Territory
        Print #filenumber, "terr " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("terr"))
        Print #filenumber, "terr " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("terr1")) + " 1"
        Print #filenumber, "terr " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("terr2")) + " 2"
        Print #filenumber, "terr " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("terr3")) + " 3"
        
        'Set Deliveries
        Print #filenumber, "wipe " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y"))
        If rsSectors.Fields("b_del") = "." Then
            Print #filenumber, "deliver bar " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("b_cut")) + " h"
        Else
            Print #filenumber, "deliver bar " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("b_cut")) + " " + rsSectors.Fields("b_del")
        End If
        If rsSectors.Fields("c_del") = "." Then
            Print #filenumber, "deliver civ " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("c_cut")) + " h"
        Else
            Print #filenumber, "deliver civ " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("c_cut")) + " " + rsSectors.Fields("c_del")
        End If
        If rsSectors.Fields("d_del") = "." Then
            Print #filenumber, "deliver dust " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("d_cut")) + " h"
        Else
            Print #filenumber, "deliver dust " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("d_cut")) + " " + rsSectors.Fields("d_del")
        End If
        If rsSectors.Fields("f_del") = "." Then
            Print #filenumber, "deliver food " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("f_cut")) + " h"
        Else
            Print #filenumber, "deliver food " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("f_cut")) + " " + rsSectors.Fields("f_del")
        End If
        If rsSectors.Fields("g_del") = "." Then
            Print #filenumber, "deliver gun " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("g_cut")) + " h"
        Else
            Print #filenumber, "deliver gun " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("g_cut")) + " " + rsSectors.Fields("g_del")
        End If
        If rsSectors.Fields("h_del") = "." Then
            Print #filenumber, "deliver hcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("h_cut")) + " h"
        Else
            Print #filenumber, "deliver hcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("h_cut")) + " " + rsSectors.Fields("h_del")
        End If
        If rsSectors.Fields("i_del") = "." Then
            Print #filenumber, "deliver iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("i_cut")) + " h"
        Else
            Print #filenumber, "deliver iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("i_cut")) + " " + rsSectors.Fields("i_del")
        End If
        If rsSectors.Fields("l_del") = "." Then
            Print #filenumber, "deliver lcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("l_cut")) + " h"
        Else
            Print #filenumber, "deliver lcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("l_cut")) + " " + rsSectors.Fields("l_del")
        End If
        If rsSectors.Fields("m_del") = "." Then
            Print #filenumber, "deliver mil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("m_cut")) + " h"
        Else
            Print #filenumber, "deliver mil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("m_cut")) + " " + rsSectors.Fields("m_del")
        End If
        If rsSectors.Fields("o_del") = "." Then
            Print #filenumber, "deliver oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("o_cut")) + " h"
        Else
            Print #filenumber, "deliver oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("o_cut")) + " " + rsSectors.Fields("o_del")
        End If
        If rsSectors.Fields("p_del") = "." Then
            Print #filenumber, "deliver pet " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("p_cut")) + " h"
        Else
            Print #filenumber, "deliver pet " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("p_cut")) + " " + rsSectors.Fields("p_del")
        End If
        If rsSectors.Fields("r_del") = "." Then
            Print #filenumber, "deliver rad " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("r_cut")) + " h"
        Else
            Print #filenumber, "deliver rad " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("r_cut")) + " " + rsSectors.Fields("r_del")
        End If
        If rsSectors.Fields("s_del") = "." Then
            Print #filenumber, "deliver shell " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("s_cut")) + " h"
        Else
            Print #filenumber, "deliver shell " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("s_cut")) + " " + rsSectors.Fields("s_del")
        End If
        If rsSectors.Fields("u_del") = "." Then
            Print #filenumber, "deliver uran " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("u_cut")) + " h"
        Else
            Print #filenumber, "deliver uran " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("u_cut")) + " " + rsSectors.Fields("u_del")
        End If
        
        'Set Distributions
        Print #filenumber, "distribute " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("dist_x")) + "," + CStr(rsSectors.Fields("dist_y"))
        Print #filenumber, "threshold bar " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("b_dist"))
        Print #filenumber, "threshold civ " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("c_dist"))
        Print #filenumber, "threshold dust " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("d_dist"))
        Print #filenumber, "threshold food " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("f_dist"))
        Print #filenumber, "threshold gun " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("g_dist"))
        Print #filenumber, "threshold hcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("h_dist"))
        Print #filenumber, "threshold iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("i_dist"))
        Print #filenumber, "threshold lcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("l_dist"))
        Print #filenumber, "threshold mil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("m_dist"))
        Print #filenumber, "threshold oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("o_dist"))
        Print #filenumber, "threshold pet " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("p_dist"))
        Print #filenumber, "threshold rad " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("r_dist"))
        Print #filenumber, "threshold shell " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("s_dist"))
        Print #filenumber, "threshold uran " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " " + CStr(rsSectors.Fields("u_dist"))
        
        rsSectors.MoveNext
    Wend
End If

       
Write #filenumber, "+++ " + CStr(nSect) + " Sectors exported"

Close #filenumber

MsgBox CStr(nLocations) + " Locations, " + CStr(nSect) + " Sectors, " + CStr(nUnit) + " Units exported to file " + FileName
Exit Sub
Empire_Error:
EmpireError "ExportNation", vbNullString, vbNullString
Close #filenumber
End Sub

Public Sub ExportSeaInformation()
Dim filenumber As Integer
Dim hldIndex As String
Dim FileName As String
Dim nLocations As Integer

filenumber = -1
On Error GoTo Empire_Error

If VerifySubDirectory("Exports", True) Then
    FileName = App.path + "\Exports\" + GameName + " Sea Information.txt"
Else
    FileName = App.path + "\" + GameName + " Sea Information.txt"
End If

filenumber = FreeFile   ' Get unused file number.
Open FileName For Output As #filenumber

hldIndex = rsSea.Index
rsSea.Index = "PrimaryKey"
rsSea.MoveFirst
While Not rsSea.EOF
    If IsNull(rsSea.Fields("ftimestamp")) Then
        rsSea.Edit
        rsSea.Fields("ftimestamp") = GetWinACEUTC
        rsSea.Update
    End If
    If IsNull(rsSea.Fields("otimestamp")) Then
        rsSea.Edit
        rsSea.Fields("otimestamp") = GetWinACEUTC
        rsSea.Update
    End If
    Write #filenumber, CInt(rsSea.Fields("x")), CInt(rsSea.Fields("y")), CInt(rsSea.Fields("fert")), CDate(rsSea.Fields("ftimestamp")), CInt(rsSea.Fields("ocontent")), CDate(rsSea.Fields("otimestamp"))
    nLocations = nLocations + 1
    rsSea.MoveNext
Wend

Close #filenumber
MsgBox CStr(nLocations) + " Locations exported to file " + FileName
rsSea.Index = hldIndex

Exit Sub

Empire_Error:
EmpireError "ExportSeaInformation", vbNullString, vbNullString
rsSea.Index = hldIndex
If filenumber > 0 Then
    Close #filenumber
End If
End Sub

Public Sub ImportIntelligence(iOffsetX As Integer, iOffsetY As Integer)
On Error GoTo Empire_Error

Dim filenumber As Integer
Dim n As Integer
Dim nSect As Integer
Dim nsect2 As Integer
Dim nUnit As Integer
Dim nunit2 As Integer
Dim FileName As String
Dim OnSectors As Boolean
Dim X As Variant
' Dim Y As Variant    8/2003 efj  removed
' Dim hldvar As Variant    8/2003 efj  removed
Dim temp() As Variant
Dim rs As Recordset
Dim count1 As Integer
Dim count2 As Integer
Dim date1 As Date
Dim date2 As Date
Dim strDate2 As String
Dim hldIndex As String

hldIndex = rsEnemyUnit.Index

ChDir App.path

Load frmCommonDialog
' Set CancelError is True
frmCommonDialog.cdb.CancelError = True
' Set flags
frmCommonDialog.cdb.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
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

FileName = frmCommonDialog.cdb.FileName
filenumber = FreeFile   ' Get unused file number.
Open FileName For Input As #filenumber

rsEnemyUnit.Index = "PrimaryKey"
OnSectors = True
Set rs = rsEnemySect
count1 = 0
count2 = 0
'first input the sector file
While Not EOF(filenumber)
    Input #filenumber, X
    If Left$(X, 3) = "+++" Then
        If OnSectors And (InStr(X, "Sectors exported") <> 0) Then '111803 rjk: Ignore headers
            OnSectors = False
            Set rs = rsEnemyUnit
            nSect = count1
            nsect2 = count2
            count1 = 0
            count2 = 0
        End If
    Else
        ReDim temp(0 To rs.Fields.Count - 1)
        temp(0) = X
        For n = 1 To rs.Fields.Count - 1
            Input #filenumber, temp(n)
            If OnSectors And n = 0 Then
                temp(n) = CStr(CInt(temp(0)) + iOffsetX)
            ElseIf OnSectors And n = 1 Then
                temp(n) = CStr(CInt(temp(1)) + iOffsetY)
            ElseIf Not OnSectors And n = 2 Then
                temp(n) = CStr(CInt(temp(2)) + iOffsetX)
            ElseIf Not OnSectors And n = 3 Then
                temp(n) = CStr(CInt(temp(3)) + iOffsetY)
            End If
        Next n
        count1 = count1 + 1
        '102903 rjk: Do not add records for ourselves
        If (CInt(temp(3)) <> CountryNumber And OnSectors) Or _
           (CInt(temp(4)) <> CountryNumber And Not OnSectors) Then
            If OnSectors Then
                rs.Seek "=", CInt(temp(0)), CInt(temp(1))
            Else
                rs.Seek "=", CInt(temp(4)), CStr(temp(9)), CInt(temp(0))
            End If
            If rs.NoMatch Then
                rs.AddNew
                For n = 0 To rs.Fields.Count - 1
                    rs.Fields(n) = temp(n)
                Next n
                rs.Update
                count2 = count2 + 1
            Else
                
                'decide which is more recent
                date1 = ConvertToLocalTimeFromWinACEUTC(rs.Fields("timestamp").Value)
                If OnSectors Then
                    strDate2 = CStr(temp(17))
                Else
                    strDate2 = CStr(temp(8))
                End If
                If Mid$(strDate2, 5, 1) <> "/" Then
                    strDate2 = Format(DateAdd("n", tz.Bias, EmpireTimeString(strDate2)), "yyyy/mm/dd hh:mm:ss")
                    If OnSectors Then
                        temp(17) = strDate2
                    Else
                        temp(8) = strDate2
                    End If
                End If
                date2 = ConvertToLocalTimeFromWinACEUTC(strDate2)
                If date1 < date2 Then
                    
                    'if export is more recent, use it
                    rs.Edit
                    For n = 0 To rs.Fields.Count - 1
                        rs.Fields(n) = temp(n)
                    Next n
                    
                    rs.Update
                    count2 = count2 + 1
                End If
            End If
        End If
    End If
Wend
nUnit = count1
nunit2 = count2

Close #filenumber

MsgBox CStr(nSect) + " Sectors read, " + CStr(nsect2) + " Sectors updated, " _
    + CStr(nUnit) + " Units read, " + CStr(nunit2) + " Units updated."

rsEnemyUnit.Index = hldIndex

Exit Sub
Empire_Error:
EmpireError "ImportIntelligence", vbNullString, vbNullString
rsEnemyUnit.Index = hldIndex

End Sub

Public Sub ImportIntelligenceTelegram(strTelegram As String, iOffsetX As Integer, iOffsetY As Integer)
On Error GoTo Empire_Error

Dim filenumber As Integer
Dim n As Integer
Dim nSect As Integer
Dim nsect2 As Integer
Dim nUnit As Integer
Dim nunit2 As Integer
Dim FileName As String
Dim OnSectors As Boolean
Dim temp() As String
Dim rs As Recordset
Dim count1 As Integer
Dim count2 As Integer
Dim date1 As Date
Dim date2 As Date
Dim strDate2 As String
Dim hldIndex As String
Dim pos As Long '110403 rjk: Changed to a Long to accomodate large Intelligence Reports.
Dim strLine As String

hldIndex = rsEnemyUnit.Index

rsEnemyUnit.Index = "PrimaryKey"
OnSectors = True
Set rs = rsEnemySect
count1 = 0
count2 = 0
'first input the sector file
pos = 1
While InStr(pos, strTelegram, vbLf) > 0
    strLine = Mid$(strTelegram, pos, InStr(pos, strTelegram, vbLf) - pos)
    pos = InStr(pos, strTelegram, vbLf) + 1
    If InStr(strLine, "+++") > 0 Then
        If OnSectors And InStr(strLine, "Sectors exported") <> 0 Then '111803 rjk: Changed processing to ignore headers
            OnSectors = False
            Set rs = rsEnemyUnit
            nSect = count1
            nsect2 = count2
            count1 = 0
            count2 = 0
        End If
    Else
        temp = Split(strLine, ",", rs.Fields.Count)
        For n = 0 To rs.Fields.Count - 1
            temp(n) = Mid$(temp(n), 2, Len(temp(n)) - 2)
            If OnSectors And n = 0 Then
                temp(n) = CStr(CInt(temp(0)) + iOffsetX)
            ElseIf OnSectors And n = 1 Then
                temp(n) = CStr(CInt(temp(1)) + iOffsetY)
            ElseIf Not OnSectors And n = 2 Then
                temp(n) = CStr(CInt(temp(2)) + iOffsetX)
            ElseIf Not OnSectors And n = 3 Then
                temp(n) = CStr(CInt(temp(3)) + iOffsetY)
            End If
        Next n
        count1 = count1 + 1
        '102903 rjk: Do not add records for ourselves
        If (CInt(temp(3)) <> CountryNumber And OnSectors) Or _
           (CInt(temp(4)) <> CountryNumber And Not OnSectors) Then
            If OnSectors Then
                rs.Seek "=", CInt(temp(0)), CInt(temp(1))
            Else
                rs.Seek "=", CInt(temp(4)), CStr(temp(9)), CInt(temp(0))
            End If
            If rs.NoMatch Then
                rs.AddNew
                For n = 0 To rs.Fields.Count - 1
                    rs.Fields(n) = temp(n)
                Next n
                rs.Update
                count2 = count2 + 1
            Else
                'decide which is more recent
                date1 = ConvertToLocalTimeFromWinACEUTC(rs.Fields("timestamp").Value)
                If OnSectors Then
                    strDate2 = CStr(temp(17))
                Else
                    strDate2 = CStr(temp(8))
                End If
                If Mid$(strDate2, 5, 1) <> "/" Then
                    strDate2 = Format(DateAdd("n", tz.Bias, EmpireTimeString(strDate2)), "yyyy/mm/dd hh:mm:ss")
                    If OnSectors Then
                        temp(17) = strDate2
                    Else
                        temp(8) = strDate2
                    End If
                End If
                date2 = ConvertToLocalTimeFromWinACEUTC(strDate2)
                If date1 < date2 Then
                    
                    'if export is more recent, use it
                    rs.Edit
                    For n = 0 To rs.Fields.Count - 1
                        rs.Fields(n) = temp(n)
                    Next n
                    
                    rs.Update
                    count2 = count2 + 1
                End If
            End If
        End If
    End If
Wend
nUnit = count1
nunit2 = count2

MsgBox CStr(nSect) + " Sectors read, " + CStr(nsect2) + " Sectors updated, " _
    + CStr(nUnit) + " Units read, " + CStr(nunit2) + " Units updated."

rsEnemyUnit.Index = hldIndex

Exit Sub
Empire_Error:
EmpireError "ImportIntelligenceTelegram", vbNullString, vbNullString
rsEnemyUnit.Index = hldIndex
End Sub

Public Sub ImportIntelligenceOffset(strTelegram As String)
If frmDrawMap.PromptUp Then
    Exit Sub
End If

frmPromptImportOffset.strTelegram = strTelegram

frmPromptImportOffset.txtOrigin = SectorString(frmDrawMap.SelX, frmDrawMap.SelY)

'Put form in proper place
frmPromptImportOffset.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptImportOffset.Width) \ 2
frmPromptImportOffset.top = frmDrawMap.top + frmDrawMap.Height - frmPromptImportOffset.Height
Set frmDrawMap.PromptForm = frmPromptImportOffset
frmDrawMap.PromptUp = True

'Start the map as the backgroud
frmDrawMap.SetFocus
'Dialog box will take it from here
frmPromptImportOffset.Show vbModeless
End Sub

Public Sub ImportSeaInformation()
Dim strFileName As String
Dim iFileNumber As Integer
Dim iUpdatedLocations As Integer, iAddedLocations As Integer
Dim iLocX As Integer, iLocY As Integer
Dim iFert As Integer, iOcontent As Integer
Dim hldIndex As String
Dim dFertTimestamp As Date, dOcontentTimestamp As Date

hldIndex = rsSea.Index
rsSea.Index = "PrimaryKey"

iFileNumber = -1
On Error GoTo Empire_Error

ChDir App.path

Load frmCommonDialog
' Set CancelError is True
frmCommonDialog.cdb.CancelError = True
' Set flags
frmCommonDialog.cdb.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
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

strFileName = frmCommonDialog.cdb.FileName

iFileNumber = FreeFile   ' Get unused file number.
Open strFileName For Input As #iFileNumber

While Not EOF(iFileNumber)
    Input #iFileNumber, iLocX, iLocY, iFert, dFertTimestamp, iOcontent, dOcontentTimestamp
    rsSea.Seek "=", iLocX, iLocY
    If rsSea.NoMatch Then
        rsSea.AddNew
        rsSea.Fields("x") = iLocX
        rsSea.Fields("y") = iLocY
        rsSea.Fields("fert") = iFert
        rsSea.Fields("ftimestamp") = dFertTimestamp
        rsSea.Fields("ocontent") = iOcontent
        rsSea.Fields("otimestamp") = dOcontentTimestamp
        iAddedLocations = iAddedLocations + 1
        rsSea.Update
    ElseIf DateDiff("s", rsSea.Fields("ftimestamp"), dFertTimestamp) > 0 Or DateDiff("s", rsSea.Fields("otimestamp"), dOcontentTimestamp) > 0 Then
        rsSea.Edit
        If DateDiff("s", rsSea.Fields("ftimestamp"), dFertTimestamp) > 0 Then
            rsSea.Fields("fert") = iFert
            rsSea.Fields("ftimestamp") = dFertTimestamp
        End If
        If DateDiff("s", rsSea.Fields("otimestamp"), dOcontentTimestamp) > 0 Then
            rsSea.Fields("ocontent") = iOcontent
            rsSea.Fields("otimestamp") = dOcontentTimestamp
        End If
        rsSea.Update
        iUpdatedLocations = iUpdatedLocations + 1
    End If
Wend
Close #iFileNumber
MsgBox CStr(iUpdatedLocations) + " locations updated and " + CStr(iAddedLocations) + " locations imported from file " + strFileName
rsSea.Index = hldIndex

Exit Sub

Empire_Error:
EmpireError "ImportSeaInformation", vbNullString, vbNullString
If iFileNumber > 0 Then
    Close #iFileNumber
End If
rsSea.Index = hldIndex
End Sub

Public Sub OffsetSectorLocation(ByRef sx As Integer, ByRef sy As Integer, offx As Integer, offy As Integer)
On Error GoTo Empire_Error

Dim MaxX As Integer
Dim MaxY As Integer

sx = sx + offx
sy = sy + offy

'Adjust x for world wrap
MaxX = rsVersion.Fields("World X") / 2
MaxY = rsVersion.Fields("World Y") / 2

'090103 rjk: Changed wrap at MaxX instead MaxX + 1 as -MaxX to +MaxX-1 is the valid range
While sx >= MaxX
    sx = sx - rsVersion.Fields("World X")
Wend

While sx < (-1 * MaxX)
    sx = sx + rsVersion.Fields("World X")
Wend

'090103 rjk: Changed wrap at MaxY instead MaxY + 1 as -MaxY to +MaxY-1 is the valid range
While sy >= MaxY
    sy = sy - rsVersion.Fields("World Y")
Wend

While sy < (-1 * MaxY)
    sy = sy + rsVersion.Fields("World Y")
Wend
Exit Sub
Empire_Error:
EmpireError "OffsetSectorLocation", vbNullString, vbNullString
End Sub

'090103 rjk: Added Function for exploring in wheel fashion
Private Function StraightPath(strPath As String) As Boolean
    Dim strFirstDirection As String
    Dim Counter As Integer
    
    If Len(strPath) = 0 Then
        StraightPath = False
        Exit Function
    End If
    If Len(strPath) = 1 Then
        StraightPath = True
        Exit Function
    End If
    
    strFirstDirection = Left$(strPath, 1)
    
    For Counter = 1 To Len(strPath)
        If strFirstDirection <> Mid$(strPath, Counter, 1) Then
            StraightPath = False
            Exit Function
        End If
    Next Counter
    StraightPath = True
End Function

'090103 rjk: Added Function for exploring along the border
Private Function AdjacentToBorder(X As Integer, Y As Integer) As Boolean
    Dim adjacentX As Integer, adjacentY As Integer
    
    adjacentX = X - 2
    adjacentY = Y
    
    OffsetSectorLocation adjacentX, adjacentY, 0, 0
    rsBmap.Seek "=", adjacentX, adjacentY
    If Not rsBmap.NoMatch Then
        'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
        If InStr(".?\/^s", rsBmap.Fields("des")) > 0 Then
            AdjacentToBorder = True
            Exit Function
        End If
    End If
    
    adjacentX = X + 2
    adjacentY = Y
    
    OffsetSectorLocation adjacentX, adjacentY, 0, 0
    rsBmap.Seek "=", adjacentX, adjacentY
    If Not rsBmap.NoMatch Then
        'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
        If InStr(".?\/^s", rsBmap.Fields("des")) > 0 Then
            AdjacentToBorder = True
            Exit Function
        End If
    End If
    adjacentX = X - 1
    adjacentY = Y - 1
    
    OffsetSectorLocation adjacentX, adjacentY, 0, 0
    rsBmap.Seek "=", adjacentX, adjacentY
    If Not rsBmap.NoMatch Then
        'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
        If InStr(".?\/^s", rsBmap.Fields("des")) > 0 Then
            AdjacentToBorder = True
            Exit Function
        End If
    End If
    adjacentX = X - 1
    adjacentY = Y + 1
    
    OffsetSectorLocation adjacentX, adjacentY, 0, 0
    rsBmap.Seek "=", adjacentX, adjacentY
    If Not rsBmap.NoMatch Then
        'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
        If InStr(".?\/^s", rsBmap.Fields("des")) > 0 Then
            AdjacentToBorder = True
            Exit Function
        End If
    End If
    adjacentX = X + 1
    adjacentY = Y - 1
    
    OffsetSectorLocation adjacentX, adjacentY, 0, 0
    rsBmap.Seek "=", adjacentX, adjacentY
    If Not rsBmap.NoMatch Then
        'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
        If InStr(".?\/^s", rsBmap.Fields("des")) > 0 Then
            AdjacentToBorder = True
            Exit Function
        End If
    End If
    adjacentX = X + 1
    adjacentY = Y + 1
    
    OffsetSectorLocation adjacentX, adjacentY, 0, 0
    rsBmap.Seek "=", adjacentX, adjacentY
    If Not rsBmap.NoMatch Then
        'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
        If InStr(".?\/^s", rsBmap.Fields("des")) > 0 Then
            AdjacentToBorder = True
            Exit Function
        End If
    End If
    
    AdjacentToBorder = False
End Function

Public Sub ExploreOut(strSect As String, strUnit As String, radius As Integer, Occur As Integer, OldSectorCount As Integer, ByVal numberToEachSector As Integer, exploreMode As enumExploreMode)

Dim range(0 To MaxExploreRadius * 4 + 5, 0 To MaxExploreRadius * 2 + 3) As Integer
Dim paths(0 To MaxExploreRadius * 4 + 5, 0 To MaxExploreRadius * 2 + 3) As String
Dim p1 As Integer
Dim StartX As Integer
Dim StartY As Integer
Dim X As Integer
Dim Y As Integer
Dim n As Integer
Dim offy As Integer
Dim offx As Integer
Dim holdy As Integer
Dim holdx As Integer
Dim strCmd As String
Dim Done As Boolean
Dim found As Boolean
Dim newsector As Boolean
Dim unknown As Boolean
Dim UnitsAvail As Integer
Dim NewSectorCount As Integer

Const OWNED_SECTOR = 2
Const CHECK_SECTOR = 1
Const EXPLORATION_SECTOR = 3
Const UNPASSABLE_SECTOR = 4
On Error Resume Next

'verify sector and set offsets
If ParseSectors(StartX, StartY, strSect) Then
    offx = StartX - (MaxExploreRadius + 2) * 2
    offy = StartY - MaxExploreRadius
Else
    Exit Sub
End If

UnitsAvail = 0

'if the sector count is the same as the last time, you are done
NewSectorCount = rsSectors.RecordCount
If Occur > 1 And NewSectorCount = OldSectorCount Then
    Exit Sub
End If

rsSectors.Seek "=", StartX, StartY
If rsSectors.NoMatch Then
    Exit Sub
Else
    If strUnit = "c" Then
        If rsSectors.Fields("mil") > 0 Then
            UnitsAvail = rsSectors.Fields("civ")
        Else
            UnitsAvail = rsSectors.Fields("civ") - 1
        End If
    Else
        If rsSectors.Fields("civ") > 0 And _
            Not (rsSectors.Fields("*").Value = "*") Then
            UnitsAvail = rsSectors.Fields("mil")
        Else
            UnitsAvail = rsSectors.Fields("mil") - 1
        End If
    End If
End If

'Build Starting Point
If radius > MaxExploreRadius Then
    radius = MaxExploreRadius
End If

Done = False
range((MaxExploreRadius + 2) * 2, MaxExploreRadius) = CHECK_SECTOR
unknown = False

'Keep going thru the grids until you have explored all
While Not Done
    Done = True
    For Y = 0 To UBound(paths, 2)
        For p1 = 0 To UBound(paths, 1) Step 2
            X = p1 + (Y Mod 2)
            If range(X, Y) = CHECK_SECTOR Then
                holdx = X + offx
                holdy = Y + offy
                OffsetSectorLocation holdx, holdy, 0, 0
                rsSectors.Seek "=", holdx, holdy
                If rsSectors.NoMatch Then
                    rsBmap.Seek "=", holdx, holdy
                    If rsBmap.NoMatch Then
                        range(X, Y) = UNPASSABLE_SECTOR
                        unknown = True
                    Else
                        ' don't go into water, unknown, wasteland, mountains, or sanctuaries
                        If InStr(".X?\/^s", rsBmap.Fields("des")) > 0 Then
                            range(X, Y) = UNPASSABLE_SECTOR
                        Else
                            '090103 rjk: Added additional exploration modes
                            Select Case exploreMode
                            Case EXP_COMPLETE
                                range(X, Y) = EXPLORATION_SECTOR
                            Case EXP_WHEEL
                                If StraightPath(paths(X, Y)) Then
                                    range(X, Y) = EXPLORATION_SECTOR
                                End If
                                If AdjacentToBorder(holdx, holdy) Then
                                    range(X, Y) = EXPLORATION_SECTOR
                                End If
                            Case EXP_BORDER
                                If AdjacentToBorder(holdx, holdy) Then
                                    range(X, Y) = EXPLORATION_SECTOR
                                End If
                            End Select
                        End If
                    End If
                Else
                   'owned Sector
                    range(X, Y) = OWNED_SECTOR
                End If
                
                If Len(paths(X, Y)) < radius And range(X, Y) <> UNPASSABLE_SECTOR Then
                    If range(X - 2, Y) = 0 Then
                        paths(X - 2, Y) = paths(X, Y) + "g"
                        range(X - 2, Y) = CHECK_SECTOR
                        Done = False
                    End If
                    If range(X - 1, Y - 1) = 0 Then
                        paths(X - 1, Y - 1) = paths(X, Y) + "y"
                        range(X - 1, Y - 1) = CHECK_SECTOR
                        Done = False
                    End If
                    If range(X + 1, Y - 1) = 0 Then
                        paths(X + 1, Y - 1) = paths(X, Y) + "u"
                        range(X + 1, Y - 1) = CHECK_SECTOR
                        Done = False
                    End If
                    If range(X + 2, Y) = 0 Then
                        paths(X + 2, Y) = paths(X, Y) + "j"
                        range(X + 2, Y) = CHECK_SECTOR
                        Done = False
                    End If
                    If range(X + 1, Y + 1) = 0 Then
                        paths(X + 1, Y + 1) = paths(X, Y) + "n"
                        range(X + 1, Y + 1) = CHECK_SECTOR
                        Done = False
                    End If
                    If range(X - 1, Y + 1) = 0 Then
                        paths(X - 1, Y + 1) = paths(X, Y) + "b"
                        range(X - 1, Y + 1) = CHECK_SECTOR
                        Done = False
                    End If
                End If
            End If
        Next p1
    Next Y
Wend

'Now build all the exploration command to the explorable sectors
strCmd = "explore " + strUnit + " " + strSect + " " + str$(numberToEachSector) + " "

newsector = False
frmEmpCmd.SubmitEmpireCommand "ex1", False
For n = 1 To radius
    found = False
    For Y = 0 To UBound(range, 2)
        For p1 = 0 To UBound(range, 1) - 1 Step 2
            X = p1 + (Y Mod 2)
            '090103 rjk - Ensure there is enough units
            If range(X, Y) = EXPLORATION_SECTOR And Len(paths(X, Y)) = n And UnitsAvail >= numberToEachSector Then
                frmEmpCmd.SubmitEmpireCommand strCmd + paths(X, Y) + "h", True
                '090103 rjk - need to subtract the number moved not just one.
                UnitsAvail = UnitsAvail - numberToEachSector
                found = True
            End If
        Next p1
    Next Y
    If found Then
        frmEmpCmd.SubmitEmpireCommand "des * ?des=- +", True
        newsector = True
    End If
Next n

If unknown And newsector And Occur < MaxExploreRadius Then
    frmEmpCmd.SubmitEmpireCommand "ex2", False
Else
    frmEmpCmd.SubmitEmpireCommand "ex4", False
End If

frmEmpCmd.SubmitEmpireCommand "db1", False
'090703 rjk: Added an equal to missed any events
frmEmpCmd.SubmitEmpireCommand "dump * ?timestamp=" + CStr(tsSectors), False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "map *", False
frmEmpCmd.SubmitEmpireCommand "bmap *", False
frmEmpCmd.SubmitEmpireCommand "db2", False

If unknown And newsector And Occur < MaxExploreRadius Then
    'the ex3 command will cause the explore out to continue to loop
    '090103 rjk: added the explore mode for the recursion call
    frmEmpCmd.SubmitEmpireCommand "ex3  " + strSect + ":" + strUnit + ":" _
        + CStr(radius) + ":" + CStr(Occur + 1) + ":" + CStr(NewSectorCount) + ":" + str$(numberToEachSector) + ":" + str(exploreMode) + ":", False
End If

End Sub

Public Function PadString(str As String, slen As Integer, Optional InFront As Boolean = True)
On Error GoTo Empire_Error
Dim strSpaces As String
Dim strPad As String
On Error Resume Next
strSpaces = "              "
While Len(strSpaces) < slen
    strSpaces = strSpaces + strSpaces
Wend

If Len(str) < slen Then
    strPad = Left$(strSpaces, slen - Len(str))
    If InFront Then
        PadString = strPad + str
    Else
        PadString = str + strPad
    End If
Else
    PadString = Left$(str, slen)
End If
Exit Function
Empire_Error:
EmpireError "PadString", vbNullString, str
End Function

Public Sub ProdSummaryReport(sx1 As Integer, sy1 As Integer, ShowStarve As Boolean, ShowOvrPop As Boolean, ShowIdle As Boolean, ShowBuild As Boolean)
On Error GoTo Empire_Error

Const PR_ITEM = 1
Const PR_ONHAND = 2
Const PR_MADE = 3
Const PR_USED = 4
Const PR_RQST = 5
Const PR_CATG_LAST = 6

Const PR_FOOD = 1
Const PR_IRON = 2
Const PR_LCM = 3
Const PR_HCM = 4
Const PR_DUST = 5
Const PR_BAR = 6
Const PR_GUN = 7
Const PR_SHELL = 8
Const PR_OIL = 9
Const PR_PET = 10
Const PR_MIL = 11
Const PR_RAD = 12
Const PR_CIV = 13
Const PR_UW = 14
Const PR_HAPPY = 15
Const PR_GRAD = 16
Const PR_TECH = 17
Const PR_RESCH = 18
Const PR_SHIP = 19
Const PR_PLANE = 20
Const PR_ARMY = 21
Const PR_CITY = 22
Const PR_FORT = 23
Const PR_COMM_LAST = 23
Const PR_COMM_ITEM = 14
Const PR_MSGC = 10

Dim ar(1 To PR_COMM_LAST, 1 To PR_CATG_LAST) As Variant
Dim sc(1 To PR_COMM_ITEM, 1 To PR_CATG_LAST) As Long
Dim ut(1 To PR_COMM_ITEM, 0 To PR_COMM_LAST) As Long
Dim cp(1 To PR_MSGC) As New Collection
Dim n As Long
Dim n2 As Long
Dim strLine As String
Dim varLine As Variant
Dim v As productionDataType
Dim strSect As String
Dim strSectHeader As String
Dim strCapt As String
Dim MaxEffGain As Integer
Dim EffGain As Single
Dim NewEff As Integer
Dim BuildType As String
Dim Target As Integer
Dim LastItem As Integer
Dim rs As Recordset

Dim Costs(1 To 7, 1 To 5) As Single
Dim CostDesc(1 To 5) As String

'Costs array
'First Dim
'1) holds full cost
'2) holds non rounded cost for current unit
'3) holds rounded (integer) cost for current unit)
'4) holds cumulative remander (used in average case)
'5) holds starting goods
'6) holds the required goods
'7) holds the expended goods

'Second Dim
'1) Cost
'2) Avail
'3) lcm
'4) hcm
'5) mil or guns

'Path cost variable for dist point total mob cost
Dim pvar As Variant
Dim pcost As Single
Dim mcost As Single
Dim total_mcost As Single
Dim pbdes As String
Dim tvar As Variant
Dim strDistPoint As String
Dim rsSect As Recordset
Dim strError1 As String
Dim strError2 As String

tvar = True
strDistPoint = vbNullString
If Not (sx1 = 0 And sy1 = 1) Then
    strDistPoint = SectorString(sx1, sy1)
    
    'find the dist point's post update packing bonus
    rsSectors.Seek "=", sx1, sy1
    If Not rsSectors.NoMatch Then
        v = Production(rsSectors)
        If v.prodValidFlag Then
            If v.NewEff < 60 Then
                pbdes = vbNullString
            Else
                pbdes = CStr(v.newdes)
            End If
        End If
    End If
End If

'Set the items
ar(PR_FOOD, PR_ITEM) = "food"
ar(PR_IRON, PR_ITEM) = "iron"
ar(PR_LCM, PR_ITEM) = "lcm"
ar(PR_HCM, PR_ITEM) = "hcm"
ar(PR_DUST, PR_ITEM) = "dust"
ar(PR_BAR, PR_ITEM) = "bar"
ar(PR_GUN, PR_ITEM) = "gun"
ar(PR_SHELL, PR_ITEM) = "shell"
ar(PR_OIL, PR_ITEM) = "oil"
ar(PR_PET, PR_ITEM) = "pet"
ar(PR_MIL, PR_ITEM) = "mil"
ar(PR_RAD, PR_ITEM) = "rad"
ar(PR_CIV, PR_ITEM) = "civ"
ar(PR_UW, PR_ITEM) = "uw"
ar(PR_HAPPY, PR_ITEM) = "stroll"
ar(PR_GRAD, PR_ITEM) = "grad"
ar(PR_TECH, PR_ITEM) = "tech"
ar(PR_RESCH, PR_ITEM) = "research"
ar(PR_SHIP, PR_ITEM) = "ships"
ar(PR_PLANE, PR_ITEM) = "planes"
ar(PR_ARMY, PR_ITEM) = "armies"
ar(PR_CITY, PR_ITEM) = "cities"
ar(PR_FORT, PR_ITEM) = "forts"

For n = 1 To PR_COMM_LAST
    ar(n, PR_ONHAND) = 0
    ar(n, PR_MADE) = 0
    ar(n, PR_USED) = 0
Next n

'Clear Event Markers
EventMarkers.Clear

'Scan production for Grads, tech units, etc.
strError1 = "Scanning for Production"
Set rsSect = rsSectors.Clone
If Not (rsSect.BOF And rsSect.EOF) Then
    rsSect.MoveFirst
    While Not rsSect.EOF
    
        If (sx1 = 0 And sy1 = 1) Or _
        (rsSect.Fields("dist_x") = sx1 And rsSect.Fields("dist_y") = sy1) Then
            strError2 = SectorString(rsSect.Fields("x"), rsSect.Fields("y"))

            'set if captured sector
            If rsSect.Fields("*").Value = "*" Then
                strCapt = "*"
            Else
                strCapt = " "
            End If
            
            strSect = SectorString(rsSect.Fields("x").Value, rsSect.Fields("y").Value)
            strSectHeader = "Sector " + PadString(strSect, 8, False) _
                   + PadString(CStr(rsSect.Fields("eff").Value), 3) + "% " + rsSect.Fields("des").Value + rsSect.Fields("sdes").Value + strCapt _
            
            'Get the post-update path cost to the sector
            pvar = BestPath(strDistPoint, strSect, MT_COMMODITY, tvar)
            pcost = pvar(2)

            'Clear the sector commodity array
            For n = 1 To PR_COMM_ITEM
                For n2 = 1 To PR_CATG_LAST
                    sc(n, n2) = 0
            Next n2, n

            'get the food needed/available
            n = FoodRequired(rsSect)
            If n > 0 Then
                ar(PR_FOOD, PR_USED) = ar(PR_FOOD, PR_USED) + n
                sc(PR_FOOD, PR_USED) = n
            End If
            
            'If there is not enough food, save a warning message
            If ShowStarve And rsSect.Fields("food").Value - n < 0 Then
                strLine = "Needs " + CStr(n - rsSect.Fields("food").Value) + " food"
                cp(1).Add strSectHeader + " - " + strLine
                MarkSector strSect, -1, strLine
            End If
            
            'get the on-hand amount
            For n = 1 To PR_COMM_ITEM
                ar(n, PR_ONHAND) = ar(n, PR_ONHAND) + rsSect.Fields(CStr(ar(n, PR_ITEM)))
                sc(n, PR_ONHAND) = rsSect.Fields(CStr(ar(n, PR_ITEM)))
            Next n
            
            'get the requested amount
            For n = 1 To PR_COMM_ITEM
                ar(n, PR_RQST) = ar(n, PR_RQST) + rsSect.Fields(Left$(CStr(ar(n, PR_ITEM)), 1) + "_dist")
                sc(n, PR_RQST) = rsSect.Fields(Left$(CStr(ar(n, PR_ITEM)), 1) + "_dist")
            Next n
            
            'handle civ and uw growth
            sc(PR_CIV, PR_MADE) = Int(NewPop(rsSect, "c"))
            sc(PR_UW, PR_MADE) = Int(NewPop(rsSect, "u"))
            ar(PR_CIV, PR_MADE) = ar(PR_CIV, PR_MADE) + sc(PR_CIV, PR_MADE)
            ar(PR_UW, PR_MADE) = ar(PR_UW, PR_MADE) + sc(PR_UW, PR_MADE)
            
            'get the production
            strError1 = "Production Report"
            v = Production(rsSect)
            If v.prodValidFlag Then
                Select Case CStr(v.newdes)
                   Case "~"
                        rsSectorType.Seek "=", CStr(v.newdes)
                        If Not rsSectorType.NoMatch Then
                            If rsSectorType.Fields("product") = "oil" Then
                                ar(PR_OIL, PR_MADE) = ar(PR_OIL, PR_MADE) + v.ProdAmount
                                sc(PR_OIL, PR_MADE) = v.ProdAmount
                            ElseIf rsSectorType.Fields("product") = "food" Then
                                ar(PR_FOOD, PR_MADE) = ar(PR_FOOD, PR_MADE) + v.ProdAmount
                                sc(PR_FOOD, PR_MADE) = v.ProdAmount
                            End If
                        End If
                   Case "m"
                        ar(PR_IRON, PR_MADE) = ar(PR_IRON, PR_MADE) + v.ProdAmount
                        sc(PR_IRON, PR_MADE) = v.ProdAmount
                    Case "g", "^"
                        ar(PR_DUST, PR_MADE) = ar(PR_DUST, PR_MADE) + v.ProdAmount
                        sc(PR_DUST, PR_MADE) = v.ProdAmount
                    Case "u"
                        ar(PR_RAD, PR_MADE) = ar(PR_RAD, PR_MADE) + v.ProdAmount
                        sc(PR_RAD, PR_MADE) = v.ProdAmount
                    Case "o"
                        ar(PR_OIL, PR_MADE) = ar(PR_OIL, PR_MADE) + v.ProdAmount
                        sc(PR_OIL, PR_MADE) = v.ProdAmount
                    Case "a"
                        ar(PR_FOOD, PR_MADE) = ar(PR_FOOD, PR_MADE) + v.ProdAmount
                        sc(PR_FOOD, PR_MADE) = v.ProdAmount
                    Case "e"
                        ar(PR_MIL, PR_MADE) = ar(PR_MIL, PR_MADE) + v.ProdAmount
                        sc(PR_MIL, PR_MADE) = v.ProdAmount
                        ar(PR_CIV, PR_USED) = ar(PR_CIV, PR_USED) + v.ProdAmount
                        sc(PR_CIV, PR_USED) = v.ProdAmount
                    Case "j"
                        ar(PR_LCM, PR_MADE) = ar(PR_LCM, PR_MADE) + v.ProdAmount
                        ar(PR_IRON, PR_USED) = ar(PR_IRON, PR_USED) + v.unitsConsumed
                        sc(PR_LCM, PR_MADE) = v.ProdAmount
                        sc(PR_IRON, PR_USED) = v.unitsConsumed
                        ut(PR_IRON, PR_LCM) = ut(PR_IRON, PR_LCM) + v.unitsConsumed
                    Case "k"
                        ar(PR_HCM, PR_MADE) = ar(PR_HCM, PR_MADE) + v.ProdAmount
                        ar(PR_IRON, PR_USED) = ar(PR_IRON, PR_USED) + v.unitsConsumed * 2
                        sc(PR_HCM, PR_MADE) = v.ProdAmount
                        sc(PR_IRON, PR_USED) = v.unitsConsumed * 2
                        ut(PR_IRON, PR_HCM) = ut(PR_IRON, PR_HCM) + v.unitsConsumed * 2
                    Case "b"
                        ar(PR_BAR, PR_MADE) = ar(PR_BAR, PR_MADE) + v.ProdAmount
                        ar(PR_DUST, PR_USED) = ar(PR_DUST, PR_USED) + v.unitsConsumed * 5
                        sc(PR_BAR, PR_MADE) = v.ProdAmount
                        sc(PR_DUST, PR_USED) = v.unitsConsumed * 5
                        ut(PR_DUST, PR_BAR) = ut(PR_DUST, PR_BAR) + v.unitsConsumed * 5
                   Case "c"
                        n = rsSect.Fields("eff")
                        If rsSect.Fields("des") <> "c" Then
                            n = 0
                        End If
                        n = CInt(v.NewEff) - n 'new efficiency- current efficiency
                        If n > 0 Then
                            If Not (rsBuild.BOF And rsBuild.EOF) Then
                                rsBuild.Seek "=", "i", v.newdes
                                If Not rsBuild.NoMatch Then
                                    If rsBuild.Fields("lcm") > 0 Then
                                        ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + (n * rsBuild.Fields("lcm"))
                                        sc(PR_LCM, PR_USED) = (n * rsBuild.Fields("lcm"))
                                        ut(PR_LCM, PR_CITY) = ut(PR_LCM, PR_CITY) + (n * rsBuild.Fields("lcm"))
                                    End If
                                    If rsBuild.Fields("hcm") > 0 Then
                                        ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + (n * rsBuild.Fields("hcm"))
                                        sc(PR_HCM, PR_USED) = (n * rsBuild.Fields("hcm"))
                                        ut(PR_HCM, PR_CITY) = ut(PR_HCM, PR_CITY) + (n * rsBuild.Fields("hcm"))
                                    End If
                                End If
                            Else
                                ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + n
                                sc(PR_LCM, PR_USED) = n
                                ut(PR_LCM, PR_CITY) = ut(PR_LCM, PR_CITY) + n
                                ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + (n * 2)
                                sc(PR_HCM, PR_USED) = n * 2
                                ut(PR_HCM, PR_CITY) = ut(PR_HCM, PR_CITY) + (n * 2)
                            End If
                            ar(PR_CITY, PR_MADE) = ar(PR_CITY, PR_MADE) + 1
                        End If
                   Case "i"
                        If frmOptions.bSPAtlantis Then
                            ar(PR_SHELL, PR_MADE) = ar(PR_SHELL, PR_MADE) + v.ProdAmount
                            sc(PR_SHELL, PR_MADE) = v.ProdAmount
                        Else
                            ar(PR_SHELL, PR_MADE) = ar(PR_SHELL, PR_MADE) + v.ProdAmount
                            ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 2
                            ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + v.unitsConsumed
                            sc(PR_SHELL, PR_MADE) = v.ProdAmount
                            sc(PR_LCM, PR_USED) = v.unitsConsumed * 2
                            sc(PR_HCM, PR_USED) = v.unitsConsumed
                            ut(PR_LCM, PR_SHELL) = ut(PR_LCM, PR_SHELL) + v.unitsConsumed * 2
                            ut(PR_HCM, PR_SHELL) = ut(PR_HCM, PR_SHELL) + v.unitsConsumed * 1
                        End If
                    Case "d"
                        ar(PR_GUN, PR_MADE) = ar(PR_GUN, PR_MADE) + v.ProdAmount
                        ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed
                        ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 5
                        ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + v.unitsConsumed * 10
                        sc(PR_GUN, PR_MADE) = v.ProdAmount
                        sc(PR_OIL, PR_USED) = v.unitsConsumed
                        sc(PR_LCM, PR_USED) = v.unitsConsumed * 5
                        sc(PR_HCM, PR_USED) = v.unitsConsumed * 10
                        ut(PR_LCM, PR_GUN) = ut(PR_LCM, PR_GUN) + v.unitsConsumed * 5
                        ut(PR_HCM, PR_GUN) = ut(PR_HCM, PR_GUN) + v.unitsConsumed * 10
                        ut(PR_OIL, PR_GUN) = ut(PR_OIL, PR_GUN) + v.unitsConsumed
                     Case "%"
                        ar(PR_PET, PR_MADE) = ar(PR_PET, PR_MADE) + v.ProdAmount
                        ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed
                        sc(PR_PET, PR_MADE) = v.ProdAmount
                        sc(PR_OIL, PR_USED) = v.unitsConsumed
                        ut(PR_OIL, PR_PET) = ut(PR_OIL, PR_PET) + v.unitsConsumed
                     Case "q" 'Heavy Metal II biofuel
                        If Not (rsSectorType.BOF And rsSectorType.EOF) Then
                            rsSectorType.Seek "=", "q"
                            If Not rsSectorType.NoMatch Then
                                If rsSectorType.Fields("use1n") > 0 Then
                                    ar(PR_FOOD, PR_USED) = ar(PR_FOOD, PR_USED) + (v.unitsConsumed * rsSectorType.Fields("use1n"))
                                    sc(PR_FOOD, PR_USED) = (v.unitsConsumed * rsSectorType.Fields("use1n"))
                                    ut(PR_FOOD, PR_OIL) = ut(PR_FOOD, PR_OIL) + (v.unitsConsumed * rsSectorType.Fields("use1n"))
                                End If
                            End If
                        Else
                            ar(PR_FOOD, PR_USED) = ar(PR_FOOD, PR_USED) + v.unitsConsumed * 5
                            sc(PR_FOOD, PR_USED) = v.unitsConsumed * 5
                            ut(PR_FOOD, PR_OIL) = ut(PR_FOOD, PR_OIL) + v.unitsConsumed * 5
                        End If
                        ar(PR_OIL, PR_MADE) = ar(PR_OIL, PR_MADE) + v.ProdAmount
                        sc(PR_OIL, PR_MADE) = v.ProdAmount
                     Case "l"
                        ar(PR_GRAD, PR_MADE) = ar(PR_GRAD, PR_MADE) + v.ProdAmount
                        ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed
                        sc(PR_LCM, PR_USED) = v.unitsConsumed
                        ut(PR_LCM, PR_GRAD) = ut(PR_LCM, PR_GRAD) + v.unitsConsumed
                    Case "y" '260104 rjk: Added for St@r W@rs only
                        ar(PR_UW, PR_MADE) = ar(PR_UW, PR_MADE) + v.ProdAmount
                        ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 2
                        ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + v.unitsConsumed
                        sc(PR_UW, PR_MADE) = v.ProdAmount
                        sc(PR_LCM, PR_USED) = v.unitsConsumed * 2
                        sc(PR_HCM, PR_USED) = v.unitsConsumed
                        ut(PR_LCM, PR_UW) = ut(PR_LCM, PR_UW) + v.unitsConsumed * 2
                        ut(PR_HCM, PR_UW) = ut(PR_HCM, PR_UW) + v.unitsConsumed * 1
                     Case "p"
                        ar(PR_HAPPY, PR_MADE) = ar(PR_HAPPY, PR_MADE) + v.ProdAmount
                        ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed
                        sc(PR_LCM, PR_USED) = v.unitsConsumed
                        ut(PR_LCM, PR_HAPPY) = ut(PR_LCM, PR_HAPPY) + v.unitsConsumed
                     Case "t"
                        If frmOptions.bSPAtlantis Then
                            ar(PR_TECH, PR_MADE) = ar(PR_TECH, PR_MADE) + v.ProdAmount
                            ar(PR_SHELL, PR_USED) = ar(PR_SHELL, PR_USED) + v.unitsConsumed * 50
                            sc(PR_SHELL, PR_USED) = v.unitsConsumed * 50
                            ut(PR_SHELL, PR_TECH) = ut(PR_SHELL, PR_TECH) + v.unitsConsumed * 50
                        Else
                            ar(PR_TECH, PR_MADE) = ar(PR_TECH, PR_MADE) + v.ProdAmount
                            ar(PR_DUST, PR_USED) = ar(PR_DUST, PR_USED) + v.unitsConsumed
                            ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed * 5
                            ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 10
                            sc(PR_DUST, PR_USED) = v.unitsConsumed
                            sc(PR_OIL, PR_USED) = v.unitsConsumed * 5
                            sc(PR_LCM, PR_USED) = v.unitsConsumed * 10
                            ut(PR_DUST, PR_TECH) = ut(PR_DUST, PR_TECH) + v.unitsConsumed
                            ut(PR_OIL, PR_TECH) = ut(PR_OIL, PR_TECH) + v.unitsConsumed * 5
                            ut(PR_LCM, PR_TECH) = ut(PR_LCM, PR_TECH) + v.unitsConsumed * 10
                        End If
                     Case "r"
                        ar(PR_RESCH, PR_MADE) = ar(PR_RESCH, PR_MADE) + v.ProdAmount
                        ar(PR_DUST, PR_USED) = ar(PR_DUST, PR_USED) + v.unitsConsumed
                        ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed * 5
                        ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 10
                        sc(PR_DUST, PR_USED) = v.unitsConsumed
                        sc(PR_OIL, PR_USED) = v.unitsConsumed * 5
                        sc(PR_LCM, PR_USED) = v.unitsConsumed * 10
                        ut(PR_DUST, PR_RESCH) = ut(PR_DUST, PR_RESCH) + v.unitsConsumed
                        ut(PR_OIL, PR_RESCH) = ut(PR_OIL, PR_RESCH) + v.unitsConsumed * 5
                        ut(PR_LCM, PR_RESCH) = ut(PR_LCM, PR_RESCH) + v.unitsConsumed * 10
                    
                    'for harbors, etc - we have to simulate builds.
                    Case "*", "h", "!", "f"
                        
                        For n = 1 To 5
                            For n2 = 1 To 7
                                Costs(n2, n) = 0
                        Next n2, n
                        
                        If CInt(v.NewEff) > 59 Then
                            Costs(5, 1) = 9999999        'Assume enough money
                            Costs(5, 2) = rsSect.Fields("avail")  'avail
                            If frmOptions.bSPAtlantis Then
                                Costs(5, 3) = rsSect.Fields("shell")  'lcm
                            Else
                                Costs(5, 3) = rsSect.Fields("lcm")  'lcm
                            End If
                            Costs(5, 4) = rsSect.Fields("hcm")  'hcm
                            CostDesc(2) = "avail"
                            If frmOptions.bSPAtlantis Then
                                CostDesc(3) = ar(PR_SHELL, PR_ITEM)
                            Else
                                CostDesc(3) = ar(PR_LCM, PR_ITEM)
                            End If
                            CostDesc(4) = ar(PR_HCM, PR_ITEM)
                            
                            n2 = 0 ' will hold count
                            Select Case CStr(v.newdes)
                                Case "h"
                                    Set rs = rsShip
                                    BuildType = "s"
                                    MaxEffGain = rsVersion.Fields("Eff gain - Ship")
                                    Target = PR_SHIP
                                    LastItem = 1
                                    'check for post update avail used
                                    If HarborAvail Then
                                         Costs(5, 2) = CInt(v.ProdAmount)
                                    End If
                                    
                                    Costs(5, 5) = 9999  'nothing
                
                                Case "!", "f"
                                    Set rs = rsLand
                                    BuildType = "l"
                                    MaxEffGain = rsVersion.Fields("Eff gain - Land")
                                    Target = PR_ARMY
                                    LastItem = PR_GUN
                                    Costs(5, 5) = rsSect.Fields("gun")
                                    CostDesc(5) = ar(LastItem, PR_ITEM)
                                    'check for post update avail used
                                    If (CStr(v.newdes) = "!" And HQAvail) Or _
                                       (CStr(v.newdes) = "f" And FortAvail) Then
                                        Costs(5, 2) = CInt(v.ProdAmount)
                                    End If
        
                                Case "*"
                                    Set rs = rsPlane
                                    BuildType = "p"
                                    MaxEffGain = rsVersion.Fields("Eff gain - Plane")
                                    Target = PR_PLANE
                                    LastItem = PR_MIL
                                    'check for post update avail used
                                    If AirportAvail Then
                                         Costs(5, 2) = CInt(v.ProdAmount)
                                    End If

                                    Costs(5, 5) = rsSect.Fields("mil")
                                    CostDesc(5) = ar(LastItem, PR_ITEM)
                            End Select
                                
                            'Now run through the units
                            If Not (rs.BOF And rs.EOF) Then
                                rs.MoveFirst
                            End If
                            While Not rs.EOF
                            
                                'Check to see if it is in this hex
                                If rsSect.Fields("x") = rs.Fields("x") And rsSect.Fields("y") = rs.Fields("y") Then
                                
                                    'check to see that it is not 100% already
                                    If rs.Fields("eff") < 100 And rs.Fields("off") = False Then
                                        'get the build requirements
                                        rsBuild.Seek "=", BuildType, rs.Fields("type")
                                        If Not rsBuild.NoMatch Then
                                            'Compute efficieny gain
                                            NewEff = rs.Fields("eff") + MaxEffGain
                                            If NewEff > 100 Then
                                                NewEff = 100
                                            End If
                                            EffGain = (NewEff - rs.Fields("eff")) / 100
                                            
                                            'load costs
                                            Costs(1, 1) = rsBuild.Fields("cost")
                                            '030404 rjk: Switched to use the avail from the server instead of calculating it,
                                            '            in case the calculation in the server is changed.
                                            '            20 + rsBuild.Fields("lcm") + rsBuild.Fields("hcm") * 2
                                            Costs(1, 2) = rsBuild.Fields("avail")
                                            Costs(1, 3) = rsBuild.Fields("lcm")
                                            Costs(1, 4) = rsBuild.Fields("hcm")
                                            Select Case CStr(v.newdes)
                                                Case "h"
                                                    Costs(1, 5) = 0
                                                Case "*"
                                                    Costs(1, 5) = rsBuild.Fields("mil")
                                                Case Else
                                                    Costs(1, 5) = rsBuild.Fields("gun")
                                            End Select
                                            For n = 1 To 5
                                               Costs(2, n) = Costs(1, n) * EffGain
                                               Costs(3, n) = CSng(Int(Costs(2, n) + 0.995))
                                               Costs(6, n) = Costs(6, n) + Costs(3, n)
                                            Next n
                                            ar(Target, PR_MADE) = ar(Target, PR_MADE) + 1
                                            n2 = n2 + 1
                                        End If
                                   End If
                                End If
                                rs.MoveNext
                            Wend
                            
                            'Put out build message
                            Dim strTemp As String
                            If Costs(6, 2) > 0 Then
                                strTemp = strSectHeader + " - building " + CStr(n2) + " " + ar(Target, PR_ITEM) + "."
                            Else
                                strTemp = strSectHeader + " - is idle."
                            End If
                            
                            If Costs(5, 2) > Costs(6, 2) Then
                                strTemp = strTemp + " " + CStr(Costs(5, 2) - Costs(6, 2)) + " avail left."
                            End If
                            cp(6).Add strTemp
                            
                            'check for shortages
                            For n = 2 To 5
                                If Costs(6, n) > Costs(5, n) Then
                                    cp(6).Add strSectHeader + " - is short " _
                                        + CStr(Costs(6, n) - Costs(5, n)) + " " + CostDesc(n) + " for builds"
                                End If
                            Next
    
                            'add to arrays
                            If frmOptions.bSPAtlantis Then
                                ar(PR_SHELL, PR_USED) = ar(PR_SHELL, PR_USED) + Costs(6, 3)
                                sc(PR_SHELL, PR_USED) = sc(PR_SHELL, PR_USED) + Costs(6, 3)
                                ut(PR_SHELL, Target) = ut(PR_SHELL, Target) + Costs(6, 3)
                            Else
                                ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + Costs(6, 3)
                                sc(PR_LCM, PR_USED) = sc(PR_LCM, PR_USED) + Costs(6, 3)
                                ut(PR_LCM, Target) = ut(PR_LCM, Target) + Costs(6, 3)
                            End If
                            ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + Costs(6, 4)
                            sc(PR_HCM, PR_USED) = sc(PR_HCM, PR_USED) + Costs(6, 4)
                            ut(PR_HCM, Target) = ut(PR_HCM, Target) + Costs(6, 4)
                            ar(LastItem, PR_USED) = ar(LastItem, PR_USED) + Costs(6, 5)
                            sc(LastItem, PR_USED) = sc(LastItem, PR_USED) + Costs(6, 5)
                            ut(LastItem, Target) = ut(LastItem, Target) + Costs(6, 5)
                        
                        End If
                        'handle new forts being built
                        If CStr(v.newdes) = "f" Then
                            n = rsSect.Fields("eff")
                            If rsSect.Fields("des") <> "f" Then
                                n = 0
                            End If
                            n = CInt(v.NewEff) - n 'new efficiency- current efficiency
                            If n > 0 Then
                                ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + n
                                ar(PR_FORT, PR_MADE) = ar(PR_FORT, PR_MADE) + 1
                                ut(PR_HCM, PR_FORT) = ut(PR_HCM, PR_FORT) + n
                            End If
                        End If
                    
                    Case Else
                End Select
            
                strError1 = "Excess Civilians"
                n = CInt(v.ExcessCiv)
                'if we are short civs, let us know about it
                If n < 0 And InStr("h*!f", rsSect.Fields("des").Value) = 0 Then
                    cp(4).Add strSectHeader + " - " + CStr(0 - n) + " civs needed"
                End If
                
                'If there are excess civs, save a warning message
                If ShowIdle And n > 10 And strCapt = " " Then
                    cp(5).Add strSectHeader + " - " + CStr(n) + " idle civs"
                    MarkSector strSect, -1, "Idle civs"
                End If
            End If
            
            strError1 = "Excess UWs"
            'if there are too many uws, save a warning message
            If ShowOvrPop And rsSect.Fields("uw") > MaxSafeUws Then
                cp(3).Add strSectHeader + " - Over MaxSafeUws by " + CStr(rsSect.Fields("uw").Value - MaxSafeUws) + " uws"
                MarkSector strSect, -1, "Over Max Safe Uws"
            End If
             
            strError1 = "City Check"
            'if there are too many civs, and its not a captured sector, save a warning message
            If ShowOvrPop And strCapt = " " Then
                If rsVersion.Fields("BIG_CITY") And rsSect.Fields("des") = "c" Then
                    If rsSect.Fields("civ") > (CInt(CLng(MaxPopulation(rsSect.Fields("des")) * 9& * CLng(rsSect.Fields("eff")) / 100&) + MaxPopulation(rsSect.Fields("des")))) Then
                        cp(2).Add strSectHeader + " - Over MaxSafeCivs by " + CStr(rsSect.Fields("civ").Value - CInt(CLng(MaxPopulation(rsSect.Fields("des")) * 9& * CLng(rsSect.Fields("eff")) / 100&) + MaxPopulation(rsSect.Fields("des")))) + " civs"
                        MarkSector strSect, -1, "Over Max Safe Civs"
                    End If
                ElseIf rsSect.Fields("civ") > MaxPopulation(rsSect.Fields("des")) Then
                    cp(2).Add strSectHeader + " - Over MaxSafeCivs by " + CStr(rsSect.Fields("civ").Value - MaxPopulation(rsSect.Fields("des"))) + " civs"
                    MarkSector strSect, -1, "Over Max Safe Civs"
                End If
            End If
            
            strError1 = "Determining Mobility"
                        'Figure out the total mob cost to ship to the sector
            For n = 1 To PR_COMM_ITEM
                If pcost > 0 And sc(n, PR_RQST) > 0 Then
                    n2 = sc(n, PR_RQST) - (sc(n, PR_ONHAND) + sc(n, PR_MADE) - sc(n, PR_USED))
                    If n2 > 0 Then
                        mcost = (pcost * MobilityWeight(Left$(ar(n, PR_ITEM), 1), n2, pbdes)) / 10#
                        total_mcost = total_mcost + mcost
                    End If
                End If
            Next n
        
        End If
        rsSect.MoveNext
    Wend

End If
            
strError1 = "Generating Report"
'Set Report
frmReport.Caption = "Area Production Summary"
frmReport.ClearReport

'Const PR_ITEM = 1
'Const PR_ONHAND = 2
'Const PR_MADE = 3
'Const PR_USED = 4
'Const PR_CATG_LAST = 6

'Set Summary Headers
strLine = PadString("ITEM", 8, False) + PadString("ON HAND", 8) + PadString("MAKING", 8) + PadString("USING", 8) _
            + PadString("DELTA", 8) + PadString("LEFT", 8) + PadString("THRSHLD", 8) + PadString("EXCESS", 8)
frmReport.AddReportLine strLine

'write Summaries
For n = 1 To PR_COMM_ITEM
    strLine = PadString(CStr(ar(n, PR_ITEM)), 8, False)
    For n2 = 2 To 4
        strLine = strLine + PadString(Format$(ar(n, n2), "###,##0"), 8)
    Next n2
    strLine = strLine + PadString(Format$(ar(n, PR_MADE) - ar(n, PR_USED), "###,##0"), 8)
    strLine = strLine + PadString(Format$(ar(n, PR_ONHAND) + ar(n, PR_MADE) - ar(n, PR_USED), "###,##0"), 8)
    strLine = strLine + PadString(Format$(ar(n, PR_RQST), "###,##0"), 8)
    strLine = strLine + PadString(Format$(ar(n, PR_ONHAND) + ar(n, PR_MADE) - ar(n, PR_USED) - ar(n, PR_RQST), "###,##0"), 8)
    
    frmReport.AddReportLine strLine
Next n

'if this is a single dist point, write out cost
If Len(strDistPoint) > 0 Then
    frmReport.AddReportLine vbNullString
    strLine = "Distribution mob. cost: " + CStr(CSng(CLng(CSng(total_mcost) * 1000) / 1000))
    frmReport.AddReportLine strLine
    frmReport.AddReportLine vbNullString
End If

'Put out item usage details
For n = 1 To PR_COMM_ITEM
    For n2 = 1 To PR_COMM_LAST
        ut(n, 0) = ut(n, 0) + ut(n, n2)
    Next n2
    If ut(n, 0) > 0 Then
        frmReport.AddReportLine vbNullString
        frmReport.AddReportLine CStr(ar(n, PR_ITEM)) + " Summary - " + Format$(ut(n, 0), "###,##0") + " used"
        For n2 = 1 To PR_COMM_LAST
            If ut(n, n2) > 0 Then
                frmReport.AddReportLine PadString(Format$(ut(n, n2), "###,##0"), 8) + " used in making" _
                    + PadString(Format$(ar(n2, PR_MADE), "###,##0"), 8) + " " + CStr(ar(n2, PR_ITEM))
            End If
        Next n2
    End If
Next n
    
frmReport.AddReportLine vbNullString
For n = 1 To PR_MSGC
    For Each varLine In cp(n)
        frmReport.AddReportLine CStr(varLine)
    Next
    frmReport.AddReportLine vbNullString
Next

frmReport.Show vbModeless
frmEmpCmd.SubmitEmpireCommand "db2", False

Exit Sub
Empire_Error:
EmpireError "ProdSummaryReport", strError1, strError2
End Sub

Public Sub AirCombatReport()
On Error GoTo Empire_Error
Dim rs As Recordset
Dim hldkey As Variant
Dim hldstatus As String
Dim statuscount As Integer
Dim nationcount As Integer
Dim timActiveDate As Date
Dim n As Integer
Dim X As Integer
Dim strTemp As String
Dim numnations As Integer
Dim aryNations() As Integer
Dim Done As Boolean
Dim bCurrent As Boolean

Set rs = rsAirCombat
ReDim aryNations(Nations.Count)

If rs.BOF And rs.EOF Then
    frmReport.AddReportLine "No air combat history found"
Else
    hldkey = rs.Index
    
    rs.Index = "TimeKey"
    rs.MoveFirst
    If timActiveDate = timNextUpdate Then
        timActiveDate = rs.Fields("nextupdate")
    Else
        timActiveDate = timNextUpdate
    End If
    n = 1
    While Not rs.EOF
        'keep track the order in which nations are found
        If rs.Fields("nation") > 0 Then '120304 rjk: Skip the invalid nations (-1)
            If aryNations(rs.Fields("nation")) = 0 Then
                aryNations(rs.Fields("nation")) = n
                n = n + 1
            End If
        
            'Mark Historical
            If timActiveDate <> rs.Fields("nextupdate") Then
                If rs.Fields("status") = "A" Or rs.Fields("status") = "D" Then
                    rs.Edit
                    rs.Fields("status") = "T"
                    rs.Fields("eff") = 100
                    rs.Update
                ElseIf rs.Fields("status") = "S" Then
                    rs.Edit
                    rs.Fields("status") = "X"
                    rs.Fields("eff") = 0
                    rs.Update
                End If
            End If
        End If
        rs.MoveNext
    Wend
    numnations = n - 1
    
    'Now do report
    rs.Index = "ReportKey"
    For X = 1 To numnations
        n = 1
        While aryNations(n) <> X
            n = n + 1
        Wend
    
        rs.Seek ">=", n  'must use >= with a partial key
        frmReport.AddReportLine "Player #" + CStr(n) + " - " _
            + Nations.NationName(n)
        frmReport.AddReportLine vbNullString
        hldstatus = vbNullString
        nationcount = 0
        Done = False
        
        While Not Done
        
            'process status change
            If rs.Fields("status") <> hldstatus Then
                If Len(hldstatus) > 0 Then
                    frmReport.AddReportLine "     Total: " + CStr(statuscount) + " planes"
                    frmReport.AddReportLine vbNullString
                End If
                Select Case rs.Fields("status")
                    Case "A"
                            frmReport.AddReportLine "Active"
                            frmReport.AddReportLine "======"
                            bCurrent = True
                    Case "D"
                            frmReport.AddReportLine "Disabled"
                            frmReport.AddReportLine "========"
                            bCurrent = True
                    Case "S"
                            frmReport.AddReportLine "Shot Down"
                            frmReport.AddReportLine "========="
                            bCurrent = True
                    Case "T"
                            frmReport.AddReportLine "Historical Active"
                            frmReport.AddReportLine "================="
                            bCurrent = False
                    Case "X"
                            frmReport.AddReportLine "Historical Shot Down"
                            frmReport.AddReportLine "===================="
                            bCurrent = False
                End Select
                hldstatus = rs.Fields("status")
                statuscount = 0
            End If
            
            'output line
            strTemp = "     " + rs.Fields("type") + " #" + PadString(CStr(rs.Fields("id")), 5, False)
            If bCurrent Then
                strTemp = strTemp + PadString(CStr(rs.Fields("eff")) + "%", 5) + "  " _
                + CStr(rs.Fields("numcombats")) + " combat"
                If rs.Fields("numcombats") > 1 Then
                    strTemp = strTemp + "s"
                End If
            End If
            frmReport.AddReportLine strTemp
            
            statuscount = statuscount + 1
            nationcount = nationcount + 1
            rs.MoveNext
            If rs.EOF Then
                Done = True
            ElseIf rs.Fields("nation") <> n Then
                Done = True
            End If
            
        Wend
        frmReport.AddReportLine "     Total: " + CStr(statuscount) + " planes"
        frmReport.AddReportLine vbNullString
        frmReport.AddReportLine "Nation Total: " + CStr(nationcount) + " planes"
        frmReport.AddReportLine vbNullString
    Next X
    rs.Index = hldkey
End If
    
frmReport.Show vbModeless
Exit Sub
Empire_Error:
EmpireError "AirCombatReport", vbNullString, vbNullString
End Sub

Public Function ApplySectorDamage(strSect As String, nDamage As Integer, Optional bAbsolute As Boolean) As String
On Error GoTo Empire_Error

If IsMissing(bAbsolute) Then
    bAbsolute = False
End If

Dim sx As Integer
Dim sy As Integer
Dim seceff As Integer
Dim sectype As String
Dim found As Boolean
Dim NewEff As Integer

If Not ParseSectors(sx, sy, strSect) Then
    ApplySectorDamage = vbNullString
    Exit Function
End If

found = True
'check for your own sector
rsSectors.Seek "=", sx, sy
If Not rsSectors.NoMatch Then
    sectype = rsSectors.Fields("des").Value
    seceff = CStr(rsSectors.Fields("eff").Value)
    If bAbsolute Then
        NewEff = seceff - nDamage
    Else
        NewEff = SectorEffAfterDamage(nDamage, sectype, seceff)
    End If
    rsSectors.Edit
    rsSectors.Fields("eff").Value = NewEff
    rsSectors.Update
Else
    'Check for Enemy Sector
    rsEnemySect.Seek "=", sx, sy
    If Not rsEnemySect.NoMatch Then
        If IsNull(rsEnemySect.Fields("des").Value) Or IsNull(rsEnemySect.Fields("eff").Value) Then
            found = False
            Exit Function
        End If
        sectype = rsEnemySect.Fields("des").Value
        seceff = rsEnemySect.Fields("eff").Value
        If bAbsolute Then
            NewEff = seceff - nDamage
        Else
            NewEff = SectorEffAfterDamage(nDamage, sectype, seceff)
        End If
        rsEnemySect.Edit
        rsEnemySect.Fields("eff").Value = NewEff
        rsEnemySect.Update
    Else
        rsBmap.Seek "=", sx, sy
        If Not rsBmap.NoMatch Then
            sectype = rsBmap.Fields("des").Value
            seceff = 100
            If bAbsolute Then
                NewEff = seceff - nDamage
            Else
                NewEff = SectorEffAfterDamage(nDamage, sectype, seceff)
            End If
        Else
            ApplySectorDamage = "No information on " + strSect
            found = False
        End If
    End If
End If

If found Then
    ApplySectorDamage = "Sector " + strSect + " efficiency reduced from " _
        + CStr(seceff) + "% to " + CStr(NewEff) + "%"
End If
Exit Function
Empire_Error:
EmpireError "ApplySectorDamage", vbNullString, strSect

End Function

Public Function ApplyShipDamage(strShip As String, nDamage As Integer, Optional bAbsolute As Boolean) As String
On Error GoTo Empire_Error
If IsMissing(bAbsolute) Then
    bAbsolute = False
End If

Dim NewEff As Integer
Dim hldIndex As String
Dim shipid As Integer
Dim seceff As Integer

shipid = CInt(strShip)

hldIndex = rsShip.Index
rsShip.Index = "PrimaryKey"
'check for your own ships
rsShip.Seek "=", shipid
If Not rsShip.NoMatch Then
    seceff = CStr(rsShip.Fields("eff").Value)
    If bAbsolute Then
        NewEff = seceff - nDamage
    Else
        'neweff = need to define raw damage routine
    End If
    rsShip.Edit
    rsShip.Fields("eff").Value = NewEff
    rsShip.Update
    rsShip.Index = hldIndex
Else
    'Check for enemy ship
    rsShip.Index = hldIndex
    hldIndex = rsEnemyUnit.Index
    rsEnemyUnit.Index = "ByID"
    rsEnemyUnit.Seek "=", "s", shipid
    If Not rsEnemySect.NoMatch Then
        seceff = rsEnemySect.Fields("eff").Value
        If seceff > 100 Then
            seceff = 100
        End If
        
        If bAbsolute Then
            NewEff = seceff - nDamage
        Else
            'need to develop this routine
        End If
        rsEnemyUnit.Edit
        rsEnemyUnit.Fields("eff").Value = NewEff
        rsEnemyUnit.Update
        rsEnemyUnit.Index = hldIndex
    Else
        'build a ship record
        rsEnemyUnit.AddNew
        seceff = 100
        If bAbsolute Then
            NewEff = seceff - nDamage
        Else
            'need to develop this routine
        End If
        rsEnemyUnit.Fields("eff").Value = NewEff
        rsEnemyUnit.Fields("id").Value = shipid
        rsEnemyUnit.Fields("class").Value = "s"
        rsEnemyUnit.Fields("owner").Value = 0
        rsEnemyUnit.Fields("x").Value = -1
        rsEnemyUnit.Fields("y").Value = 0
        rsEnemyUnit.Fields("mil").Value = -1
        rsEnemyUnit.Update
    End If
End If

ApplyShipDamage = "Ship (#" + strShip + ") efficiency reduced from " _
    + CStr(seceff) + "% to " + CStr(NewEff) + "%"

Exit Function
Empire_Error:
EmpireError "ApplyShipDamage", vbNullString, strShip

End Function

Public Sub CopySectorInfo(sx1 As Integer, sx2 As Integer, sy1 As Integer, sy2 As Integer, dx As Integer, dy As Integer, CopyType As Integer)
'On Error GoTo Empire_Error

Dim X As Integer
Dim Y As Integer
Dim tx As Integer
Dim ty As Integer
Dim n As Integer
Dim vhld() As Variant

For X = sx1 To sx2
    For Y = sy1 To sy2
        If (X + Y) Mod 2 = 0 Then
            rsSectors.Seek "=", X, Y
            If Not rsSectors.NoMatch Then
                ReDim vhld(0 To rsSectors.Fields.Count)
                For n = 0 To rsSectors.Fields.Count - 1
                    vhld(rsSectors.Fields(n).OrdinalPosition) = rsSectors.Fields(n).Value
                Next n
                tx = dx + X - sx1
                ty = dy + Y - sy1
                CopySector vhld, tx, ty, CopyType
            End If
        End If
    Next Y
Next X

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

End Sub


Public Sub CopySector(vhld As Variant, tx As Integer, ty As Integer, CopyType As Integer)
'On Error GoTo Empire_Error

Dim n As Integer
Dim strDes As String
Dim strCmd As String
Dim strCmdDes As String
Dim strCom As String
Dim intRes As Integer

rsSectors.Seek "=", tx, ty
If Not rsSectors.NoMatch Then
    rsSectors.Edit
    
    strDes = vhld(rsSectors.Fields("des").OrdinalPosition)
    If rsSectors.Fields("des") <> strDes Then
        rsSectors.Fields("des").Value = strDes
        
        'designate the sector
        strCmdDes = "designate " + SectorString(tx, ty) + " " + strDes
        frmEmpCmd.SubmitEmpireCommand strCmdDes, True
        rsSectors.Fields("eff").Value = 0
    End If
    
    If strDes <> "." Then
        intRes = vhld(rsSectors.Fields("eff").OrdinalPosition) - rsSectors.Fields("eff").Value
        If intRes <> 0 Then
            rsSectors.Fields("eff").Value = rsSectors.Fields("eff").Value + intRes
            strCmd = "setsector eff " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        
        intRes = vhld(rsSectors.Fields("mob").OrdinalPosition) - rsSectors.Fields("mob").Value
        If intRes <> 0 Then
            rsSectors.Fields("mob").Value = rsSectors.Fields("mob").Value + intRes
            strCmd = "setsector mob " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        
        intRes = vhld(rsSectors.Fields("work").OrdinalPosition) - rsSectors.Fields("work").Value
        If intRes <> 0 Then
            rsSectors.Fields("work").Value = rsSectors.Fields("work").Value + intRes
            strCmd = "setsector work " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        
        intRes = vhld(rsSectors.Fields("avail").OrdinalPosition) - rsSectors.Fields("avail").Value
        If intRes <> 0 Then
            rsSectors.Fields("avail").Value = rsSectors.Fields("avail").Value + intRes
            strCmd = "setsector avail " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        '102203 rjk: Added road / rail / defense and mines
        intRes = vhld(rsSectors.Fields("road").OrdinalPosition) - rsSectors.Fields("road").Value
        If intRes <> 0 Then
            strCmd = "edit land " + SectorString(tx, ty) + " :R " + CStr(vhld(rsSectors.Fields("road").OrdinalPosition))
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        intRes = vhld(rsSectors.Fields("rail").OrdinalPosition) - rsSectors.Fields("rail").Value
        If intRes <> 0 Then
            strCmd = "edit land " + SectorString(tx, ty) + " :r " + CStr(vhld(rsSectors.Fields("rail").OrdinalPosition))
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        intRes = vhld(rsSectors.Fields("defense").OrdinalPosition) - rsSectors.Fields("defense").Value
        If intRes <> 0 Then
            strCmd = "edit land " + SectorString(tx, ty) + " :d " + CStr(vhld(rsSectors.Fields("defense").OrdinalPosition))
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        If Not IsNull(vhld(rsSectors.Fields("sdes").OrdinalPosition)) Then
            If vhld(rsSectors.Fields("sdes").OrdinalPosition) <> rsSectors.Fields("sdes").Value Then
                If vhld(rsSectors.Fields("sdes").OrdinalPosition) = " " Then
                    strCmd = "edit land " + SectorString(tx, ty) + " :S " + CStr(vhld(rsSectors.Fields("des").OrdinalPosition))
                Else
                    strCmd = "edit land " + SectorString(tx, ty) + " :S " + CStr(vhld(rsSectors.Fields("sdes").OrdinalPosition))
                End If
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
        End If

        If Not IsNull(rsSectors.Fields("lmines")) Then '080404 rjk: Added Null check to lmines.
            If IsNull(vhld(rsSectors.Fields("lmines").OrdinalPosition)) Then
                intRes = -rsSectors.Fields("lmines").Value
            Else
                intRes = vhld(rsSectors.Fields("lmines").OrdinalPosition) - rsSectors.Fields("lmines").Value
            End If
            If intRes <> 0 Then
                strCmd = "edit land " + SectorString(tx, ty) + " :M " + CStr(vhld(rsSectors.Fields("lmines").OrdinalPosition))
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
        End If
    End If
    
    'set resources
    If CopyType > 0 Then
        If strDes <> "." Then
            intRes = vhld(rsSectors.Fields("min").OrdinalPosition)
            If intRes <> rsSectors.Fields("min").Value Then
                rsSectors.Fields("min").Value = intRes
                strCmd = "setresource iron " + SectorString(tx, ty) + " " + CStr(intRes)
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
            
            intRes = vhld(rsSectors.Fields("gold").OrdinalPosition)
            If intRes <> rsSectors.Fields("gold").Value Then
                rsSectors.Fields("gold").Value = intRes
                strCmd = "setresource gold " + SectorString(tx, ty) + " " + CStr(intRes)
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
            
            intRes = vhld(rsSectors.Fields("uran").OrdinalPosition)
            If intRes <> rsSectors.Fields("uran").Value Then
                rsSectors.Fields("uran").Value = intRes
                strCmd = "setresource uran " + SectorString(tx, ty) + " " + CStr(intRes)
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
        End If
        
        intRes = vhld(rsSectors.Fields("ocontent").OrdinalPosition)
        If intRes <> rsSectors.Fields("ocontent").Value Then
            rsSectors.Fields("ocontent").Value = intRes
            strCmd = "setresource oil " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        
        intRes = vhld(rsSectors.Fields("fert").OrdinalPosition)
        If intRes <> rsSectors.Fields("fert").Value Then
            rsSectors.Fields("fert").Value = intRes
            strCmd = "setresource fert " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
    End If
    
    'give commodities
    If CopyType > 1 And strDes <> "." Then
        strCom = "mil  gun  shellfood civ  uw   iron lcm  hcm  oil  dust pet  rad  bar  "
        For n = 0 To 13
            strCmd = Trim$(Mid$(strCom, 5 * n + 1, 5))
            intRes = vhld(rsSectors.Fields(strCmd).OrdinalPosition) - rsSectors.Fields(strCmd).Value
            If intRes <> 0 Then
                rsSectors.Fields(strCmd).Value = rsSectors.Fields(strCmd).Value + intRes
                strCmd = "give " + strCmd + " " + SectorString(tx, ty) + " " + CStr(intRes)
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
        Next n
    End If
Else
    rsSectors.AddNew
    
    rsSectors.Fields("x").Value = tx
    rsSectors.Fields("y").Value = ty
    
    strDes = vhld(rsSectors.Fields("des").OrdinalPosition)
    rsSectors.Fields("des").Value = strDes
        
        'designate the sector
    strCmdDes = "designate " + SectorString(tx, ty) + " " + strDes
    frmEmpCmd.SubmitEmpireCommand strCmdDes, True
    
    If strDes <> "." Then
        intRes = vhld(rsSectors.Fields("eff").OrdinalPosition)
        rsSectors.Fields("eff").Value = intRes
        strCmd = "setsector eff " + SectorString(tx, ty) + " " + CStr(intRes)
        frmEmpCmd.SubmitEmpireCommand strCmd, True
        
        intRes = vhld(rsSectors.Fields("mob").OrdinalPosition)
        rsSectors.Fields("mob").Value = intRes
        strCmd = "setsector mob " + SectorString(tx, ty) + " " + CStr(intRes)
        frmEmpCmd.SubmitEmpireCommand strCmd, True
        
        intRes = vhld(rsSectors.Fields("work").OrdinalPosition)
        rsSectors.Fields("work").Value = intRes
        strCmd = "setsector work " + SectorString(tx, ty) + " " + CStr(intRes)
        frmEmpCmd.SubmitEmpireCommand strCmd, True
        
        intRes = vhld(rsSectors.Fields("avail").OrdinalPosition)
        rsSectors.Fields("avail").Value = intRes
        strCmd = "setsector avail " + SectorString(tx, ty) + " " + CStr(intRes)
        frmEmpCmd.SubmitEmpireCommand strCmd, True

        strCmd = "edit land " + SectorString(tx, ty) + " :R " + CStr(vhld(rsSectors.Fields("road").OrdinalPosition))
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    
        strCmd = "edit land " + SectorString(tx, ty) + " :r " + CStr(vhld(rsSectors.Fields("rail").OrdinalPosition))
        frmEmpCmd.SubmitEmpireCommand strCmd, True

        strCmd = "edit land " + SectorString(tx, ty) + " :d " + CStr(vhld(rsSectors.Fields("defense").OrdinalPosition))
        frmEmpCmd.SubmitEmpireCommand strCmd, True
        
        If Not IsNull(vhld(rsSectors.Fields("sdes").OrdinalPosition)) Then
            If Len(vhld(rsSectors.Fields("sdes").OrdinalPosition)) > 0 Then
                If vhld(rsSectors.Fields("sdes").OrdinalPosition) = " " Then
                    strCmd = "edit land " + SectorString(tx, ty) + " :S " + CStr(vhld(rsSectors.Fields("des").OrdinalPosition))
                Else
                    strCmd = "edit land " + SectorString(tx, ty) + " :S " + CStr(vhld(rsSectors.Fields("sdes").OrdinalPosition))
                End If
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
        End If
        
        If Not IsNull(rsSectors.Fields("lmines").Value) Then
            strCmd = "edit land " + SectorString(tx, ty) + " :M " + CStr(vhld(rsSectors.Fields("lmines").OrdinalPosition))
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
    End If
    
    'set resources
    If CopyType > 0 Then
        If strDes <> "." Then
            intRes = vhld(rsSectors.Fields("min").OrdinalPosition)
            rsSectors.Fields("min").Value = intRes
            strCmd = "setresource iron " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
           
            intRes = vhld(rsSectors.Fields("gold").OrdinalPosition)
            rsSectors.Fields("gold").Value = intRes
            strCmd = "setresource gold " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
            
            intRes = vhld(rsSectors.Fields("uran").OrdinalPosition)
            rsSectors.Fields("uran").Value = intRes
            strCmd = "setresource uran " + SectorString(tx, ty) + " " + CStr(intRes)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        
        intRes = vhld(rsSectors.Fields("ocontent").OrdinalPosition)
        rsSectors.Fields("ocontent").Value = intRes
        strCmd = "setresource oil " + SectorString(tx, ty) + " " + CStr(intRes)
        frmEmpCmd.SubmitEmpireCommand strCmd, True
        
        intRes = vhld(rsSectors.Fields("fert").OrdinalPosition)
        rsSectors.Fields("fert").Value = intRes
        strCmd = "setresource fert " + SectorString(tx, ty) + " " + CStr(intRes)
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    End If
    
    'give commodities
    If CopyType > 1 And strDes <> "." Then
        strCom = "mil  gun  shellfood civ  uw   iron lcm  hcm  oil  dust pet  rad  bar  "
        For n = 0 To 13
            strCmd = Trim$(Mid$(strCom, 5 * n + 1, 5))
            intRes = vhld(rsSectors.Fields(strCmd).OrdinalPosition)
            rsSectors.Fields(strCmd).Value = intRes
            If intRes <> 0 Then
                strCmd = "give " + strCmd + " " + SectorString(tx, ty) + " " + CStr(intRes)
                frmEmpCmd.SubmitEmpireCommand strCmd, True
            End If
        Next n
    End If
End If
rsSectors.Update

End Sub


Public Function DefenseStrength(Sector As String) As Single

Dim sector_strength As Single
Dim d_dstr As Single
Dim sct_effic As Single
Dim sct_type As String
Dim sx As Integer
Dim sy As Integer

If Not ParseSectors(sx, sy, Sector) Then
    DefenseStrength = 1
    Exit Function
End If

rsEnemySect.Seek "=", sx, sy
If rsEnemySect.NoMatch Then
    DefenseStrength = 1
    Exit Function
End If
sct_effic = CSng(Val(rsEnemySect.Fields("eff")))
sct_type = rsEnemySect.Fields("des")


If sct_effic < 0 Then sct_effic = 100#
sector_strength = 1#
If sct_type = "^" Then
    sector_strength = 2#
End If

d_dstr = SectorDefenseStrength(sct_type)
   
DefenseStrength = sector_strength + (d_dstr - sector_strength) * (sct_effic / 100#)

End Function

Public Function EnemyMil(Sector As String) As Integer
Dim sx As Integer
Dim sy As Integer

If Not ParseSectors(sx, sy, Sector) Then
    EnemyMil = -1
    Exit Function
End If

rsEnemySect.Seek "=", sx, sy
If rsEnemySect.NoMatch Then
    EnemyMil = -1
    Exit Function
End If
EnemyMil = Val(rsEnemySect.Fields("mil"))
End Function

Public Function VerifySubDirectory(strDirect, Optional Create As Boolean) As Boolean

' Display the names in Path that represent directories.
Dim strPath As String
Dim strFile As String
' Dim found As Boolean    8/2003 efj  removed

If IsMissing(Create) Then
    Create = False
End If

strPath = App.path + "\"  ' Set the path.
VerifySubDirectory = False
strFile = Dir(strPath, vbDirectory)   ' Retrieve the first entry.
Do While strFile <> vbNullString And Not VerifySubDirectory   ' Start the loop.
    If (GetAttr(strPath & strFile) And vbDirectory) = vbDirectory Then
        If strDirect = strFile Then
            VerifySubDirectory = True
        End If
    End If  ' it represents a directory.
    strFile = Dir    ' Get next entry.
Loop

If (Not VerifySubDirectory) And Create Then
    MkDir strPath + strDirect
    On Error GoTo Verify_Exit
    VerifySubDirectory = True
End If

Verify_Exit:
End Function

Public Sub GetCurrentStrength(Optional tsSectors As Long)
Dim strCmd As String

If Not frmOptions.bSuppressStrength Then
    strCmd = "strength *"
    If frmOptions.bFullDumpwithoutSea Then
        strCmd = strCmd + " ?des#."
        If Not IsMissing(tsSectors) Then
            strCmd = strCmd + "&timestamp>" + CStr(tsSectors)
        End If
    ElseIf Not IsMissing(tsSectors) Then
        strCmd = strCmd + " ?timestamp>" + CStr(tsSectors)
    End If
    frmEmpCmd.SubmitEmpireCommand strCmd, False
End If
End Sub

Public Sub LoadPictures()
On Error GoTo Empire_Error
Set picGreenLight = LoadPicture(App.path + "\light.gif")
Set picRedLight = LoadPicture(App.path + "\light2.gif")
Set picBothLights = LoadPicture(App.path + "\light3.gif")
Exit Sub

Empire_Error:
EmpireError "LoadPictures", vbNullString, vbNullString
End Sub

Public Function GetWinACEUTC() As String
GetWinACEUTC = Format(DateAdd("n", tz.Bias, Now), "yyyy/mm/dd hh:mm:ss")
End Function

Public Function ConvertToLocalTimeFromWinACEUTC(strTime As String) As Date
ConvertToLocalTimeFromWinACEUTC = DateAdd("n", -tz.Bias, DateValue(strTime) + TimeValue(strTime))
End Function

Public Sub GetSectors(Optional bFull As Boolean)
Dim strCond As String

If Not bFull Then
    If frmOptions.bFullDumpwithoutSea And bDeity Then
        strCond = " ?des#.&timestamp>" + CStr(tsSectors)
    Else
        strCond = " ?timestamp>" + CStr(tsSectors)
    End If
Else
    If frmOptions.bFullDumpwithoutSea And bDeity Then
        strCond = " ?des#."
    Else
        strCond = ""
    End If
End If

If Not VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "xdump sect *" + strCond, False
Else
    frmEmpCmd.SubmitEmpireCommand "dump *" + strCond, False
End If
End Sub

Public Sub SubmitTelegramRead(bRead As Boolean, bDisplay As Boolean)
Dim strRead As String
If bRead Then
    strRead = "y"
Else
    strRead = "n"
End If
If bDeity And (VersionCheck(4, 2, 21) = VER_LESS) Then
    frmEmpCmd.SubmitEmpireCommand "read " + str$(CountryNumber) + " " + strRead, bDisplay
Else
    frmEmpCmd.SubmitEmpireCommand "read " + strRead, bDisplay
End If
End Sub

Public Function VersionCheck(iMajor As Integer, iMinor As Integer, iRevision) As enumVersion
If Not (rsVersion.BOF And rsVersion.EOF) Then
    If rsVersion.Fields("Major") > iMajor Then
        VersionCheck = VER_MORE
        Exit Function
    End If
    If rsVersion.Fields("Major") < iMajor Then
        VersionCheck = VER_LESS
        Exit Function
    End If
    If rsVersion.Fields("Minor") > iMinor Then
        VersionCheck = VER_MORE
        Exit Function
    End If
    If rsVersion.Fields("Minor") < iMinor Then
        VersionCheck = VER_LESS
        Exit Function
    End If
    If rsVersion.Fields("Revision") > iRevision Then
        VersionCheck = VER_MORE
        Exit Function
    End If
    If rsVersion.Fields("Revision") < iRevision Then
        VersionCheck = VER_LESS
        Exit Function
    End If
    VersionCheck = VER_EXACT
Else
    VersionCheck = VER_LESS
End If
End Function

Public Sub GetOContent()
Dim strShipList As String

If frmOptions.bOilContentForSeaSectors And Not (rsShip.BOF And rsShip.EOF) Then
    rsShip.MoveFirst
    While Not rsShip.EOF
        If frmDrawMap.UnitCharacteristicCheck(TYPE_S_OIL, rsShip.Fields("type")) Then
            strShipList = strShipList + CStr(rsShip.Fields("id")) + "/"
        End If
        rsShip.MoveNext
    Wend
End If
If Len(strShipList) > 0 Then
    strShipList = Left$(strShipList, Len(strShipList) - 1)
    frmEmpCmd.SubmitEmpireCommand "nav " + strShipList + " vh", False
End If
End Sub

Public Sub GetPlanes(Optional bFull As Boolean)
Dim strCond As String

If IsMissing(bFull) Then
    bFull = False
End If

If Not bFull Then
    strCond = " ?timestamp>" + CStr(tsPlane)
End If

If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "pdump *" + strCond, False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump plane *" + strCond, False
End If
End Sub

Public Sub GetLandUnits(Optional bFull As Boolean)
Dim strCond As String

If IsMissing(bFull) Then
    bFull = False
End If

If Not bFull Then
    strCond = " ?timestamp>" + CStr(tsLand)
End If

If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "ldump *" + strCond, False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump land *" + strCond, False
End If
End Sub

Public Sub GetShips(Optional bFull As Boolean)
Dim strCond As String

If IsMissing(bFull) Then
    bFull = False
End If

If Not bFull Then
    strCond = " ?timestamp>" + CStr(tsShip)
End If

If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "sdump *" + strCond, False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump ship *" + strCond, False
End If
End Sub

Public Sub GetOrders(bUpdate As Boolean)
If bUpdate Then
    frmEmpCmd.SubmitEmpireCommand "db1", False
End If
If VersionCheck(4, 3, 0) = VER_LESS Then
    DeleteAllRecords rsShipOrders
    frmEmpCmd.SubmitEmpireCommand "sorder *", False
    frmEmpCmd.SubmitEmpireCommand "qorder *", False
Else
    GetShips True
End If
If bUpdate Then
    frmEmpCmd.SubmitEmpireCommand "db2", False
End If
End Sub

Public Sub GetNukes(Optional bFull As Boolean)
Dim strCond As String

If IsMissing(bFull) Then
    bFull = False
End If

If Not bFull Then
    strCond = " ?timestamp>" + CStr(tsNuke)
End If

If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "ndump *" + strCond, False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump nuke *" + strCond, False
End If
End Sub

Public Sub GetLost(Optional bFull As Boolean)
Dim strCond As String

If IsMissing(bFull) Then
    bFull = False
End If

If Not bFull Then
    strCond = " ?timestamp>" + CStr(tsLost)
End If

If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "lost *" + strCond, False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump lost *" + strCond, False
End If
End Sub

Public Sub GetNation()
If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "nation", False
ElseIf VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
    frmEmpCmd.SubmitEmpireCommand "xdump nat *", False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump coun *", False
End If
End Sub

Public Sub GetCountryList()
If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "relation", False
ElseIf VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
    frmEmpCmd.SubmitEmpireCommand "xdump cou *", False
    frmEmpCmd.SubmitEmpireCommand "relation", False
ElseIf frmOptions.bSPAtlantis Then
    frmEmpCmd.SubmitEmpireCommand "xdump nat *", False
    frmEmpCmd.SubmitEmpireCommand "relation", False
Else
    frmEmpCmd.SubmitEmpireCommand "xdump nat *", False
End If
End Sub

Public Sub GetUpdate(bAddCommandList As Boolean)

If VersionCheck(4, 3, 10) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "update", bAddCommandList
Else
    frmEmpCmd.SubmitEmpireCommand "xdump updates 0", bAddCommandList
End If
End Sub

Public Sub StartUpdateTimer()
If VersionCheck(4, 3, 10) <> VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "xdump game *", False
End If
GetUpdate False
End Sub
Public Sub UpdateCommands()
frmDrawMap.UpdateTimer.Interval = 0
If Len(BatchScriptUpdate) > 0 Then
    frmScript.ExectuteBatchScript BatchScriptUpdate
    BatchScriptUpdate = vbNullString
End If
frmEmpCmd.SubmitEmpireCommand "db1", False
GetUpdate False
GetSectors True
GetCurrentStrength
GetOContent
GetPlanes True
GetLandUnits True
GetShips True
GetNukes True
GetLost True
If frmEmpCmd.bAutoUpdate Then
    GetNation
End If
frmEmpCmd.SubmitEmpireCommand "db2", False
End Sub

Function IsArrayEmpty(varArray As Variant) As Boolean
' Determines whether an array contains any elements.
' Returns False if it does contain elements, True
' if it does not.

Dim lngUBound As Long

On Error Resume Next
' If the array is empty, an error occurs when you
' check the array's bounds.
lngUBound = UBound(varArray)
If Err.Number <> 0 Then
   IsArrayEmpty = True
Else
   IsArrayEmpty = False
End If
End Function

Public Sub FlashLog(strText As String)
Dim strlog As String

If Not frmOptions.bFLashLogging Then
    Exit Sub
End If

If VerifySubDirectory("Game Logs", True) And FlashLogFileNumber = -1 Then
    strlog = App.path + "\Game Logs\" + GameName + " Flash " + CStr(Year(Now)) + "-" + Format$(Month(Now), "00") + "-" _
        + Format$(Day(Now), "00") + ".log"

    'open the new log
    FlashLogFileNumber = FreeFile
    'if it already exists, append to it
    If Len(Dir(strlog)) > 0 Then
        Open strlog For Append As #FlashLogFileNumber
    Else
        Open strlog For Output As #FlashLogFileNumber
    End If
    Print #FlashLogFileNumber, "+++ Flash Log opened at " + Format$(Now, "hh:mm:ss") + " +++"
End If

If FlashLogFileNumber <> -1 Then
    Print #FlashLogFileNumber, strText
End If
End Sub

Public Sub CloseFlashLog()
If FlashLogFileNumber <> -1 Then
    Print #FlashLogFileNumber, "+++ Flash Log closed at " + Format$(Now, "hh:mm:ss") + " +++"
    Close FlashLogFileNumber
End If
End Sub
