Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module WinAceRoutines
	
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
	Public Const READ_STATE_FIRSTSERVERINIT As Short = 1
	Public Const READ_STATE_USERCMDOK As Short = 2
	Public Const READ_STATE_COUNCMDOK As Short = 3
	Public Const READ_STATE_PASSCMDOK As Short = 4
	Public Const READ_STATE_KILLPROC As Short = 5
	Public Const READ_STATE_PLAYINIT As Short = 6
	Public Const READ_STATE_SECTOR_DUMP As Short = 7
	Public Const READ_STATE_PLANE_DUMP As Short = 8
	Public Const READ_STATE_LAND_DUMP As Short = 9
	Public Const READ_STATE_SHIP_DUMP As Short = 10
	Public Const READ_STATE_FULL_BMAP As Short = 11
	Public Const READ_STATE_REPORT As Short = 12
	'Public Const READ_STATE_REQ_REPORT = 13 100703 rjk: Not required as Nation list is done in relations now.
	Public Const READ_STATE_REQ_RELATIONS As Short = 14
	Public Const READ_STATE_EXPLORE_CIRCLE As Short = 15
	Public Const READ_STATE_UNIT_MOVE As Short = 16
	Public Const READ_STATE_NEXT_UPDATE As Short = 17
	Public Const READ_STATE_FIRE_SEQ As Short = 18
	Public Const READ_STATE_LAUNCH As Short = 19
	Public Const READ_STATE_BOMB As Short = 20
	Public Const READ_STATE_TELEGRAM_READ As Short = 21
	Public Const READ_STATE_TELEGRAM_WRITE As Short = 22
	Public Const READ_STATE_LOOK As Short = 23
	Public Const READ_STATE_SPY As Short = 24
	Public Const READ_STATE_RECON As Short = 25
	Public Const READ_STATE_COASTWATCH As Short = 26
	Public Const READ_STATE_ATTACK As Short = 27
	Public Const READ_STATE_VERSION As Short = 28
	Public Const READ_STATE_LOST_DUMP As Short = 29
	Public Const READ_STATE_ORIGIN As Short = 30
	Public Const READ_STATE_LISTSANC As Short = 31
	Public Const READ_STATE_STARVATION As Short = 32
	Public Const READ_STATE_SONAR As Short = 33
	Public Const READ_STATE_SHOW As Short = 34
	Public Const READ_STATE_PARA As Short = 35
	Public Const READ_STATE_BREAK As Short = 36
	Public Const READ_STATE_NATION As Short = 37
	Public Const READ_STATE_NUKE_DUMP As Short = 38
	Public Const READ_STATE_BOMBCL As Short = 39
	Public Const READ_STATE_ORDER As Short = 40
	Public Const READ_STATE_SHOWORDER As Short = 41
	Public Const READ_STATE_LANDMISSION As Short = 42
	Public Const READ_STATE_SHIPMISSION As Short = 43
	Public Const READ_STATE_PLANEMISSION As Short = 44
	Public Const READ_STATE_EDIT As Short = 45
	Public Const READ_STATE_TELEGRAM_READ_CL As Short = 46
	Public Const READ_STATE_WIRE As Short = 47
	Public Const READ_STATE_STRENGTH As Short = 48
	Public Const READ_STATE_NEWSPAPER As Short = 49 '112903 rjk: Added to gather intelligence from the newspaper.
	Public Const READ_STATE_PLAYERS As Short = 50 '120103 rjk: Added for displaying current players
	Public Const READ_STATE_SCUTTLE As Short = 51 '040404 rjk: Added for automatic answer to scuttle and scrap
	Public Const READ_STATE_XDUMP As Short = 52 '280804 rjk: Added xdump commands
	Public Const READ_STATE_OPTIONS As Short = 53 '280505 rjk: Added for login options (UTF-8).
	Public Const READ_STATE_FLASH_WRITE As Short = 54
	
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
	
	Public Const BF_NEVERABORT As Short = 0
	
	Public Const BF_ABORTONRETURNFIRE As Short = 1
	Public Const BF_ABORTREMAINING As Short = 2
	
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
	
	Private Structure SYSTEMTIME
		Dim wYear As Short
		Dim wMonth As Short
		Dim wDayOfWeek As Short
		Dim wDay As Short
		Dim wHour As Short
		Dim wMinute As Short
		Dim wSecond As Short
		Dim wMilliseconds As Short
	End Structure
	
	Public Structure TIME_ZONE_INFORMATION
		Dim Bias As Integer
		<VBFixedArray(31)> Dim StandardName() As Short
		Dim StandardDate As SYSTEMTIME
		Dim StandardBias As Integer
		<VBFixedArray(31)> Dim DaylightName() As Short
		Dim DaylightDate As SYSTEMTIME
		Dim DaylightBias As Integer
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim StandardName(31)
			ReDim DaylightName(31)
		End Sub
	End Structure
	
	'UPGRADE_WARNING: Structure TIME_ZONE_INFORMATION may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Integer
	
	Private Const TIME_ZONE_ID_INVALID As Integer = &HFFFFFFFF
	Private Const TIME_ZONE_ID_STANDARD As Integer = 1
	Private Const TIME_ZONE_ID_UNKNOWN As Integer = 0
	Private Const TIME_ZONE_ID_DAYLIGHT As Integer = 2
	
	'UPGRADE_WARNING: Arrays in structure tz may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public tz As TIME_ZONE_INFORMATION
	
	'Public Const MaxExploreRadius = 10 'why so small? drk@unxsoft.com 7/15/02 fixme
	Public Const MaxExploreRadius As Short = 20
	Enum enumExploreMode '09/01/03 rjk: Added for additional exploration modes
		EXP_COMPLETE = 1
		EXP_BORDER = 2
		EXP_WHEEL = 3
	End Enum
	
	Enum enumVersion '21/03/05 rjk: Added for Version Check Function
		VER_EXACT = -1
		VER_LESS = 0
		VER_MORE = 1
	End Enum
	
	Public MapSizeX As Short
	Public MapSizeY As Short
	Public bDeity As Boolean
	Public bUTF8 As Boolean
	Public bXMLRPC As Boolean
	Public xmlrpcServer As EmpXMLRPC
	
	Public TimeIdle As Integer
	
	Public picGreenLight As System.Drawing.Image
	Public picRedLight As System.Drawing.Image
	Public picBothLights As System.Drawing.Image
	
	Public Nations As New EmpNationList
	Public Items As EmpItems
	Public etsTable As EmpTables
	Public estsTable As EmpSymbolTables
	Public CountryNumber As Short
	
	Public BatchScript1File As String
	Public BatchScript1Time As Integer
	Public BatchScriptUpdate As String
	
	Public FlashLogFileNumber As Short
	
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
	
	Public Structure tTelegram
		Dim bEnabled As Boolean
		Dim iColumn As Short
		Dim eTelegramSound As enumTelegramSound
	End Structure
	
	'UPGRADE_WARNING: Lower bound of array tTelegramOptions was changed from totTelegram,tlWarning to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public tTelegramOptions(enumTelegramOptionType.totMOTD, enumTelegramLevel.tlHardLimit) As tTelegram
	
	Public Structure tScriptButton
		Dim strFileName As String
		Dim bSendOutputReportWindow As Boolean
		Dim strImageFileName As String
		Dim bJumpPoint As Boolean
		Dim iJumpPoint As Short
	End Structure
	Public Const NUMBER_SCRIPT_BUTTONS As Short = 10
	
	'UPGRADE_WARNING: Lower bound of array tScriptButtons was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public tScriptButtons(NUMBER_SCRIPT_BUTTONS) As tScriptButton
	
	Public Sub LoadBmap(ByRef buf As String, ByRef StartX As Short, ByRef linelen As Short)
		Dim CurY As Short
		Dim CurX As Short
		Dim X As Short
		Dim fieldname As String
		Dim BufferLen As Short
		Dim hldIndex As String
		
		On Error GoTo LoadBmap_Exit
		
		hldIndex = rsEnemySect.Index
		rsEnemySect.Index = "PrimaryKey"
		
		CurY = Val(Left(buf, 4))
		
		'Figure out how far bmap goes
		BufferLen = Len(buf)
		If linelen < BufferLen Then
			BufferLen = linelen
		End If
		
		If Left(buf, 4) = "    " Then
			Exit Sub
		End If
		
		For X = 5 To BufferLen
			If Mid(buf, X, 1) <> " " Then
				CurX = StartX + X - 6
				rsBmap.Seek("=", CurX, CurY)
				If rsBmap.NoMatch Then
					rsBmap.AddNew()
					fieldname = "x"
					rsBmap.Fields(fieldname).Value = CurX
					fieldname = "y"
					rsBmap.Fields(fieldname).Value = CurY
					fieldname = "des"
					rsBmap.Fields(fieldname).Value = Mid(buf, X, 1)
					rsBmap.Update()
				ElseIf rsBmap.Fields("des").Value <> Mid(buf, X, 1) Then 
					rsBmap.Edit()
					fieldname = "des"
					rsBmap.Fields(fieldname).Value = Mid(buf, X, 1)
					rsBmap.Update()
				End If
				If Mid(buf, X, 1) = "\" Then '020704 rjk: Clear any enemy sect intelligence if a sector is turned into a wasteland
					rsEnemySect.Seek("=", CurX, CurY)
					If Not rsEnemySect.NoMatch Then
						rsEnemySect.Delete()
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
		EmpireError("LoadBmap", fieldname, buf)
	End Sub
	
	Public Sub ImportBmap()
		On Error GoTo Empire_Error
		Dim CurY As Short
		Dim CurX As Short
		Dim X As Short
		Dim BufferLen As Short
		Dim filenumber As Short
		Dim FileName As String
		Dim bdes As String
		
		
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = True
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.ShowReadOnly = False
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = vbNullString
		
		If VerifySubDirectory("Exports", True) Then
			frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath & "\Exports"
		End If
		
		' Display the open dialog box
		On Error GoTo ImportBmap_Exit
		frmCommonDialog.cdbOpen.ShowDialog()
		On Error GoTo Empire_Error
		
		'frmCommonDialog.cdb.FileTitle
		
		FileName = frmCommonDialog.cdbOpen.FileName
		filenumber = FreeFile ' Get unused file number.
		FileOpen(filenumber, FileName, OpenMode.Input)
		
		Dim Line As Short
		Dim bmaphead1 As String
		Dim bmaphead2 As String
		Dim bmaphead3 As String
		Dim StartX As Short
		Dim AdjX As Short
		Dim AdjY As Short
		Dim temp As Short
		Dim linelen As Short
		Dim overwrite As Boolean
		Dim strLine As String
		Dim strSect As String
		
start: 
		strSect = InputBox("Enter your x,y co-ordinates for the imported bmaps origin")
		If Not ParseSectors(AdjX, AdjY, strSect) Then
			If MsgBox("Sectors not valid", MsgBoxStyle.RetryCancel + MsgBoxStyle.Critical, "Import Bmap Error") = MsgBoxResult.Retry Then
				GoTo start
			End If
			Exit Sub
		End If
		
		overwrite = (MsgBox("Do you want to overwrite your existing bmap", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Import Bmap") = MsgBoxResult.Yes)
		
		While Not EOF(filenumber)
			strLine = LineInput(filenumber)
			Line = Line + 1
			If Line = 1 Then
				bmaphead1 = strLine
				linelen = Len(RTrim(bmaphead1))
			ElseIf Line = 2 Then 
				bmaphead2 = strLine
				'Calculate the starting x
				StartX = Val(Mid(bmaphead1, 6, 1)) * 10 + Val(Mid(bmaphead2, 6, 1))
				temp = Val(Mid(bmaphead1, 7, 1)) * 10 + Val(Mid(bmaphead2, 7, 1))
				If temp < StartX Then
					StartX = -1 * StartX
				End If
			ElseIf Line = 3 And Left(strLine, 4) = "    " Then 
				'Handle three line bmap - recompute start x
				bmaphead3 = strLine
				StartX = Val(Mid(bmaphead1, 6, 1)) * 100 + Val(Mid(bmaphead2, 6, 1)) * 10 + Val(Mid(bmaphead3, 6, 1))
				temp = Val(Mid(bmaphead1, 7, 1)) * 100 + Val(Mid(bmaphead2, 7, 1)) * 10 + Val(Mid(bmaphead3, 7, 1))
				If temp < StartX Then
					StartX = -1 * StartX
				End If
			Else
				CurY = Val(Left(strLine, 4)) + AdjY
				
				'Figure out how far bmap goes
				BufferLen = Len(strLine)
				If linelen < BufferLen Then
					BufferLen = linelen
				End If
				
				If Left(strLine, 4) <> "    " Then
					For X = 5 To BufferLen
						If Mid(strLine, X, 1) <> " " Then
							CurX = StartX + X - 6 + AdjX
							
							'Adjust the x and y to compesate for
							'going over the boundry
							OffsetSectorLocation(CurX, CurY, 0, 0)
							
							'process sector by sector
							bdes = Mid(strLine, X, 1)
							If bdes = "?" Then
								bdes = "0"
							End If
							
							rsBmap.Seek("=", CurX, CurY)
							If rsBmap.NoMatch Then
								frmEmpCmd.SubmitEmpireCommand("bdes " & SectorString(CurX, CurY) & " " & bdes, False)
							ElseIf rsBmap.Fields("des").Value <> bdes Then 
								'check and make sure land and water line up with
								'current bmap
								If (InStr(".=X", rsBmap.Fields("des").Value) > 0 And InStr(".=X", bdes) = 0) Or (InStr(".=X", rsBmap.Fields("des").Value) = 0 And InStr(".=X", bdes) > 0) Then
									
									If MsgBox("Warning.  Possible error in bmap lineup. " & "Current Bmap shows '" + rsBmap.Fields("des").Value + "', imported bmap shows '" + bdes + "' Do you want to continue?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Import Bmap") = MsgBoxResult.No Then Exit Sub
								End If
								If overwrite Then
									frmEmpCmd.SubmitEmpireCommand("bdes " & SectorString(CurX, CurY) & " " & bdes, False)
									'alway mark map if minefield
								ElseIf rsBmap.Fields("des").Value = "." And bdes = "X" Then 
									frmEmpCmd.SubmitEmpireCommand("bdes " & SectorString(CurX, CurY) & " " & bdes, False)
								End If
							End If
						End If
					Next X
				End If
			End If
		End While
		Exit Sub
		
		'Error handling routine
ImportBmap_Exit: 
		Exit Sub
Empire_Error: 
		EmpireError("ImportBmap", vbNullString, vbNullString)
		
	End Sub
	
	Public Sub RefreshDataBase()
		On Error GoTo Empire_Error
		frmEmpCmd.ForceUpdates = True
		
		'Send Code to start database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		
		'If version information needs filling, request version
		If rsVersion.BOF And rsVersion.EOF Then
			frmEmpCmd.SubmitEmpireCommand("version", False)
		End If
		
		If VersionCheck(4, 3, 4) = enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
			GetNation()
			GetCountryList()
		Else
			GetCountryList()
			GetNation()
		End If
		If Not VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			RequestConfigurationTables()
		Else
			'used to calculate the max safe civ at connection when autoupdate is on
			frmEmpCmd.SubmitEmpireCommand("show sect s", False)
			'022604 rjk: Added show sect c/b to ensure the database is fill, as it starts empty
			frmEmpCmd.SubmitEmpireCommand("show sect c", False)
			frmEmpCmd.SubmitEmpireCommand("show sect b", False)
		End If
		
		'Request an update of Sector information
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		GetOContent() '110605 rjk: Added ability to get OContent for Sea Sectors
		
		'Request an update of Plane information
		GetPlanes()
		'Request an update of land information
		GetLandUnits()
		'Request an update of ship information
		GetShips()
		
		'Request an update of nuke information
		GetNukes()
		
		'Request update lost information
		GetLost()
		
		'Request an update of ship information
		frmEmpCmd.SubmitEmpireCommand("bmap *", False)
		
		'send Code to end database update
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Exit Sub
Empire_Error: 
		EmpireError("RefreshDataBase", vbNullString, vbNullString)
	End Sub
	Public Function ProcessSectorDump(ByRef strLine As String, ByRef CurX As Short, ByRef CurY As Short) As String
		
		Dim strTokens() As String
		' Dim top As Integer    8/2003 efj  removed
		Dim numTokens As Short
		' Dim i As Integer    8/2003 efj  removed
		Dim n As Short
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
		
		ParseDumpString(strLine, strTokens, bDeity)
		
		numTokens = UBound(strTokens)
		For n = LBound(strTokens) To numTokens
			If n > 3 Or n = 2 Then
				var = rsSectors.Fields(n).Name
				rsSectors.Fields(n).Value = strTokens(n)
			ElseIf n = 3 Then 
				var = rsSectors.Fields(n).Name
				If strTokens(n) = "_" Then
					strTokens(n) = " "
				End If
				rsSectors.Fields(n).Value = strTokens(n)
			ElseIf n = 0 Then 
				CurX = CShort(strTokens(n))
			Else
				CurY = CShort(strTokens(n))
				rsSectors.Seek("=", CurX, CurY)
				If rsSectors.NoMatch Then
					rsSectors.AddNew()
				Else
					rsSectors.Edit()
				End If
				var = rsSectors.Fields(0).Name
				rsSectors.Fields(0).Value = CurX
				rsSectors.Fields(1).Value = CurY
			End If
		Next n
		If Not bDeity Then '102803 rjk: Always fill in the country field
			rsSectors.Fields("country").Value = CountryNumber
		End If
		rsSectors.Update()
		hldIndex = rsSectors.Index
		ProcessSectorDump = strTokens(0) & "," & strTokens(1)
		Exit Function
		
		'Error handling routine
ProcessSectorDump_Exit: 
		EmpireError("ProcessSectorDump", var, strLine)
		ProcessSectorDump = vbNullString
	End Function
	
	Public Sub ParseStrength(ByRef strLine As String)
		Dim sx As Short
		Dim sy As Short
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
			strLine = Mid(strLine, 5)
		End If
		
		If InStr(strLine, ",") = 0 Then 'Ensure it is a sector line
			Exit Sub
		End If
		
		'Save and restore the index on entry and exit
		hldIndex = rsSectors.Index
		rsSectors.Index = "PrimaryKey"
		
		If Not ParseSectors(sx, sy, Left(strLine, 9)) Then
			EmpireError("ParseStrength", "Not valid sector", strLine)
			Exit Sub
		End If
		
		rsSectors.Seek("=", sx, sy)
		If rsSectors.NoMatch Then
			EmpireError("ParseStrength", "Sector not found", strLine)
			Exit Sub
		End If
		
		rsSectors.Edit()
		var = "units"
		rsSectors.Fields("units").Value = Val(Mid(strLine, 23, 6))
		var = "lmines"
		If Mid(strLine, 35, 1) = "?" Then '120303 rjk: display enemy mines in red "?"
			rsSectors.Fields("lmines").Value = -1
		Else
			rsSectors.Fields("lmines").Value = Val(Mid(strLine, 30, 6))
		End If
		var = "sec_mult"
		rsSectors.Fields("sec_mult").Value = Val(Mid(strLine, 37, 5))
		var = "sec_def"
		If Val(Mid(strLine, 43, 8)) > 32767 Then '112003 rjk: Added protection to not overflow the db variable (LOTR II)
			rsSectors.Fields("sec_def").Value = 32767
		Else
			rsSectors.Fields("sec_def").Value = Val(Mid(strLine, 43, 8))
		End If
		var = "runits"
		rsSectors.Fields("runits").Value = Val(Mid(strLine, 52, 9))
		var = "tot_def"
		If Val(Mid(strLine, 62, 8)) > 32767 Then '112003 rjk: Added protection to not overflow the db variable (LOTR II)
			rsSectors.Fields("tot_def").Value = 32767
		Else
			rsSectors.Fields("tot_def").Value = Val(Mid(strLine, 62, 8))
		End If
		
		rsSectors.Update()
		hldIndex = rsSectors.Index
		
		'112003 rjk: Update the sector panel if a new strength report comes in for the sector being displayed
		'112403 rjk: Moved to record is actually up to date.
		If sx = frmDrawMap.SelX And sy = frmDrawMap.SelY Then
			frmDrawMap.FillSectorBox(sx, sy)
		End If
		
		Exit Sub
		
		'Error handling routine
ParseStrength_Exit: 
		EmpireError("ParseStrength", var, strLine)
	End Sub
	
	Public Sub ProcessLostDump(ByRef strLine As String)
		
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
		
		ParseDumpString(strLine, strTokens, bDeity)
		' numTokens = UBound(strTokens)    8/2003 efj  removed
		
		Select Case strTokens(LS_TYPE)
			Case CStr(LT_SECTOR)
				rsSectors.Seek("=", CShort(strTokens(LS_X)), CShort(strTokens(LS_Y)))
				If Not rsSectors.NoMatch Then
					rsSectors.Delete()
				End If
			Case CStr(LT_SHIP)
				rsShip.Seek("=", CShort(strTokens(LS_ID)))
				If Not rsShip.NoMatch Then
					rsShip.Delete()
				End If
			Case CStr(LT_PLANE)
				rsPlane.Seek("=", CShort(strTokens(LS_ID)))
				If Not rsPlane.NoMatch Then
					rsPlane.Delete()
				End If
			Case CStr(LT_LAND)
				rsLand.Seek("=", CShort(strTokens(LS_ID)))
				If Not rsLand.NoMatch Then
					rsLand.Delete()
				End If
			Case CStr(LT_NUKE) '101403 rjk: add to process lost nukes, needed for timestamp
				rsNuke.Seek("=", CShort(strTokens(LS_ID)))
				While Not rsNuke.NoMatch
					rsNuke.Delete()
					rsNuke.Seek("=", CShort(strTokens(LS_ID)))
				End While
		End Select
		
		'101403 rjk: added to update grid when doing scrap/scuttle operations.
		If frmDrawMap.SelX = CShort(strTokens(LS_X)) And frmDrawMap.SelY = CShort(strTokens(LS_Y)) Then
			If CDbl(strTokens(LS_TYPE)) = LT_SHIP Or CDbl(strTokens(LS_TYPE)) = LT_PLANE Or CDbl(strTokens(LS_TYPE)) = LT_LAND Then
				frmDrawMap.FillGrid()
				frmDrawMap.Refresh()
			Else
				frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
			End If
		End If
		
		Exit Sub
		
		'Error handling routine
ProcessLostDump_Exit: 
		EmpireError("ProcessLostDump", vbNullString, strLine)
		
	End Sub
	
	'102703 rjk: Modified to detect if a fill grid is required.
	Public Function ProcessDump(ByRef strLine As String, ByRef rs As DAO.Recordset, ByRef strFields() As String) As Boolean
		Dim strTokens() As String
		Dim strLineFiltered As String
		Dim n As Short
		Dim id As Integer
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
		strLine = Trim(strLine)
		strLineFiltered = Left(strLine, 1)
		For n = 2 To Len(strLine)
			If Mid(strLine, n, 1) = Chr(34) Then
				bQuote = Not bQuote
				strLineFiltered = strLineFiltered & Chr(34)
			Else
				If Mid(strLine, n - 1, 1) <> " " Or Mid(strLine, n, 1) <> " " Or bQuote Then
					strLineFiltered = strLineFiltered & Mid(strLine, n, 1)
				End If
			End If
		Next n
		strTokens = Split(strLineFiltered)
		
		'200204 rjk: Check for a quoted string at the end of the list (ship name), combined if across multiple tokens
		Dim strBuild As String
		If InStr(strTokens(UBound(strTokens)), Chr(34)) > 0 Then
			n = UBound(strTokens)
			While Left(strTokens(n), 1) <> Chr(34)
				strBuild = " " & strTokens(n) & strBuild
				n = n - 1
			End While
			strTokens(n) = strTokens(n) & strBuild
			ReDim Preserve strTokens(n)
		End If
		
		If UBound(strTokens) <> UBound(strFields) Then
			EmpireError("ProcessDump", "Different number of fields than the header", strLine)
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
				id = CInt(strTokens(n)) '111403 rjk: id is long so switch from CInt to CLng
				rs.Seek("=", id)
				If rs.NoMatch Then
					rs.AddNew()
					CheckForOldEnemyUnit(id, Left(rs.Name, 1)) '111403 rjk: Ensure there is not an old enemy intelligence for this ship.
				Else
					rs.Edit()
				End If
			End If
			Select Case rs.Fields(strFields(n)).Type
				Case DAO.DataTypeEnum.dbInteger
					rs.Fields(strFields(n)).Value = CShort(strTokens(n))
				Case DAO.DataTypeEnum.dbLong
					rs.Fields(strFields(n)).Value = CInt(strTokens(n))
				Case DAO.DataTypeEnum.dbSingle
					rs.Fields(strFields(n)).Value = Val(strTokens(n)) '070304 rjk: Switch to Val for regional settings
				Case DAO.DataTypeEnum.dbText
					If strFields(n) = "name" Then
						rs.Fields(strFields(n)).Value = Mid(strTokens(n), 2, Len(strTokens(n)) - 2)
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
			rs.Fields("country").Value = strTokens(0)
		Else
			rs.Fields("country").Value = CountryNumber
		End If
		rs.Update()
		rs.Index = hldIndex
		
		Exit Function
		
		'Error handling routine
ProcessDump_Exit: 
		EmpireError("ProcessDump", var, strLine)
		
	End Function
	
	'111403 rjk: Ensure there is not an old enemy intelligence for this ship.
	Public Sub CheckForOldEnemyUnit(ByRef id As Integer, ByRef strClass As String)
		Dim hldIndex As String
		hldIndex = rsEnemyUnit.Index
		rsEnemyUnit.Index = "ByID"
		rsEnemyUnit.Seek("=", strClass, id)
		If Not rsEnemyUnit.NoMatch Then
			rsEnemyUnit.Delete()
		End If
		rsEnemyUnit.Index = hldIndex
	End Sub
	
	'101403 rjk: Changed ProcessNukeDump to accept partial dumps
	Public Sub ProcessNukeDump(ByRef strLine As String, ByRef rs As DAO.Recordset)
		'id x y num type
		Dim n As Short
		Dim numTokens As Short
		Dim strTokens() As String
		Dim hldIndex As String
		Dim var As String
		
		On Error GoTo ProcessNukeDump_Exit
		hldIndex = rs.Index
		rs.Index = "PrimaryKey"
		
		ParseDumpString(strLine, strTokens, bDeity)
		
		numTokens = UBound(strTokens)
		
		rs.Seek("=", strTokens(LBound(strTokens)), strTokens(4))
		If Not rs.NoMatch Then
			rs.Edit()
		Else
			rs.AddNew()
			rs.Fields("id").Value = strTokens(LBound(strTokens))
		End If
		
		For n = LBound(strTokens) + 1 To numTokens
			var = rs.Fields(n).Name
			rs.Fields(n).Value = strTokens(n)
		Next n
		
		If Not bDeity Then '102803 rjk: Always fill in the country field
			rsNuke.Fields("country").Value = CountryNumber
		End If
		
		rs.Update()
		
		'101403 rjk: added to update sector when doing disarm operations.
		If frmDrawMap.SelX = CShort(strTokens(LBound(strTokens) + 1)) And frmDrawMap.SelY = CShort(strTokens(LBound(strTokens) + 2)) Then
			frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
		End If
		
		rs.Index = hldIndex
		Exit Sub
		
		'Error handling routine
ProcessNukeDump_Exit: 
		
		rs.Index = hldIndex
		EmpireError("ProcessNukeDump", var, strLine)
	End Sub
	
	Public Sub BlitzSetup(ByRef Initial As Boolean, ByRef sx1 As Short, ByRef sy1 As Short, ByRef sx2 As Short, ByRef sy2 As Short)
		On Error GoTo Empire_Error
		
		Dim strSect As String
		Dim Cap As String
		Dim Harbour As String
		Dim Oil As String
		' Dim x As Integer    8/2003 efj  removed
		Dim mines As Short
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
		
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			
			If Initial Or (rsSectors.Fields("x").Value >= sx1 And rsSectors.Fields("x").Value <= sx2 And rsSectors.Fields("y").Value >= sy1 And rsSectors.Fields("y").Value <= sy2) Then
				
				rsSectors.Edit()
				
				'0,0 is a gold mine
				If Initial And rsSectors.Fields("x").Value = 0 And rsSectors.Fields("y").Value = 0 Then
					rsSectors.Fields("des").Value = "g"
					frmEmpCmd.SubmitEmpireCommand("des 0,0 g", False)
					frmEmpCmd.SubmitEmpireCommand("thr d 0,0 1", False)
					
					'2,0 is a mine
				ElseIf Initial And rsSectors.Fields("x").Value = 2 And rsSectors.Fields("y").Value = 0 Then 
					rsSectors.Fields("des").Value = "m"
					frmEmpCmd.SubmitEmpireCommand("des 2,0 m", False)
					frmEmpCmd.SubmitEmpireCommand("thr i 2,0 1", False)
					
				ElseIf Not Initial And rsSectors.Fields("des").Value = "h" Then 
					Harbour = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					
				ElseIf Not Initial And rsSectors.Fields("des").Value = "=" Then 
					
					'Dim oil well
				ElseIf rsSectors.Fields("ocontent").Value > 10 Then 
					rsSectors.Fields("des").Value = "o"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " o", False)
					frmEmpCmd.SubmitEmpireCommand("thr o " & strSect & " 1", False)
					If rsSectors.Fields("ocontent").Value > 90 And Oil = vbNullString Then
						Oil = strSect
					End If
					
					'Dim up to 3 mines
				ElseIf rsSectors.Fields("min").Value > 90 And mines < 3 Then 
					rsSectors.Fields("des").Value = "m"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " m", False)
					frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 1", False)
					mines = mines + 1
					
					'Dim up to several golds
				ElseIf rsSectors.Fields("gold").Value > 20 Then 
					rsSectors.Fields("des").Value = "g"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " g", False)
					frmEmpCmd.SubmitEmpireCommand("thr d " & strSect & " 1", False)
					
					'Finally - make it a highway
				Else
					rsSectors.Fields("des").Value = "+"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " +", False)
				End If
				
				rsSectors.Update()
			End If
			
			rsSectors.MoveNext()
		End While
		
		'capital
		Cap = SetOpenHex("c", 0, Initial, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & Cap & " c", False)
		
		If Initial Then
			'Reset Capital
			frmEmpCmd.SubmitEmpireCommand("cap " & Cap, False)
			
			'move people to oil well
			frmEmpCmd.SubmitEmpireCommand("move c 0,0 250 " & Oil, False)
			
			'build library and move people
			strSect = SetOpenHex("l", 2, Initial, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " l", False)
			frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 150", False)
			frmEmpCmd.SubmitEmpireCommand("move c 2,0 150 " & strSect, False)
			
			'build tech center and move people
			strSect = SetOpenHex("t", 2, Initial, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " t", False)
			frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 120", False)
			frmEmpCmd.SubmitEmpireCommand("thr o " & strSect & " 60", False)
			frmEmpCmd.SubmitEmpireCommand("thr d " & strSect & " 999", False)
			frmEmpCmd.SubmitEmpireCommand("move c 2,0 100 " & strSect, False)
		End If
		
		'lcm factory
		strSect = SetOpenHex("j", 2, Initial, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " j", False)
		frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 999", False)
		frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 1", False)
		frmEmpCmd.SubmitEmpireCommand("move c 0,0 250 " & strSect, False)
		
		'harbour
		If Len(Harbour) = 0 Then
			Harbour = SetOpenHex("h", 1, Initial, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & Harbour & " h", False)
			If Initial Then
				frmEmpCmd.SubmitEmpireCommand("move c 2,0 300 " & Harbour, False)
			End If
		End If
		
		'bank
		strSect = SetOpenHex("b", 0, Initial, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " b", False)
		frmEmpCmd.SubmitEmpireCommand("thr d " & strSect & " 300", False)
		
		'hcm
		strSect = SetOpenHex("k", 2, Initial, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " k", False)
		frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 999", False)
		frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 1", False)
		
		If Not Initial Then
			'enlistment center
			strSect = SetOpenHex("e", 2, Initial, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " e", False)
			frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 200", False)
			
			'another hcm
			strSect = SetOpenHex("k", 2, Initial, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " k", False)
			frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 999", False)
			frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 1", False)
			
		End If
		
		'distribution
		If Initial Then
			strSect = "* "
		Else
			strSect = CStr(sx1) & ":" & CStr(sx2) & "," & CStr(sy1) & ":" & CStr(sy2) & " "
		End If
		
		frmEmpCmd.SubmitEmpireCommand("dist " & strSect & Harbour, False)
		frmEmpCmd.SubmitEmpireCommand("thr c " & strSect & "768", False)
		frmEmpCmd.SubmitEmpireCommand("thr u " & strSect & "868", False)
		
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors(True)
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		'restore Index
		rsSectors.Index = hldIndex
		Exit Sub
Empire_Error: 
		EmpireError("BlitzSetup", vbNullString, vbNullString)
		
	End Sub
	
	
	Private Function SetOpenHex(ByRef strDes As String, ByRef Coast As Short, ByRef Initial As Boolean, ByRef sx1 As Short, ByRef sy1 As Short, ByRef sx2 As Short, ByRef sy2 As Short) As String
		
		'Coast Values
		'0 = not on coast
		'1 = on coast
		'2 = don't care
		'3 = don't care - non essential
		
		On Error GoTo Empire_Error
		Dim SectString As String
		Dim X As Short
		
		If Coast <= 2 Then
			SectString = "-+gom"
		Else
			SectString = "-+"
		End If
		
		For X = 1 To Len(SectString)
			If Len(SetOpenHex) = 0 Then
				rsSectors.MoveFirst()
				Do 
					If Not ((rsSectors.Fields("x").Value = 0 And rsSectors.Fields("y").Value = 0) Or (rsSectors.Fields("x").Value = 2 And rsSectors.Fields("y").Value = 0)) And rsSectors.Fields("des").Value = Mid(SectString, X, 1) Then
						
						If Initial Or (rsSectors.Fields("x").Value >= sx1 And rsSectors.Fields("x").Value <= sx2 And rsSectors.Fields("y").Value >= sy1 And rsSectors.Fields("y").Value <= sy2) Then
							
							Select Case Coast
								Case 0
									If rsSectors.Fields("coast").Value = 0 Then
										SetOpenHex = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
										rsSectors.Edit()
										rsSectors.Fields("des").Value = strDes
										rsSectors.Update()
										Exit Do
									End If
								Case 1
									If rsSectors.Fields("coast").Value = 1 Then
										SetOpenHex = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
										rsSectors.Edit()
										rsSectors.Fields("des").Value = strDes
										rsSectors.Update()
										Exit Do
									End If
								Case 2, 3
									SetOpenHex = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
									rsSectors.Edit()
									rsSectors.Fields("des").Value = strDes
									rsSectors.Update()
									Exit Do
							End Select
						End If
					End If
					
					rsSectors.MoveNext()
				Loop Until rsSectors.EOF
			End If
		Next X
		Exit Function
Empire_Error: 
		EmpireError("SetOpenHex", vbNullString, vbNullString)
		
	End Function
	Public Sub IslandDevelop(ByRef sx1 As Short, ByRef sy1 As Short, ByRef sx2 As Short, ByRef sy2 As Short)
		On Error GoTo Empire_Error
		
		Dim strSect As String
		' Dim Cap As String    8/2003 efj  removed
		' Dim Harbour As String    8/2003 efj  removed
		' Dim Oil As String    8/2003 efj  removed
		' Dim x As Integer    8/2003 efj  removed
		Dim mines As Short
		Dim hldIndex As String
		
		'set the index
		hldIndex = rsSectors.Index
		rsSectors.Index = "UpdateOrder"
		
		mines = 0
		' Harbour = vbnullstring    8/2003 efj  removed
		
		'Pass one - redesignate old mines to highways
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			
			If (rsSectors.Fields("x").Value >= sx1 And rsSectors.Fields("x").Value <= sx2 And rsSectors.Fields("y").Value >= sy1 And rsSectors.Fields("y").Value <= sy2) Then
				
				rsSectors.Edit()
				
				Select Case rsSectors.Fields("des").Value
					Case "g"
						If rsSectors.Fields("gold").Value = 0 Then
							rsSectors.Fields("des").Value = "+"
							strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
							frmEmpCmd.SubmitEmpireCommand("des " & strSect & " +", False)
						End If
					Case "o"
						If rsSectors.Fields("ocontent").Value < 5 Then
							rsSectors.Fields("des").Value = "+"
							strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
							frmEmpCmd.SubmitEmpireCommand("des " & strSect & " +", False)
						End If
					Case "m"
						If rsSectors.Fields("gold").Value > 0 Then
							rsSectors.Fields("des").Value = "g"
							strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
							frmEmpCmd.SubmitEmpireCommand("des " & strSect & " g", False)
							frmEmpCmd.SubmitEmpireCommand("thr d " & strSect & " 1", False)
						Else
							mines = mines + 1
						End If
				End Select
				
				rsSectors.Update()
			End If
			
			rsSectors.MoveNext()
		End While
		
		'Pass two - designate new oilwells and mines
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			
			If (rsSectors.Fields("x").Value >= sx1 And rsSectors.Fields("x").Value <= sx2 And rsSectors.Fields("y").Value >= sy1 And rsSectors.Fields("y").Value <= sy2) And rsSectors.Fields("des").Value = "+" Then
				
				rsSectors.Edit()
				
				If rsSectors.Fields("gold").Value > 0 Then
					rsSectors.Fields("des").Value = "g"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " g", False)
					frmEmpCmd.SubmitEmpireCommand("thr d " & strSect & " 1", False)
				ElseIf rsSectors.Fields("oil").Value > 5 Then 
					rsSectors.Fields("des").Value = "o"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " o", False)
					frmEmpCmd.SubmitEmpireCommand("thr o " & strSect & " 1", False)
				ElseIf rsSectors.Fields("min").Value > 80 And mines < 4 Then 
					mines = mines + 1
					rsSectors.Fields("des").Value = "m"
					strSect = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					frmEmpCmd.SubmitEmpireCommand("des " & strSect & " m", False)
					frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 1", False)
				End If
				
				rsSectors.Update()
			End If
			
			rsSectors.MoveNext()
		End While
		
		'build airport
		strSect = SetOpenHex("*", 0, False, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " *", False)
		frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 300", False)
		frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 150", False)
		frmEmpCmd.SubmitEmpireCommand("thr s " & strSect & " 100", False)
		frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 200", False)
		frmEmpCmd.SubmitEmpireCommand("thr p " & strSect & " 800", False)
		
		'build fortress
		strSect = SetOpenHex("f", 1, False, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " f", False)
		frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 200", False)
		frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 100", False)
		frmEmpCmd.SubmitEmpireCommand("thr s " & strSect & " 50", False)
		frmEmpCmd.SubmitEmpireCommand("thr g " & strSect & " 10", False)
		
		'build shell factory
		strSect = SetOpenHex("i", 2, False, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " i", False)
		frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 200", False)
		frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 100", False)
		frmEmpCmd.SubmitEmpireCommand("thr s " & strSect & " 1", False)
		
		'build defense plant
		strSect = SetOpenHex("d", 2, False, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " d", False)
		frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 100", False)
		frmEmpCmd.SubmitEmpireCommand("thr o " & strSect & " 20", False)
		frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 200", False)
		frmEmpCmd.SubmitEmpireCommand("thr g " & strSect & " 1", False)
		
		'headquarters
		strSect = SetOpenHex("!", 2, False, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " !", False)
		frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 300", False)
		frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 150", False)
		frmEmpCmd.SubmitEmpireCommand("thr s " & strSect & " 100", False)
		frmEmpCmd.SubmitEmpireCommand("thr g " & strSect & " 20", False)
		frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 200", False)
		
		'refinery
		strSect = SetOpenHex("%", 2, False, sx1, sy1, sx2, sy2)
		frmEmpCmd.SubmitEmpireCommand("des " & strSect & " %", False)
		frmEmpCmd.SubmitEmpireCommand("thr p " & strSect & " 1", False)
		frmEmpCmd.SubmitEmpireCommand("thr o " & strSect & " 20", False)
		
		'hcm
		strSect = SetOpenHex("k", 3, False, sx1, sy1, sx2, sy2)
		If Len(strSect) > 0 Then
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " k", False)
			frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 999", False)
			frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 1", False)
		End If
		
		'lcm factory
		If Len(strSect) > 0 Then
			strSect = SetOpenHex("j", 3, False, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " j", False)
			frmEmpCmd.SubmitEmpireCommand("thr i " & strSect & " 999", False)
			frmEmpCmd.SubmitEmpireCommand("thr l " & strSect & " 1", False)
		End If
		
		'build fortress
		If Len(strSect) > 0 Then
			strSect = SetOpenHex("f", 3, False, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " f", False)
			frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 200", False)
			frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 100", False)
			frmEmpCmd.SubmitEmpireCommand("thr s " & strSect & " 50", False)
			frmEmpCmd.SubmitEmpireCommand("thr g " & strSect & " 10", False)
		End If
		
		'enlistment center
		If Len(strSect) > 0 Then
			strSect = SetOpenHex("e", 3, False, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " e", False)
			frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 118", False)
		End If
		
		'build fortress
		If Len(strSect) > 0 Then
			strSect = SetOpenHex("f", 3, False, sx1, sy1, sx2, sy2)
			frmEmpCmd.SubmitEmpireCommand("des " & strSect & " f", False)
			frmEmpCmd.SubmitEmpireCommand("thr m " & strSect & " 200", False)
			frmEmpCmd.SubmitEmpireCommand("thr h " & strSect & " 100", False)
			frmEmpCmd.SubmitEmpireCommand("thr s " & strSect & " 50", False)
			frmEmpCmd.SubmitEmpireCommand("thr g " & strSect & " 10", False)
		End If
		
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors(True)
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		'restore Index
		rsSectors.Index = hldIndex
		Exit Sub
Empire_Error: 
		EmpireError("IslandDevelop", vbNullString, vbNullString)
		
	End Sub
	
	'101503 rjk: Modified to preserve information from different intelligence sources
	Public Sub AddEnemySectInfo(ByRef varInfo As Object)
		'Information is in the variant designation
		Dim n As Short
		Dim sx As Short
		Dim sy As Short
		
		'if you are not getting an array, then quit
		If Not IsArray(varInfo) Then
			Exit Sub
		End If
		
		On Error GoTo AddEnemySectInfo_Exit
		
		n = ES_X
		'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sx = varInfo(ES_X)
		n = ES_Y
		'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sy = varInfo(ES_Y)
		
		'Find out if we already have the record
		rsEnemySect.Seek("=", sx, sy)
		If rsEnemySect.NoMatch Then
			rsEnemySect.AddNew()
		Else
			rsEnemySect.Edit()
		End If
		
		'Setup the bmap record also
		rsBmap.Seek("=", sx, sy)
		If rsBmap.NoMatch Then
			rsBmap.AddNew()
			rsBmap.Fields("x").Value = sx
			rsBmap.Fields("y").Value = sy
		Else
			rsBmap.Edit()
		End If
		
		For n = 1 To UBound(varInfo)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varInfo(n) = "???" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varInfo(n) = "-1"
			End If
			
			'Put things in the right fields
			'101503 rjk: Modified to preserve information from different intelligence sources
			'101803 rjk: Removed updating EU_X and EU_Y as it should not change.
			Select Case n
				Case ES_DES
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If varInfo(n) <> "???" Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("des").Value = Left(varInfo(n), 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsBmap.Fields("des").Value = Left(varInfo(n), 1)
						rsBmap.Update()
					End If
				Case ES_X
					If rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("x").Value = CShort(varInfo(n))
					End If
				Case ES_Y
					If rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("y").Value = CShort(varInfo(n))
					End If
				Case ES_OWNER
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("owner").Value = CShort(varInfo(n))
					End If
				Case ES_OLDOWNER
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("oldowner").Value = CShort(varInfo(n))
					End If
				Case ES_EFFICIENCY
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("eff").Value = CShort(varInfo(n))
					End If
				Case ES_ROAD
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("road").Value = CShort(varInfo(n))
					End If
				Case ES_RAIL
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("rail").Value = CShort(varInfo(n))
					End If
				Case ES_DEFENSE
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("defense").Value = CShort(varInfo(n))
					End If
				Case ES_CIV
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("civ").Value = CShort(varInfo(n))
					End If
				Case ES_MIL
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("mil").Value = CShort(varInfo(n))
					End If
				Case ES_SHELL
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("shell").Value = CShort(varInfo(n))
					End If
				Case ES_GUN
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("gun").Value = CShort(varInfo(n))
					End If
				Case ES_PETROL
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("pet").Value = CShort(varInfo(n))
					End If
				Case ES_FOOD
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("food").Value = CShort(varInfo(n))
					End If
				Case ES_BARS
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("bar").Value = CShort(varInfo(n))
					End If
				Case ES_IRON
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemySect.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemySect.Fields("iron").Value = CShort(varInfo(n))
					End If
			End Select
		Next n
		rsEnemySect.Fields("timestamp").Value = GetWinACEUTC
		rsEnemySect.Update()
		
		
		Exit Sub
		
		'Error handling routine
AddEnemySectInfo_Exit: 
		Dim var As String
		Dim strLine As String
		var = CStr(n)
		For n = 1 To UBound(varInfo)
			'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = strLine & CStr(varInfo(n)) & ","
		Next n
		EmpireError("AddEnemySectInfo", var, strLine)
	End Sub
	
	'101503 rjk: Modified to preserve information from different intelligence sources
	'            Assumed that the ID and unit type is sufficient to identify record
	'            Do not use owner as it is always present.
	Public Sub AddEnemyUnitInfo(ByRef varInfo As Object)
		
		'Information is in the variant designation
		
		Dim n As Short
		Dim sowner As Short
		Dim sclass As String
		Dim sid As Short
		Dim swake As String
		Dim Index As Short
		
		'if you are not getting an array, then quit
		If Not IsArray(varInfo) Then
			Exit Sub
		End If
		
		On Error GoTo AddEnemyUnitInfo_Exit
		
		n = EU_OWNER
		'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sowner = CShort(varInfo(EU_OWNER))
		n = EU_ID
		'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sid = CShort(varInfo(EU_ID))
		n = EU_TYPE
		
		'get the class
		'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sclass = CStr(varInfo(0))
		
		'Find out if we already have the record
		'rsEnemyUnit.Index = "PrimaryKey"
		'rsEnemyUnit.Seek "=", sowner, sclass, sid
		'101503 rjk: Switch to only sclass and sid as owner is not always present
		rsEnemyUnit.Index = "ByID"
		rsEnemyUnit.Seek("=", sclass, sid)
		If rsEnemyUnit.NoMatch Then
			rsEnemyUnit.AddNew()
			'101503 rjk: move sclass and sid as they can not change, so only do once
			'record class
			rsEnemyUnit.Fields("class").Value = sclass
			rsEnemyUnit.Fields("id").Value = sid
			swake = vbNullString
			'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf Val(rsEnemyUnit.Fields("owner").Value) <> Val(varInfo(EU_OWNER)) Then 
			rsEnemyUnit.Delete() '260604 rjk: Added to deal with satellite changing owner.
			rsEnemyUnit.AddNew()
			rsEnemyUnit.Fields("class").Value = sclass
			rsEnemyUnit.Fields("id").Value = sid
			swake = vbNullString
		Else
			rsEnemyUnit.Edit()
			'for ships, keep a "wake" that shows where they have been
			swake = rsEnemyUnit.Fields("wake").Value
			If sclass = "s" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Val(varInfo(EU_X)) <> Val(rsEnemyUnit.Fields("x").Value) Or Val(varInfo(EU_Y)) <> Val(rsEnemyUnit.Fields("y").Value) Then
					If ((rsEnemyUnit.Fields("x").Value + rsEnemyUnit.Fields("y").Value) Mod 2) = 0 Then
						If Right(swake, 1) <> ";" Then 'clear wake from earlier versions
							swake = vbNullString
						End If
						swake = swake & SectorString(rsEnemyUnit.Fields("x").Value, rsEnemyUnit.Fields("y").Value) & ";"
						While Len(swake) > 50
							n = InStr(swake, ";")
							swake = Mid(swake, n + 1)
						End While
					End If
				End If
			End If
		End If
		
		'record class move to add step
		'rsEnemyUnit.Fields("class") = sclass
		
		For n = 1 To UBound(varInfo)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varInfo(n) = "???" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varInfo(n) = "-1"
			End If
			
			'Put things in the right fields
			'101503 rjk: do not erase something we already know.
			Select Case n
				'        Case EU_ID moved to add step
				'            rsEnemyUnit.Fields("id") = sid
				'101803 rjk: Removed the check on EU_X and EU_Y as should always be known.
				Case EU_X
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsEnemyUnit.Fields("x").Value = CShort(varInfo(n))
				Case EU_Y
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsEnemyUnit.Fields("y").Value = CShort(varInfo(n))
				Case EU_OWNER
					If sowner <> -1 Or rsEnemyUnit.NoMatch Then
						rsEnemyUnit.Fields("owner").Value = sowner
					End If
				Case EU_TYPE
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If varInfo(n) <> "???" Or rsEnemyUnit.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemyUnit.Fields("type").Value = varInfo(n)
					End If
				Case EU_MIL
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemyUnit.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemyUnit.Fields("mil").Value = CShort(varInfo(n))
					End If
				Case EU_TECH
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemyUnit.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemyUnit.Fields("tech").Value = CShort(varInfo(n))
					End If
				Case EU_EFF
					'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CShort(varInfo(n)) <> -1 Or rsEnemyUnit.NoMatch Then
						'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsEnemyUnit.Fields("eff").Value = CShort(varInfo(n))
					End If
			End Select
			
		Next n
		rsEnemyUnit.Fields("wake").Value = swake
		rsEnemyUnit.Fields("timestamp").Value = GetWinACEUTC
		
		rsEnemyUnit.Update()
		Exit Sub
		
		'Error handling routine
AddEnemyUnitInfo_Exit: 
		Dim var As String
		Dim strLine As String
		var = CStr(n)
		For n = 1 To UBound(varInfo)
			'UPGRADE_WARNING: Couldn't resolve default property of object varInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = strLine & CStr(varInfo(n)) & ","
		Next n
		
		EmpireError("AddEnemyUnitInfo", var, strLine)
		
	End Sub
	
	'112903 rjk: Remove the question on clearing the intelligence as there is a separate form now.
	'            Can now delete individual items.
	Public Sub ClearEnemyInfo(ByRef bSectors As Boolean, ByRef bShips As Boolean, ByRef bLands As Boolean, ByRef bPlanes As Boolean, ByRef bAirCombat As Boolean, ByRef bAllied As Boolean, ByRef bNeutral As Boolean, ByRef bEnemy As Boolean, ByRef iDays As Short)
		Dim rel As Short
		Dim dExpiry As Date
		Dim dRecord As Date
		
		dExpiry = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -iDays, Now)
		
		
		On Error GoTo Empire_Error
		'if there is no enemy information, then exit
		If (rsEnemyUnit.BOF And rsEnemyUnit.EOF) And (rsEnemySect.BOF And rsEnemySect.EOF) And (rsAirCombat.BOF And rsAirCombat.EOF) Then
			Exit Sub
		End If
		
		If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
			rsEnemyUnit.MoveFirst()
			While Not (rsEnemyUnit.EOF)
				dRecord = ConvertToLocalTimeFromWinACEUTC((rsEnemyUnit.Fields("timestamp").Value))
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If DateDiff(Microsoft.VisualBasic.DateInterval.Day, dExpiry, dRecord) <= 0 Then
					If (rsEnemyUnit.Fields("class").Value = "s" And bShips) Or (rsEnemyUnit.Fields("class").Value = "l" And bLands) Or (rsEnemyUnit.Fields("class").Value = "p" And bPlanes) Then
						rel = Nations.YourRelation(rsEnemyUnit.Fields("owner").Value)
						If ((rel = iREL_AT_WAR Or rel = iREL_HOSTILE Or rel = iREL_MOBILIZING Or rel = iREL_SITZKRIEG) And bEnemy) Or ((rel = iREL_ALLIED Or (rel = iREL_FRIENDLY And frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_ALLIED)) And bAllied) Or ((rel = iREL_NEUTRAL Or (rel = iREL_FRIENDLY And frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_NEUTRAL)) And bNeutral) Or rel = 0 Then
							rsEnemyUnit.Delete()
							rsEnemyUnit.MoveFirst()
						Else
							rsEnemyUnit.MoveNext()
						End If
					Else
						rsEnemyUnit.MoveNext()
					End If
				Else
					rsEnemyUnit.MoveNext()
				End If
			End While
		End If
		
		If bSectors Then
			DeleteAllRecordsOlderThen(rsEnemySect, dExpiry)
		End If
		
		If bAirCombat Then
			DeleteAllAirCombatRecordsOlderThen(dExpiry)
		End If
		
		Exit Sub
Empire_Error: 
		EmpireError("ClearEnemyInfo", vbNullString, vbNullString)
	End Sub
	
	Public Sub UpdateLights()
		On Error GoTo Empire_Error
		'set the green and red lights in the main window status bar
		If frmEmpCmd.CmdinProgress Or Not frmEmpCmd.ConnectedtoHost Then
			frmDrawMap.sb1.Items.Item("Lights").Image = VB6.ImageToIPictureDisp(picRedLight)
			frmDrawMap.sb1.Items.Item("Lights").ToolTipText = "Empire Server Unavailable"
			frmEmpCmd.imgLights.Image = picRedLight
		Else
			frmDrawMap.sb1.Items.Item("Lights").Image = VB6.ImageToIPictureDisp(picGreenLight)
			frmDrawMap.sb1.Items.Item("Lights").ToolTipText = "Empire Server Available"
			frmEmpCmd.imgLights.Image = picGreenLight
		End If
		Exit Sub
Empire_Error: 
		EmpireError("UpdateLights", vbNullString, vbNullString)
	End Sub
	
	Public Sub SetOwner(ByRef sx1 As Short, ByRef sy1 As Short, ByRef sx2 As Short, ByRef sy2 As Short, ByRef Nation As Short)
		On Error GoTo Empire_Error
		
		Dim hldIndex As String
		
		'set the index
		hldIndex = rsBmap.Index
		rsBmap.Index = "PrimaryKey"
		
		rsBmap.MoveFirst()
		While Not rsBmap.EOF
			
			If (rsBmap.Fields("x").Value >= sx1 And rsBmap.Fields("x").Value <= sx2 And rsBmap.Fields("y").Value >= sy1 And rsBmap.Fields("y").Value <= sy2) Then
				rsEnemySect.Seek("=", rsBmap.Fields("x"), rsBmap.Fields("y"))
				If rsBmap.Fields("des").Value <> "." And rsBmap.Fields("des").Value <> "X" Then
					If rsEnemySect.NoMatch Then
						rsEnemySect.AddNew()
						rsEnemySect.Fields("x").Value = rsBmap.Fields("x").Value
						rsEnemySect.Fields("y").Value = rsBmap.Fields("y").Value
						rsEnemySect.Fields("oldowner").Value = 0
						rsEnemySect.Fields("eff").Value = 0
						rsEnemySect.Fields("road").Value = 0
						rsEnemySect.Fields("rail").Value = 0
						rsEnemySect.Fields("defense").Value = 0
						rsEnemySect.Fields("civ").Value = 0
						rsEnemySect.Fields("mil").Value = 0
						rsEnemySect.Fields("shell").Value = 0
						rsEnemySect.Fields("gun").Value = 0
						rsEnemySect.Fields("pet").Value = 0
						rsEnemySect.Fields("food").Value = 0
						rsEnemySect.Fields("bar").Value = 0
						rsEnemySect.Fields("iron").Value = 0
						rsEnemySect.Fields("timestamp").Value = GetWinACEUTC '103003 rjk: Updated to make the timestamp format consistent
					Else
						rsEnemySect.Edit()
					End If
					rsEnemySect.Fields("owner").Value = Nation
					rsEnemySect.Fields("des").Value = rsBmap.Fields("des").Value
					
					rsEnemySect.Update()
				End If
			End If
			
			rsBmap.MoveNext()
		End While
		
		
		'restore Index
		rsBmap.Index = hldIndex
		frmDrawMap.DrawHexes()
		Exit Sub
Empire_Error: 
		EmpireError("SetOwner", vbNullString, vbNullString)
	End Sub
	
	'102203 rjk: Added as temporary, need to combine with CopyGameInfo
	'            still needs a lot of work.
	'102703 rjk: Modifiy the edit command to work with syntax supported next version of the server.
	Public Sub ExportNation()
		On Error GoTo Empire_Error
		
		Dim filenumber As Short
		Dim nLocations As Short
		Dim nSect As Short
		Dim nUnit As Short
		Dim FileName As String
		Dim strCmd As String
		Dim nID As Short
		Dim hldIndex As String
		
		If VerifySubDirectory("Exports", True) Then
			FileName = My.Application.Info.DirectoryPath & "\Exports\" & GameName & " NationScript.txt"
		Else
			FileName = My.Application.Info.DirectoryPath & "\" & GameName & " NationScript.txt"
		End If
		
		filenumber = FreeFile ' Get unused file number.
		FileOpen(filenumber, FileName, OpenMode.Output)
		
		PrintLine(filenumber, "edit country 0 T 1000")
		
		hldIndex = rsShip.Index
		rsShip.Index = "PrimaryKey"
		
		If Not (rsShip.BOF And rsShip.EOF) Then
			If VersionCheck(4, 3, 0) <> enumVersion.VER_LESS Then
				ComputeUnitCountsForShips()
			End If
			PrintLine(filenumber, "designate 0,0 h")
			PrintLine(filenumber, "setsector eff 0,0 100")
			'need to add index to ensure we start at the lowest number first
			rsShip.MoveFirst()
			nID = 0
			While Not rsShip.EOF
				If nID = rsShip.Fields("id").Value Then
					rsBuild.Seek("=", "s", rsShip.Fields("type"))
					If Not rsBuild.NoMatch Then
						'Need to supply the initial build supplies 20%
						PrintLine(filenumber, "give lcms 0,0 " & CStr(rsBuild.Fields("lcm").Value \ 5#))
						PrintLine(filenumber, "give hcms 0,0 " & CStr(rsBuild.Fields("hcm").Value \ 5#))
						PrintLine(filenumber, "setsector avail 0,0 " & CStr(rsBuild.Fields("avail").Value \ 5#))
						'May need money as well
						PrintLine(filenumber, "build ship 0,0 " + rsShip.Fields("type").Value + " 1 " + CStr(rsShip.Fields("tech").Value))
						strCmd = " U " & CStr(rsShip.Fields("id").Value)
						strCmd = strCmd & " O " & CStr(rsShip.Fields("country").Value)
						strCmd = strCmd & " L " & CStr(rsShip.Fields("x").Value) & "," & CStr(rsShip.Fields("y").Value)
						strCmd = strCmd & " E " & CStr(rsShip.Fields("eff").Value)
						strCmd = strCmd & " M " & CStr(rsShip.Fields("mob").Value)
						'strCmd= strCmd + " T " + CStr(rsShip.Fields("tech")) already in the build command
						strCmd = strCmd & " F " & CStr(rsShip.Fields("flt").Value)
						strCmd = strCmd & " H " & CStr(rsShip.Fields("he").Value)
						strCmd = strCmd & " P " & CStr(rsShip.Fields("pln").Value)
						strCmd = strCmd & " B " & CStr(rsShip.Fields("fuel").Value)
						strCmd = strCmd & " X " & CStr(rsShip.Fields("xl").Value)
						strCmd = strCmd & " Y " & CStr(rsShip.Fields("land").Value)
						'strCmd= strCmd + " R " + rsShip.Fields("retreat_path")
						'strCmd= strCmd + " W " + CStr(rsShip.Fields("retreat_flags"))
						'strCmd= strCmd + " a " + CStr(rsShip.Fields("plague_stage"))
						'strCmd= strCmd + " b " + rsShip.Fields("plague_time")
						'cargo
						strCmd = strCmd & " c " & CStr(rsShip.Fields("civ").Value)
						strCmd = strCmd & " m " & CStr(rsShip.Fields("mil").Value)
						strCmd = strCmd & " u " & CStr(rsShip.Fields("uw").Value)
						strCmd = strCmd & " f " & CStr(rsShip.Fields("food").Value)
						strCmd = strCmd & " s " & CStr(rsShip.Fields("shell").Value)
						strCmd = strCmd & " g " & CStr(rsShip.Fields("gun").Value)
						strCmd = strCmd & " p " & CStr(rsShip.Fields("petrol").Value)
						strCmd = strCmd & " i " & CStr(rsShip.Fields("iron").Value)
						strCmd = strCmd & " d " & CStr(rsShip.Fields("dust").Value)
						strCmd = strCmd & " o " & CStr(rsShip.Fields("oil").Value)
						'strCmd= strCmd + " ? " + CStr(rsShip.Fields("bar"))
						strCmd = strCmd & " l " & CStr(rsShip.Fields("lcm").Value)
						strCmd = strCmd & " h " & CStr(rsShip.Fields("hcm").Value)
						strCmd = strCmd & " r " & CStr(rsShip.Fields("rad").Value)
						PrintLine(filenumber, "edit ship 0" & strCmd)
						If Len(rsShip.Fields("name").Value) > 0 Then
							PrintLine(filenumber, "name " & CStr(rsShip.Fields("id").Value) & " " & Chr(34) + rsShip.Fields("name").Value + Chr(34))
						End If
						nUnit = nUnit + 1
						rsShip.MoveNext()
					Else
						MsgBox("Ship not found in the build information")
						Exit Sub
					End If
				Else
					rsBuild.MoveFirst()
					rsBuild.Seek("=", "s", "fb") 'needs to be more general - will not always been there
					PrintLine(filenumber, "give lcms 0,0 " & CStr(rsBuild.Fields("lcm").Value \ 5#))
					PrintLine(filenumber, "give hcms 0,0 " & CStr(rsBuild.Fields("hcm").Value \ 5#))
					PrintLine(filenumber, "setsector avail 0,0 " & CStr(rsBuild.Fields("avail").Value \ 5#))
					'May need money as well
					PrintLine(filenumber, "build ship 0,0 fb 1 ")
					PrintLine(filenumber, "edit ship 0 O 1")
				End If
				nID = nID + 1
			End While
		End If
		rsShip.Index = hldIndex
		
		hldIndex = rsPlane.Index
		rsPlane.Index = "PrimaryKey"
		
		If Not (rsPlane.BOF And rsPlane.EOF) Then
			PrintLine(filenumber, "designate 0,0 *")
			PrintLine(filenumber, "setsector eff 0,0 100")
			'need to add index to ensure we start at the lowest number first
			rsPlane.MoveFirst()
			nID = 0
			While Not rsPlane.EOF
				If nID = rsPlane.Fields("id").Value Then
					rsBuild.Seek("=", "p", rsPlane.Fields("type"))
					If Not rsBuild.NoMatch Then
						'Need to supply the initial build supplies 10%
						PrintLine(filenumber, "give lcms 0,0 " & CStr(rsBuild.Fields("lcm").Value \ 10#))
						PrintLine(filenumber, "give hcms 0,0 " & CStr(rsBuild.Fields("hcm").Value \ 10#))
						PrintLine(filenumber, "give mil 0,0 " & CStr(rsBuild.Fields("mil").Value \ 10#))
						PrintLine(filenumber, "setsector avail 0,0 " & CStr(rsBuild.Fields("avail").Value \ 10#))
						'May need money as well
						PrintLine(filenumber, "build plane 0,0 " + rsPlane.Fields("type").Value + " 1 " + CStr(rsPlane.Fields("tech").Value))
						strCmd = " U " & CStr(rsPlane.Fields("id").Value)
						strCmd = strCmd & " O " & CStr(rsPlane.Fields("country").Value)
						strCmd = strCmd & " L " & CStr(rsPlane.Fields("x").Value) & "," & CStr(rsPlane.Fields("y").Value)
						strCmd = strCmd & " e " & CStr(rsPlane.Fields("eff").Value)
						strCmd = strCmd & " M " & CStr(rsPlane.Fields("mob").Value)
						'strCmd = strCmd + " e " + CStr(rsPlane.Fields("tech")) alreay in the build command
						strCmd = strCmd & " F " & CStr(rsPlane.Fields("wing").Value)
						strCmd = strCmd & " B " & CStr(rsPlane.Fields("fuel").Value)
						PrintLine(filenumber, "edit plane 0" & strCmd)
						nUnit = nUnit + 1
						rsPlane.MoveNext()
					Else
						MsgBox("Plane not found in the build information")
						Exit Sub
					End If
				Else
					rsBuild.MoveFirst()
					rsBuild.Seek("=", "p", "f1") 'needs to be more general - will not always been there
					PrintLine(filenumber, "give lcms 0,0 " & CStr(rsBuild.Fields("lcm").Value \ 10#))
					PrintLine(filenumber, "give hcms 0,0 " & CStr(rsBuild.Fields("hcm").Value \ 10#))
					PrintLine(filenumber, "give mil 0,0 " & CStr(rsBuild.Fields("mil").Value \ 10#))
					PrintLine(filenumber, "setsector avail 0,0 " & CStr(rsBuild.Fields("avail").Value \ 10#))
					'May need money as well
					PrintLine(filenumber, "build plane 0,0 f1 1 ")
					PrintLine(filenumber, "edit plane 0 O 1")
				End If
				nID = nID + 1
			End While
		End If
		rsPlane.Index = hldIndex
		
		hldIndex = rsLand.Index
		rsLand.Index = "PrimaryKey"
		
		If Not (rsLand.BOF And rsLand.EOF) Then
			If VersionCheck(4, 3, 0) <> enumVersion.VER_LESS Then
				ComputeUnitCountsForLandUnits()
			End If
			PrintLine(filenumber, "designate 0,0 !")
			PrintLine(filenumber, "setsector eff 0,0 100")
			'need to add index to ensure we start at the lowest number first
			rsLand.MoveFirst()
			nID = 0
			While Not rsLand.EOF
				If nID = rsLand.Fields("id").Value Then
					rsBuild.Seek("=", "l", rsLand.Fields("type"))
					If Not rsBuild.NoMatch Then
						'Need to supply the initial build supplies 10%
						PrintLine(filenumber, "give lcms 0,0 " & CStr(rsBuild.Fields("lcm").Value \ 10#))
						PrintLine(filenumber, "give hcms 0,0 " & CStr(rsBuild.Fields("hcm").Value \ 10#))
						PrintLine(filenumber, "give mil 0,0 " & CStr(rsBuild.Fields("mil").Value \ 10#))
						PrintLine(filenumber, "setsector avail 0,0 " & CStr(rsBuild.Fields("avail").Value \ 10#))
						'May need money as well
						PrintLine(filenumber, "build land 0,0 " + rsLand.Fields("type").Value + " 1 " + CStr(rsLand.Fields("tech").Value))
						strCmd = " U " & CStr(rsLand.Fields("id").Value)
						strCmd = strCmd & " O " & CStr(rsLand.Fields("country").Value)
						strCmd = strCmd & " L " & CStr(rsLand.Fields("x").Value) & "," & CStr(rsLand.Fields("y").Value)
						strCmd = strCmd & " e " & CStr(rsLand.Fields("eff").Value)
						strCmd = strCmd & " M " & CStr(rsLand.Fields("mob").Value)
						'strCmd = strCmd + " t " + CStr(rsLand.Fields("tech")) alreay in the build command
						strCmd = strCmd & " a " & CStr(rsLand.Fields("army").Value)
						strCmd = strCmd & " B " & CStr(rsLand.Fields("fuel").Value)
						strCmd = strCmd & " X " & CStr(rsLand.Fields("xl").Value)
						strCmd = strCmd & " Y " & CStr(rsLand.Fields("land").Value)
						'strCmd = strCmd + " P " + rsLand.Fields("radius")
						'strCmd= strCmd + " Z " + CStr(rsLand.Fields("retreat_percentage"))
						'strCmd= strCmd + " R " + rsShip.Fields("retreat")
						'strCmd= strCmd + " W " + CStr(rsShip.Fields("retreat_flags"))
						'cargo
						'strCmd = strCmd + " c " + CStr(rsLand.Fields("civ"))
						strCmd = strCmd & " m " & CStr(rsLand.Fields("mil").Value)
						'strCmd = strCmd + " u " + CStr(rsLand.Fields("uw"))
						strCmd = strCmd & " f " & CStr(rsLand.Fields("food").Value)
						strCmd = strCmd & " s " & CStr(rsLand.Fields("shell").Value)
						strCmd = strCmd & " g " & CStr(rsLand.Fields("gun").Value)
						strCmd = strCmd & " p " & CStr(rsLand.Fields("petrol").Value)
						strCmd = strCmd & " i " & CStr(rsLand.Fields("iron").Value)
						strCmd = strCmd & " d " & CStr(rsLand.Fields("dust").Value)
						strCmd = strCmd & " o " & CStr(rsLand.Fields("oil").Value)
						'strCmd= strCmd + " ? " + CStr(rsLand.Fields("bar"))
						strCmd = strCmd & " l " & CStr(rsLand.Fields("lcm").Value)
						strCmd = strCmd & " h " & CStr(rsLand.Fields("hcm").Value)
						strCmd = strCmd & " r " & CStr(rsLand.Fields("rad").Value)
						PrintLine(filenumber, "edit land 0" & strCmd)
						nUnit = nUnit + 1
						rsLand.MoveNext()
					Else
						MsgBox("Land Unit not found in the build information")
						Exit Sub
					End If
				Else
					rsBuild.MoveFirst()
					rsBuild.Seek("=", "l", "cav") 'needs to be more general - will not always been there
					PrintLine(filenumber, "give lcms 0,0 " & CStr(rsBuild.Fields("lcm").Value \ 10#))
					PrintLine(filenumber, "give hcms 0,0 " & CStr(rsBuild.Fields("hcm").Value \ 10#))
					PrintLine(filenumber, "give mil 0,0 " & CStr(rsBuild.Fields("mil").Value \ 10#))
					PrintLine(filenumber, "setsector avail 0,0 " & CStr(rsBuild.Fields("avail").Value \ 10#))
					'May need money as well
					PrintLine(filenumber, "build land 0,0 cav 1 ")
					PrintLine(filenumber, "edit unit 0 O 1")
				End If
				nID = nID + 1
			End While
		End If
		rsLand.Index = hldIndex
		
		'need to add an index to control the order of reading
		If Not (rsBmap.BOF And rsBmap.EOF) Then
			rsBmap.MoveFirst()
			While Not rsBmap.EOF
				If Not IsNumeric(rsBmap.Fields("des").Value) Then 'need to deal sharebmap data
					nLocations = nLocations + 1
					'102803 rjk: Added bridge span and tower - needs more yet.
					If rsBmap.Fields("des").Value = "=" Then
						PrintLine(filenumber, "designate " & CStr(rsBmap.Fields("x").Value) & "," & CStr(rsBmap.Fields("y").Value) & " .")
						PrintLine(filenumber, "designate " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " #")
						PrintLine(filenumber, "give hcms " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " 100") 'needs to be read from the sector info.
						PrintLine(filenumber, "setsector avail " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " 40") 'needs to be read from the sector info.
						PrintLine(filenumber, "build b " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " g")
						'need to restore sector
					ElseIf rsBmap.Fields("des").Value = "@" Then 
						PrintLine(filenumber, "designate " & CStr(rsBmap.Fields("x").Value) & "," & CStr(rsBmap.Fields("y").Value) & " .")
						PrintLine(filenumber, "designate " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " #")
						PrintLine(filenumber, "give hcms " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " 400") 'needs to be read from the sector info.
						PrintLine(filenumber, "setsector avail " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " 160") 'needs to be read from the sector info.
						PrintLine(filenumber, "build t " & CStr(rsBmap.Fields("x").Value + 2) & "," & CStr(rsBmap.Fields("y").Value) & " g")
						'need to restore sector
					Else
						PrintLine(filenumber, "designate " & CStr(rsBmap.Fields("x").Value) & "," & CStr(rsBmap.Fields("y").Value) & " " + rsBmap.Fields("des").Value)
					End If
				Else 'set to a plain for now
					PrintLine(filenumber, "designate " & CStr(rsBmap.Fields("x").Value) & "," & CStr(rsBmap.Fields("y").Value) & " ~")
				End If
				rsBmap.MoveNext()
			End While
		End If
		
		If Not (rsSectors.BOF And rsSectors.EOF) Then
			rsSectors.MoveFirst()
			While Not rsSectors.EOF
				nSect = nSect + 1
				'SetResource Set What (iron, gold, oil, uranium, fertility)?  Set What (, , , , )?
				PrintLine(filenumber, "setresource iron " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("iron").Value))
				PrintLine(filenumber, "setresource gold " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("gold").Value))
				PrintLine(filenumber, "setresource oil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("ocontent").Value))
				PrintLine(filenumber, "setresource uranium " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("uran").Value))
				PrintLine(filenumber, "setresource fertility " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("fert").Value))
				
				PrintLine(filenumber, "edit land " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " R " & CStr(rsSectors.Fields("road").Value) & " r " & CStr(rsSectors.Fields("rail").Value) & " d " & CStr(rsSectors.Fields("defense").Value))
				
				'SetSector Give What (iron, gold, oil, uranium, fertility, owner, eff., Mob., Work, Avail., oldown, mines)?  , , , , , ., ., , ., , )?
				PrintLine(filenumber, "setsector owner " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("country").Value))
				'Print #filenumber, "setsector eff" + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -100"
				PrintLine(filenumber, "setsector eff " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("eff").Value))
				PrintLine(filenumber, "setsector mob " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("mob").Value))
				'Print #filenumber, "setsector work " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -100"
				PrintLine(filenumber, "setsector work " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("work").Value))
				'Print #filenumber, "setsector avail " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				PrintLine(filenumber, "setsector avail " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("avail").Value))
				'Print #filenumber, "setsector mines " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -65536 " 'Just a guess
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rsSectors.Fields("lmines").Value) Then
					PrintLine(filenumber, "setsector mines " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " 0")
				Else
					PrintLine(filenumber, "setsector mines " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("lmines").Value))
				End If
				
				'Extract from rsEnemySector table when "*" is set
				If rsSectors.Fields("*").Value = "*" Then
					'lookup sector in rsEnemySector table
					'Print #filenumber, "setsector " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " oldown " + )
				End If
				
				'Commodity
				'Print #filenumber, "give bar " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("bar").Value > 0 Then
					PrintLine(filenumber, "give bar " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("bar").Value))
				End If
				'Print #filenumber, "give civ " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("civ").Value > 0 Then
					PrintLine(filenumber, "give civ " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("civ").Value))
				End If
				'Print #filenumber, "give dust " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("dust").Value > 0 Then
					PrintLine(filenumber, "give dust " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("dust").Value))
				End If
				'Print #filenumber, "give food " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("food").Value > 0 Then
					PrintLine(filenumber, "give food " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("food").Value))
				End If
				'Print #filenumber, "give gun " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("gun").Value > 0 Then
					PrintLine(filenumber, "give gun " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("gun").Value))
				End If
				'Print #filenumber, "give hcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("hcm").Value > 0 Then
					PrintLine(filenumber, "give hcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("hcm").Value))
				End If
				'Print #filenumber, "give iron " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("iron").Value > 0 Then
					PrintLine(filenumber, "give iron " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("iron").Value))
				End If
				'Print #filenumber, "give lcm " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("lcm").Value > 0 Then
					PrintLine(filenumber, "give lcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("lcm").Value))
				End If
				'Print #filenumber, "give mil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("mil").Value > 0 Then
					PrintLine(filenumber, "give mil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("mil").Value))
				End If
				'Print #filenumber, "give oil " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("oil").Value > 0 Then
					PrintLine(filenumber, "give oil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("oil").Value))
				End If
				'Print #filenumber, "give pet " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("pet").Value > 0 Then
					PrintLine(filenumber, "give pet " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("pet").Value))
				End If
				'Print #filenumber, "give rad " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("rad").Value > 0 Then
					PrintLine(filenumber, "give rad " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("rad").Value))
				End If
				'Print #filenumber, "give shell " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("shell").Value > 0 Then
					PrintLine(filenumber, "give shell " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("shell").Value))
				End If
				'Print #filenumber, "give uran " + CStr(rsSectors.Fields("x")) + "," + CStr(rsSectors.Fields("y")) + " -9999"
				If rsSectors.Fields("uran").Value > 0 Then
					PrintLine(filenumber, "give uran " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("uran").Value))
				End If
				
				'Set Territory
				PrintLine(filenumber, "terr " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("terr").Value))
				PrintLine(filenumber, "terr " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("terr1").Value) & " 1")
				PrintLine(filenumber, "terr " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("terr2").Value) & " 2")
				PrintLine(filenumber, "terr " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("terr3").Value) & " 3")
				
				'Set Deliveries
				PrintLine(filenumber, "wipe " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value))
				If rsSectors.Fields("b_del").Value = "." Then
					PrintLine(filenumber, "deliver bar " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("b_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver bar " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("b_cut").Value) & " " + rsSectors.Fields("b_del").Value)
				End If
				If rsSectors.Fields("c_del").Value = "." Then
					PrintLine(filenumber, "deliver civ " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("c_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver civ " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("c_cut").Value) & " " + rsSectors.Fields("c_del").Value)
				End If
				If rsSectors.Fields("d_del").Value = "." Then
					PrintLine(filenumber, "deliver dust " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("d_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver dust " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("d_cut").Value) & " " + rsSectors.Fields("d_del").Value)
				End If
				If rsSectors.Fields("f_del").Value = "." Then
					PrintLine(filenumber, "deliver food " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("f_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver food " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("f_cut").Value) & " " + rsSectors.Fields("f_del").Value)
				End If
				If rsSectors.Fields("g_del").Value = "." Then
					PrintLine(filenumber, "deliver gun " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("g_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver gun " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("g_cut").Value) & " " + rsSectors.Fields("g_del").Value)
				End If
				If rsSectors.Fields("h_del").Value = "." Then
					PrintLine(filenumber, "deliver hcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("h_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver hcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("h_cut").Value) & " " + rsSectors.Fields("h_del").Value)
				End If
				If rsSectors.Fields("i_del").Value = "." Then
					PrintLine(filenumber, "deliver iron " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("i_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver iron " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("i_cut").Value) & " " + rsSectors.Fields("i_del").Value)
				End If
				If rsSectors.Fields("l_del").Value = "." Then
					PrintLine(filenumber, "deliver lcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("l_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver lcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("l_cut").Value) & " " + rsSectors.Fields("l_del").Value)
				End If
				If rsSectors.Fields("m_del").Value = "." Then
					PrintLine(filenumber, "deliver mil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("m_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver mil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("m_cut").Value) & " " + rsSectors.Fields("m_del").Value)
				End If
				If rsSectors.Fields("o_del").Value = "." Then
					PrintLine(filenumber, "deliver oil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("o_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver oil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("o_cut").Value) & " " + rsSectors.Fields("o_del").Value)
				End If
				If rsSectors.Fields("p_del").Value = "." Then
					PrintLine(filenumber, "deliver pet " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("p_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver pet " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("p_cut").Value) & " " + rsSectors.Fields("p_del").Value)
				End If
				If rsSectors.Fields("r_del").Value = "." Then
					PrintLine(filenumber, "deliver rad " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("r_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver rad " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("r_cut").Value) & " " + rsSectors.Fields("r_del").Value)
				End If
				If rsSectors.Fields("s_del").Value = "." Then
					PrintLine(filenumber, "deliver shell " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("s_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver shell " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("s_cut").Value) & " " + rsSectors.Fields("s_del").Value)
				End If
				If rsSectors.Fields("u_del").Value = "." Then
					PrintLine(filenumber, "deliver uran " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("u_cut").Value) & " h")
				Else
					PrintLine(filenumber, "deliver uran " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("u_cut").Value) & " " + rsSectors.Fields("u_del").Value)
				End If
				
				'Set Distributions
				PrintLine(filenumber, "distribute " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("dist_x").Value) & "," & CStr(rsSectors.Fields("dist_y").Value))
				PrintLine(filenumber, "threshold bar " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("b_dist").Value))
				PrintLine(filenumber, "threshold civ " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("c_dist").Value))
				PrintLine(filenumber, "threshold dust " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("d_dist").Value))
				PrintLine(filenumber, "threshold food " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("f_dist").Value))
				PrintLine(filenumber, "threshold gun " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("g_dist").Value))
				PrintLine(filenumber, "threshold hcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("h_dist").Value))
				PrintLine(filenumber, "threshold iron " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("i_dist").Value))
				PrintLine(filenumber, "threshold lcm " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("l_dist").Value))
				PrintLine(filenumber, "threshold mil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("m_dist").Value))
				PrintLine(filenumber, "threshold oil " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("o_dist").Value))
				PrintLine(filenumber, "threshold pet " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("p_dist").Value))
				PrintLine(filenumber, "threshold rad " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("r_dist").Value))
				PrintLine(filenumber, "threshold shell " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("s_dist").Value))
				PrintLine(filenumber, "threshold uran " & CStr(rsSectors.Fields("x").Value) & "," & CStr(rsSectors.Fields("y").Value) & " " & CStr(rsSectors.Fields("u_dist").Value))
				
				rsSectors.MoveNext()
			End While
		End If
		
		
		WriteLine(filenumber, "+++ " & CStr(nSect) & " Sectors exported")
		
		FileClose(filenumber)
		
		MsgBox(CStr(nLocations) & " Locations, " & CStr(nSect) & " Sectors, " & CStr(nUnit) & " Units exported to file " & FileName)
		Exit Sub
Empire_Error: 
		EmpireError("ExportNation", vbNullString, vbNullString)
		FileClose(filenumber)
	End Sub
	
	Public Sub ExportSeaInformation()
		Dim filenumber As Short
		Dim hldIndex As String
		Dim FileName As String
		Dim nLocations As Short
		
		filenumber = -1
		On Error GoTo Empire_Error
		
		If VerifySubDirectory("Exports", True) Then
			FileName = My.Application.Info.DirectoryPath & "\Exports\" & GameName & " Sea Information.txt"
		Else
			FileName = My.Application.Info.DirectoryPath & "\" & GameName & " Sea Information.txt"
		End If
		
		filenumber = FreeFile ' Get unused file number.
		FileOpen(filenumber, FileName, OpenMode.Output)
		
		hldIndex = rsSea.Index
		rsSea.Index = "PrimaryKey"
		rsSea.MoveFirst()
		While Not rsSea.EOF
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsSea.Fields("ftimestamp").Value) Then
				rsSea.Edit()
				rsSea.Fields("ftimestamp").Value = GetWinACEUTC
				rsSea.Update()
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsSea.Fields("otimestamp").Value) Then
				rsSea.Edit()
				rsSea.Fields("otimestamp").Value = GetWinACEUTC
				rsSea.Update()
			End If
			WriteLine(filenumber, CShort(rsSea.Fields("x").Value), CShort(rsSea.Fields("y").Value), CShort(rsSea.Fields("fert").Value), CDate(rsSea.Fields("ftimestamp").Value), CShort(rsSea.Fields("ocontent").Value), CDate(rsSea.Fields("otimestamp").Value))
			nLocations = nLocations + 1
			rsSea.MoveNext()
		End While
		
		FileClose(filenumber)
		MsgBox(CStr(nLocations) & " Locations exported to file " & FileName)
		rsSea.Index = hldIndex
		
		Exit Sub
		
Empire_Error: 
		EmpireError("ExportSeaInformation", vbNullString, vbNullString)
		rsSea.Index = hldIndex
		If filenumber > 0 Then
			FileClose(filenumber)
		End If
	End Sub
	
	Public Sub ImportIntelligence(ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		On Error GoTo Empire_Error
		
		Dim filenumber As Short
		Dim n As Short
		Dim nSect As Short
		Dim nsect2 As Short
		Dim nUnit As Short
		Dim nunit2 As Short
		Dim FileName As String
		Dim OnSectors As Boolean
		Dim X As Object
		' Dim Y As Variant    8/2003 efj  removed
		' Dim hldvar As Variant    8/2003 efj  removed
		Dim temp() As Object
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim count1 As Short
		Dim count2 As Short
		Dim date1 As Date
		Dim date2 As Date
		Dim strDate2 As String
		Dim hldIndex As String
		
		hldIndex = rsEnemyUnit.Index
		
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = True
		' Set flags
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.ShowReadOnly = False
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_NOTE: Variable frmCommonDialog.cdb was renamed frmCommonDialog.cdbOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="94ADAC4D-C65D-414F-A061-8FDC6B83C5EC"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckPathExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.CheckFileExists = True
		frmCommonDialog.cdbOpen.CheckPathExists = True
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = vbNullString
		If VerifySubDirectory("Exports", True) Then
			frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath & "\Exports"
		End If
		
		' Display the open dialog box
		frmCommonDialog.cdbOpen.ShowDialog()
		
		FileName = frmCommonDialog.cdbOpen.FileName
		filenumber = FreeFile ' Get unused file number.
		FileOpen(filenumber, FileName, OpenMode.Input)
		
		rsEnemyUnit.Index = "PrimaryKey"
		OnSectors = True
		rs = rsEnemySect
		count1 = 0
		count2 = 0
		'first input the sector file
		While Not EOF(filenumber)
			Input(filenumber, X)
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Left(X, 3) = "+++" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If OnSectors And (InStr(X, "Sectors exported") <> 0) Then '111803 rjk: Ignore headers
					OnSectors = False
					rs = rsEnemyUnit
					nSect = count1
					nsect2 = count2
					count1 = 0
					count2 = 0
				End If
			Else
				ReDim temp(rs.Fields.Count - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object temp(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				temp(0) = X
				For n = 1 To rs.Fields.Count - 1
					Input(filenumber, temp(n))
					If OnSectors And n = 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						temp(n) = CStr(CShort(temp(0)) + iOffsetX)
					ElseIf OnSectors And n = 1 Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						temp(n) = CStr(CShort(temp(1)) + iOffsetY)
					ElseIf Not OnSectors And n = 2 Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						temp(n) = CStr(CShort(temp(2)) + iOffsetX)
					ElseIf Not OnSectors And n = 3 Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						temp(n) = CStr(CShort(temp(3)) + iOffsetY)
					End If
				Next n
				count1 = count1 + 1
				'102903 rjk: Do not add records for ourselves
				'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (CShort(temp(3)) <> CountryNumber And OnSectors) Or (CShort(temp(4)) <> CountryNumber And Not OnSectors) Then
					If OnSectors Then
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rs.Seek("=", CShort(temp(0)), CShort(temp(1)))
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rs.Seek("=", CShort(temp(4)), CStr(temp(9)), CShort(temp(0)))
					End If
					If rs.NoMatch Then
						rs.AddNew()
						For n = 0 To rs.Fields.Count - 1
							'UPGRADE_WARNING: Couldn't resolve default property of object temp(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							rs.Fields(n).Value = temp(n)
						Next n
						rs.Update()
						count2 = count2 + 1
					Else
						
						'decide which is more recent
						date1 = ConvertToLocalTimeFromWinACEUTC((rs.Fields("timestamp").Value))
						If OnSectors Then
							'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strDate2 = CStr(temp(17))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object temp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strDate2 = CStr(temp(8))
						End If
						If Mid(strDate2, 5, 1) <> "/" Then
							strDate2 = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, tz.Bias, EmpireTimeString(strDate2)), "yyyy/mm/dd hh:mm:ss")
							If OnSectors Then
								'UPGRADE_WARNING: Couldn't resolve default property of object temp(17). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								temp(17) = strDate2
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object temp(8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								temp(8) = strDate2
							End If
						End If
						date2 = ConvertToLocalTimeFromWinACEUTC(strDate2)
						If date1 < date2 Then
							
							'if export is more recent, use it
							rs.Edit()
							For n = 0 To rs.Fields.Count - 1
								'UPGRADE_WARNING: Couldn't resolve default property of object temp(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								rs.Fields(n).Value = temp(n)
							Next n
							
							rs.Update()
							count2 = count2 + 1
						End If
					End If
				End If
			End If
		End While
		nUnit = count1
		nunit2 = count2
		
		FileClose(filenumber)
		
		MsgBox(CStr(nSect) & " Sectors read, " & CStr(nsect2) & " Sectors updated, " & CStr(nUnit) & " Units read, " & CStr(nunit2) & " Units updated.")
		
		rsEnemyUnit.Index = hldIndex
		
		Exit Sub
Empire_Error: 
		EmpireError("ImportIntelligence", vbNullString, vbNullString)
		rsEnemyUnit.Index = hldIndex
		
	End Sub
	
	Public Sub ImportIntelligenceTelegram(ByRef strTelegram As String, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		On Error GoTo Empire_Error
		
		Dim filenumber As Short
		Dim n As Short
		Dim nSect As Short
		Dim nsect2 As Short
		Dim nUnit As Short
		Dim nunit2 As Short
		Dim FileName As String
		Dim OnSectors As Boolean
		Dim temp() As String
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim count1 As Short
		Dim count2 As Short
		Dim date1 As Date
		Dim date2 As Date
		Dim strDate2 As String
		Dim hldIndex As String
		Dim pos As Integer '110403 rjk: Changed to a Long to accomodate large Intelligence Reports.
		Dim strLine As String
		
		hldIndex = rsEnemyUnit.Index
		
		rsEnemyUnit.Index = "PrimaryKey"
		OnSectors = True
		rs = rsEnemySect
		count1 = 0
		count2 = 0
		'first input the sector file
		pos = 1
		While InStr(pos, strTelegram, vbLf) > 0
			strLine = Mid(strTelegram, pos, InStr(pos, strTelegram, vbLf) - pos)
			pos = InStr(pos, strTelegram, vbLf) + 1
			If InStr(strLine, "+++") > 0 Then
				If OnSectors And InStr(strLine, "Sectors exported") <> 0 Then '111803 rjk: Changed processing to ignore headers
					OnSectors = False
					rs = rsEnemyUnit
					nSect = count1
					nsect2 = count2
					count1 = 0
					count2 = 0
				End If
			Else
				temp = Split(strLine, ",", rs.Fields.Count)
				For n = 0 To rs.Fields.Count - 1
					temp(n) = Mid(temp(n), 2, Len(temp(n)) - 2)
					If OnSectors And n = 0 Then
						temp(n) = CStr(CShort(temp(0)) + iOffsetX)
					ElseIf OnSectors And n = 1 Then 
						temp(n) = CStr(CShort(temp(1)) + iOffsetY)
					ElseIf Not OnSectors And n = 2 Then 
						temp(n) = CStr(CShort(temp(2)) + iOffsetX)
					ElseIf Not OnSectors And n = 3 Then 
						temp(n) = CStr(CShort(temp(3)) + iOffsetY)
					End If
				Next n
				count1 = count1 + 1
				'102903 rjk: Do not add records for ourselves
				If (CShort(temp(3)) <> CountryNumber And OnSectors) Or (CShort(temp(4)) <> CountryNumber And Not OnSectors) Then
					If OnSectors Then
						rs.Seek("=", CShort(temp(0)), CShort(temp(1)))
					Else
						rs.Seek("=", CShort(temp(4)), CStr(temp(9)), CShort(temp(0)))
					End If
					If rs.NoMatch Then
						rs.AddNew()
						For n = 0 To rs.Fields.Count - 1
							rs.Fields(n).Value = temp(n)
						Next n
						rs.Update()
						count2 = count2 + 1
					Else
						'decide which is more recent
						date1 = ConvertToLocalTimeFromWinACEUTC((rs.Fields("timestamp").Value))
						If OnSectors Then
							strDate2 = CStr(temp(17))
						Else
							strDate2 = CStr(temp(8))
						End If
						If Mid(strDate2, 5, 1) <> "/" Then
							strDate2 = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, tz.Bias, EmpireTimeString(strDate2)), "yyyy/mm/dd hh:mm:ss")
							If OnSectors Then
								temp(17) = strDate2
							Else
								temp(8) = strDate2
							End If
						End If
						date2 = ConvertToLocalTimeFromWinACEUTC(strDate2)
						If date1 < date2 Then
							
							'if export is more recent, use it
							rs.Edit()
							For n = 0 To rs.Fields.Count - 1
								rs.Fields(n).Value = temp(n)
							Next n
							
							rs.Update()
							count2 = count2 + 1
						End If
					End If
				End If
			End If
		End While
		nUnit = count1
		nunit2 = count2
		
		MsgBox(CStr(nSect) & " Sectors read, " & CStr(nsect2) & " Sectors updated, " & CStr(nUnit) & " Units read, " & CStr(nunit2) & " Units updated.")
		
		rsEnemyUnit.Index = hldIndex
		
		Exit Sub
Empire_Error: 
		EmpireError("ImportIntelligenceTelegram", vbNullString, vbNullString)
		rsEnemyUnit.Index = hldIndex
	End Sub
	
	Public Sub ImportIntelligenceOffset(ByRef strTelegram As String)
		If frmDrawMap.PromptUp Then
			Exit Sub
		End If
		
		frmPromptImportOffset.strTelegram = strTelegram
		
		frmPromptImportOffset.txtOrigin.Text = SectorString((frmDrawMap.SelX), (frmDrawMap.SelY))
		
		'Put form in proper place
		frmPromptImportOffset.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptImportOffset.Width)) \ 2)
		frmPromptImportOffset.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptImportOffset.Height))
		frmDrawMap.PromptForm = frmPromptImportOffset
		frmDrawMap.PromptUp = True
		
		'Start the map as the backgroud
		frmDrawMap.Activate()
		'Dialog box will take it from here
		frmPromptImportOffset.Show()
	End Sub
	
	Public Sub ImportSeaInformation()
		Dim strFileName As String
		Dim iFileNumber As Short
		Dim iUpdatedLocations, iAddedLocations As Short
		Dim iLocX, iLocY As Short
		Dim iFert, iOcontent As Short
		Dim hldIndex As String
		Dim dFertTimestamp, dOcontentTimestamp As Date
		
		hldIndex = rsSea.Index
		rsSea.Index = "PrimaryKey"
		
		iFileNumber = -1
		On Error GoTo Empire_Error
		
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = True
		' Set flags
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.ShowReadOnly = False
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_NOTE: Variable frmCommonDialog.cdb was renamed frmCommonDialog.cdbOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="94ADAC4D-C65D-414F-A061-8FDC6B83C5EC"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckPathExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.CheckFileExists = True
		frmCommonDialog.cdbOpen.CheckPathExists = True
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = vbNullString
		If VerifySubDirectory("Exports", True) Then
			frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath & "\Exports"
		End If
		
		' Display the open dialog box
		frmCommonDialog.cdbOpen.ShowDialog()
		
		strFileName = frmCommonDialog.cdbOpen.FileName
		
		iFileNumber = FreeFile ' Get unused file number.
		FileOpen(iFileNumber, strFileName, OpenMode.Input)
		
		While Not EOF(iFileNumber)
			Input(iFileNumber, iLocX)
			Input(iFileNumber, iLocY)
			Input(iFileNumber, iFert)
			Input(iFileNumber, dFertTimestamp)
			Input(iFileNumber, iOcontent)
			Input(iFileNumber, dOcontentTimestamp)
			rsSea.Seek("=", iLocX, iLocY)
			If rsSea.NoMatch Then
				rsSea.AddNew()
				rsSea.Fields("x").Value = iLocX
				rsSea.Fields("y").Value = iLocY
				rsSea.Fields("fert").Value = iFert
				rsSea.Fields("ftimestamp").Value = dFertTimestamp
				rsSea.Fields("ocontent").Value = iOcontent
				rsSea.Fields("otimestamp").Value = dOcontentTimestamp
				iAddedLocations = iAddedLocations + 1
				rsSea.Update()
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
			ElseIf DateDiff(Microsoft.VisualBasic.DateInterval.Second, rsSea.Fields("ftimestamp").Value, dFertTimestamp) > 0 Or DateDiff(Microsoft.VisualBasic.DateInterval.Second, rsSea.Fields("otimestamp").Value, dOcontentTimestamp) > 0 Then 
				rsSea.Edit()
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If DateDiff(Microsoft.VisualBasic.DateInterval.Second, rsSea.Fields("ftimestamp").Value, dFertTimestamp) > 0 Then
					rsSea.Fields("fert").Value = iFert
					rsSea.Fields("ftimestamp").Value = dFertTimestamp
				End If
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If DateDiff(Microsoft.VisualBasic.DateInterval.Second, rsSea.Fields("otimestamp").Value, dOcontentTimestamp) > 0 Then
					rsSea.Fields("ocontent").Value = iOcontent
					rsSea.Fields("otimestamp").Value = dOcontentTimestamp
				End If
				rsSea.Update()
				iUpdatedLocations = iUpdatedLocations + 1
			End If
		End While
		FileClose(iFileNumber)
		MsgBox(CStr(iUpdatedLocations) & " locations updated and " & CStr(iAddedLocations) & " locations imported from file " & strFileName)
		rsSea.Index = hldIndex
		
		Exit Sub
		
Empire_Error: 
		EmpireError("ImportSeaInformation", vbNullString, vbNullString)
		If iFileNumber > 0 Then
			FileClose(iFileNumber)
		End If
		rsSea.Index = hldIndex
	End Sub
	
	Public Sub OffsetSectorLocation(ByRef sx As Short, ByRef sy As Short, ByRef offx As Short, ByRef offy As Short)
		On Error GoTo Empire_Error
		
		Dim MaxX As Short
		Dim MaxY As Short
		
		sx = sx + offx
		sy = sy + offy
		
		'Adjust x for world wrap
		MaxX = rsVersion.Fields("World X").Value / 2
		MaxY = rsVersion.Fields("World Y").Value / 2
		
		'090103 rjk: Changed wrap at MaxX instead MaxX + 1 as -MaxX to +MaxX-1 is the valid range
		While sx >= MaxX
			sx = sx - rsVersion.Fields("World X").Value
		End While
		
		While sx < (-1 * MaxX)
			sx = sx + rsVersion.Fields("World X").Value
		End While
		
		'090103 rjk: Changed wrap at MaxY instead MaxY + 1 as -MaxY to +MaxY-1 is the valid range
		While sy >= MaxY
			sy = sy - rsVersion.Fields("World Y").Value
		End While
		
		While sy < (-1 * MaxY)
			sy = sy + rsVersion.Fields("World Y").Value
		End While
		Exit Sub
Empire_Error: 
		EmpireError("OffsetSectorLocation", vbNullString, vbNullString)
	End Sub
	
	'090103 rjk: Added Function for exploring in wheel fashion
	Private Function StraightPath(ByRef strPath As String) As Boolean
		Dim strFirstDirection As String
		Dim Counter As Short
		
		If Len(strPath) = 0 Then
			StraightPath = False
			Exit Function
		End If
		If Len(strPath) = 1 Then
			StraightPath = True
			Exit Function
		End If
		
		strFirstDirection = Left(strPath, 1)
		
		For Counter = 1 To Len(strPath)
			If strFirstDirection <> Mid(strPath, Counter, 1) Then
				StraightPath = False
				Exit Function
			End If
		Next Counter
		StraightPath = True
	End Function
	
	'090103 rjk: Added Function for exploring along the border
	Private Function AdjacentToBorder(ByRef X As Short, ByRef Y As Short) As Boolean
		Dim adjacentX, adjacentY As Short
		
		adjacentX = X - 2
		adjacentY = Y
		
		OffsetSectorLocation(adjacentX, adjacentY, 0, 0)
		rsBmap.Seek("=", adjacentX, adjacentY)
		If Not rsBmap.NoMatch Then
			'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
			If InStr(".?\/^s", rsBmap.Fields("des").Value) > 0 Then
				AdjacentToBorder = True
				Exit Function
			End If
		End If
		
		adjacentX = X + 2
		adjacentY = Y
		
		OffsetSectorLocation(adjacentX, adjacentY, 0, 0)
		rsBmap.Seek("=", adjacentX, adjacentY)
		If Not rsBmap.NoMatch Then
			'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
			If InStr(".?\/^s", rsBmap.Fields("des").Value) > 0 Then
				AdjacentToBorder = True
				Exit Function
			End If
		End If
		adjacentX = X - 1
		adjacentY = Y - 1
		
		OffsetSectorLocation(adjacentX, adjacentY, 0, 0)
		rsBmap.Seek("=", adjacentX, adjacentY)
		If Not rsBmap.NoMatch Then
			'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
			If InStr(".?\/^s", rsBmap.Fields("des").Value) > 0 Then
				AdjacentToBorder = True
				Exit Function
			End If
		End If
		adjacentX = X - 1
		adjacentY = Y + 1
		
		OffsetSectorLocation(adjacentX, adjacentY, 0, 0)
		rsBmap.Seek("=", adjacentX, adjacentY)
		If Not rsBmap.NoMatch Then
			'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
			If InStr(".?\/^s", rsBmap.Fields("des").Value) > 0 Then
				AdjacentToBorder = True
				Exit Function
			End If
		End If
		adjacentX = X + 1
		adjacentY = Y - 1
		
		OffsetSectorLocation(adjacentX, adjacentY, 0, 0)
		rsBmap.Seek("=", adjacentX, adjacentY)
		If Not rsBmap.NoMatch Then
			'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
			If InStr(".?\/^s", rsBmap.Fields("des").Value) > 0 Then
				AdjacentToBorder = True
				Exit Function
			End If
		End If
		adjacentX = X + 1
		adjacentY = Y + 1
		
		OffsetSectorLocation(adjacentX, adjacentY, 0, 0)
		rsBmap.Seek("=", adjacentX, adjacentY)
		If Not rsBmap.NoMatch Then
			'looking for edge (water, unknown, wasteland, mountains, or sanctuaries)
			If InStr(".?\/^s", rsBmap.Fields("des").Value) > 0 Then
				AdjacentToBorder = True
				Exit Function
			End If
		End If
		
		AdjacentToBorder = False
	End Function
	
	Public Sub ExploreOut(ByRef strSect As String, ByRef strUnit As String, ByRef radius As Short, ByRef Occur As Short, ByRef OldSectorCount As Short, ByVal numberToEachSector As Short, ByRef exploreMode As enumExploreMode)
		
		Dim range(MaxExploreRadius * 4 + 5, MaxExploreRadius * 2 + 3) As Short
		Dim paths(MaxExploreRadius * 4 + 5, MaxExploreRadius * 2 + 3) As String
		Dim p1 As Short
		Dim StartX As Short
		Dim StartY As Short
		Dim X As Short
		Dim Y As Short
		Dim n As Short
		Dim offy As Short
		Dim offx As Short
		Dim holdy As Short
		Dim holdx As Short
		Dim strCmd As String
		Dim Done As Boolean
		Dim found As Boolean
		Dim newsector As Boolean
		Dim unknown As Boolean
		Dim UnitsAvail As Short
		Dim NewSectorCount As Short
		
		Const OWNED_SECTOR As Short = 2
		Const CHECK_SECTOR As Short = 1
		Const EXPLORATION_SECTOR As Short = 3
		Const UNPASSABLE_SECTOR As Short = 4
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
		
		rsSectors.Seek("=", StartX, StartY)
		If rsSectors.NoMatch Then
			Exit Sub
		Else
			If strUnit = "c" Then
				If rsSectors.Fields("mil").Value > 0 Then
					UnitsAvail = rsSectors.Fields("civ").Value
				Else
					UnitsAvail = rsSectors.Fields("civ").Value - 1
				End If
			Else
				If rsSectors.Fields("civ").Value > 0 And Not (rsSectors.Fields("*").Value = "*") Then
					UnitsAvail = rsSectors.Fields("mil").Value
				Else
					UnitsAvail = rsSectors.Fields("mil").Value - 1
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
						OffsetSectorLocation(holdx, holdy, 0, 0)
						rsSectors.Seek("=", holdx, holdy)
						If rsSectors.NoMatch Then
							rsBmap.Seek("=", holdx, holdy)
							If rsBmap.NoMatch Then
								range(X, Y) = UNPASSABLE_SECTOR
								unknown = True
							Else
								' don't go into water, unknown, wasteland, mountains, or sanctuaries
								If InStr(".X?\/^s", rsBmap.Fields("des").Value) > 0 Then
									range(X, Y) = UNPASSABLE_SECTOR
								Else
									'090103 rjk: Added additional exploration modes
									Select Case exploreMode
										Case enumExploreMode.EXP_COMPLETE
											range(X, Y) = EXPLORATION_SECTOR
										Case enumExploreMode.EXP_WHEEL
											If StraightPath(paths(X, Y)) Then
												range(X, Y) = EXPLORATION_SECTOR
											End If
											If AdjacentToBorder(holdx, holdy) Then
												range(X, Y) = EXPLORATION_SECTOR
											End If
										Case enumExploreMode.EXP_BORDER
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
								paths(X - 2, Y) = paths(X, Y) & "g"
								range(X - 2, Y) = CHECK_SECTOR
								Done = False
							End If
							If range(X - 1, Y - 1) = 0 Then
								paths(X - 1, Y - 1) = paths(X, Y) & "y"
								range(X - 1, Y - 1) = CHECK_SECTOR
								Done = False
							End If
							If range(X + 1, Y - 1) = 0 Then
								paths(X + 1, Y - 1) = paths(X, Y) & "u"
								range(X + 1, Y - 1) = CHECK_SECTOR
								Done = False
							End If
							If range(X + 2, Y) = 0 Then
								paths(X + 2, Y) = paths(X, Y) & "j"
								range(X + 2, Y) = CHECK_SECTOR
								Done = False
							End If
							If range(X + 1, Y + 1) = 0 Then
								paths(X + 1, Y + 1) = paths(X, Y) & "n"
								range(X + 1, Y + 1) = CHECK_SECTOR
								Done = False
							End If
							If range(X - 1, Y + 1) = 0 Then
								paths(X - 1, Y + 1) = paths(X, Y) & "b"
								range(X - 1, Y + 1) = CHECK_SECTOR
								Done = False
							End If
						End If
					End If
				Next p1
			Next Y
		End While
		
		'Now build all the exploration command to the explorable sectors
		strCmd = "explore " & strUnit & " " & strSect & " " & Str(numberToEachSector) & " "
		
		newsector = False
		frmEmpCmd.SubmitEmpireCommand("ex1", False)
		For n = 1 To radius
			found = False
			For Y = 0 To UBound(range, 2)
				For p1 = 0 To UBound(range, 1) - 1 Step 2
					X = p1 + (Y Mod 2)
					'090103 rjk - Ensure there is enough units
					If range(X, Y) = EXPLORATION_SECTOR And Len(paths(X, Y)) = n And UnitsAvail >= numberToEachSector Then
						frmEmpCmd.SubmitEmpireCommand(strCmd & paths(X, Y) & "h", True)
						'090103 rjk - need to subtract the number moved not just one.
						UnitsAvail = UnitsAvail - numberToEachSector
						found = True
					End If
				Next p1
			Next Y
			If found Then
				frmEmpCmd.SubmitEmpireCommand("des * ?des=- +", True)
				newsector = True
			End If
		Next n
		
		If unknown And newsector And Occur < MaxExploreRadius Then
			frmEmpCmd.SubmitEmpireCommand("ex2", False)
		Else
			frmEmpCmd.SubmitEmpireCommand("ex4", False)
		End If
		
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'090703 rjk: Added an equal to missed any events
		frmEmpCmd.SubmitEmpireCommand("dump * ?timestamp=" & CStr(tsSectors), False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("map *", False)
		frmEmpCmd.SubmitEmpireCommand("bmap *", False)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		If unknown And newsector And Occur < MaxExploreRadius Then
			'the ex3 command will cause the explore out to continue to loop
			'090103 rjk: added the explore mode for the recursion call
			frmEmpCmd.SubmitEmpireCommand("ex3  " & strSect & ":" & strUnit & ":" & CStr(radius) & ":" & CStr(Occur + 1) & ":" & CStr(NewSectorCount) & ":" & Str(numberToEachSector) & ":" & Str(exploreMode) & ":", False)
		End If
		
	End Sub
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function PadString(ByRef str_Renamed As String, ByRef slen As Short, Optional ByRef InFront As Boolean = True) As Object
		On Error GoTo Empire_Error
		Dim strSpaces As String
		Dim strPad As String
		On Error Resume Next
		strSpaces = "              "
		While Len(strSpaces) < slen
			strSpaces = strSpaces & strSpaces
		End While
		
		If Len(str_Renamed) < slen Then
			strPad = Left(strSpaces, slen - Len(str_Renamed))
			If InFront Then
				'UPGRADE_WARNING: Couldn't resolve default property of object PadString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PadString = strPad & str_Renamed
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object PadString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PadString = str_Renamed & strPad
			End If
		Else
			PadString = Left(str_Renamed, slen)
		End If
		Exit Function
Empire_Error: 
		EmpireError("PadString", vbNullString, str_Renamed)
	End Function
	
	Public Sub ProdSummaryReport(ByRef sx1 As Short, ByRef sy1 As Short, ByRef ShowStarve As Boolean, ByRef ShowOvrPop As Boolean, ByRef ShowIdle As Boolean, ByRef ShowBuild As Boolean)
		On Error GoTo Empire_Error
		
		Const PR_ITEM As Short = 1
		Const PR_ONHAND As Short = 2
		Const PR_MADE As Short = 3
		Const PR_USED As Short = 4
		Const PR_RQST As Short = 5
		Const PR_CATG_LAST As Short = 6
		
		Const PR_FOOD As Short = 1
		Const PR_IRON As Short = 2
		Const PR_LCM As Short = 3
		Const PR_HCM As Short = 4
		Const PR_DUST As Short = 5
		Const PR_BAR As Short = 6
		Const PR_GUN As Short = 7
		Const PR_SHELL As Short = 8
		Const PR_OIL As Short = 9
		Const PR_PET As Short = 10
		Const PR_MIL As Short = 11
		Const PR_RAD As Short = 12
		Const PR_CIV As Short = 13
		Const PR_UW As Short = 14
		Const PR_HAPPY As Short = 15
		Const PR_GRAD As Short = 16
		Const PR_TECH As Short = 17
		Const PR_RESCH As Short = 18
		Const PR_SHIP As Short = 19
		Const PR_PLANE As Short = 20
		Const PR_ARMY As Short = 21
		Const PR_CITY As Short = 22
		Const PR_FORT As Short = 23
		Const PR_COMM_LAST As Short = 23
		Const PR_COMM_ITEM As Short = 14
		Const PR_MSGC As Short = 10
		
		'UPGRADE_WARNING: Lower bound of array ar was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim ar(PR_COMM_LAST, PR_CATG_LAST) As Object
		'UPGRADE_WARNING: Lower bound of array sc was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim sc(PR_COMM_ITEM, PR_CATG_LAST) As Integer
		'UPGRADE_WARNING: Lower bound of array ut was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim ut(PR_COMM_ITEM, PR_COMM_LAST) As Integer
		'UPGRADE_WARNING: Lower bound of array cp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		'UPGRADE_WARNING: Arrays can't be declared with New. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC9D3AE5-6B95-4B43-91C7-28276302A5E8"'
		Dim cp(PR_MSGC) As New Collection
		Dim n As Integer
		Dim n2 As Integer
		Dim strLine As String
		Dim varLine As Object
		Dim v As productionDataType
		Dim strSect As String
		Dim strSectHeader As String
		Dim strCapt As String
		Dim MaxEffGain As Short
		Dim EffGain As Single
		Dim NewEff As Short
		Dim BuildType As String
		Dim Target As Short
		Dim LastItem As Short
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		
		'UPGRADE_WARNING: Lower bound of array Costs was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim Costs(7, 5) As Single
		'UPGRADE_WARNING: Lower bound of array CostDesc was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim CostDesc(5) As String
		
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
		Dim pvar As Object
		Dim pcost As Single
		Dim mcost As Single
		Dim total_mcost As Single
		Dim pbdes As String
		Dim tvar As Object
		Dim strDistPoint As String
		'UPGRADE_WARNING: Arrays in structure rsSect may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rsSect As DAO.Recordset
		Dim strError1 As String
		Dim strError2 As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object tvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tvar = True
		strDistPoint = vbNullString
		If Not (sx1 = 0 And sy1 = 1) Then
			strDistPoint = SectorString(sx1, sy1)
			
			'find the dist point's post update packing bonus
			rsSectors.Seek("=", sx1, sy1)
			If Not rsSectors.NoMatch Then
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				v = Production(rsSectors)
				If v.prodValidFlag Then
					If v.NewEff < 60 Then
						pbdes = vbNullString
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pbdes = CStr(v.newdes)
					End If
				End If
			End If
		End If
		
		'Set the items
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FOOD, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_FOOD, PR_ITEM) = "food"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_IRON, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_IRON, PR_ITEM) = "iron"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_LCM, PR_ITEM) = "lcm"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_HCM, PR_ITEM) = "hcm"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_DUST, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_DUST, PR_ITEM) = "dust"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_BAR, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_BAR, PR_ITEM) = "bar"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_GUN, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_GUN, PR_ITEM) = "gun"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHELL, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_SHELL, PR_ITEM) = "shell"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_OIL, PR_ITEM) = "oil"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_PET, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_PET, PR_ITEM) = "pet"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_MIL, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_MIL, PR_ITEM) = "mil"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_RAD, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_RAD, PR_ITEM) = "rad"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_CIV, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_CIV, PR_ITEM) = "civ"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_UW, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_UW, PR_ITEM) = "uw"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HAPPY, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_HAPPY, PR_ITEM) = "stroll"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_GRAD, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_GRAD, PR_ITEM) = "grad"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_TECH, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_TECH, PR_ITEM) = "tech"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_RESCH, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_RESCH, PR_ITEM) = "research"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHIP, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_SHIP, PR_ITEM) = "ships"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_PLANE, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_PLANE, PR_ITEM) = "planes"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_ARMY, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_ARMY, PR_ITEM) = "armies"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_CITY, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_CITY, PR_ITEM) = "cities"
		'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FORT, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ar(PR_FORT, PR_ITEM) = "forts"
		
		For n = 1 To PR_COMM_LAST
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_ONHAND). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ar(n, PR_ONHAND) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ar(n, PR_MADE) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ar(n, PR_USED) = 0
		Next n
		
		'Clear Event Markers
		EventMarkers.Clear()
		
		'Scan production for Grads, tech units, etc.
		strError1 = "Scanning for Production"
		rsSect = rsSectors.Clone
		Dim strTemp As String
		If Not (rsSect.BOF And rsSect.EOF) Then
			rsSect.MoveFirst()
			While Not rsSect.EOF
				
				If (sx1 = 0 And sy1 = 1) Or (rsSect.Fields("dist_x").Value = sx1 And rsSect.Fields("dist_y").Value = sy1) Then
					strError2 = SectorString(rsSect.Fields("x").Value, rsSect.Fields("y").Value)
					
					'set if captured sector
					If rsSect.Fields("*").Value = "*" Then
						strCapt = "*"
					Else
						strCapt = " "
					End If
					
					strSect = SectorString((rsSect.Fields("x").Value), (rsSect.Fields("y").Value))
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(CStr(rsSect.Fields(eff).Value), 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSectHeader = "Sector " + PadString(strSect, 8, False) + PadString(CStr(rsSect.Fields("eff").Value), 3) + "% " + rsSect.Fields("des").Value + rsSect.Fields("sdes").Value + strCapt
					'Get the post-update path cost to the sector
					'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pvar = BestPath(strDistPoint, strSect, EmpCommon.enumMobType.MT_COMMODITY, tvar)
					'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pcost = pvar(2)
					
					'Clear the sector commodity array
					For n = 1 To PR_COMM_ITEM
						For n2 = 1 To PR_CATG_LAST
							sc(n, n2) = 0
						Next n2
					Next n
					
					'get the food needed/available
					n = FoodRequired(rsSect)
					If n > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FOOD, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ar(PR_FOOD, PR_USED) = ar(PR_FOOD, PR_USED) + n
						sc(PR_FOOD, PR_USED) = n
					End If
					
					'If there is not enough food, save a warning message
					If ShowStarve And rsSect.Fields("food").Value - n < 0 Then
						strLine = "Needs " & CStr(n - rsSect.Fields("food").Value) & " food"
						cp(1).Add(strSectHeader & " - " & strLine)
						MarkSector(strSect, -1, strLine)
					End If
					
					'get the on-hand amount
					For n = 1 To PR_COMM_ITEM
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_ONHAND). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ar(n, PR_ONHAND) = ar(n, PR_ONHAND) + rsSect.Fields(CStr(ar(n, PR_ITEM))).Value
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sc(n, PR_ONHAND) = rsSect.Fields(CStr(ar(n, PR_ITEM))).Value
					Next n
					
					'get the requested amount
					For n = 1 To PR_COMM_ITEM
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_RQST). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ar(n, PR_RQST) = ar(n, PR_RQST) + rsSect.Fields(Left(CStr(ar(n, PR_ITEM)), 1) & "_dist").Value
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sc(n, PR_RQST) = rsSect.Fields(Left(CStr(ar(n, PR_ITEM)), 1) & "_dist").Value
					Next n
					
					'handle civ and uw growth
					sc(PR_CIV, PR_MADE) = Int(NewPop(rsSect, "c"))
					sc(PR_UW, PR_MADE) = Int(NewPop(rsSect, "u"))
					'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_CIV, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ar(PR_CIV, PR_MADE) = ar(PR_CIV, PR_MADE) + sc(PR_CIV, PR_MADE)
					'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_UW, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ar(PR_UW, PR_MADE) = ar(PR_UW, PR_MADE) + sc(PR_UW, PR_MADE)
					
					'get the production
					strError1 = "Production Report"
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSect)
					If v.prodValidFlag Then
						'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Select Case CStr(v.newdes)
							Case "~"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								rsSectorType.Seek("=", CStr(v.newdes))
								If Not rsSectorType.NoMatch Then
									If rsSectorType.Fields("product").Value = "oil" Then
										'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_OIL, PR_MADE) = ar(PR_OIL, PR_MADE) + v.ProdAmount
										'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										sc(PR_OIL, PR_MADE) = v.ProdAmount
									ElseIf rsSectorType.Fields("product").Value = "food" Then 
										'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FOOD, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_FOOD, PR_MADE) = ar(PR_FOOD, PR_MADE) + v.ProdAmount
										'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										sc(PR_FOOD, PR_MADE) = v.ProdAmount
									End If
								End If
							Case "m"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_IRON, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_IRON, PR_MADE) = ar(PR_IRON, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_IRON, PR_MADE) = v.ProdAmount
							Case "g", "^"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_DUST, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_DUST, PR_MADE) = ar(PR_DUST, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_DUST, PR_MADE) = v.ProdAmount
							Case "u"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_RAD, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_RAD, PR_MADE) = ar(PR_RAD, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_RAD, PR_MADE) = v.ProdAmount
							Case "o"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_OIL, PR_MADE) = ar(PR_OIL, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_OIL, PR_MADE) = v.ProdAmount
							Case "a"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FOOD, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_FOOD, PR_MADE) = ar(PR_FOOD, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_FOOD, PR_MADE) = v.ProdAmount
							Case "e"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_MIL, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_MIL, PR_MADE) = ar(PR_MIL, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_MIL, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_CIV, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_CIV, PR_USED) = ar(PR_CIV, PR_USED) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_CIV, PR_USED) = v.ProdAmount
							Case "j"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_LCM, PR_MADE) = ar(PR_LCM, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_IRON, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_IRON, PR_USED) = ar(PR_IRON, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_LCM, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_IRON, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_IRON, PR_LCM) = ut(PR_IRON, PR_LCM) + v.unitsConsumed
							Case "k"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_HCM, PR_MADE) = ar(PR_HCM, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_IRON, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_IRON, PR_USED) = ar(PR_IRON, PR_USED) + v.unitsConsumed * 2
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_HCM, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_IRON, PR_USED) = v.unitsConsumed * 2
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_IRON, PR_HCM) = ut(PR_IRON, PR_HCM) + v.unitsConsumed * 2
							Case "b"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_BAR, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_BAR, PR_MADE) = ar(PR_BAR, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_DUST, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_DUST, PR_USED) = ar(PR_DUST, PR_USED) + v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_BAR, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_DUST, PR_USED) = v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_DUST, PR_BAR) = ut(PR_DUST, PR_BAR) + v.unitsConsumed * 5
							Case "c"
								n = rsSect.Fields("eff").Value
								If rsSect.Fields("des").Value <> "c" Then
									n = 0
								End If
								n = CShort(v.NewEff) - n 'new efficiency- current efficiency
								If n > 0 Then
									If Not (rsBuild.BOF And rsBuild.EOF) Then
										rsBuild.Seek("=", "i", v.newdes)
										If Not rsBuild.NoMatch Then
											If rsBuild.Fields("lcm").Value > 0 Then
												'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + (n * rsBuild.Fields("lcm").Value)
												sc(PR_LCM, PR_USED) = (n * rsBuild.Fields("lcm").Value)
												ut(PR_LCM, PR_CITY) = ut(PR_LCM, PR_CITY) + (n * rsBuild.Fields("lcm").Value)
											End If
											If rsBuild.Fields("hcm").Value > 0 Then
												'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + (n * rsBuild.Fields("hcm").Value)
												sc(PR_HCM, PR_USED) = (n * rsBuild.Fields("hcm").Value)
												ut(PR_HCM, PR_CITY) = ut(PR_HCM, PR_CITY) + (n * rsBuild.Fields("hcm").Value)
											End If
										End If
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + n
										sc(PR_LCM, PR_USED) = n
										ut(PR_LCM, PR_CITY) = ut(PR_LCM, PR_CITY) + n
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + (n * 2)
										sc(PR_HCM, PR_USED) = n * 2
										ut(PR_HCM, PR_CITY) = ut(PR_HCM, PR_CITY) + (n * 2)
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_CITY, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_CITY, PR_MADE) = ar(PR_CITY, PR_MADE) + 1
								End If
							Case "i"
								If frmOptions.bSPAtlantis Then
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHELL, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_SHELL, PR_MADE) = ar(PR_SHELL, PR_MADE) + v.ProdAmount
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_SHELL, PR_MADE) = v.ProdAmount
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHELL, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_SHELL, PR_MADE) = ar(PR_SHELL, PR_MADE) + v.ProdAmount
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 2
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + v.unitsConsumed
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_SHELL, PR_MADE) = v.ProdAmount
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_LCM, PR_USED) = v.unitsConsumed * 2
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_HCM, PR_USED) = v.unitsConsumed
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_LCM, PR_SHELL) = ut(PR_LCM, PR_SHELL) + v.unitsConsumed * 2
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_HCM, PR_SHELL) = ut(PR_HCM, PR_SHELL) + v.unitsConsumed * 1
								End If
							Case "d"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_GUN, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_GUN, PR_MADE) = ar(PR_GUN, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + v.unitsConsumed * 10
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_GUN, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_OIL, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_LCM, PR_USED) = v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_HCM, PR_USED) = v.unitsConsumed * 10
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_LCM, PR_GUN) = ut(PR_LCM, PR_GUN) + v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_HCM, PR_GUN) = ut(PR_HCM, PR_GUN) + v.unitsConsumed * 10
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_OIL, PR_GUN) = ut(PR_OIL, PR_GUN) + v.unitsConsumed
							Case "%"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_PET, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_PET, PR_MADE) = ar(PR_PET, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_PET, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_OIL, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_OIL, PR_PET) = ut(PR_OIL, PR_PET) + v.unitsConsumed
							Case "q" 'Heavy Metal II biofuel
								If Not (rsSectorType.BOF And rsSectorType.EOF) Then
									rsSectorType.Seek("=", "q")
									If Not rsSectorType.NoMatch Then
										If rsSectorType.Fields("use1n").Value > 0 Then
											'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FOOD, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ar(PR_FOOD, PR_USED) = ar(PR_FOOD, PR_USED) + (v.unitsConsumed * rsSectorType.Fields("use1n").Value)
											'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											sc(PR_FOOD, PR_USED) = (v.unitsConsumed * rsSectorType.Fields("use1n").Value)
											'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ut(PR_FOOD, PR_OIL) = ut(PR_FOOD, PR_OIL) + (v.unitsConsumed * rsSectorType.Fields("use1n").Value)
										End If
									End If
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FOOD, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_FOOD, PR_USED) = ar(PR_FOOD, PR_USED) + v.unitsConsumed * 5
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_FOOD, PR_USED) = v.unitsConsumed * 5
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_FOOD, PR_OIL) = ut(PR_FOOD, PR_OIL) + v.unitsConsumed * 5
								End If
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_OIL, PR_MADE) = ar(PR_OIL, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_OIL, PR_MADE) = v.ProdAmount
							Case "l"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_GRAD, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_GRAD, PR_MADE) = ar(PR_GRAD, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_LCM, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_LCM, PR_GRAD) = ut(PR_LCM, PR_GRAD) + v.unitsConsumed
							Case "y" '260104 rjk: Added for St@r W@rs only
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_UW, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_UW, PR_MADE) = ar(PR_UW, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 2
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_UW, PR_MADE) = v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_LCM, PR_USED) = v.unitsConsumed * 2
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_HCM, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_LCM, PR_UW) = ut(PR_LCM, PR_UW) + v.unitsConsumed * 2
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_HCM, PR_UW) = ut(PR_HCM, PR_UW) + v.unitsConsumed * 1
							Case "p"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HAPPY, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_HAPPY, PR_MADE) = ar(PR_HAPPY, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_LCM, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_LCM, PR_HAPPY) = ut(PR_LCM, PR_HAPPY) + v.unitsConsumed
							Case "t"
								If frmOptions.bSPAtlantis Then
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_TECH, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_TECH, PR_MADE) = ar(PR_TECH, PR_MADE) + v.ProdAmount
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHELL, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_SHELL, PR_USED) = ar(PR_SHELL, PR_USED) + v.unitsConsumed * 50
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_SHELL, PR_USED) = v.unitsConsumed * 50
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_SHELL, PR_TECH) = ut(PR_SHELL, PR_TECH) + v.unitsConsumed * 50
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_TECH, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_TECH, PR_MADE) = ar(PR_TECH, PR_MADE) + v.ProdAmount
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_DUST, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_DUST, PR_USED) = ar(PR_DUST, PR_USED) + v.unitsConsumed
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed * 5
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 10
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_DUST, PR_USED) = v.unitsConsumed
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_OIL, PR_USED) = v.unitsConsumed * 5
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sc(PR_LCM, PR_USED) = v.unitsConsumed * 10
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_DUST, PR_TECH) = ut(PR_DUST, PR_TECH) + v.unitsConsumed
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_OIL, PR_TECH) = ut(PR_OIL, PR_TECH) + v.unitsConsumed * 5
									'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ut(PR_LCM, PR_TECH) = ut(PR_LCM, PR_TECH) + v.unitsConsumed * 10
								End If
							Case "r"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_RESCH, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_RESCH, PR_MADE) = ar(PR_RESCH, PR_MADE) + v.ProdAmount
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_DUST, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_DUST, PR_USED) = ar(PR_DUST, PR_USED) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_OIL, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_OIL, PR_USED) = ar(PR_OIL, PR_USED) + v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + v.unitsConsumed * 10
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_DUST, PR_USED) = v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_OIL, PR_USED) = v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sc(PR_LCM, PR_USED) = v.unitsConsumed * 10
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_DUST, PR_RESCH) = ut(PR_DUST, PR_RESCH) + v.unitsConsumed
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_OIL, PR_RESCH) = ut(PR_OIL, PR_RESCH) + v.unitsConsumed * 5
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ut(PR_LCM, PR_RESCH) = ut(PR_LCM, PR_RESCH) + v.unitsConsumed * 10
								
								'for harbors, etc - we have to simulate builds.
							Case "*", "h", "!", "f"
								
								For n = 1 To 5
									For n2 = 1 To 7
										Costs(n2, n) = 0
									Next n2
								Next n
								
								If CShort(v.NewEff) > 59 Then
									Costs(5, 1) = 9999999 'Assume enough money
									Costs(5, 2) = rsSect.Fields("avail").Value 'avail
									If frmOptions.bSPAtlantis Then
										Costs(5, 3) = rsSect.Fields("shell").Value 'lcm
									Else
										Costs(5, 3) = rsSect.Fields("lcm").Value 'lcm
									End If
									Costs(5, 4) = rsSect.Fields("hcm").Value 'hcm
									CostDesc(2) = "avail"
									If frmOptions.bSPAtlantis Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHELL, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										CostDesc(3) = ar(PR_SHELL, PR_ITEM)
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										CostDesc(3) = ar(PR_LCM, PR_ITEM)
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									CostDesc(4) = ar(PR_HCM, PR_ITEM)
									
									n2 = 0 ' will hold count
									'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Select Case CStr(v.newdes)
										Case "h"
											rs = rsShip
											BuildType = "s"
											MaxEffGain = rsVersion.Fields("Eff gain - Ship").Value
											Target = PR_SHIP
											LastItem = 1
											'check for post update avail used
											If HarborAvail Then
												'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Costs(5, 2) = CShort(v.ProdAmount)
											End If
											
											Costs(5, 5) = 9999 'nothing
											
										Case "!", "f"
											rs = rsLand
											BuildType = "l"
											MaxEffGain = rsVersion.Fields("Eff gain - Land").Value
											Target = PR_ARMY
											LastItem = PR_GUN
											Costs(5, 5) = rsSect.Fields("gun").Value
											'UPGRADE_WARNING: Couldn't resolve default property of object ar(LastItem, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											CostDesc(5) = ar(LastItem, PR_ITEM)
											'check for post update avail used
											'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If (CStr(v.newdes) = "!" And HQAvail) Or (CStr(v.newdes) = "f" And FortAvail) Then
												'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Costs(5, 2) = CShort(v.ProdAmount)
											End If
											
										Case "*"
											rs = rsPlane
											BuildType = "p"
											MaxEffGain = rsVersion.Fields("Eff gain - Plane").Value
											Target = PR_PLANE
											LastItem = PR_MIL
											'check for post update avail used
											If AirportAvail Then
												'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Costs(5, 2) = CShort(v.ProdAmount)
											End If
											
											Costs(5, 5) = rsSect.Fields("mil").Value
											'UPGRADE_WARNING: Couldn't resolve default property of object ar(LastItem, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											CostDesc(5) = ar(LastItem, PR_ITEM)
									End Select
									
									'Now run through the units
									If Not (rs.BOF And rs.EOF) Then
										rs.MoveFirst()
									End If
									While Not rs.EOF
										
										'Check to see if it is in this hex
										If rsSect.Fields("x").Value = rs.Fields("x").Value And rsSect.Fields("y").Value = rs.Fields("y").Value Then
											
											'check to see that it is not 100% already
											If rs.Fields("eff").Value < 100 And rs.Fields("off").Value = False Then
												'get the build requirements
												rsBuild.Seek("=", BuildType, rs.Fields("type"))
												If Not rsBuild.NoMatch Then
													'Compute efficieny gain
													NewEff = rs.Fields("eff").Value + MaxEffGain
													If NewEff > 100 Then
														NewEff = 100
													End If
													EffGain = (NewEff - rs.Fields("eff").Value) / 100
													
													'load costs
													Costs(1, 1) = rsBuild.Fields("cost").Value
													'030404 rjk: Switched to use the avail from the server instead of calculating it,
													'            in case the calculation in the server is changed.
													'            20 + rsBuild.Fields("lcm") + rsBuild.Fields("hcm") * 2
													Costs(1, 2) = rsBuild.Fields("avail").Value
													Costs(1, 3) = rsBuild.Fields("lcm").Value
													Costs(1, 4) = rsBuild.Fields("hcm").Value
													'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													Select Case CStr(v.newdes)
														Case "h"
															Costs(1, 5) = 0
														Case "*"
															Costs(1, 5) = rsBuild.Fields("mil").Value
														Case Else
															Costs(1, 5) = rsBuild.Fields("gun").Value
													End Select
													For n = 1 To 5
														Costs(2, n) = Costs(1, n) * EffGain
														Costs(3, n) = CSng(Int(Costs(2, n) + 0.995))
														Costs(6, n) = Costs(6, n) + Costs(3, n)
													Next n
													'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object ar(Target, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													ar(Target, PR_MADE) = ar(Target, PR_MADE) + 1
													n2 = n2 + 1
												End If
											End If
										End If
										rs.MoveNext()
									End While
									
									'Put out build message
									If Costs(6, 2) > 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(Target, PR_ITEM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										strTemp = strSectHeader & " - building " & CStr(n2) & " " + ar(Target, PR_ITEM) + "."
									Else
										strTemp = strSectHeader & " - is idle."
									End If
									
									If Costs(5, 2) > Costs(6, 2) Then
										strTemp = strTemp & " " & CStr(Costs(5, 2) - Costs(6, 2)) & " avail left."
									End If
									cp(6).Add(strTemp)
									
									'check for shortages
									For n = 2 To 5
										If Costs(6, n) > Costs(5, n) Then
											cp(6).Add(strSectHeader & " - is short " & CStr(Costs(6, n) - Costs(5, n)) & " " & CostDesc(n) & " for builds")
										End If
									Next 
									
									'add to arrays
									If frmOptions.bSPAtlantis Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_SHELL, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_SHELL, PR_USED) = ar(PR_SHELL, PR_USED) + Costs(6, 3)
										sc(PR_SHELL, PR_USED) = sc(PR_SHELL, PR_USED) + Costs(6, 3)
										ut(PR_SHELL, Target) = ut(PR_SHELL, Target) + Costs(6, 3)
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_LCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_LCM, PR_USED) = ar(PR_LCM, PR_USED) + Costs(6, 3)
										sc(PR_LCM, PR_USED) = sc(PR_LCM, PR_USED) + Costs(6, 3)
										ut(PR_LCM, Target) = ut(PR_LCM, Target) + Costs(6, 3)
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + Costs(6, 4)
									sc(PR_HCM, PR_USED) = sc(PR_HCM, PR_USED) + Costs(6, 4)
									ut(PR_HCM, Target) = ut(PR_HCM, Target) + Costs(6, 4)
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ar(LastItem, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ar(LastItem, PR_USED) = ar(LastItem, PR_USED) + Costs(6, 5)
									sc(LastItem, PR_USED) = sc(LastItem, PR_USED) + Costs(6, 5)
									ut(LastItem, Target) = ut(LastItem, Target) + Costs(6, 5)
									
								End If
								'handle new forts being built
								'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If CStr(v.newdes) = "f" Then
									n = rsSect.Fields("eff").Value
									If rsSect.Fields("des").Value <> "f" Then
										n = 0
									End If
									n = CShort(v.NewEff) - n 'new efficiency- current efficiency
									If n > 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_HCM, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_HCM, PR_USED) = ar(PR_HCM, PR_USED) + n
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object ar(PR_FORT, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ar(PR_FORT, PR_MADE) = ar(PR_FORT, PR_MADE) + 1
										ut(PR_HCM, PR_FORT) = ut(PR_HCM, PR_FORT) + n
									End If
								End If
								
							Case Else
						End Select
						
						strError1 = "Excess Civilians"
						n = CShort(v.ExcessCiv)
						'if we are short civs, let us know about it
						If n < 0 And InStr("h*!f", rsSect.Fields("des").Value) = 0 Then
							cp(4).Add(strSectHeader & " - " & CStr(0 - n) & " civs needed")
						End If
						
						'If there are excess civs, save a warning message
						If ShowIdle And n > 10 And strCapt = " " Then
							cp(5).Add(strSectHeader & " - " & CStr(n) & " idle civs")
							MarkSector(strSect, -1, "Idle civs")
						End If
					End If
					
					strError1 = "Excess UWs"
					'if there are too many uws, save a warning message
					If ShowOvrPop And rsSect.Fields("uw").Value > MaxSafeUws Then
						cp(3).Add(strSectHeader & " - Over MaxSafeUws by " & CStr(rsSect.Fields("uw").Value - MaxSafeUws) & " uws")
						MarkSector(strSect, -1, "Over Max Safe Uws")
					End If
					
					strError1 = "City Check"
					'if there are too many civs, and its not a captured sector, save a warning message
					If ShowOvrPop And strCapt = " " Then
						If rsVersion.Fields("BIG_CITY").Value And rsSect.Fields("des").Value = "c" Then
							If rsSect.Fields("civ").Value > (CShort(CInt(MaxPopulation(rsSect.Fields("des").Value) * 9 * CInt(rsSect.Fields("eff").Value) / 100) + MaxPopulation(rsSect.Fields("des").Value))) Then
								cp(2).Add(strSectHeader & " - Over MaxSafeCivs by " & CStr(rsSect.Fields("civ").Value - CShort(CInt(MaxPopulation(rsSect.Fields("des").Value) * 9 * CInt(rsSect.Fields("eff").Value) / 100) + MaxPopulation(rsSect.Fields("des").Value))) & " civs")
								MarkSector(strSect, -1, "Over Max Safe Civs")
							End If
						ElseIf rsSect.Fields("civ").Value > MaxPopulation(rsSect.Fields("des").Value) Then 
							cp(2).Add(strSectHeader & " - Over MaxSafeCivs by " & CStr(rsSect.Fields("civ").Value - MaxPopulation(rsSect.Fields("des").Value)) & " civs")
							MarkSector(strSect, -1, "Over Max Safe Civs")
						End If
					End If
					
					strError1 = "Determining Mobility"
					'Figure out the total mob cost to ship to the sector
					For n = 1 To PR_COMM_ITEM
						If pcost > 0 And sc(n, PR_RQST) > 0 Then
							n2 = sc(n, PR_RQST) - (sc(n, PR_ONHAND) + sc(n, PR_MADE) - sc(n, PR_USED))
							If n2 > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mcost = (pcost * MobilityWeight(Left(ar(n, PR_ITEM), 1), n2, pbdes)) / 10#
								total_mcost = total_mcost + mcost
							End If
						End If
					Next n
					
				End If
				rsSect.MoveNext()
			End While
			
		End If
		
		strError1 = "Generating Report"
		'Set Report
		frmReport.Text = "Area Production Summary"
		frmReport.ClearReport()
		
		'Const PR_ITEM = 1
		'Const PR_ONHAND = 2
		'Const PR_MADE = 3
		'Const PR_USED = 4
		'Const PR_CATG_LAST = 6
		
		'Set Summary Headers
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(EXCESS, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(THRSHLD, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(LEFT, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(DELTA, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(USING, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(MAKING, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(ON HAND, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strLine = PadString("ITEM", 8, False) + PadString("ON HAND", 8) + PadString("MAKING", 8) + PadString("USING", 8) + PadString("DELTA", 8) + PadString("LEFT", 8) + PadString("THRSHLD", 8) + PadString("EXCESS", 8)
		frmReport.AddReportLine(strLine)
		
		'write Summaries
		For n = 1 To PR_COMM_ITEM
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = PadString(CStr(ar(n, PR_ITEM)), 8, False)
			For n2 = 2 To 4
				'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strLine = strLine + PadString(VB6.Format(ar(n, n2), "###,##0"), 8)
			Next n2
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = strLine + PadString(VB6.Format(ar(n, PR_MADE) - ar(n, PR_USED), "###,##0"), 8)
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = strLine + PadString(VB6.Format(ar(n, PR_ONHAND) + ar(n, PR_MADE) - ar(n, PR_USED), "###,##0"), 8)
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = strLine + PadString(VB6.Format(ar(n, PR_RQST), "###,##0"), 8)
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_RQST). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_USED). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(n, PR_MADE). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strLine = strLine + PadString(VB6.Format(ar(n, PR_ONHAND) + ar(n, PR_MADE) - ar(n, PR_USED) - ar(n, PR_RQST), "###,##0"), 8)
			
			frmReport.AddReportLine(strLine)
		Next n
		
		'if this is a single dist point, write out cost
		If Len(strDistPoint) > 0 Then
			frmReport.AddReportLine(vbNullString)
			strLine = "Distribution mob. cost: " & CStr(CSng(CInt(CSng(total_mcost) * 1000) / 1000))
			frmReport.AddReportLine(strLine)
			frmReport.AddReportLine(vbNullString)
		End If
		
		'Put out item usage details
		For n = 1 To PR_COMM_ITEM
			For n2 = 1 To PR_COMM_LAST
				ut(n, 0) = ut(n, 0) + ut(n, n2)
			Next n2
			If ut(n, 0) > 0 Then
				frmReport.AddReportLine(vbNullString)
				'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmReport.AddReportLine(CStr(ar(n, PR_ITEM)) & " Summary - " & VB6.Format(ut(n, 0), "###,##0") & " used")
				For n2 = 1 To PR_COMM_LAST
					If ut(n, n2) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object ar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object PadString(Format$(ar(n2, PR_MADE), ###,##0), 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						frmReport.AddReportLine(PadString(VB6.Format(ut(n, n2), "###,##0"), 8) + " used in making" + PadString(VB6.Format(ar(n2, PR_MADE), "###,##0"), 8) + " " + CStr(ar(n2, PR_ITEM)))
					End If
				Next n2
			End If
		Next n
		
		frmReport.AddReportLine(vbNullString)
		For n = 1 To PR_MSGC
			For	Each varLine In cp(n)
				'UPGRADE_WARNING: Couldn't resolve default property of object varLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmReport.AddReportLine(CStr(varLine))
			Next varLine
			frmReport.AddReportLine(vbNullString)
		Next 
		
		frmReport.Show()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Exit Sub
Empire_Error: 
		EmpireError("ProdSummaryReport", strError1, strError2)
	End Sub
	
	Public Sub AirCombatReport()
		On Error GoTo Empire_Error
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim hldkey As Object
		Dim hldstatus As String
		Dim statuscount As Short
		Dim nationcount As Short
		Dim timActiveDate As Date
		Dim n As Short
		Dim X As Short
		Dim strTemp As String
		Dim numnations As Short
		Dim aryNations() As Short
		Dim Done As Boolean
		Dim bCurrent As Boolean
		
		rs = rsAirCombat
		ReDim aryNations(Nations.Count)
		
		If rs.BOF And rs.EOF Then
			frmReport.AddReportLine("No air combat history found")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object hldkey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldkey = rs.Index
			
			rs.Index = "TimeKey"
			rs.MoveFirst()
			If timActiveDate = timNextUpdate Then
				timActiveDate = rs.Fields("nextupdate").Value
			Else
				timActiveDate = timNextUpdate
			End If
			n = 1
			While Not rs.EOF
				'keep track the order in which nations are found
				If rs.Fields("nation").Value > 0 Then '120304 rjk: Skip the invalid nations (-1)
					If aryNations(rs.Fields("nation").Value) = 0 Then
						aryNations(rs.Fields("nation").Value) = n
						n = n + 1
					End If
					
					'Mark Historical
					If timActiveDate <> rs.Fields("nextupdate").Value Then
						If rs.Fields("status").Value = "A" Or rs.Fields("status").Value = "D" Then
							rs.Edit()
							rs.Fields("status").Value = "T"
							rs.Fields("eff").Value = 100
							rs.Update()
						ElseIf rs.Fields("status").Value = "S" Then 
							rs.Edit()
							rs.Fields("status").Value = "X"
							rs.Fields("eff").Value = 0
							rs.Update()
						End If
					End If
				End If
				rs.MoveNext()
			End While
			numnations = n - 1
			
			'Now do report
			rs.Index = "ReportKey"
			For X = 1 To numnations
				n = 1
				While aryNations(n) <> X
					n = n + 1
				End While
				
				rs.Seek(">=", n) 'must use >= with a partial key
				frmReport.AddReportLine("Player #" & CStr(n) & " - " & Nations.NationName(n))
				frmReport.AddReportLine(vbNullString)
				hldstatus = vbNullString
				nationcount = 0
				Done = False
				
				While Not Done
					
					'process status change
					If rs.Fields("status").Value <> hldstatus Then
						If Len(hldstatus) > 0 Then
							frmReport.AddReportLine("     Total: " & CStr(statuscount) & " planes")
							frmReport.AddReportLine(vbNullString)
						End If
						Select Case rs.Fields("status").Value
							Case "A"
								frmReport.AddReportLine("Active")
								frmReport.AddReportLine("======")
								bCurrent = True
							Case "D"
								frmReport.AddReportLine("Disabled")
								frmReport.AddReportLine("========")
								bCurrent = True
							Case "S"
								frmReport.AddReportLine("Shot Down")
								frmReport.AddReportLine("=========")
								bCurrent = True
							Case "T"
								frmReport.AddReportLine("Historical Active")
								frmReport.AddReportLine("=================")
								bCurrent = False
							Case "X"
								frmReport.AddReportLine("Historical Shot Down")
								frmReport.AddReportLine("====================")
								bCurrent = False
						End Select
						hldstatus = rs.Fields("status").Value
						statuscount = 0
					End If
					
					'output line
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(CStr(rs.Fields(id)), 5, False). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strTemp = "     " + rs.Fields("type").Value + " #" + PadString(CStr(rs.Fields("id").Value), 5, False)
					If bCurrent Then
						'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strTemp = strTemp + PadString(CStr(rs.Fields("eff").Value) & "%", 5) + "  " + CStr(rs.Fields("numcombats").Value) + " combat"
						If rs.Fields("numcombats").Value > 1 Then
							strTemp = strTemp & "s"
						End If
					End If
					frmReport.AddReportLine(strTemp)
					
					statuscount = statuscount + 1
					nationcount = nationcount + 1
					rs.MoveNext()
					If rs.EOF Then
						Done = True
					ElseIf rs.Fields("nation").Value <> n Then 
						Done = True
					End If
					
				End While
				frmReport.AddReportLine("     Total: " & CStr(statuscount) & " planes")
				frmReport.AddReportLine(vbNullString)
				frmReport.AddReportLine("Nation Total: " & CStr(nationcount) & " planes")
				frmReport.AddReportLine(vbNullString)
			Next X
			'UPGRADE_WARNING: Couldn't resolve default property of object hldkey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rs.Index = hldkey
		End If
		
		frmReport.Show()
		Exit Sub
Empire_Error: 
		EmpireError("AirCombatReport", vbNullString, vbNullString)
	End Sub
	
	Public Function ApplySectorDamage(ByRef strSect As String, ByRef nDamage As Short, Optional ByRef bAbsolute As Boolean = False) As String
		On Error GoTo Empire_Error
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bAbsolute) Then
			bAbsolute = False
		End If
		
		Dim sx As Short
		Dim sy As Short
		Dim seceff As Short
		Dim sectype As String
		Dim found As Boolean
		Dim NewEff As Short
		
		If Not ParseSectors(sx, sy, strSect) Then
			ApplySectorDamage = vbNullString
			Exit Function
		End If
		
		found = True
		'check for your own sector
		rsSectors.Seek("=", sx, sy)
		If Not rsSectors.NoMatch Then
			sectype = rsSectors.Fields("des").Value
			seceff = CShort(CStr(rsSectors.Fields("eff").Value))
			If bAbsolute Then
				NewEff = seceff - nDamage
			Else
				NewEff = SectorEffAfterDamage(nDamage, sectype, seceff)
			End If
			rsSectors.Edit()
			rsSectors.Fields("eff").Value = NewEff
			rsSectors.Update()
		Else
			'Check for Enemy Sector
			rsEnemySect.Seek("=", sx, sy)
			If Not rsEnemySect.NoMatch Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rsEnemySect.Fields("des").Value) Or IsDbNull(rsEnemySect.Fields("eff").Value) Then
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
				rsEnemySect.Edit()
				rsEnemySect.Fields("eff").Value = NewEff
				rsEnemySect.Update()
			Else
				rsBmap.Seek("=", sx, sy)
				If Not rsBmap.NoMatch Then
					sectype = rsBmap.Fields("des").Value
					seceff = 100
					If bAbsolute Then
						NewEff = seceff - nDamage
					Else
						NewEff = SectorEffAfterDamage(nDamage, sectype, seceff)
					End If
				Else
					ApplySectorDamage = "No information on " & strSect
					found = False
				End If
			End If
		End If
		
		If found Then
			ApplySectorDamage = "Sector " & strSect & " efficiency reduced from " & CStr(seceff) & "% to " & CStr(NewEff) & "%"
		End If
		Exit Function
Empire_Error: 
		EmpireError("ApplySectorDamage", vbNullString, strSect)
		
	End Function
	
	Public Function ApplyShipDamage(ByRef strShip As String, ByRef nDamage As Short, Optional ByRef bAbsolute As Boolean = False) As String
		On Error GoTo Empire_Error
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bAbsolute) Then
			bAbsolute = False
		End If
		
		Dim NewEff As Short
		Dim hldIndex As String
		Dim shipid As Short
		Dim seceff As Short
		
		shipid = CShort(strShip)
		
		hldIndex = rsShip.Index
		rsShip.Index = "PrimaryKey"
		'check for your own ships
		rsShip.Seek("=", shipid)
		If Not rsShip.NoMatch Then
			seceff = CShort(CStr(rsShip.Fields("eff").Value))
			If bAbsolute Then
				NewEff = seceff - nDamage
			Else
				'neweff = need to define raw damage routine
			End If
			rsShip.Edit()
			rsShip.Fields("eff").Value = NewEff
			rsShip.Update()
			rsShip.Index = hldIndex
		Else
			'Check for enemy ship
			rsShip.Index = hldIndex
			hldIndex = rsEnemyUnit.Index
			rsEnemyUnit.Index = "ByID"
			rsEnemyUnit.Seek("=", "s", shipid)
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
				rsEnemyUnit.Edit()
				rsEnemyUnit.Fields("eff").Value = NewEff
				rsEnemyUnit.Update()
				rsEnemyUnit.Index = hldIndex
			Else
				'build a ship record
				rsEnemyUnit.AddNew()
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
				rsEnemyUnit.Update()
			End If
		End If
		
		ApplyShipDamage = "Ship (#" & strShip & ") efficiency reduced from " & CStr(seceff) & "% to " & CStr(NewEff) & "%"
		
		Exit Function
Empire_Error: 
		EmpireError("ApplyShipDamage", vbNullString, strShip)
		
	End Function
	
	Public Sub CopySectorInfo(ByRef sx1 As Short, ByRef sx2 As Short, ByRef sy1 As Short, ByRef sy2 As Short, ByRef dx As Short, ByRef dy As Short, ByRef CopyType As Short)
		'On Error GoTo Empire_Error
		
		Dim X As Short
		Dim Y As Short
		Dim tx As Short
		Dim ty As Short
		Dim n As Short
		Dim vhld() As Object
		
		For X = sx1 To sx2
			For Y = sy1 To sy2
				If (X + Y) Mod 2 = 0 Then
					rsSectors.Seek("=", X, Y)
					If Not rsSectors.NoMatch Then
						ReDim vhld(rsSectors.Fields.Count)
						For n = 0 To rsSectors.Fields.Count - 1
							'UPGRADE_WARNING: Couldn't resolve default property of object vhld(rsSectors.Fields().OrdinalPosition). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vhld(rsSectors.Fields(n).OrdinalPosition) = rsSectors.Fields(n).Value
						Next n
						tx = dx + X - sx1
						ty = dy + Y - sy1
						CopySector(vhld, tx, ty, CopyType)
					End If
				End If
			Next Y
		Next X
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
	End Sub
	
	
	Public Sub CopySector(ByRef vhld As Object, ByRef tx As Short, ByRef ty As Short, ByRef CopyType As Short)
		'On Error GoTo Empire_Error
		
		Dim n As Short
		Dim strDes As String
		Dim strCmd As String
		Dim strCmdDes As String
		Dim strCom As String
		Dim intRes As Short
		
		rsSectors.Seek("=", tx, ty)
		If Not rsSectors.NoMatch Then
			rsSectors.Edit()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strDes = vhld(rsSectors.Fields("des").OrdinalPosition)
			If rsSectors.Fields("des").Value <> strDes Then
				rsSectors.Fields("des").Value = strDes
				
				'designate the sector
				strCmdDes = "designate " & SectorString(tx, ty) & " " & strDes
				frmEmpCmd.SubmitEmpireCommand(strCmdDes, True)
				rsSectors.Fields("eff").Value = 0
			End If
			
			If strDes <> "." Then
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("eff").OrdinalPosition) - rsSectors.Fields("eff").Value
				If intRes <> 0 Then
					rsSectors.Fields("eff").Value = rsSectors.Fields("eff").Value + intRes
					strCmd = "setsector eff " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("mob").OrdinalPosition) - rsSectors.Fields("mob").Value
				If intRes <> 0 Then
					rsSectors.Fields("mob").Value = rsSectors.Fields("mob").Value + intRes
					strCmd = "setsector mob " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("work").OrdinalPosition) - rsSectors.Fields("work").Value
				If intRes <> 0 Then
					rsSectors.Fields("work").Value = rsSectors.Fields("work").Value + intRes
					strCmd = "setsector work " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("avail").OrdinalPosition) - rsSectors.Fields("avail").Value
				If intRes <> 0 Then
					rsSectors.Fields("avail").Value = rsSectors.Fields("avail").Value + intRes
					strCmd = "setsector avail " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				'102203 rjk: Added road / rail / defense and mines
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("road").OrdinalPosition) - rsSectors.Fields("road").Value
				If intRes <> 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strCmd = "edit land " & SectorString(tx, ty) & " :R " & CStr(vhld(rsSectors.Fields("road").OrdinalPosition))
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("rail").OrdinalPosition) - rsSectors.Fields("rail").Value
				If intRes <> 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strCmd = "edit land " & SectorString(tx, ty) & " :r " & CStr(vhld(rsSectors.Fields("rail").OrdinalPosition))
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("defense").OrdinalPosition) - rsSectors.Fields("defense").Value
				If intRes <> 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strCmd = "edit land " & SectorString(tx, ty) & " :d " & CStr(vhld(rsSectors.Fields("defense").OrdinalPosition))
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(vhld(rsSectors.Fields("sdes").OrdinalPosition)) Then
					If vhld(rsSectors.Fields("sdes").OrdinalPosition) <> rsSectors.Fields("sdes").Value Then
						'UPGRADE_WARNING: Couldn't resolve default property of object vhld(rsSectors.Fields(sdes).OrdinalPosition). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If vhld(rsSectors.Fields("sdes").OrdinalPosition) = " " Then
							'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strCmd = "edit land " & SectorString(tx, ty) & " :S " & CStr(vhld(rsSectors.Fields("des").OrdinalPosition))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strCmd = "edit land " & SectorString(tx, ty) & " :S " & CStr(vhld(rsSectors.Fields("sdes").OrdinalPosition))
						End If
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
				End If
				
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rsSectors.Fields("lmines").Value) Then '080404 rjk: Added Null check to lmines.
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(vhld(rsSectors.Fields("lmines").OrdinalPosition)) Then
						intRes = -rsSectors.Fields("lmines").Value
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						intRes = vhld(rsSectors.Fields("lmines").OrdinalPosition) - rsSectors.Fields("lmines").Value
					End If
					If intRes <> 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strCmd = "edit land " & SectorString(tx, ty) & " :M " & CStr(vhld(rsSectors.Fields("lmines").OrdinalPosition))
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
				End If
			End If
			
			'set resources
			If CopyType > 0 Then
				If strDes <> "." Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields("min").OrdinalPosition)
					If intRes <> rsSectors.Fields("min").Value Then
						rsSectors.Fields("min").Value = intRes
						strCmd = "setresource iron " & SectorString(tx, ty) & " " & CStr(intRes)
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields("gold").OrdinalPosition)
					If intRes <> rsSectors.Fields("gold").Value Then
						rsSectors.Fields("gold").Value = intRes
						strCmd = "setresource gold " & SectorString(tx, ty) & " " & CStr(intRes)
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields("uran").OrdinalPosition)
					If intRes <> rsSectors.Fields("uran").Value Then
						rsSectors.Fields("uran").Value = intRes
						strCmd = "setresource uran " & SectorString(tx, ty) & " " & CStr(intRes)
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("ocontent").OrdinalPosition)
				If intRes <> rsSectors.Fields("ocontent").Value Then
					rsSectors.Fields("ocontent").Value = intRes
					strCmd = "setresource oil " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("fert").OrdinalPosition)
				If intRes <> rsSectors.Fields("fert").Value Then
					rsSectors.Fields("fert").Value = intRes
					strCmd = "setresource fert " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
			End If
			
			'give commodities
			If CopyType > 1 And strDes <> "." Then
				strCom = "mil  gun  shellfood civ  uw   iron lcm  hcm  oil  dust pet  rad  bar  "
				For n = 0 To 13
					strCmd = Trim(Mid(strCom, 5 * n + 1, 5))
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields(strCmd).OrdinalPosition) - rsSectors.Fields(strCmd).Value
					If intRes <> 0 Then
						rsSectors.Fields(strCmd).Value = rsSectors.Fields(strCmd).Value + intRes
						strCmd = "give " & strCmd & " " & SectorString(tx, ty) & " " & CStr(intRes)
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
				Next n
			End If
		Else
			rsSectors.AddNew()
			
			rsSectors.Fields("x").Value = tx
			rsSectors.Fields("y").Value = ty
			
			'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strDes = vhld(rsSectors.Fields("des").OrdinalPosition)
			rsSectors.Fields("des").Value = strDes
			
			'designate the sector
			strCmdDes = "designate " & SectorString(tx, ty) & " " & strDes
			frmEmpCmd.SubmitEmpireCommand(strCmdDes, True)
			
			If strDes <> "." Then
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("eff").OrdinalPosition)
				rsSectors.Fields("eff").Value = intRes
				strCmd = "setsector eff " & SectorString(tx, ty) & " " & CStr(intRes)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("mob").OrdinalPosition)
				rsSectors.Fields("mob").Value = intRes
				strCmd = "setsector mob " & SectorString(tx, ty) & " " & CStr(intRes)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("work").OrdinalPosition)
				rsSectors.Fields("work").Value = intRes
				strCmd = "setsector work " & SectorString(tx, ty) & " " & CStr(intRes)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("avail").OrdinalPosition)
				rsSectors.Fields("avail").Value = intRes
				strCmd = "setsector avail " & SectorString(tx, ty) & " " & CStr(intRes)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCmd = "edit land " & SectorString(tx, ty) & " :R " & CStr(vhld(rsSectors.Fields("road").OrdinalPosition))
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCmd = "edit land " & SectorString(tx, ty) & " :r " & CStr(vhld(rsSectors.Fields("rail").OrdinalPosition))
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCmd = "edit land " & SectorString(tx, ty) & " :d " & CStr(vhld(rsSectors.Fields("defense").OrdinalPosition))
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(vhld(rsSectors.Fields("sdes").OrdinalPosition)) Then
					If Len(vhld(rsSectors.Fields("sdes").OrdinalPosition)) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object vhld(rsSectors.Fields(sdes).OrdinalPosition). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If vhld(rsSectors.Fields("sdes").OrdinalPosition) = " " Then
							'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strCmd = "edit land " & SectorString(tx, ty) & " :S " & CStr(vhld(rsSectors.Fields("des").OrdinalPosition))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strCmd = "edit land " & SectorString(tx, ty) & " :S " & CStr(vhld(rsSectors.Fields("sdes").OrdinalPosition))
						End If
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
				End If
				
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rsSectors.Fields("lmines").Value) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strCmd = "edit land " & SectorString(tx, ty) & " :M " & CStr(vhld(rsSectors.Fields("lmines").OrdinalPosition))
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
			End If
			
			'set resources
			If CopyType > 0 Then
				If strDes <> "." Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields("min").OrdinalPosition)
					rsSectors.Fields("min").Value = intRes
					strCmd = "setresource iron " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields("gold").OrdinalPosition)
					rsSectors.Fields("gold").Value = intRes
					strCmd = "setresource gold " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields("uran").OrdinalPosition)
					rsSectors.Fields("uran").Value = intRes
					strCmd = "setresource uran " & SectorString(tx, ty) & " " & CStr(intRes)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("ocontent").OrdinalPosition)
				rsSectors.Fields("ocontent").Value = intRes
				strCmd = "setresource oil " & SectorString(tx, ty) & " " & CStr(intRes)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intRes = vhld(rsSectors.Fields("fert").OrdinalPosition)
				rsSectors.Fields("fert").Value = intRes
				strCmd = "setresource fert " & SectorString(tx, ty) & " " & CStr(intRes)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			End If
			
			'give commodities
			If CopyType > 1 And strDes <> "." Then
				strCom = "mil  gun  shellfood civ  uw   iron lcm  hcm  oil  dust pet  rad  bar  "
				For n = 0 To 13
					strCmd = Trim(Mid(strCom, 5 * n + 1, 5))
					'UPGRADE_WARNING: Couldn't resolve default property of object vhld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intRes = vhld(rsSectors.Fields(strCmd).OrdinalPosition)
					rsSectors.Fields(strCmd).Value = intRes
					If intRes <> 0 Then
						strCmd = "give " & strCmd & " " & SectorString(tx, ty) & " " & CStr(intRes)
						frmEmpCmd.SubmitEmpireCommand(strCmd, True)
					End If
				Next n
			End If
		End If
		rsSectors.Update()
		
	End Sub
	
	
	Public Function DefenseStrength(ByRef Sector As String) As Single
		
		Dim sector_strength As Single
		Dim d_dstr As Single
		Dim sct_effic As Single
		Dim sct_type As String
		Dim sx As Short
		Dim sy As Short
		
		If Not ParseSectors(sx, sy, Sector) Then
			DefenseStrength = 1
			Exit Function
		End If
		
		rsEnemySect.Seek("=", sx, sy)
		If rsEnemySect.NoMatch Then
			DefenseStrength = 1
			Exit Function
		End If
		sct_effic = CSng(Val(rsEnemySect.Fields("eff").Value))
		sct_type = rsEnemySect.Fields("des").Value
		
		
		If sct_effic < 0 Then sct_effic = 100#
		sector_strength = 1#
		If sct_type = "^" Then
			sector_strength = 2#
		End If
		
		d_dstr = SectorDefenseStrength(sct_type)
		
		DefenseStrength = sector_strength + (d_dstr - sector_strength) * (sct_effic / 100#)
		
	End Function
	
	Public Function EnemyMil(ByRef Sector As String) As Short
		Dim sx As Short
		Dim sy As Short
		
		If Not ParseSectors(sx, sy, Sector) Then
			EnemyMil = -1
			Exit Function
		End If
		
		rsEnemySect.Seek("=", sx, sy)
		If rsEnemySect.NoMatch Then
			EnemyMil = -1
			Exit Function
		End If
		EnemyMil = Val(rsEnemySect.Fields("mil").Value)
	End Function
	
	Public Function VerifySubDirectory(ByRef strDirect As Object, Optional ByRef Create As Boolean = False) As Boolean
		
		' Display the names in Path that represent directories.
		Dim strPath As String
		Dim strFile As String
		' Dim found As Boolean    8/2003 efj  removed
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(Create) Then
			Create = False
		End If
		
		strPath = My.Application.Info.DirectoryPath & "\" ' Set the path.
		VerifySubDirectory = False
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		strFile = Dir(strPath, FileAttribute.Directory) ' Retrieve the first entry.
		Do While strFile <> vbNullString And Not VerifySubDirectory ' Start the loop.
			If (GetAttr(strPath & strFile) And FileAttribute.Directory) = FileAttribute.Directory Then
				'UPGRADE_WARNING: Couldn't resolve default property of object strDirect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If strDirect = strFile Then
					VerifySubDirectory = True
				End If
			End If ' it represents a directory.
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			strFile = Dir() ' Get next entry.
		Loop 
		
		If (Not VerifySubDirectory) And Create Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strDirect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MkDir(strPath + strDirect)
			On Error GoTo Verify_Exit
			VerifySubDirectory = True
		End If
		
Verify_Exit: 
	End Function
	
	Public Sub GetCurrentStrength(Optional ByRef tsSectors As Integer = 0)
		Dim strCmd As String
		
		If Not frmOptions.bSuppressStrength Then
			strCmd = "strength *"
			If frmOptions.bFullDumpwithoutSea Then
				strCmd = strCmd & " ?des#."
				'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
				If Not IsNothing(tsSectors) Then
					strCmd = strCmd & "&timestamp>" & CStr(tsSectors)
				End If
				'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			ElseIf Not IsNothing(tsSectors) Then 
				strCmd = strCmd & " ?timestamp>" & CStr(tsSectors)
			End If
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		End If
	End Sub
	
	Public Sub LoadPictures()
		On Error GoTo Empire_Error
		picGreenLight = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\light.gif")
		picRedLight = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\light2.gif")
		picBothLights = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\light3.gif")
		Exit Sub
		
Empire_Error: 
		EmpireError("LoadPictures", vbNullString, vbNullString)
	End Sub
	
	Public Function GetWinACEUTC() As String
		GetWinACEUTC = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, tz.Bias, Now), "yyyy/mm/dd hh:mm:ss")
	End Function
	
	Public Function ConvertToLocalTimeFromWinACEUTC(ByRef strTime As String) As Date
		'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		ConvertToLocalTimeFromWinACEUTC = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, -tz.Bias, System.Date.FromOADate(DateValue(strTime).ToOADate + TimeValue(strTime).ToOADate))
	End Function
	
	Public Sub GetSectors(Optional ByRef bFull As Boolean = False)
		Dim strCond As String
		
		If Not bFull Then
			If frmOptions.bFullDumpwithoutSea And bDeity Then
				strCond = " ?des#.&timestamp>" & CStr(tsSectors)
			Else
				strCond = " ?timestamp>" & CStr(tsSectors)
			End If
		Else
			If frmOptions.bFullDumpwithoutSea And bDeity Then
				strCond = " ?des#."
			Else
				strCond = ""
			End If
		End If
		
		If Not VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("xdump sect *" & strCond, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("dump *" & strCond, False)
		End If
	End Sub
	
	Public Sub SubmitTelegramRead(ByRef bRead As Boolean, ByRef bDisplay As Boolean)
		Dim strRead As String
		If bRead Then
			strRead = "y"
		Else
			strRead = "n"
		End If
		If bDeity And (VersionCheck(4, 2, 21) = enumVersion.VER_LESS) Then
			frmEmpCmd.SubmitEmpireCommand("read " & Str(CountryNumber) & " " & strRead, bDisplay)
		Else
			frmEmpCmd.SubmitEmpireCommand("read " & strRead, bDisplay)
		End If
	End Sub
	
	Public Function VersionCheck(ByRef iMajor As Short, ByRef iMinor As Short, ByRef iRevision As Object) As enumVersion
		If Not (rsVersion.BOF And rsVersion.EOF) Then
			If rsVersion.Fields("Major").Value > iMajor Then
				VersionCheck = enumVersion.VER_MORE
				Exit Function
			End If
			If rsVersion.Fields("Major").Value < iMajor Then
				VersionCheck = enumVersion.VER_LESS
				Exit Function
			End If
			If rsVersion.Fields("Minor").Value > iMinor Then
				VersionCheck = enumVersion.VER_MORE
				Exit Function
			End If
			If rsVersion.Fields("Minor").Value < iMinor Then
				VersionCheck = enumVersion.VER_LESS
				Exit Function
			End If
			If rsVersion.Fields("Revision").Value > iRevision Then
				VersionCheck = enumVersion.VER_MORE
				Exit Function
			End If
			If rsVersion.Fields("Revision").Value < iRevision Then
				VersionCheck = enumVersion.VER_LESS
				Exit Function
			End If
			VersionCheck = enumVersion.VER_EXACT
		Else
			VersionCheck = enumVersion.VER_LESS
		End If
	End Function
	
	Public Sub GetOContent()
		Dim strShipList As String
		
		If frmOptions.bOilContentForSeaSectors And Not (rsShip.BOF And rsShip.EOF) Then
			rsShip.MoveFirst()
			While Not rsShip.EOF
				If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_S_OIL, rsShip.Fields("type").Value) Then
					strShipList = strShipList & CStr(rsShip.Fields("id").Value) & "/"
				End If
				rsShip.MoveNext()
			End While
		End If
		If Len(strShipList) > 0 Then
			strShipList = Left(strShipList, Len(strShipList) - 1)
			frmEmpCmd.SubmitEmpireCommand("nav " & strShipList & " vh", False)
		End If
	End Sub
	
	Public Sub GetPlanes(Optional ByRef bFull As Boolean = False)
		Dim strCond As String
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bFull) Then
			bFull = False
		End If
		
		If Not bFull Then
			strCond = " ?timestamp>" & CStr(tsPlane)
		End If
		
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("pdump *" & strCond, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump plane *" & strCond, False)
		End If
	End Sub
	
	Public Sub GetLandUnits(Optional ByRef bFull As Boolean = False)
		Dim strCond As String
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bFull) Then
			bFull = False
		End If
		
		If Not bFull Then
			strCond = " ?timestamp>" & CStr(tsLand)
		End If
		
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("ldump *" & strCond, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump land *" & strCond, False)
		End If
	End Sub
	
	Public Sub GetShips(Optional ByRef bFull As Boolean = False)
		Dim strCond As String
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bFull) Then
			bFull = False
		End If
		
		If Not bFull Then
			strCond = " ?timestamp>" & CStr(tsShip)
		End If
		
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("sdump *" & strCond, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump ship *" & strCond, False)
		End If
	End Sub
	
	Public Sub GetOrders(ByRef bUpdate As Boolean)
		If bUpdate Then
			frmEmpCmd.SubmitEmpireCommand("db1", False)
		End If
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			DeleteAllRecords(rsShipOrders)
			frmEmpCmd.SubmitEmpireCommand("sorder *", False)
			frmEmpCmd.SubmitEmpireCommand("qorder *", False)
		Else
			GetShips(True)
		End If
		If bUpdate Then
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		End If
	End Sub
	
	Public Sub GetNukes(Optional ByRef bFull As Boolean = False)
		Dim strCond As String
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bFull) Then
			bFull = False
		End If
		
		If Not bFull Then
			strCond = " ?timestamp>" & CStr(tsNuke)
		End If
		
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("ndump *" & strCond, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump nuke *" & strCond, False)
		End If
	End Sub
	
	Public Sub GetLost(Optional ByRef bFull As Boolean = False)
		Dim strCond As String
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bFull) Then
			bFull = False
		End If
		
		If Not bFull Then
			strCond = " ?timestamp>" & CStr(tsLost)
		End If
		
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("lost *" & strCond, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump lost *" & strCond, False)
		End If
	End Sub
	
	Public Sub GetNation()
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("nation", False)
		ElseIf VersionCheck(4, 3, 4) = enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then 
			frmEmpCmd.SubmitEmpireCommand("xdump nat *", False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump coun *", False)
		End If
	End Sub
	
	Public Sub GetCountryList()
		If VersionCheck(4, 3, 0) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("relation", False)
		ElseIf VersionCheck(4, 3, 4) = enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then 
			frmEmpCmd.SubmitEmpireCommand("xdump cou *", False)
			frmEmpCmd.SubmitEmpireCommand("relation", False)
		ElseIf frmOptions.bSPAtlantis Then 
			frmEmpCmd.SubmitEmpireCommand("xdump nat *", False)
			frmEmpCmd.SubmitEmpireCommand("relation", False)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump nat *", False)
		End If
	End Sub
	
	Public Sub GetUpdate(ByRef bAddCommandList As Boolean)
		
		If VersionCheck(4, 3, 10) = enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("update", bAddCommandList)
		Else
			frmEmpCmd.SubmitEmpireCommand("xdump updates 0", bAddCommandList)
		End If
	End Sub
	
	Public Sub StartUpdateTimer()
		If VersionCheck(4, 3, 10) <> enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("xdump game *", False)
		End If
		GetUpdate(False)
	End Sub
	Public Sub UpdateCommands()
		'UPGRADE_WARNING: Timer property UpdateTimer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
		frmDrawMap.UpdateTimer.Interval = 0
		If Len(BatchScriptUpdate) > 0 Then
			frmScript.ExectuteBatchScript(BatchScriptUpdate)
			BatchScriptUpdate = vbNullString
		End If
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetUpdate(False)
		GetSectors(True)
		GetCurrentStrength()
		GetOContent()
		GetPlanes(True)
		GetLandUnits(True)
		GetShips(True)
		GetNukes(True)
		GetLost(True)
		If frmEmpCmd.bAutoUpdate Then
			GetNation()
		End If
		frmEmpCmd.SubmitEmpireCommand("db2", False)
	End Sub
	
	Function IsArrayEmpty(ByRef varArray As Object) As Boolean
		' Determines whether an array contains any elements.
		' Returns False if it does contain elements, True
		' if it does not.
		
		Dim lngUBound As Integer
		
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
	
	Public Sub FlashLog(ByRef strText As String)
		Dim strlog As String
		
		If Not frmOptions.bFLashLogging Then
			Exit Sub
		End If
		
		If VerifySubDirectory("Game Logs", True) And FlashLogFileNumber = -1 Then
			strlog = My.Application.Info.DirectoryPath & "\Game Logs\" & GameName & " Flash " & CStr(Year(Now)) & "-" & VB6.Format(Month(Now), "00") & "-" & VB6.Format(VB.Day(Now), "00") & ".log"
			
			'open the new log
			FlashLogFileNumber = FreeFile
			'if it already exists, append to it
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Len(Dir(strlog)) > 0 Then
				FileOpen(FlashLogFileNumber, strlog, OpenMode.Append)
			Else
				FileOpen(FlashLogFileNumber, strlog, OpenMode.Output)
			End If
			PrintLine(FlashLogFileNumber, "+++ Flash Log opened at " & VB6.Format(Now, "hh:mm:ss") & " +++")
		End If
		
		If FlashLogFileNumber <> -1 Then
			PrintLine(FlashLogFileNumber, strText)
		End If
	End Sub
	
	Public Sub CloseFlashLog()
		If FlashLogFileNumber <> -1 Then
			PrintLine(FlashLogFileNumber, "+++ Flash Log closed at " & VB6.Format(Now, "hh:mm:ss") & " +++")
			FileClose(FlashLogFileNumber)
		End If
	End Sub
End Module