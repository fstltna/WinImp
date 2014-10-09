Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptCopyGameDB
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'112004 rjk: Fixed LandUnits
	'030405 rjk: Fixed RestoreDummyProductionSector, had references to rsSectors.
	'160705 rjk: Added ability to get OContent for Sea Sectors
	'180206 rjk: Replace ndump, ldump, sdump, pdump with GetNukes, GetLandUnits, GetShips and GetPlanes.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'090607 rjk: Added Land Unit counts and Plane counts for Server 4.3.6.
	'120108 rjk: Removed unused local variable strCmd from CopyGameInfo.
	'            Removed unused local variables iSecX and iSecY from CreatePlanes, CreateLandUnits
	'            and CreateShips.
	
	
	'UPGRADE_WARNING: Event cbGameName.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cbGameName.Change was upgraded to cbGameName.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cbGameName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbGameName.TextChanged
		If Len(Trim(cbGameName.Text)) = 0 Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cbGameName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbGameName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbGameName.SelectedIndexChanged
		If Len(Trim(cbGameName.Text)) = 0 Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		CopyGameInfo()
		Me.Close()
	End Sub
	
	Private Sub CopyGameInfo()
		Dim iOffsetX As Short
		Dim iOffsetY As Short
		Dim dbEmpireCopy As DAO.Database
		Dim DatabaseName As String
		
		iOffsetX = Val(txtOffsetX.Text)
		iOffsetY = Val(txtOffsetY.Text)
		
		DatabaseName = Trim(cbGameName.Text)
		
		'Allow the user to specifiy the path
		If InStr(DatabaseName, "\") = 0 Then
			DatabaseName = My.Application.Info.DirectoryPath & "\Games\" & DatabaseName & ".mdb"
		End If
		
		dbEmpireCopy = DAODBEngine_definst.OpenDatabase(DatabaseName, False, False)
		
		'Copy sectors and bmap
		If chkItems(5).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			CopySectors(dbEmpireCopy, iOffsetX, iOffsetY)
		End If
		
		If chkItems(10).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			UpdateNation(dbEmpireCopy, iOffsetX, iOffsetY)
		End If
		
		'Copy Ships
		If chkItems(6).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			CreateHarbor()
			CreateShips(dbEmpireCopy, iOffsetX, iOffsetY)
			CleanupShips()
		End If
		
		'Copy Plane
		If chkItems(7).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			CreateAirport()
			CreatePlanes(dbEmpireCopy, iOffsetX, iOffsetY)
			CleanupPlanes()
		End If
		
		'Copy Land Units
		If chkItems(8).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			CreateHeadquarters()
			CreateLandUnits(dbEmpireCopy, iOffsetX, iOffsetY)
			CleanupLandUnits()
		End If
		
		If chkItems(6).CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkItems(7).CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkItems(8).CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkItems(9).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			RestoreDummyProductionSector(dbEmpireCopy, iOffsetX, iOffsetY)
		End If
		
		dbEmpireCopy.Close()
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		GetCurrentStrength(tsSectors)
		GetOContent()
		
		If chkItems(6).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			GetShips()
		End If
		If chkItems(7).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			GetLandUnits()
		End If
		If chkItems(8).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			GetPlanes()
		End If
		If chkItems(9).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			GetNukes()
		End If
		frmEmpCmd.SubmitEmpireCommand("db2", False)
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptCopyGameDB.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptCopyGameDB_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim FileName As String ' Walking filename variable.
		Dim strDirectoryName As String
		
		strDirectoryName = My.Application.Info.DirectoryPath & "\Games\*.mdb"
		
		' Search through this directory and sum file sizes.
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(strDirectoryName, vbNormal Or FileAttribute.Archive)
		While Len(FileName) <> 0
			cbGameName.Items.Add(VB.Left(FileName, Len(FileName) - 4))
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir() ' Get next file.
		End While
	End Sub
	
	Private Sub CopySectors(ByRef dbEmpireCopy As DAO.Database, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		'UPGRADE_WARNING: Arrays in structure rsB may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rsB As DAO.Recordset
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim iSecX As Short
		Dim iSecY As Short
		Dim strCmd As String
		Dim strDes As String
		Dim strCom As String
		Dim n As Short
		
		rs = dbEmpireCopy.OpenRecordset("Sectors")
		rs.Index = "PrimaryKey"
		If Not (rs.BOF And rs.EOF) Then
			rs.MoveFirst()
		End If
		
		While Not rs.EOF
			CopySectorAndOwner(rs, iOffsetX, iOffsetY)
			rs.MoveNext()
		End While
		
		rsB = dbEmpireCopy.OpenRecordset("Bmap")
		rsB.Index = "PrimaryKey"
		If Not (rsB.BOF And rsB.EOF) Then
			rsB.MoveFirst()
		End If
		
		While Not rsB.EOF
			iSecX = rsB.Fields("x").Value
			iSecY = rsB.Fields("y").Value
			rs.Seek("=", iSecX, iSecY)
			If rs.NoMatch Then
				OffsetSectorLocation(iSecX, iSecY, iOffsetX, iOffsetY)
				strDes = rsB.Fields("des").Value
				strCmd = "des " & SectorString(iSecX, iSecY) & " " & strDes
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			End If
			rsB.MoveNext()
		End While
		
		'wipe out commodities from sea sectors
		strCom = "mil  gun  shellfood civ  uw   iron lcm  hcm  oil  dust pet  rad  bar  "
		For n = 0 To 13
			strCmd = Trim(Mid(strCom, 5 * n + 1, 5))
			strCmd = "give " & strCmd & " * ?des=.&" & strCmd & ">0 -9999"
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		Next n
		
		rs.Close()
		rsB.Close()
	End Sub
	
	Private Sub CreateHarbor()
		'Setup a dummy harbor
		CreateDummyProductionSector("h")
		'scuttle any ships that already exist
		ScuttleUnits("ship")
	End Sub
	
	Private Sub CreateHeadquarters()
		'Setup a dummy headquarters
		CreateDummyProductionSector("!")
		'scuttle any land units that already exist
		ScuttleUnits("land")
	End Sub
	
	
	Private Sub CreateAirport()
		'Setup a dummy airport
		CreateDummyProductionSector("*")
		'scuttle any planes that already exist
		ScuttleUnits("plane")
	End Sub
	
	Private Sub CreateShips(ByRef dbEmpireCopy As DAO.Database, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		
		rs = dbEmpireCopy.OpenRecordset("ship")
		rs.Index = "PrimaryKey"
		
		If Not (rs.BOF And rs.EOF) Then
			If VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
				ComputeUnitCountsForShips()
			End If
			CreateUnits(rs, "ship", "fb")
			AdjustShips(rs, iOffsetX, iOffsetY)
		End If
		rs.Close()
	End Sub
	
	Private Sub CreateLandUnits(ByRef dbEmpireCopy As DAO.Database, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		
		rs = dbEmpireCopy.OpenRecordset("land")
		rs.Index = "PrimaryKey"
		
		If Not (rs.BOF And rs.EOF) Then
			If VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
				ComputeUnitCountsForLandUnits()
			End If
			CreateUnits(rs, "land", "cav", "unit")
			AdjustLandUnits(rs, iOffsetX, iOffsetY)
		End If
		rs.Close()
	End Sub
	
	Private Sub CreatePlanes(ByRef dbEmpireCopy As DAO.Database, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		
		rs = dbEmpireCopy.OpenRecordset("plane")
		rs.Index = "PrimaryKey"
		
		If Not (rs.BOF And rs.EOF) Then
			CreateUnits(rs, "plane", "f1")
			AdjustPlanes(rs, iOffsetX, iOffsetY)
		End If
	End Sub
	
	Private Sub AdjustShips(ByRef rs As DAO.Recordset, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		Dim iSecX As Short
		Dim iSecY As Short
		Dim strCmd As String
		
		rs.MoveFirst()
		While Not rs.EOF
			iSecX = rs.Fields("x").Value
			iSecY = rs.Fields("y").Value
			OffsetSectorLocation(iSecX, iSecY, iOffsetX, iOffsetY)
			strCmd = " L " & SectorString(iSecX, iSecY)
			If rs.Fields("tech").Value > 0 Then
				strCmd = strCmd & " T " & CStr(rs.Fields("tech").Value)
			End If
			If rs.Fields("mob").Value > 0 Then
				strCmd = strCmd & " M " & CStr(rs.Fields("mob").Value)
			End If
			If rs.Fields("land").Value > 0 Then
				strCmd = strCmd & " Y " & CStr(rs.Fields("land").Value)
			End If
			If rs.Fields("eff").Value > 0 Then
				strCmd = strCmd & " E " & CStr(rs.Fields("eff").Value)
			End If
			If rs.Fields("pln").Value > 0 Then
				strCmd = strCmd & " P " & CStr(rs.Fields("pln").Value)
			End If
			If rs.Fields("xl").Value > 0 Then
				strCmd = strCmd & " X " & CStr(rs.Fields("xl").Value)
			End If
			If rs.Fields("he").Value > 0 Then
				strCmd = strCmd & " H " & CStr(rs.Fields("he").Value)
			End If
			If rs.Fields("mil").Value > 0 Then
				strCmd = strCmd & " m " & CStr(rs.Fields("mil").Value)
			End If
			If rs.Fields("civ").Value > 0 Then
				strCmd = strCmd & " c " & CStr(rs.Fields("civ").Value)
			End If
			If rs.Fields("uw").Value > 0 Then
				strCmd = strCmd & " u " & CStr(rs.Fields("uw").Value)
			End If
			If rs.Fields("food").Value > 0 Then
				strCmd = strCmd & " f " & CStr(rs.Fields("food").Value)
			End If
			If rs.Fields("gun").Value > 0 Then
				strCmd = strCmd & " g " & CStr(rs.Fields("gun").Value)
			End If
			If rs.Fields("shell").Value > 0 Then
				strCmd = strCmd & " s " & CStr(rs.Fields("shell").Value)
			End If
			If rs.Fields("iron").Value > 0 Then
				strCmd = strCmd & " i " & CStr(rs.Fields("iron").Value)
			End If
			If rs.Fields("dust").Value > 0 Then
				strCmd = strCmd & " d " & CStr(rs.Fields("dust").Value)
			End If
			If rs.Fields("oil").Value > 0 Then
				strCmd = strCmd & " o " & CStr(rs.Fields("oil").Value)
			End If
			If rs.Fields("petrol").Value > 0 Then
				strCmd = strCmd & " p " & CStr(rs.Fields("petrol").Value)
			End If
			If rs.Fields("lcm").Value > 0 Then
				strCmd = strCmd & " l " & CStr(rs.Fields("lcm").Value)
			End If
			If rs.Fields("hcm").Value > 0 Then
				strCmd = strCmd & " h " & CStr(rs.Fields("hcm").Value)
			End If
			If rs.Fields("rad").Value > 0 Then
				strCmd = strCmd & " r " & CStr(rs.Fields("rad").Value)
			End If
			strCmd = "edit ship " & CStr(rs.Fields("id").Value) & strCmd
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			rs.MoveNext()
		End While
	End Sub
	
	Private Sub AdjustLandUnits(ByRef rs As DAO.Recordset, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		Dim iSecX As Short
		Dim iSecY As Short
		Dim strCmd As String
		
		rs.MoveFirst()
		
		While Not rs.EOF
			iSecX = rs.Fields("x").Value
			iSecY = rs.Fields("y").Value
			OffsetSectorLocation(iSecX, iSecY, iOffsetX, iOffsetY)
			strCmd = " L " & SectorString(iSecX, iSecY)
			strCmd = strCmd & " t " & CStr(rs.Fields("tech").Value)
			strCmd = strCmd & " M " & CStr(rs.Fields("mob").Value)
			strCmd = strCmd & " e " & CStr(rs.Fields("eff").Value)
			strCmd = strCmd & " F " & CStr(rs.Fields("fort").Value)
			strCmd = strCmd & " B " & CStr(rs.Fields("fuel").Value)
			strCmd = strCmd & " S " & CStr(rs.Fields("ship").Value)
			strCmd = strCmd & " Y " & CStr(rs.Fields("land").Value)
			strCmd = strCmd & " X " & CStr(rs.Fields("xl").Value)
			strCmd = strCmd & " m " & CStr(rs.Fields("mil").Value)
			strCmd = strCmd & " f " & CStr(rs.Fields("food").Value)
			strCmd = strCmd & " g " & CStr(rs.Fields("gun").Value)
			strCmd = strCmd & " s " & CStr(rs.Fields("shell").Value)
			strCmd = strCmd & " i " & CStr(rs.Fields("iron").Value)
			strCmd = strCmd & " d " & CStr(rs.Fields("dust").Value)
			strCmd = strCmd & " o " & CStr(rs.Fields("oil").Value)
			strCmd = strCmd & " p " & CStr(rs.Fields("petrol").Value)
			strCmd = strCmd & " l " & CStr(rs.Fields("lcm").Value)
			strCmd = strCmd & " h " & CStr(rs.Fields("hcm").Value)
			strCmd = strCmd & " r " & CStr(rs.Fields("rad").Value)
			strCmd = "edit unit " & CStr(rs.Fields("id").Value) & strCmd
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			rs.MoveNext()
		End While
	End Sub
	
	Private Sub AdjustPlanes(ByRef rs As DAO.Recordset, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		Dim iSecX As Short
		Dim iSecY As Short
		Dim strCmd As String
		
		rs.MoveFirst()
		While Not rs.EOF
			iSecX = rs.Fields("x").Value
			iSecY = rs.Fields("y").Value
			OffsetSectorLocation(iSecX, iSecY, iOffsetX, iOffsetY)
			strCmd = " l " & SectorString(iSecX, iSecY)
			strCmd = strCmd & " t " & CStr(rs.Fields("tech").Value)
			strCmd = strCmd & " m " & CStr(rs.Fields("mob").Value)
			strCmd = strCmd & " e " & CStr(rs.Fields("eff").Value)
			strCmd = strCmd & " a " & CStr(rs.Fields("att").Value)
			strCmd = strCmd & " d " & CStr(rs.Fields("def").Value)
			strCmd = strCmd & " r " & CStr(rs.Fields("range").Value)
			strCmd = strCmd & " s " & CStr(rs.Fields("ship").Value)
			strCmd = strCmd & " l " & CStr(rs.Fields("land").Value)
			strCmd = "edit plane " & CStr(rs.Fields("id").Value) & strCmd
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			rs.MoveNext()
		End While
	End Sub
	
	Private Sub CleanupShips()
		'now scrap the filler Ships
		frmEmpCmd.SubmitEmpireCommand("sc1", False)
		frmEmpCmd.SubmitEmpireCommand("scuttle ship * ?type=fb&eff=20&xloc=0&yloc=0&own=98", False)
	End Sub
	
	Private Sub CleanupLandUnits()
		'now scrap the filler Land Units
		frmEmpCmd.SubmitEmpireCommand("sc1", False)
		frmEmpCmd.SubmitEmpireCommand("scuttle land * ?type=cav&eff=10&xloc=0&yloc=0&own=98", False)
	End Sub
	
	Private Sub CleanupPlanes()
		'now scrap the filler Planes
		frmEmpCmd.SubmitEmpireCommand("sc1", False)
		frmEmpCmd.SubmitEmpireCommand("scuttle plane * ?type=f1&eff=10&xloc=0&yloc=0&own=98", False)
	End Sub
	
	Private Sub CreateDummyProductionSector(ByRef strDes As String)
		Dim strCmd As String
		
		strCmd = "designate 0,0 " & strDes
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "setsector eff 0,0 100"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "setsector work 0,0 100"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "setsector oldown 0,0 0"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "setsector own 0,0 0"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
	End Sub
	
	Private Sub RestoreDummyProductionSector(ByRef dbEmpireCopy As DAO.Database, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim strCmd As String
		Dim strCom As String
		Dim strComm As String
		Dim iIndex As Short
		Dim iSecDistX As Short
		Dim iSecDistY As Short
		
		rs = dbEmpireCopy.OpenRecordset("Sectors")
		rs.Index = "PrimaryKey"
		rs.Seek("=", -iOffsetX, -iOffsetY)
		If rs.NoMatch Then
			strCmd = "designate 0,0 ."
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		Else
			strCmd = "setsector oldown 0,0 " & CStr(rs.Fields("Country").Value)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "setsector own 0,0 " & CStr(rs.Fields("Country").Value)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "setsector mob 0,0 " & CStr(rs.Fields("mob").Value)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "setsector work 0,0 " & CStr(rs.Fields("work").Value)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "setsector avail 0,0 " & CStr(rs.Fields("avail").Value)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			frmEmpCmd.SubmitEmpireCommand("designate 0,0 " + rs.Fields("des").Value, False)
			strCom = "mil  gun  shelllcm  hcm  "
			For iIndex = 0 To 4
				strComm = Trim(Mid(strCom, 5 * iIndex + 1, 5))
				strCmd = "give " & strComm & " 0,0 ?" & strComm & ">0 -9999"
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				If rs.Fields(strComm).Value > 0 Then
					strCmd = "give " & strComm & " 0,0 " & CStr(rs.Fields(strComm).Value)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
			Next iIndex
			strCom = "eff  work avail"
			For iIndex = 0 To 2
				strComm = Trim(Mid(strCom, 5 * iIndex + 1, 5))
				strCmd = "setsector " & strComm & " 0,0 -9999"
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
				If rs.Fields(strComm).Value > 0 Then
					strCmd = "setsector " & strComm & " 0,0 " & CStr(rs.Fields(strComm).Value)
					frmEmpCmd.SubmitEmpireCommand(strCmd, False)
				End If
			Next iIndex
			iSecDistX = rs.Fields("dist_x").Value
			iSecDistY = rs.Fields("dist_y").Value
			strCmd = "edit land 0,0 D " & SectorString(iSecDistX, iSecDistY)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			If rs.Fields("off").Value = 0 Then
				frmEmpCmd.SubmitEmpireCommand("start 0,0", False)
			Else
				frmEmpCmd.SubmitEmpireCommand("stop 0,0", False)
			End If
			SetThreshold("0,0", "c", rs)
			SetThreshold("0,0", "m", rs)
			SetThreshold("0,0", "u", rs)
			SetThreshold("0,0", "s", rs)
			SetThreshold("0,0", "g", rs)
			SetThreshold("0,0", "l", rs)
			SetThreshold("0,0", "h", rs)
			SetThreshold("0,0", "o", rs)
			SetThreshold("0,0", "p", rs)
			SetThreshold("0,0", "f", rs)
			SetThreshold("0,0", "d", rs)
			SetThreshold("0,0", "b", rs)
			SetThreshold("0,0", "i", rs)
			SetThreshold("0,0", "r", rs)
			strCmd = "terr 0,0 " & CStr(rs.Fields("terr").Value) & " 0"
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "terr 0,0 " & CStr(rs.Fields("terr1").Value) & " 1"
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "terr 0,0 " & CStr(rs.Fields("terr2").Value) & " 2"
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "terr 0,0 " & CStr(rs.Fields("terr3").Value) & " 3"
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		End If
		rs.Close()
	End Sub
	
	Private Sub UpdateNation(ByRef dbEmpireCopy As DAO.Database, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim strCmd As String
		Dim CountryNumber As Short
		Dim iSecX As Short
		Dim iSecY As Short
		
		rs = dbEmpireCopy.OpenRecordset("Nations")
		If Not (rs.BOF And rs.EOF) Then
			rs.MoveFirst()
			CountryNumber = Val(rs.Fields("Number").Value)
			rs.Close()
		Else
			rs.Close()
			Exit Sub
		End If
		
		rs = dbEmpireCopy.OpenRecordset("Nation")
		If Not (rs.BOF And rs.EOF) Then
			rs.MoveFirst()
			strCmd = "add " & CStr(CountryNumber) & " " & CStr(CountryNumber) & " " & CStr(CountryNumber) & " new i"
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			strCmd = "edit country " & CStr(CountryNumber)
			strCmd = strCmd & " E " & CStr(rs.Fields("Education").Value)
			strCmd = strCmd & " T " & CStr(rs.Fields("TechLevel").Value)
			strCmd = strCmd & " H " & CStr(rs.Fields("Happiness").Value)
			strCmd = strCmd & " R " & CStr(rs.Fields("Research").Value)
			If ParseSectors(iSecX, iSecY, rs.Fields("CapSect").Value) Then
				OffsetSectorLocation(iSecX, iSecY, iOffsetX, iOffsetY)
				strCmd = strCmd & " c " & CStr(iSecX) & "," & CStr(iSecY)
			End If
			
			strCmd = strCmd & " m " & CStr(rs.Fields("MilReserves").Value)
			strCmd = strCmd & " o " & CStr(iOffsetX) & "," & CStr(iOffsetY)
			strCmd = strCmd & " s 5 "
			strCmd = strCmd & " M 25000 "
			strCmd = strCmd & " b 640 "
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		End If
		rs.Close()
	End Sub
	
	Private Sub CreateUnits(ByRef rs As DAO.Recordset, ByRef strUnit As String, ByRef strDefaultUnit As String, Optional ByRef strEdit As String = "")
		Dim strCmd As String
		Dim nUnit As Short
		Dim iDefaultMil As Short
		Dim iDefaultGuns As Short
		Dim iDefaultAvail As Short
		Dim iDefaultHcms As Short
		Dim iDefaultLcms As Short
		
		If Len(strEdit) = 0 Then
			strEdit = strUnit
		End If
		
		If strUnit = "unit" Then
			rsBuild.Seek("=", "l", strDefaultUnit)
		Else
			rsBuild.Seek("=", VB.Left(strUnit, 1), strDefaultUnit)
		End If
		If rsBuild.NoMatch Then
			EmpireError("CreateUnits", "Can not find default unit characteristics", strDefaultUnit)
			Exit Sub
		End If
		iDefaultMil = rsBuild.Fields("mil").Value
		iDefaultGuns = rsBuild.Fields("gun").Value
		iDefaultAvail = rsBuild.Fields("avail").Value
		iDefaultHcms = rsBuild.Fields("hcm").Value
		iDefaultLcms = rsBuild.Fields("lcm").Value
		
		rs.MoveFirst()
		
		nUnit = 0
		While Not rs.EOF
			'need a filler Unit
			If nUnit < rs.Fields("id").Value Then
				SupplySector(iDefaultMil, iDefaultGuns, iDefaultAvail, iDefaultHcms, iDefaultLcms)
				strCmd = "build " & strUnit & " 0,0 " & strDefaultUnit
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
				strCmd = "edit " & strEdit & " " & CStr(nUnit) & " O 98"
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
				'need to build the Unit
			ElseIf nUnit = rs.Fields("id").Value Then 
				If strUnit = "unit" Then
					rsBuild.Seek("=", "l", rs.Fields("type"))
				Else
					rsBuild.Seek("=", VB.Left(strUnit, 1), rs.Fields("type"))
				End If
				If rsBuild.NoMatch Then
					EmpireError("CreateUnits", "Can not find unit characteristics", rs.Fields("type").Value)
					Exit Sub
				End If
				SupplySector(CShort(rsBuild.Fields("mil").Value), CShort(rsBuild.Fields("gun").Value), CShort(rsBuild.Fields("avail").Value), CShort(rsBuild.Fields("hcm").Value), CShort(rsBuild.Fields("lcm").Value))
				strCmd = "build " & strUnit & " 0,0 " + rs.Fields("type").Value + " 1 " + CStr(rs.Fields("tech").Value)
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
				strCmd = "edit " & strEdit & " " & CStr(nUnit) & " O " & CStr(rs.Fields("country").Value)
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
				rs.MoveNext()
			End If
			nUnit = nUnit + 1
		End While
	End Sub
	
	Private Sub SupplySector(ByRef iMil As Short, ByRef iGuns As Short, ByRef iAvail As Short, ByRef iHcms As Short, ByRef iLcms As Short)
		If iMil > 0 Then
			frmEmpCmd.SubmitEmpireCommand("give mil 0,0 " & CStr(iMil), False)
		End If
		If iGuns > 0 Then
			frmEmpCmd.SubmitEmpireCommand("give gun 0,0 " & CStr(iGuns), False)
		End If
		If iHcms > 0 Then
			frmEmpCmd.SubmitEmpireCommand("give hcm 0,0 " & CStr(iHcms), False)
		End If
		If iLcms > 0 Then
			frmEmpCmd.SubmitEmpireCommand("give lcm 0,0 " & CStr(iLcms), False)
		End If
		If iAvail > 0 Then
			frmEmpCmd.SubmitEmpireCommand("setsector avail 0,0 " & CStr(iAvail), False)
		End If
	End Sub
	
	Private Sub ScuttleUnits(ByRef strUnit As String)
		Dim strCmd As String
		
		strCmd = "sc1"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		
		strCmd = "scuttle " & strUnit & " *"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
	End Sub
	
	Private Sub CopySectorAndOwner(ByRef rs As DAO.Recordset, ByRef iOffsetX As Short, ByRef iOffsetY As Short)
		Dim vhld() As Object
		Dim iSecX As Short
		Dim iSecY As Short
		Dim iSecDistX As Short
		Dim iSecDistY As Short
		Dim strCmd As String
		Dim strSector As String
		Dim n As Short
		Dim secowner As Short
		
		ReDim vhld(rs.Fields.Count)
		
		For n = 0 To rs.Fields.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object vhld(rs.Fields().OrdinalPosition). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vhld(rs.Fields(n).OrdinalPosition) = rs.Fields(n).Value
		Next n
		iSecX = rs.Fields("x").Value
		iSecY = rs.Fields("y").Value
		OffsetSectorLocation(iSecX, iSecY, iOffsetX, iOffsetY)
		CopySector(vhld, iSecX, iSecY, 3)
		
		strSector = SectorString(iSecX, iSecY)
		secowner = rs.Fields("country").Value
		strCmd = "setsector own " & strSector & " " & CStr(secowner)
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		'    secowner = rs.Fields("oldown")
		strCmd = "setsector oldown " & strSector & " " & CStr(secowner)
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		iSecDistX = rs.Fields("dist_x").Value
		iSecDistY = rs.Fields("dist_y").Value
		OffsetSectorLocation(iSecDistX, iSecDistY, iOffsetX, iOffsetY)
		strCmd = "edit land " & strSector & " D " & SectorString(iSecDistX, iSecDistY)
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		If rs.Fields("off").Value = 0 Then
			frmEmpCmd.SubmitEmpireCommand("start " & strSector, False)
		Else
			frmEmpCmd.SubmitEmpireCommand("stop " & strSector, False)
		End If
		SetThreshold(strSector, "c", rs)
		SetThreshold(strSector, "m", rs)
		SetThreshold(strSector, "u", rs)
		SetThreshold(strSector, "s", rs)
		SetThreshold(strSector, "g", rs)
		SetThreshold(strSector, "l", rs)
		SetThreshold(strSector, "h", rs)
		SetThreshold(strSector, "o", rs)
		SetThreshold(strSector, "p", rs)
		SetThreshold(strSector, "f", rs)
		SetThreshold(strSector, "d", rs)
		SetThreshold(strSector, "b", rs)
		SetThreshold(strSector, "i", rs)
		SetThreshold(strSector, "r", rs)
		strCmd = "terr " & strSector & " " & CStr(rs.Fields("terr").Value) & " 0"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "terr " & strSector & " " & CStr(rs.Fields("terr1").Value) & " 1"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "terr " & strSector & " " & CStr(rs.Fields("terr2").Value) & " 2"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		strCmd = "terr " & strSector & " " & CStr(rs.Fields("terr3").Value) & " 3"
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
	End Sub
	
	Private Sub SetThreshold(ByRef strSector As String, ByRef strComm As String, ByRef rs As DAO.Recordset)
		Dim strCmd As String
		
		strCmd = "threshold " & strComm & " " & strSector & " " & CStr(rs.Fields(strComm & "_dist").Value)
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		If rs.Fields(strComm & "_del").Value = "." Then
			strCmd = "deliver " & strComm & " " & strSector & " " & CStr(rs.Fields(strComm & "_cut").Value) & " h"
		Else
			strCmd = "deliver " & strComm & " " & strSector & " " & CStr(rs.Fields(strComm & "_cut").Value) & " " & CStr(rs.Fields(strComm & "_del").Value)
		End If
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
	End Sub
End Class