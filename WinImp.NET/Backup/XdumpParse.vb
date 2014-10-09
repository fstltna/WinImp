Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module XdumpParse
	
	'Change Log:
	'270804 rjk: Created to Parse xdump output.
	'111104 rjk: Added the parsing order info from the xdump ship.
	'200505 rjk: Switched mobquota to mquota to match in server for 4.2.21.
	'110805 rjk: Fixed land chr and plane chr to account for the removal of I_NONE item.
	'100905 rjk: Added maxpop to chr sect.
	'180206 rjk: Added AddMissions, AddOrders, DirectionString, MaxPopulation, GetType,
	'            UpdateTechLevelForShow, GetTechLevelForShow, RequestConfigurationTables,
	'            RequestMetaTables, ParseXDTableName, ParseXDMeta, ParseXDLost,
	'            ParseXDCountries, ParseXDNation and ParseXDMetaHeader functions.
	'            Modify the rest of the functions to support the switch to meta data
	'            and the changes to the columns for 4.3.0.
	'120306 rjk: Fixed Speed and Firing Ranges.  Fixed land description.
	'            Fixed Unit Counts for Ships.  Removed FillUnitGrid from LandChr.
	'            Fixed the 0 item counts from the ship capabilities.
	'200306 rjk: Flush the existing Build records before reload tables.
	'210306 rjk: Flush the existing unit and sector records before reload tables.
	'230306 rjk: Fix so meta tables are requested always requested independent of AutoUpdate.
	'            Add attack and defense values to ships, land units and planes.
	'090406 rjk: Fixed no current record when loading the xdump ver.
	'            Occurs when a breaking a new country.
	'200406 rjk: Fixed ostr and dstr to use Val for SectorChr.
	'            Fixed level in Product parsing.
	'            Fixed Sector (defense) is it is incorrectly overwrite by Product (level).
	'230406 rjk: Finished the Unit Grid fields.
	'300406 rjk: Updated the Show Tool with new values.
	'            Fixed Exit sub to restore the headers.
	'050506 rjk: Added Theme game for South Pacific: Search for Atlantis
	'            Fixed the Infrastructure parsing.
	'060506 rjk: Added Nuke changes for 4.3.3 server.
	'140506 rjk: Added Avail calculation for Planes/Land Units/Ships.
	'210506 rjk: Fixed Nuke Type to deal with only field size of 5 characters.
	'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun.
	'            Switched NSC_STRINGY a variable and extract the value from meta-type table.
	'            SP: Atlantis also uses changes for xdump nat and coun.
	'230506 rjk: Incorporate 4.3.4 server changes for fleet/army/wing.
	'250506 rjk: Added Four SP: Atlantis infrastructures to the database,
	'            use runway -> min, radar -> gld, fort -> fert nad navigate -> oil.
	'            Change MaxPopulation to use shell industry for SP: Atlantis
	'270506 rjk: Add timestamp checks for Lost processing to prevent a lost item
	'            from deleting regained item.
	'290506 rjk: Add timestamp storage to the units and sector files.
	'100606 rjk: Move product effic from ProductChr to SectorChr for server 4.3.6.
	'            Added off (new budget functions) for units for server 4.3.6.
	'120606 rjk: Fixed missions types.  Added Symbol table support.
	'150606 rjk: Convert iNscSTRINGY to Symbol table.  Convert GetNation to use
	'            XD_LEVEL symbol table.
	'200606 rjk: Added sector-navigation, ship-chr-flags symbol tables.
	'            Added processing deity terr for 4.3.6 servers.
	'            Incorporate mobility (mob0, mob1, remove mcst) for 4.3.6 servers.
	'            Added support for nav field in the sector chr.
	'            Added support the canal capability flags in the ship flags.
	'            Added support for the new enable field for infrastructure for 4.3.6 servers.
	'            Updated symbol type values to support Long (ship flags to high for integers).
	'250606 rjk: Added the storing of off for units.
	'030706 rjk: Fixed ParseXDNuke to add the unit grid instead of the sector panel.
	'            The nukes have been moved from the text in the sector panel to be a
	'            grid tab.
	'090706 rjk: Added elev and track for HeavyMetal theme game.
	'120706 rjk: Add some support RAILWAYS in HeavyMetal theme game.
	'130706 rjk: Fixed support for ASW planes, sign problem with &H8000.
	'140706 rjk: Add some support RAILWAYS in HeavyMetal theme game,
	'            done with the adjacent sectors.
	'150706 rjk: Added AUTO_POWER for HeavyMeta theme game.
	'160706 rjk: Fixed the ship RFR (real firing range) for 4.3.X servers.
	'            Disable error for missing characteristics for ships and lands.
	'            Can occur when you tech falls back.
	'210706 rjk: Fixed UpdateRailTrack to deal with 0 highway sectors.
	'            Fixed Production because the level type was set incorrectly in ParseXDProduct.
	'041006 rjk: Added previous location checks to planes, land units and nukes for map update.
	'181006 rjk: Added OldEnemyUnit check.
	'241006 rjk: Fixed "wing" for planes.
	'091206 rjk: Fixed Sector Chr to load the packing factors for mobility calculations.
	'161206 rjk: Automatically detect the change to XDump capable server and
	'            request the meta tables.
	'221206 rjk: Fixed ParseXDVersion and ParseXDSectorChr to use Val for floating
	'            points number, fixes regional problems.
	'040307 rjk: Added SECT_MOB_ACCESS, UNIT_MOB_ACCESS and unit_mob_neg_factor
	'            to ParseXDVersion for server 3.4.10.
	'200307 rjk: Change AUTO_POWER from a game option for Heavy Metal to standard option.
	'100607 rjk: Delete enemy intelligence when we own the sector.
	'180707 rjk: Added access fields for ParseXDNation and ParseXDCountries for server
	'            4.3.10
	'            Added update_demand to ParseXDVersion for server 4.3.10.
	'290807 rjk: Added xdump updates and xdump game for Time to Update / Auto Update
	'160907 rjk: Only request meta entries for updates and game for servers 4.3.10.
	'221007 rjk: Fixed "Fire Range"/fire_range_factor.
	'111207 rjk: Fixed the cargo index in AddOrders.
	'131207 rjk: Added end cargo fields to Ship orders, rename cargo to start cargo
	'251207 rjk: Swap cargo start and end to match qorder.
	'281207 rjk: Swap dest's for order to match sorder.
	'010108 rjk: Add elev to xdump sect for 4.3.11 servers.
	'030108 rjk: Add storage of elev in the sector table.
	'120108 rjk: Removed unused local variable eiItem from GetNationType.
	'            Removed unused local variable iMaxPop from MaxPopulation.
	'310308 rjk: Added terrain to ParseXDSectorChr
	'080408 rjk: Added maxnoc for 4.3.13 to ParseXDVersion
	'240508 rjk: Use to city packing value from sect-chr to set WinACE's version big city flag.
	'010608 rjk: Enhanced MaxPopulation to evaluate by sector
	'140608 rjk: Added packing_type to sector chr to fix the packing factor for mobility calculations.
	'120708 rjk: Add Items.UpdateDatabase for ParseItemChar, fixes food mobility costs.
	'160708 rjk: Added terr0 for 4.3.15 server in ParseXDSector
	'270808 rjk: Fixed ParseXDPlane, did not work when the rsShip happen to be add the end of the table.
	'270908 rjk: Added timeused to ParseXDCountries and ParseXDNation, replace minused.
	'160209 rjk: Added RAILWAYS option to ParseXDVersion, originally part of Heavy Metal/Plastic games.
	'            Added down status to game information
	'180709 rjk: Added maint to ParseSectorChr.  New field for 4.3.23.
	'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.
	'090612 rjk: Added max_idle_visitor and login_grace_time to version parsing.
	
	Public Enum enumXdumpType
		XD_UNKNOWN
		XD_COUNTRIES
		XD_GAME
		XD_ITEM_CHR
		XD_INFRA_CHR
		XD_LAND_CHR
		XD_LAND
		XD_LEVEL
		XD_LOST
		XD_META
		XD_META_TYPE
		XD_MISSIONS
		XD_NATION
		XD_NATION_RELATIONSHIPS
		XD_NUKE_CHR
		XD_NUKE
		XD_PLANE_CHR
		XD_PLANE
		XD_PRODUCT_CHR
		XD_SECTOR_CHR
		XD_SECTOR_NAVIGATION
		XD_SECTOR
		XD_SHIP_CHR_FLAGS
		XD_SHIP_CHR
		XD_SHIP
		XD_UPDATES
		XD_VERSION
	End Enum
	
	Private Enum enumNationStatus
		STAT_UNUSED
		STAT_NEW
		STAT_VIS
		STAT_SANCT
		STAT_ACTIVE
		STAT_GOD
	End Enum
	
	Private xdMeta As Boolean
	Private xdType As enumXdumpType
	Private xdMetaType As enumXdumpType
	Private tmStamp As Integer
	Private iRowCount As Short
	
	Public Sub ParseXdump(ByRef strLine As String, ByRef bUpdateTimeStamp As Boolean)
		Dim strParams() As String
		
		'check for an announcement or tele message in the middle of a dump
		If InStr(strLine, "You have") > 0 Then
			Exit Sub
		End If
		
		If Left(strLine, 6) = "XDUMP " Then
			xdType = enumXdumpType.XD_UNKNOWN
			tmStamp = 0
			strParams = Split(strLine, " ")
			iRowCount = 0
			If UBound(strParams) = 3 Then
				ParseXDMetaHeader((strParams(2)))
				xdMeta = True
				tmStamp = CInt(strParams(3))
			ElseIf UBound(strParams) = 2 Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xdType = ParseXDTableName(strParams(1))
				If xdType <> enumXdumpType.XD_UNKNOWN Then
					tmStamp = CInt(strParams(2))
					If bUpdateTimeStamp Then
						Select Case xdType
							Case enumXdumpType.XD_LAND
								tsLand = tmStamp
							Case enumXdumpType.XD_NUKE
								tsNuke = tmStamp
							Case enumXdumpType.XD_PLANE
								tsPlane = tmStamp
							Case enumXdumpType.XD_SECTOR
								tsSectors = tmStamp
							Case enumXdumpType.XD_SHIP
								tsShip = tmStamp
							Case enumXdumpType.XD_LOST
								tsLost = tmStamp
							Case enumXdumpType.XD_LAND_CHR
								DeleteBuildRecords("l")
							Case enumXdumpType.XD_PLANE_CHR
								DeleteBuildRecords("p")
							Case enumXdumpType.XD_SHIP_CHR
								DeleteBuildRecords("s")
							Case enumXdumpType.XD_SECTOR_CHR
								DeleteBuildRecords("i")
							Case enumXdumpType.XD_COUNTRIES
								DeleteAllRecords(rsNations)
						End Select
					End If
					If frmEmpCmd.ForceUpdates Then
						Select Case xdType
							Case enumXdumpType.XD_LAND
								DeleteAllRecords(rsLand)
							Case enumXdumpType.XD_NUKE
								DeleteAllRecords(rsNuke)
							Case enumXdumpType.XD_PLANE
								DeleteAllRecords(rsPlane)
							Case enumXdumpType.XD_SECTOR
								DeleteAllRecords(rsSectors)
							Case enumXdumpType.XD_SHIP
								DeleteAllRecords(rsShip)
						End Select
					End If
					Select Case xdType
						Case enumXdumpType.XD_LAND_CHR
							UpdateTechLevelForShow("l", "b")
							UpdateTechLevelForShow("l", "c")
							UpdateTechLevelForShow("l", "s")
						Case enumXdumpType.XD_NUKE_CHR
							UpdateTechLevelForShow("n", "b")
							UpdateTechLevelForShow("n", "s")
						Case enumXdumpType.XD_PLANE_CHR
							UpdateTechLevelForShow("p", "b")
							UpdateTechLevelForShow("p", "c")
							UpdateTechLevelForShow("p", "s")
						Case enumXdumpType.XD_SECTOR_CHR
							UpdateTechLevelForShow("c", "b")
							UpdateTechLevelForShow("c", "h")
							UpdateTechLevelForShow("c", "s")
						Case enumXdumpType.XD_SHIP_CHR
							UpdateTechLevelForShow("s", "b")
							UpdateTechLevelForShow("s", "c")
							UpdateTechLevelForShow("s", "s")
						Case enumXdumpType.XD_VERSION
							UpdateTechLevelForShow("b", "b")
							UpdateTechLevelForShow("t", "b")
						Case enumXdumpType.XD_LEVEL, enumXdumpType.XD_META_TYPE, enumXdumpType.XD_MISSIONS, enumXdumpType.XD_SECTOR_NAVIGATION, enumXdumpType.XD_SHIP_CHR_FLAGS
							ParseXDSymbolTableHeader(xdType, strParams(1))
					End Select
				Else
					Debug.Print("ParseXdump:Invalid xdump type:" & CStr(xdType) & ", " & strLine)
				End If
			Else
				EmpireError("ParseXdump", "Invalid xdump type:" & CStr(xdType), strLine)
			End If
		ElseIf Left(strLine, 12) = "Usage: xdump" Then 
			Exit Sub
		ElseIf Left(strLine, 1) = "/" Then 
			Select Case xdType
				Case enumXdumpType.XD_COUNTRIES
					frmChat.ControlFlashForm()
				Case enumXdumpType.XD_SHIP, enumXdumpType.XD_SECTOR, enumXdumpType.XD_LAND, enumXdumpType.XD_NUKE, enumXdumpType.XD_PLANE
					frmDrawMap.Refresh()
					If rsVersion.Fields("RAILWAYS").Value And frmEmpCmd.ForceUpdates Then
						UpdateRailTrack()
					End If
				Case enumXdumpType.XD_VERSION, enumXdumpType.XD_SHIP_CHR, enumXdumpType.XD_SECTOR_CHR, enumXdumpType.XD_LAND_CHR, enumXdumpType.XD_PLANE_CHR, enumXdumpType.XD_PRODUCT_CHR, enumXdumpType.XD_ITEM_CHR, enumXdumpType.XD_INFRA_CHR
					If xdType = enumXdumpType.XD_ITEM_CHR Then
						Items.UpdateDatabase()
					End If
					If frmToolShow.Visible Then
						frmToolShow.FillGrid()
					End If
					
			End Select
			xdType = enumXdumpType.XD_UNKNOWN
			xdMeta = False
		ElseIf xdMeta Then 
			ParseXDMeta(xdMetaType, strLine)
		Else
			Select Case xdType
				Case enumXdumpType.XD_COUNTRIES
					ParseXDCountries(strLine)
				Case enumXdumpType.XD_GAME
					ParseXDGame(strLine)
				Case enumXdumpType.XD_ITEM_CHR
					ParseXDItemChr(strLine)
				Case enumXdumpType.XD_INFRA_CHR
					ParseXDInfraChr(strLine)
				Case enumXdumpType.XD_LAND
					ParseXDLand(strLine)
				Case enumXdumpType.XD_LOST
					ParseXDLost(strLine)
				Case enumXdumpType.XD_LAND_CHR
					ParseXDLandChr(strLine)
				Case enumXdumpType.XD_NATION
					ParseXDNation(strLine)
				Case enumXdumpType.XD_NATION_RELATIONSHIPS
					ParseXDNationRelationships(strLine)
				Case enumXdumpType.XD_NUKE
					ParseXDNuke(strLine)
				Case enumXdumpType.XD_NUKE_CHR
					ParseXDNukeChr(strLine)
				Case enumXdumpType.XD_PLANE
					ParseXDPlane(strLine)
				Case enumXdumpType.XD_PLANE_CHR
					ParseXDPlaneChr(strLine)
				Case enumXdumpType.XD_SECTOR
					ParseXDSector(strLine)
				Case enumXdumpType.XD_SECTOR_CHR
					ParseXDSectorChr(strLine)
				Case enumXdumpType.XD_SHIP
					ParseXDShip(strLine)
				Case enumXdumpType.XD_SHIP_CHR
					ParseXDShipChr(strLine)
				Case enumXdumpType.XD_PRODUCT_CHR
					ParseXDProductChr(strLine)
				Case enumXdumpType.XD_UPDATES
					ParseXDUpdates(strLine)
				Case enumXdumpType.XD_VERSION
					ParseXDVersion(strLine)
				Case enumXdumpType.XD_LEVEL, enumXdumpType.XD_META_TYPE, enumXdumpType.XD_MISSIONS, enumXdumpType.XD_SECTOR_NAVIGATION, enumXdumpType.XD_SHIP_CHR_FLAGS
					ParseXDSymbolTable(strLine)
				Case enumXdumpType.XD_UNKNOWN
					'        EmpireError "ParseXdump", "Invalid xdump unknown type", strLine
				Case Else
					EmpireError("ParseXdump", "No Parser type:" & CStr(xdType), strLine)
			End Select
		End If
	End Sub
	
	Private Sub ParseXDSector(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iOwnerIndex As Short
		Dim iXlocIndex As Short
		Dim iYlocIndex As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		
		'Save and restore the index on entry and exit
		hldIndex = rsSectors.Index
		rsSectors.Index = "PrimaryKey"
		
		On Error GoTo ParseXDSector_Exit
		
		etTable = etsTable(enumXdumpType.XD_SECTOR)
		If etTable Is Nothing Then
			EmpireError("ParseXDSector", "The meta data was not found for sector", strLine)
			rsSectors.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDSector", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsSectors.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Sector information to get a record.
		iOwnerIndex = etTable.GetZeroIndex("owner")
		iXlocIndex = etTable.GetZeroIndex("xloc")
		iYlocIndex = etTable.GetZeroIndex("yloc")
		
		If iXlocIndex = -1 Or iYlocIndex = -1 Then
			EmpireError("ParseXDSector", "The record identifier fields not present", CStr(iXlocIndex) & ":" & CStr(iYlocIndex))
			rsSectors.Index = hldIndex
			Exit Sub
		End If
		
		rsSectors.Seek("=", strParams(iXlocIndex), strParams(iYlocIndex))
		If rsSectors.NoMatch Then
			rsSectors.AddNew()
		Else
			rsSectors.Edit()
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "owner"
					rsSectors.Fields("Country").Value = strParams(iIndex)
				Case "xloc"
					rsSectors.Fields("x").Value = strParams(iIndex)
				Case "yloc"
					rsSectors.Fields("y").Value = strParams(iIndex)
				Case "des"
					rsSectors.Fields("des").Value = GetSectorDes(Val(strParams(iIndex)))
				Case "newdes"
					rsSectors.Fields("sdes").Value = GetSectorDes(Val(strParams(iIndex)))
					If etTable.GetZeroIndex("des") <> -1 Then
						If strParams(etTable.GetZeroIndex("des")) = strParams(iIndex) Then
							rsSectors.Fields("sdes").Value = " "
						End If
					End If
				Case "effic"
					rsSectors.Fields("eff").Value = strParams(iIndex)
				Case "mobil"
					rsSectors.Fields("mob").Value = strParams(iIndex)
				Case "oldown"
					If strParams(iIndex) = strParams(iOwnerIndex) Then
						rsSectors.Fields("*").Value = " "
					Else
						rsSectors.Fields("*").Value = "*"
					End If
				Case "off", "min", "gold", "fert", "ocontent", "uran", "work", "avail", "terr", "terr1", "terr2", "terr3", "uw", "food", "shell", "gun", "iron", "dust", "bar", "oil", "lcm", "hcm", "rad", "road", "rail", "fallout", "c_dist", "m_dist", "u_dist", "f_dist", "s_dist", "g_dist", "p_dist", "i_dist", "d_dist", "b_dist", "o_dist", "l_dist", "h_dist", "r_dist", "timestamp", "elev"
					rsSectors.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "terr0"
					rsSectors.Fields("terr").Value = strParams(iIndex)
				Case "petrol"
					rsSectors.Fields("pet").Value = strParams(iIndex)
				Case "civil"
					rsSectors.Fields("civ").Value = strParams(iIndex)
				Case "milit"
					rsSectors.Fields("mil").Value = strParams(iIndex)
				Case "dfense"
					rsSectors.Fields("defense").Value = strParams(iIndex)
				Case "coastal"
					rsSectors.Fields("coast").Value = strParams(iIndex)
				Case "xdist"
					rsSectors.Fields("dist_x").Value = strParams(iIndex)
				Case "ydist"
					rsSectors.Fields("dist_y").Value = strParams(iIndex)
				Case "mines"
					rsSectors.Fields("lmines").Value = strParams(iIndex)
				Case "c_del", "m_del", "u_del", "f_del", "s_del", "g_del", "p_del", "i_del", "d_del", "b_del", "o_del", "l_del", "h_del", "r_del"
					rsSectors.Fields(etTable.Column(iHeaderIndex).Name).Value = DirectionString(Val(strParams(iIndex)) And 7)
					rsSectors.Fields(Left(etTable.Column(iHeaderIndex).Name, 2) & "cut").Value = Val(strParams(iIndex)) And &HFFF8
				Case "pstage", "ptime", "che", "che_target", "access", "loyal"
					'South Pacific: The Search For Atlantis
				Case "frt"
					If frmOptions.bSPAtlantis Then
						rsSectors.Fields("fert").Value = strParams(iIndex)
					Else
						EmpireError("ParseXDSector", "Field not understood:" & strErrorField, strLine)
					End If
				Case "runway"
					If frmOptions.bSPAtlantis Then
						rsSectors.Fields("min").Value = strParams(iIndex)
					Else
						EmpireError("ParseXDSector", "Field not understood:" & strErrorField, strLine)
					End If
				Case "rdar"
					If frmOptions.bSPAtlantis Then
						rsSectors.Fields("gold").Value = strParams(iIndex)
					Else
						EmpireError("ParseXDSector", "Field not understood:" & strErrorField, strLine)
					End If
				Case "nvigate"
					If frmOptions.bSPAtlantis Then
						rsSectors.Fields("ocontent").Value = strParams(iIndex)
					Else
						EmpireError("ParseXDSector", "Field not understood:" & strErrorField, strLine)
					End If
				Case "dterr"
					If bDeity And VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						'In the future stored the deity territory
					Else
						EmpireError("ParseXDSector", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDSector", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		If iOwnerIndex = -1 Then
			rsSectors.Fields("country").Value = CountryNumber
		End If
		
		rsSectors.Update()
		rsSectors.Index = hldIndex
		
		If Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY Then
			frmDrawMap.FillSectorBox(Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex)))
			If Not frmDrawMap.IsAnyTabActive Then
				frmDrawMap.DisplayFirstSectorPanel()
			End If
		End If
		
		If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
			hldIndex = rsEnemySect.Index
			rsEnemySect.Index = "PrimaryKey"
			rsEnemySect.Seek("=", Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex)))
			If Not rsEnemySect.NoMatch Then
				rsEnemySect.Delete()
			End If
			rsEnemySect.Index = hldIndex
		End If
		
		Exit Sub
		
		'Error handling routine
ParseXDSector_Exit: 
		EmpireError("ParseXDSector", strErrorField, strLine)
		rsSectors.Index = hldIndex
	End Sub
	
	Private Function GetSectorDes(ByRef iDchrIndex As Short) As String
		If Not (rsSectorType.BOF And rsSectorType.EOF) Then
			rsSectorType.MoveFirst()
			While Not rsSectorType.EOF
				If rsSectorType.Fields("dchr_i").Value = iDchrIndex Then
					GetSectorDes = rsSectorType.Fields("des").Value
					Exit Function
				End If
				rsSectorType.MoveNext()
			End While
		End If
		GetSectorDes = "?"
	End Function
	
	Private Sub ParseXDSectorChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iDesIndex As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim hldBuildIndex As String
		Dim etTable As EmpTable
		Dim iCost As Short
		Dim iBuild As Short
		Dim iLcms As Short
		Dim iHcms As Short
		
		'Save and restore the index on entry and exit
		hldIndex = rsSectorType.Index
		rsSectorType.Index = "PrimaryKey"
		
		On Error GoTo ParseXDSectorChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_SECTOR_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDSectorChr", "The meta data was not found for sect-chr", strLine)
			rsSectorType.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDSectorChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsSectorType.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Sector information to get a record.
		iDesIndex = etTable.GetZeroIndex("mnem")
		
		If iDesIndex = -1 Then
			EmpireError("ParseXDSectorChr", "The record identifier fields not present", CStr(iDesIndex))
			rsSectorType.Index = hldIndex
			Exit Sub
		End If
		strParams(iDesIndex) = RemoveEscapeSequencesAndQuotes(strParams(iDesIndex))
		
		rsSectorType.Seek("=", strParams(iDesIndex))
		If rsSectorType.NoMatch Then
			rsSectorType.AddNew()
		Else
			rsSectorType.Edit()
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					rsSectorType.Fields("desc").Value = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
				Case "mnem"
					rsSectorType.Fields("des").Value = strParams(iIndex)
				Case "mcst"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						EmpireError("ParseXDSectorChr", "Field not understood:" & strErrorField, strLine)
					Else
						rsSectorType.Fields("mcost").Value = strParams(iIndex)
					End If
				Case "ostr"
					rsSectorType.Fields("offense").Value = Val(strParams(iIndex))
				Case "dstr"
					rsSectorType.Fields("defense").Value = Val(strParams(iIndex))
				Case "flg", "value", "uid", "maint"
				Case "pkg"
					rsSectorType.Fields("pack_civ").Value = Items.FindByLetter("c").Packing(Val(strParams(iIndex)))
					rsSectorType.Fields("pack_mil").Value = Items.FindByLetter("m").Packing(Val(strParams(iIndex)))
					rsSectorType.Fields("pack_uw").Value = Items.FindByLetter("u").Packing(Val(strParams(iIndex)))
					rsSectorType.Fields("pack_bar").Value = Items.FindByLetter("b").Packing(Val(strParams(iIndex)))
					rsSectorType.Fields("pack_other").Value = Items.FindByLetter("l").Packing(Val(strParams(iIndex)))
					If strParams(iDesIndex) = "c" Then
						rsVersion.MoveFirst()
						rsVersion.Edit()
						If Val(strParams(iIndex)) = EmpItem.enumItemPacking.PACKING_URBAN Then
							rsVersion.Fields("BIG_CITY").Value = True
						Else
							rsVersion.Fields("BIG_CITY").Value = False
						End If
						rsVersion.Update()
					End If
					rsSectorType.Fields("pack_type").Value = Val(strParams(iIndex))
				Case "cost"
					iCost = CShort(strParams(iIndex))
				Case "lcms"
					iLcms = CShort(strParams(iIndex))
				Case "hcms"
					iHcms = CShort(strParams(iIndex))
				Case "build"
					iBuild = CShort(strParams(iIndex))
					'        If strParams(iIndex) >= 0 Then
					'        End If
					'               dchr[x].d_mnem, dchr[x].d_cost, dchr[x].d_build,
					'           dchr[x].d_lcms, dchr[x].d_hcms);
				Case "peffic"
					If VersionCheck(4, 3, 6) = WinAceRoutines.enumVersion.VER_LESS Then
						EmpireError("ParseXDSectorChr", "Field not understood:" & strErrorField, strLine)
					Else
						rsSectorType.Fields("eff").Value = strParams(iIndex)
					End If
				Case "prd"
					rsSectorType.Fields("pchr_i").Value = strParams(iIndex)
				Case "maxpop", "nav"
					rsSectorType.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "terrain"
					rsSectorType.Fields(etTable.Column(iHeaderIndex).Name).Value = Val(strParams(iIndex))
				Case "mob0"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						' do nothing for the moment unable I figure this field does
						rsSectorType.Fields("mcost").Value = Val(strParams(iIndex))
					Else
						EmpireError("ParseXDSectorChr", "Field not understood:" & strErrorField, strLine)
					End If
				Case "mob1"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						' do nothing for the moment unable I figure this field does
						rsSectorType.Fields("mcost100").Value = Val(strParams(iIndex))
					Else
						EmpireError("ParseXDSectorChr", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDSectorChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		If iCost >= 0 Then
			If iCost > 0 Or iLcms > 0 Or iHcms > 0 Or iBuild <> 1 Then
				strErrorField = "Build Infrastructure"
				hldBuildIndex = rsBuild.Index
				rsBuild.Index = "PrimaryKey"
				
				If rsBuild.BOF And rsBuild.EOF Then
					rsBuild.AddNew()
					rsBuild.Fields("type").Value = "i"
				Else
					rsBuild.Seek("=", "i", strParams(iDesIndex))
					If rsBuild.NoMatch Then
						rsBuild.AddNew()
						rsBuild.Fields("type").Value = "i"
					Else
						rsBuild.Edit()
					End If
				End If
				rsBuild.Fields("id").Value = strParams(iDesIndex)
				rsBuild.Fields("cost").Value = iCost
				rsBuild.Fields("lcm").Value = iLcms
				rsBuild.Fields("hcm").Value = iHcms
				rsBuild.Fields("stat1").Value = iBuild
				rsBuild.Update()
				rsBuild.Index = hldBuildIndex
			End If
		End If
		
		rsSectorType.Fields("dchr_i").Value = iRowCount
		rsSectorType.Fields("pcode").Value = ""
		rsSectorType.Update()
		rsSectorType.Index = hldIndex
		iRowCount = iRowCount + 1
		Exit Sub
		
		'Error handling routine
ParseXDSectorChr_Exit: 
		EmpireError("ParseXDSectorChr", strErrorField, strLine)
		rsSectorType.Index = hldIndex
	End Sub
	
	Private Sub ParseXDPlane(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iOwnerIndex As Short
		Dim iIDIndex As Short
		Dim iXlocIndex As Short
		Dim iYlocIndex As Short
		Dim iOpxIndex As Short
		Dim iOpyIndex As Short
		Dim iRadiusIndex As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim hldChrIndex As String
		Dim etTable As EmpTable
		Dim iPrevXloc As Short
		Dim iPrevYloc As Short
		
		'Save and restore the index on entry and exit
		hldIndex = rsPlane.Index
		rsPlane.Index = "PrimaryKey"
		
		On Error GoTo ParseXDPlane_Exit
		
		etTable = etsTable(enumXdumpType.XD_PLANE)
		If etTable Is Nothing Then
			EmpireError("ParseXDPlane", "The meta data was not found for plane", strLine)
			rsPlane.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDPlane", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsPlane.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Sector information to get a record.
		iOwnerIndex = etTable.GetZeroIndex("owner")
		iIDIndex = etTable.GetZeroIndex("uid")
		iXlocIndex = etTable.GetZeroIndex("xloc")
		iYlocIndex = etTable.GetZeroIndex("yloc")
		iOpxIndex = etTable.GetZeroIndex("opx")
		iOpyIndex = etTable.GetZeroIndex("opy")
		iRadiusIndex = etTable.GetZeroIndex("radius")
		
		If iIDIndex = -1 Or iOwnerIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or iRadiusIndex = -1 Or iOpxIndex = -1 Or iOpxIndex = -1 Then
			EmpireError("ParseXDPlane", "One of the following fields is missing: owner, id, xloc, yloc, opx, opy, rad", "")
			rsPlane.Index = hldIndex
			Exit Sub
		End If
		
		If rsPlane.BOF And rsPlane.EOF Then
			rsPlane.AddNew()
		Else
			rsPlane.Seek("=", strParams(iIDIndex))
			If rsPlane.NoMatch Then
				rsPlane.AddNew()
			Else
				rsPlane.Edit()
			End If
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "owner"
					rsPlane.Fields("Country").Value = strParams(iIndex)
				Case "uid"
					rsPlane.Fields("id").Value = strParams(iIndex)
				Case "type"
					rsPlane.Fields("type").Value = GetUnitType(Val(strParams(iIndex)), "p")
				Case "wing"
					If estsTable(enumXdumpType.XD_META_TYPE)((etTable.Column(iHeaderIndex).ColType)).Name = "c" Then
						strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
						If strParams(iIndex) = "" Or strParams(iIndex) = " " Then 'replace space with ~
							rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = "~"
						Else
							rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = Left(strParams(iIndex), 1)
						End If
					Else
						If CDbl(strParams(iIndex)) = 32 Then 'replace space with ~
							rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = "~"
						Else
							rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = Chr(CInt(strParams(iIndex)))
						End If
					End If
				Case "xloc"
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(rsPlane.Fields("x").Value) Then
						iPrevXloc = -1
					Else
						iPrevXloc = rsPlane.Fields("x").Value
					End If
					rsPlane.Fields("x").Value = strParams(iIndex)
				Case "yloc"
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(rsPlane.Fields("y").Value) Then
						iPrevYloc = -1
					Else
						iPrevYloc = rsPlane.Fields("y").Value
					End If
					rsPlane.Fields("y").Value = strParams(iIndex)
				Case "effic"
					rsPlane.Fields("eff").Value = strParams(iIndex)
				Case "mobil"
					rsPlane.Fields("mob").Value = strParams(iIndex)
				Case "harden"
					rsPlane.Fields("hard").Value = strParams(iIndex)
				Case "range"
					rsPlane.Fields("react").Value = strParams(iIndex)
				Case "tech"
					rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					hldChrIndex = rsBuild.Index
					rsBuild.Index = "PrimaryKey"
					
					rsBuild.Seek("=", "p", rsPlane.Fields("type"))
					If Not rsBuild.NoMatch Then
						rsPlane.Fields("att").Value = Fix(ComputePlaneAttDef(rsBuild.Fields("base att").Value, rsBuild.Fields("tech").Value, rsPlane.Fields("tech").Value))
						rsPlane.Fields("def").Value = Fix(ComputePlaneAttDef(rsBuild.Fields("base def").Value, rsBuild.Fields("tech").Value, rsPlane.Fields("tech").Value))
						rsPlane.Fields("acc").Value = Fix(ComputePlaneAccuracy(rsBuild.Fields("base acc").Value, rsBuild.Fields("tech").Value, rsPlane.Fields("tech").Value))
						rsPlane.Fields("load").Value = Fix(ComputePlaneLoad(rsBuild.Fields("base load").Value, rsBuild.Fields("tech").Value, rsPlane.Fields("tech").Value))
						'            'need to fix for limited range planes
						rsPlane.Fields("range").Value = Fix(ComputePlaneRange(rsBuild.Fields("base spd").Value, rsBuild.Fields("tech").Value, rsPlane.Fields("tech").Value))
						rsPlane.Fields("fuel").Value = rsBuild.Fields("stat6").Value
						'        Else
						'            EmpireError "ParseXDPlane", "Plane Characteristics not found:" + rsPlane.Fields("type"), strLine
					End If
					rsBuild.Index = hldChrIndex
				Case "ship", "land", "timestamp"
					rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "nuketype"
					rsPlane.Fields("nuke").Value = GetUnitType(Val(strParams(iIndex)), "n")
				Case "opx", "opy", "radius", "access", "theta"
				Case "mission"
					AddMission(Val(strParams(iIndex)), Val(strParams(iIDIndex)), Val(strParams(iOwnerIndex)), "p", SectorString(Val(strParams(iOpxIndex)), Val(strParams(iOpyIndex))), Val(strParams(iRadiusIndex)))
				Case "flags"
					If (strParams(iIndex) And &H1) <> 0 Then
						rsPlane.Fields("laun").Value = "Y"
					Else
						rsPlane.Fields("laun").Value = "N"
					End If
					If (strParams(iIndex) And &H2) <> 0 Then
						rsPlane.Fields("orb").Value = "Y"
					Else
						rsPlane.Fields("orb").Value = "N"
					End If
					If (strParams(iIndex) And &H4) <> 0 Then
						rsPlane.Fields("grd").Value = "A"
					Else
						rsPlane.Fields("grd").Value = "G"
					End If
				Case "off"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						rsPlane.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					Else
						EmpireError("ParseXDPlane", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDPlane", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		If iOwnerIndex = -1 Then
			rsPlane.Fields("country").Value = CountryNumber
		End If
		
		rsPlane.Update()
		rsPlane.Index = hldIndex
		
		If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
			frmDrawMap.FillGrid()
		End If
		
		CheckForOldEnemyUnit(CShort(strParams(iIDIndex)), "p")
		Exit Sub
		
		'Error handling routine
ParseXDPlane_Exit: 
		EmpireError("ParseXDPlane", strErrorField, strLine)
		rsPlane.Index = hldIndex
	End Sub
	
	Private Sub ParseXDPlaneChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iCapIndex As Short
		Dim iIDIndex As Short
		Dim iTechIndex As Short
		Dim iPos As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		Dim lFlags As Integer
		
		'Save and restore the index on entry and exit
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		
		On Error GoTo ParseXDPlaneChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_PLANE_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDPlaneChr", "The meta data was not found for plane-chr", strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDPlaneChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Index information to get a record.
		iIDIndex = etTable.GetZeroIndex("name")
		strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
		iTechIndex = etTable.GetZeroIndex("tech")
		
		If iIDIndex = -1 Then
			EmpireError("ParseXDPlaneChr", "The record identifier fields not present", CStr(iIDIndex))
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		iPos = InStr(strParams(iIDIndex), " ")
		
		If rsBuild.BOF And rsBuild.EOF Then
			rsBuild.AddNew()
			rsBuild.Fields("type").Value = "p"
		Else
			rsBuild.Seek("=", "p", Left(strParams(iIDIndex), iPos - 1))
			If rsBuild.NoMatch Then
				rsBuild.AddNew()
				rsBuild.Fields("type").Value = "p"
			Else
				rsBuild.Edit()
			End If
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					If iPos > 0 Then
						rsBuild.Fields("id").Value = Left(strParams(iIndex), iPos - 1)
						rsBuild.Fields("desc").Value = LTrim(Mid(strParams(iIndex), iPos + 1, Len(strParams(iIndex)) - iPos))
					Else
						EmpireError("ParseXDPlaneChr", "Incorrect format of name", strLine)
					End If
				Case "cost", "tech"
					rsBuild.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "l_build"
					rsBuild.Fields("lcm").Value = strParams(iIndex)
				Case "h_build"
					rsBuild.Fields("hcm").Value = strParams(iIndex)
				Case "crew"
					rsBuild.Fields("mil").Value = strParams(iIndex)
				Case "flags"
					iCapIndex = 1
					lFlags = CInt(strParams(iIndex))
					If (lFlags And &H1) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "bomber"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H2) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "tactical"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H4) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "intercept"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H8) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "cargo"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H10) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "VTOL"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H20) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "missile"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H40) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "light"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H80) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "spy"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H100) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "image"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H200) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "satellite"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H400) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "stealth"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H800) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "SDI"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H1000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "half-stealth"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H2000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "x-light"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H4000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "helo"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H8000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "ASW"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H10000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "para"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H20000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "escort"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H40000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "mine"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H80000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "sweep"
						iCapIndex = iCapIndex + 1
					End If
					If (lFlags And &H100000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "marine"
						iCapIndex = iCapIndex + 1
					End If
					rsBuild.Fields("cap" & CStr(iCapIndex)).Value = ""
				Case "acc"
					rsBuild.Fields("base acc").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDPlaneChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
						rsBuild.Fields("stat1").Value = Fix(ComputePlaneAccuracy(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "load"
					rsBuild.Fields("base load").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDPlaneChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
						rsBuild.Fields("stat2").Value = Fix(ComputePlaneLoad(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "att"
					rsBuild.Fields("base att").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDPlaneChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
						rsBuild.Fields("stat3").Value = Fix(ComputePlaneAttDef(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "def"
					rsBuild.Fields("base def").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDPlaneChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
						rsBuild.Fields("stat4").Value = Fix(ComputePlaneAttDef(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "range"
					rsBuild.Fields("base spd").Value = CShort(strParams(iIndex))
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDPlaneChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
						rsBuild.Fields("stat5").Value = Fix(ComputePlaneRange(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "fuel"
					rsBuild.Fields("stat6").Value = strParams(iIndex)
				Case "stealth"
					rsBuild.Fields("stat7").Value = strParams(iIndex)
				Case "type"
					rsBuild.Fields("chr_i").Value = strParams(iIndex)
				Case Else
					EmpireError("ParseXDPlaneChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		rsBuild.Fields("avail").Value = CalculateAvail(CShort(rsBuild.Fields("lcm").Value), CShort(rsBuild.Fields("hcm").Value))
		
		rsBuild.Update()
		rsBuild.Index = hldIndex
		Exit Sub
		
		'Error handling routine
ParseXDPlaneChr_Exit: 
		EmpireError("ParseXDPlaneChr", strErrorField, strLine)
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub ParseXDShip(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iArrayIndex As Short
		Dim iHeaderIndex As Short
		Dim iOwnerIndex As Short
		Dim iIDIndex As Short
		Dim iXlocIndex As Short
		Dim iYlocIndex As Short
		Dim iOpxIndex As Short
		Dim iOpyIndex As Short
		Dim iRadiusIndex As Short
		Dim strErrorField As String
		Dim hldShipIndex As String
		Dim hldChrIndex As String
		Dim iXstartIndex As Short
		Dim iYstartIndex As Short
		Dim iXendIndex As Short
		Dim iYendIndex As Short
		Dim iTypeIndex As Short
		Dim strStartCargo(6) As String
		Dim strEndCargo(6) As String
		Dim iStart(6) As Short
		Dim iEnd(6) As Short
		Dim etTable As EmpTable
		Dim bAutoNav As Boolean
		Dim iPrevXloc As Short
		Dim iPrevYloc As Short
		
		'Save and restore the index on entry and exit
		hldShipIndex = rsShip.Index
		rsShip.Index = "PrimaryKey"
		
		On Error GoTo ParseXDShip_Exit
		
		etTable = etsTable(enumXdumpType.XD_SHIP)
		If etTable Is Nothing Then
			EmpireError("ParseXDShip", "The meta data was not found for ship", strLine)
			rsShip.Index = hldShipIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDShip", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsShip.Index = hldShipIndex
			Exit Sub
		End If
		
		'Get the Sector information to get a record.
		iOwnerIndex = etTable.GetZeroIndex("owner")
		iIDIndex = etTable.GetZeroIndex("uid")
		iXlocIndex = etTable.GetZeroIndex("xloc")
		iYlocIndex = etTable.GetZeroIndex("yloc")
		iOpxIndex = etTable.GetZeroIndex("opx")
		iOpyIndex = etTable.GetZeroIndex("opy")
		iRadiusIndex = etTable.GetZeroIndex("radius")
		iXstartIndex = etTable.GetZeroIndex("xstart")
		iYstartIndex = etTable.GetZeroIndex("ystart")
		iXendIndex = etTable.GetZeroIndex("xend")
		iYendIndex = etTable.GetZeroIndex("yend")
		iTypeIndex = etTable.GetZeroIndex("type")
		
		If iOwnerIndex = -1 Or iIDIndex = -1 Or iOpxIndex = -1 Or iOpyIndex = -1 Or iRadiusIndex = -1 Or iXlocIndex = -11 Or iYlocIndex = -1 Or iTypeIndex = -1 Or iXstartIndex = -1 Or iYstartIndex = -1 Or iXendIndex = -1 Or iYendIndex = -1 Then
			EmpireError("ParseXDShip", "One of the following fields is missing: owner, id, xloc, yloc, opx, opy, rad, type" & ", startx, starty, endx or endy", CStr(iIDIndex))
			rsShip.Index = hldShipIndex
			Exit Sub
		End If
		
		If rsShip.BOF And rsShip.EOF Then
			rsShip.AddNew()
		Else
			rsShip.Seek("=", strParams(iIDIndex))
			If rsShip.NoMatch Then
				rsShip.AddNew()
			Else
				rsShip.Edit()
			End If
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "owner"
					rsShip.Fields("Country").Value = strParams(iIndex)
				Case "uid"
					rsShip.Fields("id").Value = strParams(iIndex)
				Case "type"
					rsShip.Fields("type").Value = GetUnitType(Val(strParams(iIndex)), "s")
				Case "fleet"
					If estsTable(enumXdumpType.XD_META_TYPE)((etTable.Column(iHeaderIndex).ColType)).Name = "c" Then
						strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
						If strParams(iIndex) = "" Or strParams(iIndex) = " " Then 'replace space with ~
							rsShip.Fields("flt").Value = "~"
						Else
							rsShip.Fields("flt").Value = Left(strParams(iIndex), 1)
						End If
					Else
						If CDbl(strParams(iIndex)) = 32 Then 'replace space with ~
							rsShip.Fields("flt").Value = "~"
						Else
							rsShip.Fields("flt").Value = Chr(CInt(strParams(iIndex)))
						End If
					End If
				Case "xloc"
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(rsShip.Fields("x").Value) Then
						iPrevXloc = -1
					Else
						iPrevXloc = rsShip.Fields("x").Value
					End If
					rsShip.Fields("x").Value = strParams(iIndex)
				Case "yloc"
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(rsShip.Fields("y").Value) Then
						iPrevYloc = 0
					Else
						iPrevYloc = rsShip.Fields("y").Value
					End If
					rsShip.Fields("y").Value = strParams(iIndex)
				Case "effic"
					rsShip.Fields("eff").Value = strParams(iIndex)
				Case "mobil"
					rsShip.Fields("mob").Value = strParams(iIndex)
				Case "nplane"
					rsShip.Fields("pln").Value = strParams(iIndex)
				Case "nland"
					rsShip.Fields("land").Value = strParams(iIndex)
				Case "nchoppers"
					rsShip.Fields("he").Value = strParams(iIndex)
				Case "nxlight"
					rsShip.Fields("xl").Value = strParams(iIndex)
				Case "civil"
					rsShip.Fields("civ").Value = strParams(iIndex)
				Case "milit"
					rsShip.Fields("mil").Value = strParams(iIndex)
				Case "xbuilt"
					rsShip.Fields("origx").Value = strParams(iIndex)
				Case "ybuilt"
					rsShip.Fields("origy").Value = strParams(iIndex)
				Case "name"
					rsShip.Fields(etTable.Column(iHeaderIndex).Name).Value = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
				Case "tech"
					rsShip.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					hldChrIndex = rsBuild.Index
					rsBuild.Index = "PrimaryKey"
					
					rsBuild.Seek("=", "s", GetUnitType(Val(strParams(iTypeIndex)), "s"))
					If Not rsBuild.NoMatch Then
						rsShip.Fields("def").Value = Fix(ComputeShipDefense(rsBuild.Fields("base def").Value, rsBuild.Fields("tech").Value, rsShip.Fields("tech").Value))
						rsShip.Fields("spd").Value = Fix(ComputeShipSpeed(rsBuild.Fields("base spd").Value, rsBuild.Fields("tech").Value, rsShip.Fields("tech").Value))
						rsShip.Fields("vis").Value = Fix(ComputeShipVisibility(rsBuild.Fields("base vul").Value, rsBuild.Fields("tech").Value, rsShip.Fields("tech").Value))
						rsShip.Fields("rng").Value = Fix(ComputeShipRange(rsBuild.Fields("base frnge").Value, rsBuild.Fields("tech").Value, rsShip.Fields("tech").Value))
						rsShip.Fields("fir").Value = Fix(ComputeShipFiringRange(rsBuild.Fields("base load").Value, rsBuild.Fields("tech").Value, rsShip.Fields("tech").Value))
						'        Else
						'            EmpireError "ParseXDShip", "Ship Characteristics not found:" + rsShip.Fields("type"), strLine
					End If
					rsBuild.Index = hldChrIndex
				Case "timestamp", "uw", "food", "fuel", "shell", "gun", "petrol", "iron", "dust", "bar", "oil", "lcm", "hcm", "rad"
					rsShip.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "opx", "opy", "radius", "flags", "pstage", "ptime", "access", "builder", "mquota", "path", "follow", "rflags", "rpath"
				Case "autonav"
					If Val(strParams(iIndex)) > 0 Then
						bAutoNav = True
					Else
						bAutoNav = False
					End If
				Case "mission"
					AddMission(Val(strParams(iIndex)), Val(strParams(iIDIndex)), Val(strParams(iOwnerIndex)), "s", SectorString(Val(strParams(iOpxIndex)), Val(strParams(iOpyIndex))), Val(strParams(iRadiusIndex)))
				Case "xstart"
				Case "ystart"
				Case "xend"
				Case "yend"
					'    Case "cargostart"
				Case "cargoend"
					
					For iArrayIndex = 1 To etTable.Column(iHeaderIndex).Length
						If strParams(iIndex) = "-1" Then
							strStartCargo(iArrayIndex) = ""
						Else
							strStartCargo(iArrayIndex) = Items.FindByChrIndex(Val(strParams(iIndex))).Letter
						End If
						iIndex = iIndex + 1
					Next 
					iIndex = iIndex - 1
					'    Case "cargoend"
				Case "cargostart"
					For iArrayIndex = 1 To etTable.Column(iHeaderIndex).Length
						If strParams(iIndex) = "-1" Then
							strEndCargo(iArrayIndex) = ""
						Else
							strEndCargo(iArrayIndex) = Items.FindByChrIndex(Val(strParams(iIndex))).Letter
						End If
						iIndex = iIndex + 1
					Next 
					iIndex = iIndex - 1
					'    Case "amtstart"
				Case "amtend"
					For iArrayIndex = 1 To etTable.Column(iHeaderIndex).Length
						iStart(iArrayIndex) = Val(strParams(iIndex))
						iIndex = iIndex + 1
					Next 
					iIndex = iIndex - 1
					'    Case "amtend"
				Case "amtstart"
					For iArrayIndex = 1 To etTable.Column(iHeaderIndex).Length
						iEnd(iArrayIndex) = Val(strParams(iIndex))
						iIndex = iIndex + 1
					Next 
					iIndex = iIndex - 1
				Case "off"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						rsShip.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					Else
						EmpireError("ParseXDShip", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDShip", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		If iOwnerIndex = -1 Then
			rsShip.Fields("country").Value = CountryNumber
		End If
		
		rsShip.Update()
		rsShip.Index = hldShipIndex
		
		AddOrders(bAutoNav, Val(strParams(iIDIndex)), Val(strParams(iOwnerIndex)), 0, 0, SectorString(Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex))), SectorString(Val(strParams(iXstartIndex)), Val(strParams(iYstartIndex))), SectorString(Val(strParams(iXendIndex)), Val(strParams(iYendIndex))), strStartCargo, strEndCargo, iStart, iEnd)
		
		If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
			frmDrawMap.FillGrid()
		End If
		
		CheckForOldEnemyUnit(CShort(strParams(iIDIndex)), "s")
		Exit Sub
		
		'Error handling routine
ParseXDShip_Exit: 
		EmpireError("ParseXDShip", strErrorField, strLine)
		rsShip.Index = hldShipIndex
	End Sub
	
	Private Sub ParseXDShipChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iCapIndex As Short
		Dim iIDIndex As Short
		Dim iTechIndex As Short
		Dim iTypeIndex As Short
		Dim iPos As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		
		'Save and restore the index on entry and exit
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		
		On Error GoTo ParseXDShipChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_SHIP_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDShipChr", "The meta data was not found for ship-chr", strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDShipChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Index information to get a record.
		iIDIndex = etTable.GetZeroIndex("name")
		strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
		iTechIndex = etTable.GetZeroIndex("tech")
		
		If iIDIndex = -1 Then
			EmpireError("ParseXDShipChr", "The record identifier field not present", CStr(iIDIndex))
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		If iTechIndex = -1 Then
			EmpireError("ParseXDShipChr", "The tech field not present", CStr(iTechIndex))
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		iPos = InStr(strParams(iIDIndex), " ")
		If rsBuild.BOF And rsBuild.EOF Then
			rsBuild.AddNew()
			rsBuild.Fields("type").Value = "s"
		Else
			rsBuild.Seek("=", "s", Left(strParams(iIDIndex), iPos - 1))
			If rsBuild.NoMatch Then
				rsBuild.AddNew()
				rsBuild.Fields("type").Value = "s"
			Else
				rsBuild.Edit()
			End If
		End If
		iCapIndex = 1
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					If iPos > 0 Then
						rsBuild.Fields("id").Value = Left(strParams(iIndex), iPos - 1)
						rsBuild.Fields("desc").Value = LTrim(Mid(strParams(iIndex), iPos + 1, Len(strParams(iIndex)) - iPos))
					Else
						EmpireError("ParseXDShipChr", "Incorrect format of name", strLine)
					End If
				Case "cost", "tech"
					rsBuild.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "l_build"
					rsBuild.Fields("lcm").Value = strParams(iIndex)
				Case "h_build"
					rsBuild.Fields("hcm").Value = strParams(iIndex)
				Case "civil"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "c"
						iCapIndex = iCapIndex + 1
					End If
				Case "milit"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "m"
						iCapIndex = iCapIndex + 1
					End If
				Case "shell"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "s"
						iCapIndex = iCapIndex + 1
					End If
				Case "gun"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "g"
						iCapIndex = iCapIndex + 1
					End If
				Case "petrol"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "p"
						iCapIndex = iCapIndex + 1
					End If
				Case "iron"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "i"
						iCapIndex = iCapIndex + 1
					End If
				Case "dust"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "d"
						iCapIndex = iCapIndex + 1
					End If
				Case "bar"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "b"
						iCapIndex = iCapIndex + 1
					End If
				Case "food"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "f"
						iCapIndex = iCapIndex + 1
					End If
				Case "oil"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "o"
						iCapIndex = iCapIndex + 1
					End If
				Case "lcm"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "l"
						iCapIndex = iCapIndex + 1
					End If
				Case "hcm"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "h"
						iCapIndex = iCapIndex + 1
					End If
				Case "uw"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "u"
						iCapIndex = iCapIndex + 1
					End If
				Case "rad"
					If Len(strParams(iIndex)) > 0 And CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "r"
						iCapIndex = iCapIndex + 1
					End If
				Case "flags"
					If (strParams(iIndex) And &H1) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "fish"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H2) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "torp"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H4) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "dchrg"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H8) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "plane"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H10) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "miss"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H20) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "oil"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H40) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "sonar"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H80) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "mine"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H100) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "sweep"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H200) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "sub"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H400) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "spy"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H800) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "land"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H1000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "sub-torp"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H2000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "trade"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H4000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "semi-land"
						iCapIndex = iCapIndex + 1
					End If
					'        If (strParams(iIndex) And &H8000&) <> 0 Then
					'            '#define M_XLIGHT    bit(15) /* can hold xlight planes */
					'        End If
					'        If (strParams(iIndex) And &H10000) <> 0 Then
					'            '#define M_CHOPPER   bit(16) /* can hold choppers */
					'        End If
					If (strParams(iIndex) And &H20000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "oiler"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H40000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "supply"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H80000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "canal"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H100000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "anti-missile"
						iCapIndex = iCapIndex + 1
					End If
				Case "armor"
					rsBuild.Fields("base def").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDShipChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
						rsBuild.Fields("stat1").Value = Fix(ComputeShipDefense(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "speed"
					rsBuild.Fields("base spd").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDShipChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
						rsBuild.Fields("stat2").Value = Fix(ComputeShipSpeed(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "visib"
					rsBuild.Fields("base vul").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDShipChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
						rsBuild.Fields("stat3").Value = Fix(ComputeShipVisibility(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "vrnge"
					rsBuild.Fields("stat4").Value = strParams(iIndex)
				Case "frnge"
					rsBuild.Fields("base frnge").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDShipChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
						rsBuild.Fields("stat5").Value = Fix(ComputeShipRange(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "glim"
					rsBuild.Fields("base load").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDShipChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
						rsBuild.Fields("stat6").Value = Fix(ComputeShipFiringRange(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "nland"
					rsBuild.Fields("stat7").Value = strParams(iIndex)
				Case "nplanes"
					rsBuild.Fields("stat8").Value = strParams(iIndex)
				Case "nchoppers"
					rsBuild.Fields("stat9").Value = strParams(iIndex)
				Case "nxlight"
					rsBuild.Fields("stat10").Value = strParams(iIndex)
				Case "fuelc"
					rsBuild.Fields("stat11").Value = strParams(iIndex)
				Case "fuelu"
					rsBuild.Fields("stat12").Value = strParams(iIndex)
				Case "type"
					rsBuild.Fields("chr_i").Value = strParams(iIndex)
				Case Else
					EmpireError("ParseXDShipChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		rsBuild.Fields("avail").Value = CalculateAvail(CShort(rsBuild.Fields("lcm").Value), CShort(rsBuild.Fields("hcm").Value))
		
		rsBuild.Update()
		rsBuild.Index = hldIndex
		Exit Sub
		
		'Error handling routine
ParseXDShipChr_Exit: 
		EmpireError("ParseXDShipChr", strErrorField, strLine)
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub ParseXDLand(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iOwnerIndex As Short
		Dim iIDIndex As Short
		Dim iXlocIndex As Short
		Dim iYlocIndex As Short
		Dim iOpxIndex As Short
		Dim iOpyIndex As Short
		Dim iRadiusIndex As Short
		Dim iTypeIndex As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim hldChrIndex As String
		Dim etTable As EmpTable
		Dim iPrevXloc As Short
		Dim iPrevYloc As Short
		
		'Save and restore the index on entry and exit
		hldIndex = rsLand.Index
		rsLand.Index = "PrimaryKey"
		
		On Error GoTo ParseXDLand_Exit
		
		etTable = etsTable(enumXdumpType.XD_LAND)
		If etTable Is Nothing Then
			EmpireError("ParseXDLand", "The meta data was not found for land", strLine)
			rsLand.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDLand", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsLand.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Sector information to get a record.
		iOwnerIndex = etTable.GetZeroIndex("owner")
		iIDIndex = etTable.GetZeroIndex("uid")
		iXlocIndex = etTable.GetZeroIndex("xloc")
		iYlocIndex = etTable.GetZeroIndex("yloc")
		iOpxIndex = etTable.GetZeroIndex("opx")
		iOpyIndex = etTable.GetZeroIndex("opy")
		iRadiusIndex = etTable.GetZeroIndex("radius")
		iTypeIndex = etTable.GetZeroIndex("type")
		
		If iIDIndex = -1 Or iOwnerIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or iOpxIndex = -1 Or iOpyIndex = -1 Or iRadiusIndex = -1 Or iTypeIndex = -1 Then
			EmpireError("ParseXDLand", "One of the following fields is missing: owner, id, xloc, yloc, opx, opy, rad, tech", "")
			rsLand.Index = hldIndex
			Exit Sub
		End If
		
		If rsLand.BOF And rsLand.EOF Then
			rsLand.AddNew()
		Else
			rsLand.Seek("=", strParams(iIDIndex))
			If rsLand.NoMatch Then
				rsLand.AddNew()
			Else
				rsLand.Edit()
			End If
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "owner"
					rsLand.Fields("Country").Value = strParams(iIndex)
				Case "uid"
					rsLand.Fields("id").Value = strParams(iIndex)
				Case "type"
					rsLand.Fields("type").Value = GetUnitType(Val(strParams(iIndex)), "l")
				Case "army"
					If estsTable(enumXdumpType.XD_META_TYPE)((etTable.Column(iHeaderIndex).ColType)).Name = "c" Then
						strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
						If strParams(iIndex) = "" Or strParams(iIndex) = " " Then 'replace space with ~
							rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = "~"
						Else
							rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = Left(strParams(iIndex), 1)
						End If
					Else
						If CDbl(strParams(iIndex)) = 32 Then 'replace space with ~
							rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = "~"
						Else
							rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = Chr(CInt(strParams(iIndex)))
						End If
					End If
				Case "xloc"
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(rsLand.Fields("x").Value) Then
						iPrevXloc = -1
					Else
						iPrevXloc = rsLand.Fields("x").Value
					End If
					rsLand.Fields("x").Value = strParams(iIndex)
				Case "yloc"
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDbNull(rsLand.Fields("y").Value) Then
						iPrevYloc = -1
					Else
						iPrevYloc = rsLand.Fields("y").Value
					End If
					rsLand.Fields("y").Value = strParams(iIndex)
				Case "effic"
					rsLand.Fields("eff").Value = strParams(iIndex)
				Case "mobil"
					rsLand.Fields("mob").Value = strParams(iIndex)
				Case "ship"
					rsLand.Fields("ship").Value = strParams(iIndex)
				Case "harden"
					rsLand.Fields("fort").Value = strParams(iIndex)
				Case "nxlight"
					rsLand.Fields("xl").Value = strParams(iIndex)
				Case "civil"
					rsLand.Fields("civ").Value = strParams(iIndex)
				Case "milit"
					rsLand.Fields("mil").Value = strParams(iIndex)
				Case "nxlight"
					rsLand.Fields("xl").Value = strParams(iIndex)
				Case "retreat"
					rsLand.Fields("retr").Value = strParams(iIndex)
				Case "name"
					rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
				Case "tech"
					rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					hldChrIndex = rsBuild.Index
					rsBuild.Index = "PrimaryKey"
					
					rsBuild.Seek("=", "l", GetUnitType(Val(strParams(iTypeIndex)), "l"))
					If Not rsBuild.NoMatch Then
						rsLand.Fields("att").Value = System.Math.Round(ComputeLandAttDef(rsBuild.Fields("base att").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value), 1)
						rsLand.Fields("def").Value = System.Math.Round(ComputeLandAttDef(rsBuild.Fields("base def").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value), 1)
						rsLand.Fields("vul").Value = Fix(ComputeLandVulerability(rsBuild.Fields("base vul").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						rsLand.Fields("spd").Value = Fix(ComputeLandSpeed(rsBuild.Fields("base spd").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						rsLand.Fields("vis").Value = rsBuild.Fields("stat5").Value
						rsLand.Fields("spy").Value = rsBuild.Fields("stat6").Value
						rsLand.Fields("radius").Value = rsBuild.Fields("stat7").Value
						rsLand.Fields("frg").Value = Fix(ComputeLandFiringRange(rsBuild.Fields("base frnge").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						rsLand.Fields("acc").Value = Fix(ComputeLandAccuracy(rsBuild.Fields("base acc").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						rsLand.Fields("dam").Value = Fix(ComputeLandDamage(rsBuild.Fields("base dam").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						rsLand.Fields("amm").Value = Fix(ComputeLandAmmo(rsBuild.Fields("base load").Value, rsBuild.Fields("base dam").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						rsLand.Fields("aaf").Value = Fix(ComputeLandAAF(rsBuild.Fields("base aaf").Value, rsBuild.Fields("tech").Value, rsLand.Fields("tech").Value))
						'    lp->lnd_fuelc = (int)LND_FC(lcp->l_fuelc, tech_diff);
						'    lp->lnd_fuelu = (int)LND_FU(lcp->l_fuelu, tech_diff);
						'    lp->lnd_maxlight = (int)LND_XPL(lcp->l_nxlight, tech_diff);
						'    lp->lnd_maxland = (int)LND_MXL(lcp->l_mxland, tech_diff);
						'        Else
						'            EmpireError "ParseXDLand", "Land Characteristics not found:" + rsLand.Fields("type"), strLine
					End If
					rsBuild.Index = hldChrIndex
				Case "timestamp", "react", "uw", "food", "fuel", "shell", "gun", "petrol", "iron", "dust", "bar", "oil", "lcm", "hcm", "rad", "nland", "land"
					rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "opx", "opy", "radius", "flags", "pstage", "ptime", "access", "rpath", "rflags"
				Case "mission"
					AddMission(Val(strParams(iIndex)), Val(strParams(iIDIndex)), Val(strParams(iOwnerIndex)), "l", SectorString(Val(strParams(iOpxIndex)), Val(strParams(iOpyIndex))), Val(strParams(iRadiusIndex)))
				Case "off"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						rsLand.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					Else
						EmpireError("ParseXDLand", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDLand", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		If iOwnerIndex = -1 Then
			rsLand.Fields("country").Value = CountryNumber
		End If
		
		rsLand.Update()
		rsLand.Index = hldIndex
		
		If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
			frmDrawMap.FillGrid()
		End If
		
		CheckForOldEnemyUnit(CShort(strParams(iIDIndex)), "l")
		Exit Sub
		
		'Error handling routine
ParseXDLand_Exit: 
		EmpireError("ParseXDLand", strErrorField, strLine)
		rsLand.Index = hldIndex
	End Sub
	
	Private Sub ParseXDLandChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iCapIndex As Short
		Dim iIDIndex As Short
		Dim iTechIndex As Short
		Dim iDamageIndex As Short
		Dim iPos As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		
		'Save and restore the index on entry and exit
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		
		On Error GoTo ParseXDLandChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_LAND_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDLandChr", "The meta data was not found for item", strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDLandChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		'Get the ID information to get a record.
		iIDIndex = etTable.GetZeroIndex("name")
		strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
		iTechIndex = etTable.GetZeroIndex("tech")
		iDamageIndex = etTable.GetZeroIndex("dam")
		
		If iIDIndex = -1 Then
			EmpireError("ParseXDLandChr", "The record identifier field not present", CStr(iIDIndex))
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		iPos = InStr(strParams(iIDIndex), " ")
		'
		If rsBuild.BOF And rsBuild.EOF Then
			rsBuild.AddNew()
			rsBuild.Fields("type").Value = "l"
		Else
			rsBuild.Seek("=", "l", Left(strParams(iIDIndex), iPos - 1))
			If rsBuild.NoMatch Then
				rsBuild.AddNew()
				rsBuild.Fields("type").Value = "l"
			Else
				rsBuild.Edit()
			End If
		End If
		
		rsBuild.Fields("gun").Value = 0
		
		iHeaderIndex = 1
		iCapIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					If iPos > 0 Then
						rsBuild.Fields("id").Value = Left(strParams(iIndex), iPos - 1)
						rsBuild.Fields("desc").Value = LTrim(Right(strParams(iIndex), Len(strParams(iIndex)) - iPos))
					Else
						EmpireError("ParseXDLandChr", "Incorrect format of name", strLine)
					End If
				Case "cost", "tech"
					rsBuild.Fields(etTable.Column(iHeaderIndex).Name).Value = Val(strParams(iIndex))
				Case "l_build"
					rsBuild.Fields("lcm").Value = Val(strParams(iIndex))
				Case "h_build"
					rsBuild.Fields("hcm").Value = Val(strParams(iIndex))
				Case "g_build"
					rsBuild.Fields("gun").Value = Val(strParams(iIndex))
				Case "s_build"
					'shells not used yet
				Case "civil"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "c"
						iCapIndex = iCapIndex + 1
					End If
				Case "milit"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "m"
						iCapIndex = iCapIndex + 1
					End If
				Case "shell"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "s"
						iCapIndex = iCapIndex + 1
					End If
				Case "gun"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "g"
						iCapIndex = iCapIndex + 1
					End If
				Case "petrol"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "p"
						iCapIndex = iCapIndex + 1
					End If
				Case "iron"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "i"
						iCapIndex = iCapIndex + 1
					End If
				Case "dust"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "d"
						iCapIndex = iCapIndex + 1
					End If
				Case "bar"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "b"
						iCapIndex = iCapIndex + 1
					End If
				Case "food"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "f"
						iCapIndex = iCapIndex + 1
					End If
				Case "oil"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "o"
						iCapIndex = iCapIndex + 1
					End If
				Case "lcm"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "l"
						iCapIndex = iCapIndex + 1
					End If
				Case "hcm"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "h"
						iCapIndex = iCapIndex + 1
					End If
				Case "uw"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "u"
						iCapIndex = iCapIndex + 1
					End If
				Case "rad"
					If CDbl(strParams(iIndex)) > 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = strParams(iIndex) & "r"
						iCapIndex = iCapIndex + 1
					End If
				Case "flags"
					If (strParams(iIndex) And &H1) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "xlight"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H2) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "engineer"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H4) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "supply"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H8) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "security"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H10) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "light"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H20) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "marine"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H40) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "recon"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H80) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "radar"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H100) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "assault"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H200) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "flak"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H400) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "spy"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H800) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "train"
						iCapIndex = iCapIndex + 1
					End If
					If (strParams(iIndex) And &H1000) <> 0 Then
						rsBuild.Fields("cap" & CStr(iCapIndex)).Value = "heavy"
						iCapIndex = iCapIndex + 1
					End If
				Case "att"
					rsBuild.Fields("base att").Value = Val(strParams(iIndex))
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat1").Value = System.Math.Round(ComputeLandAttDef(Val(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()), 1)
						'            rsBuild.Fields("stat1") = Fix(ComputeLandAttDef(Val(strParams(iIndex)), CInt(strParams(iTechIndex)), CInt(rsNation.Fields("TechLevel"))) * 10) / 10
					End If
				Case "def"
					rsBuild.Fields("base def").Value = Val(strParams(iIndex))
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat2").Value = System.Math.Round(ComputeLandAttDef(Val(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()), 1)
						'            rsBuild.Fields("stat2") = Fix(ComputeLandAttDef(Val(strParams(iIndex)), CInt(strParams(iTechIndex)), CInt(rsNation.Fields("TechLevel"))) * 10) / 10
					End If
				Case "vul"
					rsBuild.Fields("base vul").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat3").Value = Fix(ComputeLandVulerability(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "spd"
					rsBuild.Fields("base spd").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat4").Value = Fix(ComputeLandSpeed(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "vis"
					rsBuild.Fields("stat5").Value = strParams(iIndex)
				Case "spy"
					rsBuild.Fields("stat6").Value = strParams(iIndex)
				Case "rmax"
					rsBuild.Fields("stat7").Value = strParams(iIndex)
				Case "frg"
					rsBuild.Fields("base frnge").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat8").Value = Fix(ComputeLandFiringRange(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "acc"
					rsBuild.Fields("base acc").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat9").Value = Fix(ComputeLandAccuracy(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "dam"
					rsBuild.Fields("base dam").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat10").Value = Fix(ComputeLandDamage(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "ammo"
					rsBuild.Fields("base load").Value = strParams(iIndex)
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") And CheckIndex(iDamageIndex, "ParseXDLandChr", "damage") Then
						rsBuild.Fields("stat11").Value = Fix(ComputeLandAmmo(CShort(strParams(iIndex)), CShort(strParams(iDamageIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "aaf"
					rsBuild.Fields("base aaf").Value = strParams(iIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iTechIndex, ParseXDLandChr, tech). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
						rsBuild.Fields("stat12").Value = Fix(ComputeLandAAF(CShort(strParams(iIndex)), CShort(strParams(iTechIndex)), GetTechLevelForShow()))
					End If
				Case "fuelc"
					rsBuild.Fields("stat13").Value = strParams(iIndex)
				Case "fuelu"
					rsBuild.Fields("stat14").Value = strParams(iIndex)
				Case "nxlight"
					rsBuild.Fields("stat15").Value = strParams(iIndex)
				Case "nland"
					rsBuild.Fields("stat16").Value = strParams(iIndex)
				Case "type"
					rsBuild.Fields("chr_i").Value = strParams(iIndex)
				Case Else
					EmpireError("ParseXDLandChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		rsBuild.Fields("avail").Value = CalculateAvail(CShort(rsBuild.Fields("lcm").Value), CShort(rsBuild.Fields("hcm").Value))
		
		rsBuild.Update()
		rsBuild.Index = hldIndex
		Exit Sub
		
		'Error handling routine
ParseXDLandChr_Exit: 
		EmpireError("ParseXDLandChr", strErrorField, strLine)
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub ParseXDNuke(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iOwnerIndex As Short
		Dim iIDIndex As Short
		Dim iTypeIndex As Short
		Dim iXlocIndex As Short
		Dim iYlocIndex As Short
		Dim iOffIndex As Short
		Dim iTypeLoop As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		Dim iPrevXloc As Short
		Dim iPrevYloc As Short
		
		hldIndex = rsNuke.Index
		rsNuke.Index = "PrimaryKey"
		
		On Error GoTo ParseXDNuke_Exit
		
		etTable = etsTable(enumXdumpType.XD_NUKE)
		If etTable Is Nothing Then
			EmpireError("ParseXDNuke", "The meta data was not found for nuke", strLine)
			rsNuke.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDNuke", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsNuke.Index = hldIndex
			Exit Sub
		End If
		
		'Get the Sector information to get a record.
		iOwnerIndex = etTable.GetZeroIndex("owner")
		iIDIndex = etTable.GetZeroIndex("uid")
		If VersionCheck(4, 3, 3) <> WinAceRoutines.enumVersion.VER_LESS Then
			iTypeIndex = etTable.GetZeroIndex("type")
		Else
			iTypeIndex = etTable.GetZeroIndex("types")
		End If
		iXlocIndex = etTable.GetZeroIndex("xloc")
		iYlocIndex = etTable.GetZeroIndex("yloc")
		iOffIndex = etTable.GetZeroIndex("off")
		
		If iIDIndex = -1 Or iTypeIndex = -1 Or iOwnerIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or (iOffIndex = -1 And VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS) Then
			EmpireError("ParseXDNuke", "One of the following fields is missing: ", "owner, uid, type(s), xloc, yloc")
			rsNuke.Index = hldIndex
			Exit Sub
		End If
		If VersionCheck(4, 3, 3) <> WinAceRoutines.enumVersion.VER_LESS Then
			rsNuke.Seek("=", strParams(iIDIndex))
			If rsNuke.NoMatch Then
				rsNuke.AddNew()
			Else
				rsNuke.Edit()
			End If
			strErrorField = "owner"
			rsNuke.Fields("country").Value = strParams(iOwnerIndex)
			strErrorField = "id"
			rsNuke.Fields("id").Value = strParams(iIDIndex)
			strErrorField = "x"
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsNuke.Fields("x").Value) Then
				iPrevXloc = -1
			Else
				iPrevXloc = rsNuke.Fields("x").Value
			End If
			rsNuke.Fields("x").Value = strParams(iXlocIndex)
			strErrorField = "y"
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsNuke.Fields("y").Value) Then
				iPrevYloc = -1
			Else
				iPrevYloc = rsNuke.Fields("y").Value
			End If
			rsNuke.Fields("y").Value = strParams(iYlocIndex)
			strErrorField = "type"
			rsNuke.Fields("type").Value = GetUnitType(CShort(strParams(iTypeIndex)), "n")
			rsNuke.Fields("num").Value = 1
			If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
				rsNuke.Fields("off").Value = strParams(iOffIndex)
			End If
			rsNuke.Update()
		Else
			For iTypeLoop = 0 To etTable.Column("types").Length - 1
				If CDbl(strParams(iTypeIndex + iTypeLoop)) > 0 Then
					'Save and restore the index on entry and exit
					rsNuke.Seek("=", strParams(iIDIndex), GetUnitType(iTypeLoop, "n"))
					If rsNuke.NoMatch Then
						rsNuke.AddNew()
					Else
						rsNuke.Edit()
					End If
					strErrorField = GetUnitType(iTypeLoop, "n")
					rsNuke.Fields("country").Value = strParams(iOwnerIndex)
					rsNuke.Fields("id").Value = strParams(iIDIndex)
					rsNuke.Fields("x").Value = strParams(iXlocIndex)
					rsNuke.Fields("y").Value = strParams(iYlocIndex)
					rsNuke.Fields("type").Value = GetUnitType(iTypeLoop, "n")
					rsNuke.Fields("num").Value = strParams(iTypeIndex + iTypeLoop)
					rsNuke.Update()
				Else
					rsNuke.Seek("=", strParams(iIDIndex), GetUnitType(iTypeLoop, "n"))
					If Not rsNuke.NoMatch Then
						rsNuke.Delete()
					End If
				End If
			Next 
		End If
		rsNuke.Index = hldIndex
		
		If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
			frmDrawMap.FillGrid()
		End If
		
		CheckForOldEnemyUnit(CShort(strParams(iIDIndex)), "n")
		Exit Sub
		
		'Error handling routine
ParseXDNuke_Exit: 
		EmpireError("ParseXDNuke", strErrorField, strLine)
		rsNuke.Index = hldIndex
	End Sub
	
	Private Sub ParseXDNukeChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iIDIndex As Short
		Dim iTechIndex As Short
		Dim iPos As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		Dim sDRNukeConst As Single
		
		'Save and restore the index on entry and exit
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		
		On Error GoTo ParseXDNukeChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_NUKE_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDNukeChr", "The meta data was not found for nuke-chr", strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDNukeChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		'Get the ID information to get a record.
		iIDIndex = etTable.GetZeroIndex("name")
		strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
		iTechIndex = etTable.GetZeroIndex("tech")
		
		If iIDIndex = -1 Then
			EmpireError("ParseXDNukeChr", "The record identifier field not present", CStr(iIDIndex))
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		iPos = InStr(strParams(iIDIndex), " ")
		'
		If rsBuild.BOF And rsBuild.EOF Then
			rsBuild.AddNew()
			rsBuild.Fields("type").Value = "n"
		Else
			rsBuild.Seek("=", "n", Left(Left(strParams(iIDIndex), iPos - 1), 5))
			If rsBuild.NoMatch Then
				rsBuild.AddNew()
				rsBuild.Fields("type").Value = "n"
			Else
				rsBuild.Edit()
			End If
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					If iPos > 0 Then
						rsBuild.Fields("id").Value = Left(Left(strParams(iIndex), iPos - 1), 5)
						rsBuild.Fields("desc").Value = LTrim(Mid(strParams(iIndex), iPos + 1, Len(strParams(iIndex)) - iPos))
					Else
						EmpireError("ParseXDNukeChr", "Incorrect format of name", strLine)
					End If
				Case "cost", "avail"
					rsBuild.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
				Case "l_build"
					rsBuild.Fields("lcm").Value = strParams(iIndex)
				Case "h_build"
					rsBuild.Fields("hcm").Value = strParams(iIndex)
				Case "o_build"
					rsBuild.Fields("oil").Value = strParams(iIndex)
				Case "r_build"
					rsBuild.Fields("rad").Value = strParams(iIndex)
				Case "flags"
					If (strParams(iIndex) And &H1) <> 0 Then
						rsBuild.Fields("cap1").Value = "neutron"
					Else
						rsBuild.Fields("cap1").Value = ""
					End If
				Case "blast"
					rsBuild.Fields("stat1").Value = strParams(iIndex)
				Case "dam"
					rsBuild.Fields("stat2").Value = strParams(iIndex)
				Case "weight"
					rsBuild.Fields("stat3").Value = strParams(iIndex)
				Case "tech"
					rsBuild.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
					rsBuild.Fields("stat4").Value = strParams(iIndex)
					If rsVersion.BOF And rsVersion.EOF Then
						sDRNukeConst = 0#
					Else
						sDRNukeConst = rsVersion.Fields("DRNUKE Const").Value
					End If
					If sDRNukeConst > 0# Then
						rsBuild.Fields("stat5").Value = Int(CDbl(strParams(iIndex)) * sDRNukeConst) + 1
					Else
						rsBuild.Fields("stat5").Value = 0
					End If
				Case "type"
					rsBuild.Fields("chr_i").Value = strParams(iIndex)
				Case Else
					EmpireError("ParseXDNukeChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		rsBuild.Update()
		rsBuild.Index = hldIndex
		Exit Sub
		
		'Error handling routine
ParseXDNukeChr_Exit: 
		EmpireError("ParseXDNukeChr", strErrorField, strLine)
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub ParseXDItemChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIDIndex As Short
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim strErrorField As String
		Dim eiChr As EmpItem
		Dim bNew As Boolean
		Dim etTable As EmpTable
		
		On Error GoTo ParseXDItemChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_ITEM_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDItemChr", "The meta data was not found for item-chr", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDItemChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		'Get the ID information to get a record.
		iIDIndex = etTable.GetZeroIndex("mnem")
		
		If iIDIndex = -1 Then
			EmpireError("ParseXDItemChr", "The record identifier field not present", CStr(iIDIndex))
			Exit Sub
		End If
		
		strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
		
		If strParams(iIDIndex) = "0" Then 'skip "unused" row
			Exit Sub
		End If
		
		eiChr = Items.FindByLetter(strParams(iIDIndex))
		If eiChr Is Nothing Then
			eiChr = New EmpItem
			bNew = True
		Else
			bNew = False
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					eiChr.ItemName = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
				Case "value"
					eiChr.Value = CShort(strParams(iIndex))
				Case "sell"
					If CDbl(strParams(iIndex)) = 0 Then
						eiChr.Sellable = False
					Else
						eiChr.Sellable = True
					End If
				Case "lbs"
					eiChr.Weight = CShort(strParams(iIndex))
				Case "uid"
					eiChr.ChrIndex = CShort(strParams(iIndex))
				Case "mnem"
					eiChr.Letter = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
				Case "pkg"
					eiChr.Packing(EmpItem.enumItemPacking.PACKING_INEFFICIENT) = CShort(strParams(iIndex))
					eiChr.Packing(EmpItem.enumItemPacking.PACKING_REGULAR) = CShort(strParams(iIndex + 1))
					eiChr.Packing(EmpItem.enumItemPacking.PACKING_WAREHOUSE) = CShort(strParams(iIndex + 2))
					eiChr.Packing(EmpItem.enumItemPacking.PACKING_URBAN) = CShort(strParams(iIndex + 3))
					eiChr.Packing(EmpItem.enumItemPacking.PACKING_BANK) = CShort(strParams(iIndex + 4))
					iIndex = iIndex + 4
				Case "melt_denom"
				Case Else
					EmpireError("ParseXDItemChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		eiChr.ChrIndex = iRowCount
		If bNew Then
			eiChr.IntelligenceNames = eiChr.ItemName
			eiChr.ProductName = eiChr.ItemName
			eiChr.FormName = eiChr.ItemName
			eiChr.ConditionalName = Left(eiChr.ItemName, 3)
			eiChr.DatabaseName = Left(eiChr.ItemName, 3)
			eiChr.SectorPanelLabel = UCase(Left(eiChr.ItemName, 1)) & Mid(eiChr.ItemName, 2, 2) & ":"
			eiChr.DistributionPanelLabel = UCase(Left(eiChr.ItemName, 1)) & Mid(eiChr.ItemName, 2, 2)
			Items.AddItem((eiChr))
		End If
		
		iRowCount = iRowCount + 1
		Exit Sub
		
		'Error handling routine
ParseXDItemChr_Exit: 
		EmpireError("ParseXDItemChr", strErrorField, strLine)
	End Sub
	
	Private Sub ParseXDInfraChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iIDIndex As Short
		Dim iPos As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		Dim bEnable As Boolean
		
		'Save and restore the index on entry and exit
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		
		On Error GoTo ParseXDInfraChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_INFRA_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDInfraChr", "The meta data was not found for infra", strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDInfraChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		'Get the ID information to get a record.
		iIDIndex = etTable.GetZeroIndex("name")
		If iIDIndex = -1 Then
			EmpireError("ParseXDInfraChr", "The record identifier field not present", CStr(iIDIndex))
			rsBuild.Index = hldIndex
			Exit Sub
		End If
		
		strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
		iPos = InStr(strParams(iIDIndex), " ")
		If iPos > 0 Then
			strParams(iIDIndex) = Left(strParams(iIDIndex), iPos - 1)
		End If
		If Len(strParams(iIDIndex)) > 5 Then
			strParams(iIDIndex) = Left(strParams(iIDIndex), 5)
		End If
		
		If rsBuild.BOF And rsBuild.EOF Then
			rsBuild.AddNew()
			rsBuild.Fields("type").Value = "i"
		Else
			rsBuild.Seek("=", "i", strParams(iIDIndex))
			If rsBuild.NoMatch Then
				rsBuild.AddNew()
				rsBuild.Fields("type").Value = "i"
			Else
				rsBuild.Edit()
			End If
		End If
		
		iHeaderIndex = 1
		bEnable = True
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "name"
					rsBuild.Fields("id").Value = strParams(iIndex)
				Case "dcost"
					rsBuild.Fields("cost").Value = strParams(iIndex)
				Case "lcms"
					rsBuild.Fields("lcm").Value = strParams(iIndex)
				Case "hcms"
					rsBuild.Fields("hcm").Value = strParams(iIndex)
				Case "mcost"
					rsBuild.Fields("stat2").Value = strParams(iIndex)
				Case "enable"
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						If strParams(iIndex) = "0" Then
							If Not rsBuild.NoMatch Then
								rsBuild.Delete()
							Else
								rsBuild.CancelUpdate()
							End If
							Exit Sub
						End If
					Else
						EmpireError("ParseXDLandChr", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDLandChr", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		
		rsBuild.Update()
		rsBuild.Index = hldIndex
		Exit Sub
		
		'Command: xdump meta infra
		'XDUMP meta infrastructure 1145750521
		'"name" 3 4 0 -1
		'"lcms" 6 0 0 -1
		'"hcms" 6 0 0 -1
		'"dcost" 6 0 0 -1
		'"mcost" 6 0 0 -1
		
		'Error handling routine
ParseXDInfraChr_Exit: 
		EmpireError("ParseXDInfraChr", strErrorField, strLine)
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub ParseXDProductChr(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		
		'Save and restore the index on entry and exit
		hldIndex = rsSectorType.Index
		rsSectorType.Index = "PrimaryKey"
		
		On Error GoTo ParseXDProductChr_Exit
		
		etTable = etsTable(enumXdumpType.XD_PRODUCT_CHR)
		If etTable Is Nothing Then
			EmpireError("ParseXDProductChr", "The meta data was not found for product-chr", strLine)
			rsSectorType.Index = hldIndex
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDProductChr", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			rsSectorType.Index = hldIndex
			Exit Sub
		End If
		
		rsSectorType.MoveFirst()
		While Not rsSectorType.EOF And iRowCount <> 0
			If iRowCount = rsSectorType.Fields("pchr_i").Value Then
				rsSectorType.Edit()
				iHeaderIndex = 1
				For iIndex = 0 To UBound(strParams)
					strErrorField = etTable.Column(iHeaderIndex).Name
					Select Case etTable.Column(iHeaderIndex).Name
						Case "sname"
							rsSectorType.Fields("product").Value = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
						Case "name"
						Case "ctype"
							rsSectorType.Fields("use1s").Value = GetType_Renamed(Val(strParams(iIndex)))
							iIndex = iIndex + 1
							rsSectorType.Fields("use2s").Value = GetType_Renamed(Val(strParams(iIndex)))
							iIndex = iIndex + 1
							rsSectorType.Fields("use3s").Value = GetType_Renamed(Val(strParams(iIndex)))
						Case "camt"
							rsSectorType.Fields("use1n").Value = strParams(iIndex)
							iIndex = iIndex + 1
							rsSectorType.Fields("use2n").Value = strParams(iIndex)
							iIndex = iIndex + 1
							rsSectorType.Fields("use3n").Value = strParams(iIndex)
						Case "nlndx"
							rsSectorType.Fields("level").Value = GetNationType(Val(strParams(iIndex)))
						Case "cost"
							rsSectorType.Fields(etTable.Column(iHeaderIndex).Name).Value = strParams(iIndex)
						Case "nrndx", "level", "uid"
						Case "type"
							rsSectorType.Fields("pcode").Value = GetType_Renamed(Val(strParams(iIndex)))
						Case "nlmin"
							rsSectorType.Fields("min").Value = strParams(iIndex)
						Case "effic"
							If VersionCheck(4, 3, 6) = WinAceRoutines.enumVersion.VER_LESS Then
								rsSectorType.Fields("eff").Value = strParams(iIndex)
							Else
								EmpireError("ParseXDProductChr", "Field not understood:" & strErrorField, strLine)
							End If
						Case "nllag"
							rsSectorType.Fields("lag").Value = strParams(iIndex)
						Case "nrdep"
							rsSectorType.Fields("dep").Value = strParams(iIndex)
						Case Else
							EmpireError("ParseXDProductChr", "Field not understood:" & strErrorField, strLine)
					End Select
					iHeaderIndex = iHeaderIndex + 1
				Next iIndex
				rsSectorType.Update()
			End If
			rsSectorType.MoveNext()
		End While
		rsSectorType.Index = hldIndex
		iRowCount = iRowCount + 1
		Exit Sub
		
		'Error handling routine
ParseXDProductChr_Exit: 
		EmpireError("ParseXDProductChr", strErrorField, strLine)
		rsSectorType.Index = hldIndex
	End Sub
	
	Private Sub ParseXDGame(ByRef strLine As String)
		Dim strParams() As String
		Dim etTable As EmpTable
		Dim strErrorField As String
		Dim iIndex As Short
		Static iTurn As Short
		
		On Error GoTo ParseXDGame_Exit
		
		etTable = etsTable(enumXdumpType.XD_GAME)
		If etTable Is Nothing Then
			EmpireError("ParseXDGame", "The meta data was not found for game", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDGame", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iIndex + 1).Name
			Select Case etTable.Column(iIndex + 1).Name
				Case "turn"
					If iTurn = 0 Then
						iTurn = CShort(strParams(iIndex))
					ElseIf iTurn <> CShort(strParams(iIndex)) And frmEmpCmd.bAutoUpdate Then 
						iTurn = CShort(strParams(iIndex))
						UpdateCommands()
					End If
				Case "upd_disable"
					frmEmpCmd.bUpdateEnabled = Not CBool(strParams(iIndex))
				Case "tick", "rt", "down"
					'ignore
				Case Else
					EmpireError("ParseXDGame", "Field not understood:" & strErrorField, strLine)
			End Select
		Next 
		Exit Sub
		
ParseXDGame_Exit: 
		EmpireError("ParseXDGame", strErrorField, strLine)
	End Sub
	
	
	Private Sub ParseXDUpdates(ByRef strLine As String)
		Dim strParams() As String
		Dim etTable As EmpTable
		Dim strErrorField As String
		Dim iIndex As Short
		Dim lTimeUpdate As Integer
		Dim lTimeCurrent As Integer
		
		If Not frmEmpCmd.bSetUpdateTime Then
			Exit Sub
		End If
		frmEmpCmd.bSetUpdateTime = False
		
		On Error GoTo ParseXDUpdates_Exit
		
		etTable = etsTable(enumXdumpType.XD_UPDATES)
		If etTable Is Nothing Then
			EmpireError("ParseXDUpdates", "The meta data was not found for update", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDUpdates", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iIndex + 1).Name
			Select Case etTable.Column(iIndex + 1).Name
				Case "time"
					lTimeUpdate = CInt(strParams(iIndex))
					If lTimeUpdate > 0 Then
						lTimeCurrent = CInt(tmStamp)
						frmDrawMap.TimerAtUpdate = VB.Timer()
						frmDrawMap.SecondsToUpdate = lTimeUpdate - lTimeCurrent
						frmDrawMap.UpdateTimer.Interval = 1000
						frmDrawMap.UpdateTimer.Enabled = True
					Else
						frmDrawMap.TimerAtUpdate = 0
						frmDrawMap.SecondsToUpdate = 0
					End If
				Case Else
					EmpireError("ParseXDUpdates", "Field not understood:" & strErrorField, strLine)
			End Select
		Next 
		Exit Sub
		
ParseXDUpdates_Exit: 
		EmpireError("ParseXDUpdates", strErrorField, strLine)
	End Sub
	
	Private Sub ParseXDVersion(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iETUupdate As Short
		Dim iBuildBridgeHcm As Short
		Dim iBuildTowerHcm As Short
		Dim strErrorField As String
		Dim etTable As EmpTable
		Dim bXDumpServer As Boolean
		
		On Error GoTo ParseXDVersion_Exit
		
		etTable = etsTable(enumXdumpType.XD_VERSION)
		If etTable Is Nothing Then
			EmpireError("ParseXDVersion", "The meta data was not found for version", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDVersion", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		iETUupdate = etTable.GetZeroIndex("etu_per_update")
		iBuildBridgeHcm = etTable.GetZeroIndex("buil_bh")
		iBuildTowerHcm = etTable.GetZeroIndex("buil_tower_bh")
		
		If VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
			bXDumpServer = True
		Else
			bXDumpServer = False
		End If
		
		If rsVersion.RecordCount = 0 Then
			rsVersion.AddNew()
		Else
			rsVersion.MoveFirst()
			rsVersion.Edit()
		End If
		iHeaderIndex = 1
		Dim strVersions() As String
		Dim iPos As Short
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "privname", "privlog", "update_policy", "update_window", "update_times", "blitz_time", "update_demandpolicy", "update_wantmin", "update_missed", "update_demandtimes", "game_days", "game_hours", "ALL_BLEED", "AUTO_POWER", "BLITZ", "BRIDGETOWERS", "DEFENSE_INFRA", "DEMANDUPDATE", "DRNUKE", "EASY_BRIDGES", "FALLOUT", "FUEL", "GODNEWS", "GUINEA_PIGS", "HIDDEN", "INTERDICT_ATT", "LANDSPIES", "LOANS", "LOSE_CONTACT", "MARKET", "MOB_ACCESS", "NOMOBCOST", "NO_FORT_FIRE", "NO_PLAGUE", "PINPOINTMISSILE", "RES_POP", "SAIL", "SECT_MOB_ACCESS", "SHOWPLANE", "SLOW_WAR", "SUPER_BARS", "TECH_POP", "TRADESHIPS", "TREATIES", "UNIT_MOB_ACCESS", "m_m_p_d", "max_btus", "max_idle", "players_at_00", "war_cost", "ally_factor", "money_land", "money_plane", "money_ship", "torpedo_damage", "fort_max_interdiction_range", "land_max_interdiction_range", "ship_max_interdiction_range", "flakscale", "combat_mob", "people_damage", "unit_damage", "collateral_dam", "assault_penalty", "sect_mob_neg_factor", "unit_mob_neg_factor", "mission_mob_cost", "decay_per_etu", "fallout_spread", "MARK_DELAY", "TRADE_DELAY", "buytax", "tradetax", "trade_1_dist", "trade_2_dist", "trade_3_dist", "trade_1", "trade_2", "trade_3", "trade_ally_bonus", "trade_ally_cut", "fuel_mult", "update_demand", "maxnoc", "max_idle_visitor", "login_grace_time"
				Case "version"
					strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
					iPos = InStrRev(strParams(iIndex), " ")
					If Len(strParams(iIndex)) > 1 And iPos > 0 Then
						strVersions = Split(Right(strParams(iIndex), Len(strParams(iIndex)) - iPos + 1), ".")
						If UBound(strVersions) = 2 Then
							rsVersion.Fields("Major").Value = strVersions(0)
							rsVersion.Fields("Minor").Value = strVersions(1)
							rsVersion.Fields("Revision").Value = strVersions(2)
						Else
							EmpireError("ParseXDVersion", "Wrong number of fields in version:" & strErrorField, strLine)
						End If
					Else
						EmpireError("ParseXDVersion", "Unexpected format in version:" & strErrorField, strLine)
					End If
				Case "WORLD_X"
					rsVersion.Fields("World X").Value = strParams(iIndex)
				Case "WORLD_Y"
					rsVersion.Fields("World Y").Value = strParams(iIndex)
				Case "s_p_etu"
					rsVersion.Fields("Empire Time Unit").Value = strParams(iIndex)
				Case "etu_per_update"
					rsVersion.Fields("ETU per update").Value = strParams(iIndex)
				Case "btu_build_rate"
					rsVersion.Fields("Civs for BTUs").Value = (1# / (Val(strParams(iIndex)) * 100#))
				Case "fgrate"
					rsVersion.Fields("Food grow").Value = Val(strParams(iIndex)) * 100#
				Case "fcrate"
					rsVersion.Fields("Food harvest").Value = Val(strParams(iIndex)) * 1000#
				Case "obrate"
					rsVersion.Fields("Civ Birthrate").Value = Val(strParams(iIndex)) * 1000#
				Case "uwbrate"
					rsVersion.Fields("Uw Birthrate").Value = Val(strParams(iIndex)) * 1000#
				Case "eatrate"
					rsVersion.Fields("Adult food").Value = Val(strParams(iIndex)) * 1000#
				Case "babyeat"
					rsVersion.Fields("Baby food").Value = Val(strParams(iIndex)) * 1000#
				Case "bankint"
					rsVersion.Fields("Bank Interest").Value = Val(strParams(iIndex)) * 1000#
				Case "money_civ"
					rsVersion.Fields("Civ Taxes").Value = Val(strParams(iIndex)) * 1000#
				Case "money_uw"
					rsVersion.Fields("UW Taxes").Value = Val(strParams(iIndex)) * 1000#
				Case "money_mil"
					rsVersion.Fields("Mil Salary").Value = -Val(strParams(iIndex)) * 1000#
				Case "money_res"
					rsVersion.Fields("Reserve Salary").Value = -Val(strParams(iIndex)) * 1000#
				Case "rollover_avail_max"
					rsVersion.Fields("Rollover Avail").Value = strParams(iIndex)
				Case "hap_cons"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						rsVersion.Fields("Civs per stroller").Value = Val(strParams(iIndex)) / Val(strParams(iETUupdate))
					End If
				Case "edu_cons"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						rsVersion.Fields("Civs per graduate").Value = Val(strParams(iIndex)) / Val(strParams(iETUupdate))
					End If
				Case "hap_avg"
					rsVersion.Fields("Happy average").Value = Val(strParams(iIndex))
				Case "edu_avg"
					rsVersion.Fields("Edu average").Value = Val(strParams(iIndex))
				Case "tech_log_base"
					rsVersion.Fields("Tech base").Value = Val(strParams(iIndex))
				Case "easy_tech"
					rsVersion.Fields("Easy tech").Value = Val(strParams(iIndex))
				Case "level_age_rate"
					rsVersion.Fields("Tech Decay Rate").Value = 1
					rsVersion.Fields("Tech Decay Time").Value = Val(strParams(iIndex))
				Case "fire_range_factor"
					rsVersion.Fields("Fire range").Value = strParams(iIndex)
				Case "sect_mob_max"
					rsVersion.Fields("Max mob - Sector").Value = strParams(iIndex)
				Case "ship_mob_max"
					rsVersion.Fields("Max mob - Ship").Value = strParams(iIndex)
				Case "plane_mob_max"
					rsVersion.Fields("Max mob - Plane").Value = strParams(iIndex)
				Case "land_mob_max"
					rsVersion.Fields("Max mob - Land").Value = strParams(iIndex)
				Case "sect_mob_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						rsVersion.Fields("Mob gain - Sector").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
					End If
				Case "ship_mob_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						rsVersion.Fields("Mob gain - Ship").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
					End If
				Case "plane_mob_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						rsVersion.Fields("Mob gain - Plane").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
					End If
				Case "land_mob_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						rsVersion.Fields("Mob gain - Land").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
					End If
				Case "ship_grow_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						If Val(strParams(iIndex)) * Val(strParams(iETUupdate)) > 100 Then
							rsVersion.Fields("Eff gain - Ship").Value = 100#
						Else
							rsVersion.Fields("Eff gain - Ship").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
						End If
					End If
				Case "plane_grow_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						If Val(strParams(iIndex)) * Val(strParams(iETUupdate)) > 100 Then
							rsVersion.Fields("Eff gain - Plane").Value = 100#
						Else
							rsVersion.Fields("Eff gain - Plane").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
						End If
					End If
				Case "land_grow_scale"
					'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex(iETUupdate, ParseXDVersion, ETU per update). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
						If Val(strParams(iIndex)) * Val(strParams(iETUupdate)) > 100 Then
							rsVersion.Fields("Eff gain - Land").Value = 100#
						Else
							rsVersion.Fields("Eff gain - Land").Value = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
						End If
					End If
				Case "buil_bh", "buil_tower_bh"
				Case "buil_bt"
					'pr("Bridges require %g tech,", buil_bt);
					'pr(" %d hcm,", buil_bh);
					'pr(" %d workers,\n", 0);
					ParseShowText("Bridges require " & CStr(Int(Val(strParams(iIndex)))) & " tech, " & CStr(strParams(iBuildBridgeHcm)) & " hcm, 0 workers", "b", "b", 2)
				Case "buil_bc"
					'pr("%d available workforce, and cost $%g\n",
					'  (SCT_BLD_WORK(0, buil_bh) * SCT_MINEFF + 99) / 100,
					'  buil_bc);
					ParseShowText(CStr(Int((SectorBuildWork(0, CShort(strParams(iBuildBridgeHcm))) * 20 + 99) / 100)) & " available workforce, and cost $" & CStr(Int(Val(strParams(iIndex)))), "b", "b", 3)
				Case "buil_tower_bt"
					'pr("Bridge Towers require %g tech,", buil_tower_bt);
					'pr(" %d hcm,", buil_tower_bh);
					'pr(" %d workers,\n", 0);
					ParseShowText("Bridge Towers require " & CStr(Int(Val(strParams(iIndex)))) & " tech, " & CStr(strParams(iBuildTowerHcm)) & " hcm, 0 workers", "t", "b", 2)
				Case "buil_tower_bc"
					'pr("%d available workforce, and cost $%g\n",
					'   (SCT_BLD_WORK(0, buil_tower_bh) * SCT_MINEFF + 99) / 100,
					'   buil_tower_bc);
					ParseShowText(CStr(Int((SectorBuildWork(0, CShort(strParams(iBuildTowerHcm))) * 20 + 99) / 100)) & " available workforce, and cost $" & CStr(Int(Val(strParams(iIndex)))), "t", "b", 3)
				Case "NOFOOD"
					If CDbl(strParams(iIndex)) = 1 Then
						rsVersion.Fields("Food Needed").Value = "N"
					Else
						rsVersion.Fields("Food Needed").Value = "Y"
					End If
				Case "BIG_CITY", "GO_RENEW", "RAILWAYS"
					If CDbl(strParams(iIndex)) = 1 Then
						rsVersion.Fields(etTable.Column(iHeaderIndex).Name).Value = True
					Else
						rsVersion.Fields(etTable.Column(iHeaderIndex).Name).Value = False
					End If
				Case "drnuke_const"
					rsVersion.Fields("DRNUKE Const").Value = Val(strParams(iIndex))
				Case "LIMBER"
					If frmOptions.bHeavyMetal Then
						'Assume the bHeavyMetal has set both of this options for the time being
					Else
						EmpireError("ParseXDVersion", "Field not understood:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDVersion", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		rsVersion.Update()
		If Not bXDumpServer And VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
			RequestMetaTables()
		End If
		Exit Sub
		
		'Error handling routine
ParseXDVersion_Exit: 
		EmpireError("ParseXDVersion", strErrorField, strLine)
	End Sub
	
	Public Function GetIndexStringArray(ByRef strArray() As String, ByRef strMatch As String) As Short
		Dim iIndex As Short
		Dim iAdjustedIndex As Short
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(strArray) Then
			GetIndexStringArray = -1
			Exit Function
		End If
		
		iAdjustedIndex = 0
		For iIndex = 0 To UBound(strArray)
			If iIndex < UBound(strArray) Then
				If strArray(iIndex) = "item" And strArray(iIndex + 1) = "15" Then
					iAdjustedIndex = iAdjustedIndex + 14
					iIndex = iIndex + 1
				ElseIf strArray(iIndex) = "pkg" And strArray(iIndex + 1) = "5" Then 
					iAdjustedIndex = iAdjustedIndex + 4
					iIndex = iIndex + 1
				Else
					If strArray(iIndex) = strMatch Then
						GetIndexStringArray = iAdjustedIndex
						Exit Function
					End If
				End If
			Else
				If strArray(iIndex) = strMatch Then
					GetIndexStringArray = iAdjustedIndex
					Exit Function
				End If
			End If
			iAdjustedIndex = iAdjustedIndex + 1
		Next iIndex
		GetIndexStringArray = -1
	End Function
	
	Private Function GetUnitType(ByRef eiChrIndex As Short, ByRef strUnit As String) As String
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("chr_i").Value = eiChrIndex And rsBuild.Fields("type").Value = strUnit Then
					GetUnitType = rsBuild.Fields("id").Value
					Exit Function
				End If
				rsBuild.MoveNext()
			End While
		End If
		GetUnitType = "????"
	End Function
	
	Public Function RemoveEscapeSequencesAndQuotes(ByRef strParam As String) As String
		Dim iStartPos As Short
		Dim iFoundPos As Short
		Dim strOutput As String
		
		iFoundPos = InStr(strParam, "\")
		
		If iFoundPos = 0 Then
			If InStr(strParam, Chr(34)) = 1 Then
				strParam = Mid(strParam, 2, Len(strParam) - 2)
			End If
			RemoveEscapeSequencesAndQuotes = strParam
			Exit Function
		End If
		
		iStartPos = 1
		While iFoundPos <> 0
			strOutput = strOutput & Mid(strParam, iStartPos, iFoundPos - iStartPos)
			strOutput = strOutput & Chr(CInt("&" & Mid(strParam, iFoundPos + 1, 3)))
			iStartPos = iFoundPos + 4
			iFoundPos = InStr(iStartPos, strParam, "\")
		End While
		strParam = strOutput & Mid(strParam, iStartPos)
		If InStr(strParam, Chr(34)) = 1 Then
			strParam = Mid(strParam, 2, Len(strParam) - 2)
		End If
		RemoveEscapeSequencesAndQuotes = strParam
	End Function
	
	Private Function CheckIndex(ByRef iIndex As Short, ByRef strFunction As String, ByRef strField As String) As Object
		If iIndex < 0 Then
			EmpireError(strFunction, "The " & strField & " field not present", CStr(iIndex))
			'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CheckIndex = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object CheckIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CheckIndex = True
		End If
	End Function
	
	Private Sub ParseXDMetaHeader(ByRef strType As String)
		Dim xdType As enumXdumpType
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xdType = ParseXDTableName(strType)
		etsTable.AddUpdate(xdType, strType)
		xdMetaType = xdType
	End Sub
	
	Private Sub ParseXDSymbolTableHeader(ByRef xdSymbolTable As enumXdumpType, ByRef strType As String)
		estsTable.AddUpdate(xdSymbolTable, strType)
	End Sub
	
	Private Sub ParseXDNation(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iCnumIndex As Short
		Dim iXcapIndex As Short
		Dim strErrorField As String
		Dim etTable As EmpTable
		
		On Error GoTo ParseXDNation_Exit
		
		etTable = etsTable(enumXdumpType.XD_NATION)
		If etTable Is Nothing Then
			EmpireError("ParseXDNation", "The meta data was not found for nat", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDNation", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		iCnumIndex = etTable.GetZeroIndex("cnum")
		iXcapIndex = etTable.GetZeroIndex("xcap")
		
		If iCnumIndex = -1 Or iXcapIndex = -1 Then
			EmpireError("ParseXDNation", "One of the following is missing: country number or X location for the capital", "")
			Exit Sub
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "cnum"
					CountryNumber = Val(strParams(iIndex))
				Case "cname"
					If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
						frmEmpCmd.CountryName = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
						Nations.AddNation(Val(strParams(iCnumIndex)), (frmEmpCmd.CountryName))
						If rsNations.BOF And rsNations.EOF Then
							rsNations.AddNew()
						Else
							rsNations.Seek("=", Val(strParams(iCnumIndex)))
							If rsNations.NoMatch Then
								rsNations.AddNew()
							Else
								rsNations.Edit()
							End If
						End If
						rsNations.Fields("Name").Value = frmEmpCmd.CountryName
						rsNations.Fields("Number").Value = Val(strParams(iCnumIndex))
						rsNations.Update()
					Else
						EmpireError("ParseXDNation", "Field not understood:" & strErrorField, strLine)
					End If
				Case "xcap"
				Case "ycap"
					CapSect = SectorString(Val(strParams(iXcapIndex)), Val(strParams(iIndex)))
				Case "access"
					If VersionCheck(4, 3, 10) = WinAceRoutines.enumVersion.VER_LESS Then
						EmpireError("ParseXDCountries", "Access field invalid:" & strErrorField, strLine)
					End If
				Case "milreserve"
					MilReserves = Val(strParams(iIndex))
				Case "education"
					Education = Val(strParams(iIndex))
				Case "tech"
					TechLevel = Val(strParams(iIndex))
				Case "research"
					Research = Val(strParams(iIndex))
					'safepop = (int)((double)poplimit / (1.0 + obrate * (double)etu_per_update));
					'uwpop = (int)((double)poplimit / (1.0 + uwbrate * (double)etu_per_update));
					'safepop++;
					'if (((double)safepop * (1.0 + obrate * (double)etu_per_update)) >
					'    ((double)poplimit + 0.0000001))
					'safepop--;
					'uwpop++;
					'if (((double)uwpop * (1.0 + uwbrate * (double)etu_per_update)) >
					'   ((double)poplimit + 0.0000001))
					'uwpop--;
					'
					Maxpop = MaxPopulation()
					MaxSafeCivs = Int(CSng(Maxpop) / (1# + (Val(rsVersion.Fields("Civ Birthrate").Value) / 1000#) * Val(rsVersion.Fields("ETU per update").Value)))
					If MaxSafeCivs > Maxpop Then
						MaxSafeCivs = Maxpop
					End If
					MaxSafeUws = Int(CSng(Maxpop) / (1# + (Val(rsVersion.Fields("Uw Birthrate").Value) / 1000#) * Val(rsVersion.Fields("ETU per update").Value)))
					If MaxSafeUws > Maxpop Then
						MaxSafeUws = Maxpop
					End If
				Case "happiness"
					Happiness = Val(strParams(iIndex))
				Case "max"
				Case "stat", "passwd", "ip", "hostname", "userid", "xorg", "yorg", "dayno", "update", "missed", "tgms", "ann", "minused", "timeused", "btu", "money", "login", "logout", "newstim", "annotim", "flags"
				Case "relations", "contacts", "rejects"
					If VersionCheck(4, 3, 4) <> WinAceRoutines.enumVersion.VER_LESS Or frmOptions.bSPAtlantis Then
						EmpireError("ParseXDNation", "Field not understood:" & strErrorField, strLine)
					End If
					iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
				Case Else
					EmpireError("ParseXDNation", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		SaveNationVariables()
		Exit Sub
		
		'Error handling routine
ParseXDNation_Exit: 
		EmpireError("ParseXDNation", strErrorField, strLine)
	End Sub
	
	Private Sub ParseXDCountries(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iHeaderIndex As Short
		Dim iCnumIndex As Short
		Dim iStatusIndex As Short
		Dim iRelIndex As Short
		Dim iStatus As Short
		Dim strErrorField As String
		Dim etTable As EmpTable
		
		On Error GoTo ParseXDCountries_Exit
		
		etTable = etsTable(enumXdumpType.XD_COUNTRIES)
		If etTable Is Nothing Then
			EmpireError("ParseXDCountries", "The meta data was not found for cou", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDCountries", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		iCnumIndex = etTable.GetZeroIndex("cnum")
		iStatusIndex = etTable.GetZeroIndex("stat")
		
		If iCnumIndex = -1 And iStatusIndex = -1 Then
			EmpireError("ParseXDCountries", "One of the following is missing: country number or status", "")
			Exit Sub
		End If
		
		iHeaderIndex = 1
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iHeaderIndex).Name
			Select Case etTable.Column(iHeaderIndex).Name
				Case "cname"
					iStatus = Val(strParams(iStatusIndex))
					Select Case iStatus
						Case enumNationStatus.STAT_SANCT, enumNationStatus.STAT_ACTIVE, enumNationStatus.STAT_GOD
							strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
							Nations.AddNation(Val(strParams(iCnumIndex)), strParams(iIndex))
							If rsNations.BOF And rsNations.EOF Then
								rsNations.AddNew()
							Else
								rsNations.Seek("=", Val(strParams(iCnumIndex)))
								If rsNations.NoMatch Then
									rsNations.AddNew()
								Else
									rsNations.Edit()
								End If
							End If
							rsNations.Fields("Name").Value = strParams(iIndex)
							rsNations.Fields("Number").Value = Val(strParams(iCnumIndex))
							rsNations.Update()
					End Select
				Case "cnum", "stat", "cname", "passwd", "ip", "hostname", "userid", "xcap", "ycap", "xorg", "yorg", "dayno", "update", "missed", "tgms", "ann", "minused", "timeused", "btu", "milreserve", "money", "login", "logout", "newstim", "annotim", "tech", "research", "education", "happiness", "flags"
				Case "relations" ' 7 0 99 38
					If VersionCheck(4, 3, 4) <> WinAceRoutines.enumVersion.VER_LESS Then
						For iRelIndex = 0 To etTable.Column(iHeaderIndex).Length - 1
							If Nations.NationName(iRelIndex) <> vbNullString Then
								Nations.Relation(Val(strParams(iCnumIndex)), iRelIndex) = CShort(strParams(iIndex + iRelIndex))
							End If
						Next 
					End If
					iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
				Case "contacts" ' 6 1 99 -1
					iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
				Case "rejects" ' 6 1 99 -1
					iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
				Case "access" ' 6 0 0 -1
					If VersionCheck(4, 3, 10) = WinAceRoutines.enumVersion.VER_LESS Or Not bDeity Then
						EmpireError("ParseXDCountries", "Access field invalid:" & strErrorField, strLine)
					End If
				Case Else
					EmpireError("ParseXDCountries", "Field not understood:" & strErrorField, strLine)
			End Select
			iHeaderIndex = iHeaderIndex + 1
		Next iIndex
		Exit Sub
		
		'Error handling routine
ParseXDCountries_Exit: 
		EmpireError("ParseXDCountries", strErrorField, strLine)
	End Sub
	
	Private Sub ParseXDNationRelationships(ByRef strLine As String)
		Dim strParams() As String
		Dim etTable As EmpTable
		Dim strErrorField As String
		Dim iIndex As Short
		Dim iValue As Short
		
		On Error GoTo ParseXDNationRelationships_Exit
		
		etTable = etsTable(enumXdumpType.XD_NATION_RELATIONSHIPS)
		If etTable Is Nothing Then
			EmpireError("ParseXDNationRelationships", "The meta data was not found for nation relationships", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDNationRelationships", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iIndex + 1).Name
			Select Case etTable.Column(iIndex + 1).Name
				Case "value"
					iValue = CShort(strParams(iIndex))
				Case "name"
					Select Case RemoveEscapeSequencesAndQuotes(strParams(iIndex))
						Case "unknown"
							iREL_AT_WAR = -1
							iREL_SITZKRIEG = -1 'SLOW_WAR only
							iREL_MOBILIZING = -1 'SLOW_WAR only
							iREL_HOSTILE = -1
							iREL_NEUTRAL = -1
							iREL_FRIENDLY = -1
							iREL_ALLIED = -1
						Case "at-war"
							iREL_AT_WAR = iValue
						Case "sitzkrieg"
							iREL_SITZKRIEG = iValue
						Case "mobilizing"
							iREL_MOBILIZING = iValue
						Case "hostile"
							iREL_HOSTILE = iValue
						Case "neutral"
							iREL_NEUTRAL = iValue
						Case "friendly"
							iREL_FRIENDLY = iValue
						Case "allied"
							iREL_ALLIED = iValue
						Case Else
							EmpireError("ParseXDNationRelationships", "Expected Relationship: " & strErrorField, strLine)
					End Select
				Case Else
					EmpireError("ParseXDNationRelationships", "Field not understood:" & strErrorField, strLine)
			End Select
		Next 
		Exit Sub
		
ParseXDNationRelationships_Exit: 
		EmpireError("ParseXDNationRelationships", strErrorField, strLine)
	End Sub
	
	Private Sub ParseXDLost(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iTypeIndex As Short
		Dim iIDIndex As Short
		Dim iXlocIndex As Short
		Dim iYlocIndex As Short
		Dim iTimestamp As Integer
		Dim strErrorField As String
		Dim hldIndex As String
		Dim etTable As EmpTable
		'UPGRADE_WARNING: Arrays in structure rsLost may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rsLost As DAO.Recordset
		
		hldIndex = ""
		
		On Error GoTo ParseXDLost_Exit
		
		etTable = etsTable(enumXdumpType.XD_LOST)
		If etTable Is Nothing Then
			EmpireError("ParseXDLost", "The meta data was not found for lost", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDLost", "The number of fields does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		iTypeIndex = etTable.GetZeroIndex("type")
		iIDIndex = etTable.GetZeroIndex("id")
		iXlocIndex = etTable.GetZeroIndex("x")
		iYlocIndex = etTable.GetZeroIndex("y")
		iTimestamp = etTable.GetZeroIndex("timestamp")
		
		If iTypeIndex = -1 Or iIDIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or iTimestamp = -1 Then
			EmpireError("ParseXDLost", "One of the following: type, id, x, y or timestamp is not present", "")
			Exit Sub
		End If
		
		Select Case strParams(iTypeIndex)
			Case CStr(LT_SECTOR)
				rsLost = rsSectors
				strErrorField = "Invalid Sector"
				hldIndex = rsLost.Index
				rsLost.Index = "PrimaryKey"
				rsLost.Seek("=", CShort(strParams(iXlocIndex)), CShort(strParams(iYlocIndex)))
				If Not rsLost.NoMatch Then
					If CInt(strParams(iTimestamp)) > rsLost.Fields("timestamp").Value Then
						rsLost.Delete()
					End If
				End If
			Case CStr(LT_SHIP)
				rsLost = rsShip
				strErrorField = "Invalid Ship"
				hldIndex = rsLost.Index
				rsLost.Index = "PrimaryKey"
				rsLost.Seek("=", CShort(strParams(iIDIndex)))
				If Not rsLost.NoMatch Then
					If CInt(strParams(iTimestamp)) > rsLost.Fields("timestamp").Value Then
						rsLost.Delete()
					End If
				End If
			Case CStr(LT_PLANE)
				rsLost = rsPlane
				strErrorField = "Invalid Plane"
				hldIndex = rsLost.Index
				rsLost.Index = "PrimaryKey"
				rsLost.Seek("=", CShort(strParams(iIDIndex)))
				If Not rsLost.NoMatch Then
					If CInt(strParams(iTimestamp)) > rsLost.Fields("timestamp").Value Then
						rsLost.Delete()
					End If
				End If
			Case CStr(LT_LAND)
				rsLost = rsLand
				strErrorField = "Invalid Land"
				hldIndex = rsLost.Index
				rsLost.Index = "PrimaryKey"
				rsLost.Seek("=", CShort(strParams(iIDIndex)))
				If Not rsLost.NoMatch Then
					If CInt(strParams(iTimestamp)) > rsLost.Fields("timestamp").Value Then
						rsLost.Delete()
					End If
				End If
			Case CStr(LT_NUKE)
				rsLost = rsNuke
				strErrorField = "Invalid Nuke"
				hldIndex = rsLost.Index
				rsLost.Index = "PrimaryKey"
				rsLost.Seek("=", CShort(strParams(iIDIndex)))
				While Not rsLost.NoMatch
					rsLost.Delete()
					rsLost.Seek("=", CShort(strParams(iIDIndex)))
				End While
		End Select
		
		rsLost.Index = hldIndex
		
		If Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY Then
			frmDrawMap.FillSectorBox(Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex)))
			If Not frmDrawMap.IsAnyTabActive Then
				frmDrawMap.DisplayFirstSectorPanel()
			End If
		End If
		
		Exit Sub
		
		'Error handling routine
ParseXDLost_Exit: 
		EmpireError("ParseXDLost", strErrorField, strLine)
		If Len(hldIndex) > 0 Then
			rsLost.Index = hldIndex
		End If
	End Sub
	
	Private Sub ParseXDMeta(ByRef xdTable As enumXdumpType, ByRef strLine As String)
		Dim etTable As EmpTable
		
		etTable = etsTable.Table(xdTable)
		If Not etTable Is Nothing Then
			etTable.AddUpdateColumn(strLine)
		End If
	End Sub
	
	Private Sub ParseXDSymbolTable(ByRef strLine As String)
		Dim strParams() As String
		Dim iIndex As Short
		Dim iValueIndex As Short
		Dim strErrorField As String
		Dim etTable As EmpTable
		
		On Error GoTo ParseXDSymbolTable_Exit
		
		etTable = etsTable(xdType)
		If etTable Is Nothing Then
			EmpireError("ParseXDSymbolTable", "The meta data was not found for table (" & CStr(xdType) & ")", strLine)
			Exit Sub
		End If
		
		strParams = Split(strLine, " ")
		
		If UBound(strParams) + 1 <> etTable.ParameterCount Then
			EmpireError("ParseXDSymbolTable", "The number of fields for table " & etTable.Name & " does not match the header", CStr(etTable.ParameterCount) & ":" & strLine)
			Exit Sub
		End If
		
		iValueIndex = etTable.GetZeroIndex("value")
		
		If iValueIndex = -1 Then
			EmpireError("ParseXDSymbolTable", "The record identifier field not present", CStr(iValueIndex))
			Exit Sub
		End If
		
		For iIndex = 0 To UBound(strParams)
			strErrorField = etTable.Column(iIndex + 1).Name
			Select Case etTable.Column(iIndex + 1).Name
				Case "name"
					estsTable(xdType).AddUpdateSymbolEntry(CInt(strParams(iValueIndex)), RemoveEscapeSequencesAndQuotes(strParams(iIndex)))
				Case "value"
			End Select
		Next 
		Exit Sub
		
		'Error handling routine
ParseXDSymbolTable_Exit: 
		EmpireError("ParseXDSymbolTable", strErrorField, strLine)
	End Sub
	
	Private Function ParseXDTableName(ByRef strName As String) As Object
		Select Case strName
			Case "country"
				If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ParseXDTableName = enumXdumpType.XD_COUNTRIES
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ParseXDTableName = enumXdumpType.XD_NATION
				End If
			Case "game"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_GAME
			Case "item"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_ITEM_CHR
			Case "infrastructure"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_INFRA_CHR
			Case "land"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_LAND
			Case "land-chr"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_LAND_CHR
			Case "level"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_LEVEL
			Case "lost"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_LOST
			Case "meta"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_META
			Case "missions"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_MISSIONS
			Case "nat"
				If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ParseXDTableName = enumXdumpType.XD_NATION
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ParseXDTableName = enumXdumpType.XD_COUNTRIES
				End If
			Case "nation-relationships"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_NATION_RELATIONSHIPS
			Case "nuke"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_NUKE
			Case "nuke-chr"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_NUKE_CHR
			Case "sect"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_SECTOR
			Case "sect-chr"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_SECTOR_CHR
			Case "sector-navigation"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_SECTOR_NAVIGATION
			Case "ship"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_SHIP
			Case "ship-chr"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_SHIP_CHR
			Case "ship-chr-flags"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_SHIP_CHR_FLAGS
			Case "plane"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_PLANE
			Case "plane-chr"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_PLANE_CHR
			Case "product"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_PRODUCT_CHR
			Case "updates"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_UPDATES
			Case "version"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_VERSION
			Case "meta-type"
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_META_TYPE
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseXDTableName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseXDTableName = enumXdumpType.XD_UNKNOWN
		End Select
	End Function
	
	Public Sub RequestMetaTables()
		If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
			Exit Sub
		End If
		frmEmpCmd.SubmitEmpireCommand("xdump meta meta", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta ver", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta item", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta infra", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta product", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta sect-chr", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta land-chr", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta ship-chr", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta nuke-chr", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta plane-chr", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta sect", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta land", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta ship", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta nuke", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta plane", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta nat", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta lost", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta cou", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta meta-type", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta missions", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta level", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta sector-navigation", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta ship-chr-flags", False)
		If VersionCheck(4, 3, 10) <> WinAceRoutines.enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("xdump meta updates", False)
			frmEmpCmd.SubmitEmpireCommand("xdump meta game", False)
		End If
		frmEmpCmd.SubmitEmpireCommand("xdump meta nation-relationships", False)
		frmEmpCmd.SubmitEmpireCommand("xdump meta-type *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump missions *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump level *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump sector-navigation *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump ship-chr-flags *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump nation-relationships *", False)
	End Sub
	
	Public Sub RequestConfigurationTables()
		frmEmpCmd.SubmitEmpireCommand("xdump ver *", False)
		GetNation()
		frmEmpCmd.SubmitEmpireCommand("xdump item *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump sect-chr *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump product *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump infra *", False)
		RequestUnitConfigurationTables()
	End Sub
	
	Public Sub RequestUnitConfigurationTables()
		frmEmpCmd.SubmitEmpireCommand("xdump land-chr *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump ship-chr *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump nuke-chr *", False)
		frmEmpCmd.SubmitEmpireCommand("xdump plane-chr *", False)
	End Sub
	
	Private Function GetTechLevelForShow() As Short
		If Len(frmToolShow.txtTech.Text) > 0 Then
			GetTechLevelForShow = CShort(frmToolShow.txtTech.Text)
		ElseIf bDeity Then 
			GetTechLevelForShow = 1000
		Else
			If rsNation.BOF And rsNation.EOF Then
				GetTechLevelForShow = 0
			Else
				rsNation.MoveFirst()
				GetTechLevelForShow = CShort(rsNation.Fields("TechLevel").Value)
			End If
		End If
	End Function
	
	Private Sub UpdateTechLevelForShow(ByRef strUnit As String, ByRef strReport As String)
		Dim strText As String
		
		strText = "Printing for tech level '" & CStr(GetTechLevelForShow) & "'"
		ParseShowText(strText, strUnit, strReport, 1)
	End Sub
	
	'UPGRADE_NOTE: GetType was upgraded to GetType_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function GetType_Renamed(ByRef iIndex As Short) As String
		Dim eiItem As EmpItem
		
		If iIndex = -1 Then
			GetType_Renamed = ""
			Exit Function
		End If
		
		eiItem = Items.FindByChrIndex(iIndex)
		
		If eiItem Is Nothing Then
			GetType_Renamed = ""
		Else
			GetType_Renamed = eiItem.Letter
		End If
	End Function
	
	Private Function GetNationType(ByRef lIndex As Integer) As String
		Select Case Left(estsTable(enumXdumpType.XD_LEVEL)(lIndex).Name, 3)
			Case "non"
				GetNationType = ""
			Case "tec"
				GetNationType = "tech"
			Case Else
				GetNationType = Left(estsTable(enumXdumpType.XD_LEVEL)(lIndex).Name, 3)
		End Select
	End Function
	
	Public Function MaxPopulation(Optional ByRef strDes As String = "") As Short
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not IsNothing(strDes) Then
			rsSectorType.Seek("=", strDes)
		ElseIf frmOptions.bSPAtlantis Then 
			rsSectorType.Seek("=", "i")
		Else
			rsSectorType.Seek("=", "m")
		End If
		If Not rsSectorType.NoMatch Then
			MaxPopulation = rsSectorType.Fields("maxpop").Value
		Else
			MaxPopulation = 999
		End If
	End Function
	
	Private Function DirectionString(ByRef iValue As Short) As String
		Select Case iValue
			Case 0
				DirectionString = "."
			Case 1
				DirectionString = "u"
			Case 2
				DirectionString = "j"
			Case 3
				DirectionString = "n"
			Case 4
				DirectionString = "b"
			Case 5
				DirectionString = "g"
			Case 6
				DirectionString = "y"
			Case Else
				DirectionString = "?"
		End Select
	End Function
	
	Private Sub AddMission(ByRef lMission As Integer, ByRef iID As Short, ByRef iOwner As Short, ByRef strType As String, ByRef strSector As String, ByRef iRadius As Short)
		If lMission > 0 Then
			If rsMissions.BOF And rsMissions.EOF Then
				rsMissions.AddNew()
			Else
				rsMissions.Seek("=", strType, iID)
				If Not rsMissions.NoMatch Then
					rsMissions.Edit()
				Else
					rsMissions.AddNew()
				End If
			End If
			rsMissions.Fields("id").Value = iID
			rsMissions.Fields("owner").Value = iOwner
			rsMissions.Fields("type").Value = strType
			rsMissions.Fields("op sector").Value = strSector
			rsMissions.Fields("mission").Value = estsTable(enumXdumpType.XD_MISSIONS)(lMission).Name
			rsMissions.Fields("radius").Value = iRadius
			rsMissions.Update()
		ElseIf Not (rsMissions.BOF And rsMissions.EOF) Then 
			rsMissions.Seek("=", strType, iID)
			If Not rsMissions.NoMatch Then
				rsMissions.Delete()
			End If
		End If
	End Sub
	
	Private Sub AddOrders(ByRef bAutoNav As Boolean, ByRef iID As Short, ByRef iOwner As Short, ByRef iLength As Short, ByRef iETA As Short, ByRef strCurSector As String, ByRef strStartSector As String, ByRef strEndSector As String, ByRef strStartCargo() As String, ByRef strEndCargo() As String, ByRef iStart() As Short, ByRef iEnd() As Short)
		Dim iIndex As Short
		
		If bAutoNav Then
			If rsShipOrders.BOF And rsShipOrders.EOF Then
				rsShipOrders.AddNew()
			Else
				rsShipOrders.Seek("=", iID)
				If Not rsShipOrders.NoMatch Then
					rsShipOrders.Edit()
				Else
					rsShipOrders.AddNew()
				End If
			End If
			rsShipOrders.Fields("id").Value = iID
			rsShipOrders.Fields("curr sector").Value = strCurSector
			rsShipOrders.Fields("start sector").Value = strEndSector
			rsShipOrders.Fields("end sector").Value = strStartSector
			rsShipOrders.Fields("length").Value = iLength
			rsShipOrders.Fields("eta").Value = iETA
			rsShipOrders.Fields("owner").Value = iOwner
			For iIndex = 1 To 6
				rsShipOrders.Fields("start cargo " & CStr(iIndex)).Value = strStartCargo(iIndex)
				rsShipOrders.Fields("end cargo " & CStr(iIndex)).Value = strEndCargo(iIndex)
				rsShipOrders.Fields("start level " & CStr(iIndex)).Value = iStart(iIndex)
				rsShipOrders.Fields("end level " & CStr(iIndex)).Value = iEnd(iIndex)
			Next 
			rsShipOrders.Update()
		ElseIf Not (rsShipOrders.BOF And rsShipOrders.EOF) Then 
			rsShipOrders.Seek("=", iID)
			If Not rsShipOrders.NoMatch Then
				rsShipOrders.Delete()
			End If
		End If
		If frmDrawMap.PromptUp And frmDrawMap.PromptForm Is frmPromptOrder Then
			frmPromptOrder.OrderLoadData()
		End If
	End Sub
	
	Private Function CalculateAvail(ByRef iLCM As Short, ByRef iHCM As Short) As Short
		CalculateAvail = 20 + iLCM + (2 * iHCM)
	End Function
	
	'HeavyMetal modification
	Private Sub UpdateRailTrack()
		Dim xas, xa, ya, yas As Object
		Dim iIndex As Short
		Dim strSect() As String
		Dim iCount As Short
		Dim vSect As Object
		Dim iSecX As Short
		Dim iSecY As Short
		
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xa = New Object(){-2, -1, 1, 2, 1, -1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ya. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ya = New Object(){0, -1, -1, 0, 1, 1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xas. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xas = New Object(){1, 2, 1, -1, -2, -1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object yas. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		yas = New Object(){-1, 0, 1, 1, 0, -1}
		
		'#define SCT_HAS_MAJOR_RAILWAY(sp) \
		'    (dchr[(sp)->sct_type].d_mob1 == 0 && (sp)->sct_effic >= SCT_RAIL_EFF)
		'void
		'set_railway(struct sctstr *sp)
		'{
		'    int i;
		'    struct sctstr *nsp;
		
		'    sp->sct_track = !!SCT_HAS_MAJOR_RAILWAY(sp);
		'    for (i = DIR_FIRST; i <= DIR_LAST; i++) {
		'    nsp = getsectp(sp->sct_x + diroff[i][0],
		'               sp->sct_y + diroff[i][1]);
		'    sp->sct_track +=
		'        nsp->sct_own == sp->sct_own && SCT_HAS_MAJOR_RAILWAY(nsp);
		'    }
		'}
		
		iCount = 0
		
		If rsSectors.BOF And rsSectors.EOF Then
			Exit Sub
		End If
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			rsSectorType.Seek("=", rsSectors.Fields("des"))
			If rsSectorType.NoMatch Then
				Exit Sub
			End If
			rsSectors.Edit()
			If rsSectorType.Fields("mcost100").Value = 0 And rsSectors.Fields("eff").Value >= 5 Then
				rsSectors.Fields("rail").Value = 1
				iCount = iCount + 1
				'UPGRADE_WARNING: Lower bound of array strSect was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve strSect(iCount)
				strSect(iCount) = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
			Else
				rsSectors.Fields("rail").Value = 0
			End If
			rsSectors.Update()
			rsSectors.MoveNext()
		End While
		If iCount = 0 Then
			Exit Sub
		End If
		
		For	Each vSect In strSect
			For iIndex = 0 To 5
				'UPGRADE_WARNING: Couldn't resolve default property of object vSect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ParseSectors(iSecX, iSecY, CStr(vSect)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ya(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object xa(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsSectors.Seek("=", iSecX + CShort(xa(iIndex)), iSecY + CShort(ya(iIndex)))
					If Not rsSectors.NoMatch Then
						rsSectorType.Seek("=", rsSectors.Fields("des"))
						If rsSectorType.NoMatch Then
							Exit Sub
						End If
						If rsSectors.Fields("eff").Value >= 60 Then
							rsSectors.Edit()
							rsSectors.Fields("rail").Value = rsSectors.Fields("rail").Value + 1
							rsSectors.Update()
						End If
					End If
				End If
			Next iIndex
		Next vSect
		
	End Sub
End Module