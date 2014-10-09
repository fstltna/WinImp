Option Strict Off
Option Explicit On
Module EmpDataBase
	
	'Change Log:
	'090703 drk: removed the ship stats table, as it was static (never refreshed from server)
	'            and redundant anyway.
	'101703 rjk: Added Function to update db structure.
	'101703 rjk: Added fields for strength in the rsSectors table
	'101703 rjk: Added country field in nuke table
	'102503 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
	'120303 rjk: Switched options to frmOptions and boolean options
	'120703 rjk: Added rsItem support different commodities (theme games)
	'121303 rjk: Removed GetCommmodityValuefromDB, part of Item changes
	'010204 rjk: Add PrimaryKey (id) for nuke Table if it is missing
	'            Remove ShipMerchant.
	'050204 rjk: Currently CheckDBFields does not need to do anything as QueryDef changes for land unit required to new database.
	'070304 rjk: Switched Enemy timestamp to UTC
	'270304 rjk: Added 'Rollover Avail' to Version table
	'270304 rjk: Added DeleteAllRecords function and switched code over to use for clearing tables
	'210404 rjk: Remove the local copy of tz to use the global tz in UpdateEnemyTimeStamp.
	'300704 rjk: DeleteAllRecordsOlderThen to support clear function to Days range.
	'290804 rjk: Added index to Telegram header to look for duplicates.
	'220305 rjk: Added 'Version' to Version table
	'281205 rjk: Added max. eff. gain limits for sect to the version table.
	'110206 rjk: Added 'drnuke_const' to Version table
	'170206 rjk: Added Index for Country Number in Nations table
	'120306 rjk: Added Base Speed and Base Firing Range to Build Info table
	'200306 rjk: New function DeleteBuildRecords.
	'230306 rjk: Added Base Attack and Base Defense to Build Info table
	'210406 rjk: Added Base Acc, Base Load, Base Vul, Base Dam and Base AAF
	'            to Build Info table
	'230406 rjk: Added DeleteAllAirCombatRecordsOlderThen Function.
	'280506 rjk: Add timestamps to sector, ship, land, plane and nuke tables
	'120606 rjk: Added support for symbol tables.
	'180606 rjk: Added mcost at 100% to the Sector Type table for 4.3.6 servers
	'190606 rjk: Added nav field to the sector type.
	'200606 rjk: Added nuke id Index, in 4.3.6 servers, there are not multiple
	'            types per nuke id.
	'250606 rjk: Added off to ship, land, plane and nuke tables.
	'230806 rjk: Modified rsSea index to be PrimaryKey from Primary Key
	'230806 rjk: Added timestamps for oil and fert for Sea sectors table.
	'020907 rjk: Replace TimeZoneAdj with a database field in version table
	'131207 rjk: Added end cargo fields to Ship orders, rename cargo to start cargo
	'030108 rjk: Add elevation to the sector database.
	'120109 rjk: Removed unused local variable n. Open_error is removed, not used in OpenEmpireDB.
	'140608 rjk: Added packing_type to sector chr tofix the packing factor for mobility calculations.
	'160209 rjk: Added RAILWAYS option, originally part of Heavy Metal/Plastic games.
	'            Added Terrain to the sector type table.
	
	'database
	Public dbsName As String 'Hold empire database name
	Public GameName As String
	Public dbEmpire As DAO.Database
	'UPGRADE_WARNING: Arrays in structure rsSectors may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsSectors As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsBmap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsBmap As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsShip may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsShip As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsShipList may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsShipList As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsLandList may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsLandList As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsPlaneList may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsPlaneList As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsLand may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsLand As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsPlane may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsPlane As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsNuke may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsNuke As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsTeleHead may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsTeleHead As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsTeleBody may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsTeleBody As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsEnemySect may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsEnemySect As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsEnemyUnit may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsEnemyUnit As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsVersion may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsVersion As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsNation may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsNation As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsNations may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsNations As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsBuild may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsBuild As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsSea may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsSea As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsAirCombat may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsAirCombat As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsShipOrders may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsShipOrders As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsMissions may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsMissions As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsSectorType may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsSectorType As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsShowText may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsShowText As DAO.Recordset
	'UPGRADE_WARNING: Arrays in structure rsItems may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public rsItems As DAO.Recordset '120703 rjk: Added to support different commodities (theme games)
	
	Public IsDBOpen As Boolean
	
	'Nation Variables
	Public Education As Single
	Public TechLevel As Single
	Public Happiness As Single
	Public Research As Single
	Public CapSect As String
	Public Maxpop As Integer
	Public MaxSafeCivs As Integer
	Public MaxSafeUws As Integer
	Public HapNeeded As Single
	Public MilReserves As Integer
	Public HarborAvail As Boolean
	Public AirportAvail As Boolean
	Public FortAvail As Boolean
	Public HQAvail As Boolean
	
	Private Sub CheckDBFields(ByRef dbEmpire As DAO.Database)
		Dim idxNew As DAO.Index
		Dim iLoop As Short
		
		'270304 rjk: Added 'Rollover Avail' to Version table
		If dbEmpire.TableDefs("Version Info").Fields.Count = 39 Then
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("Rollover Avail", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Version Info").Fields("Rollover Avail").DefaultValue = 0
			If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
				rsVersion = dbEmpire.OpenRecordset("Version Info")
				rsVersion.MoveFirst()
				While Not rsVersion.EOF
					rsVersion.Edit()
					rsVersion.Fields("Rollover Avail").Value = 0
					rsVersion.Update()
					rsVersion.MoveNext()
				End While
			End If
		End If
		'220305 rjk: Added 'Version' to Version table
		If dbEmpire.TableDefs("Version Info").Fields.Count = 40 Then
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("Major", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Version Info").Fields("Major").DefaultValue = 0
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("Minor", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Version Info").Fields("Minor").DefaultValue = 0
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("Revision", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Version Info").Fields("Revision").DefaultValue = 0
			If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
				rsVersion = dbEmpire.OpenRecordset("Version Info")
				rsVersion.MoveFirst()
				While Not rsVersion.EOF
					rsVersion.Edit()
					rsVersion.Fields("Major").Value = 0
					rsVersion.Fields("Minor").Value = 0
					rsVersion.Fields("Revision").Value = 0
					rsVersion.Update()
					rsVersion.MoveNext()
				End While
			End If
		End If
		'181205 rjk: Added 'Eff gain - Sect' to Version table
		If dbEmpire.TableDefs("Version Info").Fields.Count = 43 Then
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("Eff gain - Sect", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Version Info").Fields("Eff gain - Sect").DefaultValue = 100
			If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
				rsVersion = dbEmpire.OpenRecordset("Version Info")
				rsVersion.MoveFirst()
				While Not rsVersion.EOF
					rsVersion.Edit()
					rsVersion.Fields("Eff gain - Sect").Value = 100
					rsVersion.Update()
					rsVersion.MoveNext()
				End While
			End If
		End If
		'110206 rjk: Added 'drnuke_const' to Version table
		If dbEmpire.TableDefs("Version Info").Fields.Count = 44 Then
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("DRNUKE Const", DAO.DataTypeEnum.dbSingle))
			dbEmpire.TableDefs("Version Info").Fields("DRNUKE Const").DefaultValue = 0#
			If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
				rsVersion = dbEmpire.OpenRecordset("Version Info")
				rsVersion.MoveFirst()
				While Not rsVersion.EOF
					rsVersion.Edit()
					rsVersion.Fields("DRNUKE Const").Value = 0#
					rsVersion.Update()
					rsVersion.MoveNext()
				End While
			End If
		End If
		'020907 rjk: Added 'timezoneadj' to Version table - difference between local time
		'            and server time
		If dbEmpire.TableDefs("Version Info").Fields.Count = 45 Then
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("Time Zone Adj", DAO.DataTypeEnum.dbLong))
			dbEmpire.TableDefs("Version Info").Fields("Time Zone Adj").DefaultValue = 0
			If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
				rsVersion = dbEmpire.OpenRecordset("Version Info")
				rsVersion.MoveFirst()
				While Not rsVersion.EOF
					rsVersion.Edit()
					rsVersion.Fields("Time Zone Adj").Value = 0
					rsVersion.Update()
					rsVersion.MoveNext()
				End While
			End If
		End If
		'160209 rjk: Added RAILWAYS options, originally part of the Heavy Metal/Plastic
		If dbEmpire.TableDefs("Version Info").Fields.Count = 46 Then
			dbEmpire.TableDefs("Version Info").Fields.Append(dbEmpire.TableDefs("Version Info").CreateField("RAILWAYS", DAO.DataTypeEnum.dbBoolean))
			dbEmpire.TableDefs("Version Info").Fields("RAILWAYS").DefaultValue = False
			If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
				rsVersion = dbEmpire.OpenRecordset("Version Info")
				rsVersion.MoveFirst()
				While Not rsVersion.EOF
					rsVersion.Edit()
					rsVersion.Fields("RAILWAYS").Value = False
					rsVersion.Update()
					rsVersion.MoveNext()
				End While
			End If
		End If
		'270804 rjk: Added Index to Sector Type table
		If dbEmpire.TableDefs("Sector Type").Fields.Count = 25 Then
			dbEmpire.TableDefs("Sector Type").Fields.Append(dbEmpire.TableDefs("Sector Type").CreateField("dchr_i", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Sector Type").Fields("dchr_i").DefaultValue = 0
			If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
				rsSectorType = dbEmpire.OpenRecordset("Sector Type")
				rsSectorType.MoveFirst()
				While Not rsSectorType.EOF
					rsSectorType.Edit()
					rsSectorType.Fields("dchr_i").Value = 0
					rsSectorType.Update()
					rsSectorType.MoveNext()
				End While
			End If
		End If
		'100904 rjk: Added Product Index to Sector Type table
		If dbEmpire.TableDefs("Sector Type").Fields.Count = 26 Then
			dbEmpire.TableDefs("Sector Type").Fields.Append(dbEmpire.TableDefs("Sector Type").CreateField("pchr_i", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Sector Type").Fields("pchr_i").DefaultValue = 0
			If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
				rsSectorType = dbEmpire.OpenRecordset("Sector Type")
				rsSectorType.MoveFirst()
				While Not rsSectorType.EOF
					rsSectorType.Edit()
					rsSectorType.Fields("pchr_i").Value = 0
					rsSectorType.Update()
					rsSectorType.MoveNext()
				End While
			End If
		End If
		
		'180606 rjk: Added mcost at 100% to the Sector Type table for 4.3.6 servers
		If dbEmpire.TableDefs("Sector Type").Fields.Count = 27 Then
			dbEmpire.TableDefs("Sector Type").Fields.Append(dbEmpire.TableDefs("Sector Type").CreateField("mcost100", DAO.DataTypeEnum.dbSingle))
			dbEmpire.TableDefs("Sector Type").Fields("mcost100").DefaultValue = -2#
			If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
				rsSectorType = dbEmpire.OpenRecordset("Sector Type")
				rsSectorType.MoveFirst()
				While Not rsSectorType.EOF
					rsSectorType.Edit()
					rsSectorType.Fields("mcost100").Value = -2#
					rsSectorType.Update()
					rsSectorType.MoveNext()
				End While
			End If
		End If
		
		'190606 rjk: Added nav
		If dbEmpire.TableDefs("Sector Type").Fields.Count = 28 Then
			dbEmpire.TableDefs("Sector Type").Fields.Append(dbEmpire.TableDefs("Sector Type").CreateField("nav", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Sector Type").Fields("nav").DefaultValue = 0
			If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
				rsSectorType = dbEmpire.OpenRecordset("Sector Type")
				rsSectorType.MoveFirst()
				While Not rsSectorType.EOF
					rsSectorType.Edit()
					rsSectorType.Fields("nav").Value = 0
					rsSectorType.Update()
					rsSectorType.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("Sector Type").Fields.Count = 29 Then
			dbEmpire.TableDefs("Sector Type").Fields.Append(dbEmpire.TableDefs("Sector Type").CreateField("pack_type", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Sector Type").Fields("pack_type").DefaultValue = 1
			If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
				rsSectorType = dbEmpire.OpenRecordset("Sector Type")
				rsSectorType.MoveFirst()
				While Not rsSectorType.EOF
					rsSectorType.Edit()
					rsSectorType.Fields("pack_type").Value = 1
					rsSectorType.Update()
					rsSectorType.MoveNext()
				End While
			End If
		End If
		'160209 rjk: Added Terrain to the sector type table
		If dbEmpire.TableDefs("Sector Type").Fields.Count = 30 Then
			dbEmpire.TableDefs("Sector Type").Fields.Append(dbEmpire.TableDefs("Sector Type").CreateField("terrain", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Sector Type").Fields("terrain").DefaultValue = 0
			If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
				rsSectorType = dbEmpire.OpenRecordset("Sector Type")
				rsSectorType.MoveFirst()
				While Not rsSectorType.EOF
					rsSectorType.Edit()
					rsSectorType.Fields("terrain").Value = 0
					rsSectorType.Update()
					rsSectorType.MoveNext()
				End While
			End If
		End If
		
		'280804 rjk: Added Index to Build Info table
		If dbEmpire.TableDefs("Build Info").Fields.Count = 52 Then
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("chr_i", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("chr_i").DefaultValue = 0
			If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
				rsBuild = dbEmpire.OpenRecordset("Build Info")
				rsBuild.MoveFirst()
				While Not rsBuild.EOF
					rsBuild.Edit()
					rsBuild.Fields("chr_i").Value = 0
					rsBuild.Update()
					rsBuild.MoveNext()
				End While
			End If
		End If
		'120306 rjk: Added Base Speed and Base Firing Range to Build Info table
		If dbEmpire.TableDefs("Build Info").Fields.Count = 53 Then
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base spd", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base spd").DefaultValue = 0
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base frnge", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base frnge").DefaultValue = 0
			If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
				rsBuild = dbEmpire.OpenRecordset("Build Info")
				rsBuild.MoveFirst()
				While Not rsBuild.EOF
					rsBuild.Edit()
					rsBuild.Fields("base spd").Value = 0
					rsBuild.Fields("base frnge").Value = 0
					rsBuild.Update()
					rsBuild.MoveNext()
				End While
			End If
		End If
		'230306 rjk: Added Base Attack and Base Defense to Build Info table
		If dbEmpire.TableDefs("Build Info").Fields.Count = 55 Then
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base att", DAO.DataTypeEnum.dbSingle))
			dbEmpire.TableDefs("Build Info").Fields("base att").DefaultValue = 0#
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base def", DAO.DataTypeEnum.dbSingle))
			dbEmpire.TableDefs("Build Info").Fields("base def").DefaultValue = 0#
			If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
				rsBuild = dbEmpire.OpenRecordset("Build Info")
				rsBuild.MoveFirst()
				While Not rsBuild.EOF
					rsBuild.Edit()
					rsBuild.Fields("base att").Value = 0#
					rsBuild.Fields("base def").Value = 0#
					rsBuild.Update()
					rsBuild.MoveNext()
				End While
			End If
		End If
		'210406 rjk: Added Base Accuracy to Build Info table
		If dbEmpire.TableDefs("Build Info").Fields.Count = 57 Then
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base acc", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base acc").DefaultValue = 0#
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base load", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base load").DefaultValue = 0#
			If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
				rsBuild = dbEmpire.OpenRecordset("Build Info")
				rsBuild.MoveFirst()
				While Not rsBuild.EOF
					rsBuild.Edit()
					rsBuild.Fields("base acc").Value = 0#
					rsBuild.Fields("base load").Value = 0#
					rsBuild.Update()
					rsBuild.MoveNext()
				End While
			End If
		End If
		'220406 rjk: Added Base Accuracy to Build Info table
		If dbEmpire.TableDefs("Build Info").Fields.Count = 59 Then
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base vul", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base vul").DefaultValue = 0#
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base dam", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base dam").DefaultValue = 0#
			dbEmpire.TableDefs("Build Info").Fields.Append(dbEmpire.TableDefs("Build Info").CreateField("base aaf", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Build Info").Fields("base aaf").DefaultValue = 0#
			If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
				rsBuild = dbEmpire.OpenRecordset("Build Info")
				rsBuild.MoveFirst()
				While Not rsBuild.EOF
					rsBuild.Edit()
					rsBuild.Fields("base vul").Value = 0#
					rsBuild.Fields("base dam").Value = 0#
					rsBuild.Fields("base aaf").Value = 0#
					rsBuild.Update()
					rsBuild.MoveNext()
				End While
			End If
		End If
		'300804 rjk: Add Index for Title to the Telegram Header table
		If dbEmpire.TableDefs("Telegram Header").Indexes.Count = 3 Then
			idxNew = dbEmpire.TableDefs("Telegram Header").CreateIndex("Title")
			With idxNew
				.Primary = False
				.Unique = False
				'UPGRADE_WARNING: Couldn't resolve default property of object idxNew.Fields.Append. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Fields.Append(.CreateField("Title"))
			End With
			
			dbEmpire.TableDefs("Telegram Header").Indexes.Append(idxNew)
			dbEmpire.TableDefs("Telegram Header").Indexes.Refresh()
		End If
		'050904 rjk: Added Index to Items table
		If dbEmpire.TableDefs("Items").Fields.Count = 17 Then
			dbEmpire.TableDefs("Items").Fields.Append(dbEmpire.TableDefs("Items").CreateField("pack_ie", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Items").Fields("pack_ie").DefaultValue = 1
			dbEmpire.TableDefs("Items").Fields.Append(dbEmpire.TableDefs("Items").CreateField("chr_i", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Items").Fields("chr_i").DefaultValue = 0
			If dbEmpire.TableDefs("Items").RecordCount > 0 Then
				rsItems = dbEmpire.OpenRecordset("Items")
				rsItems.MoveFirst()
				While Not rsItems.EOF
					rsItems.Edit()
					rsItems.Fields("pack_ie").Value = 1
					rsItems.Fields("chr_i").Value = 0
					rsItems.Update()
					rsItems.MoveNext()
				End While
			End If
		End If
		'110605 rjk: Added location Index to Mission table for reacting planes.
		If dbEmpire.TableDefs("Missions").Indexes.Count = 1 Then
			idxNew = dbEmpire.TableDefs("Missions").CreateIndex("Mission")
			With idxNew
				.Primary = False
				.Unique = False
				'UPGRADE_WARNING: Couldn't resolve default property of object idxNew.Fields.Append. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Fields.Append(.CreateField("type"))
				'UPGRADE_WARNING: Couldn't resolve default property of object idxNew.Fields.Append. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Fields.Append(.CreateField("mission"))
			End With
			
			dbEmpire.TableDefs("Missions").Indexes.Append(idxNew)
			dbEmpire.TableDefs("Missions").Indexes.Refresh()
		End If
		'170206 rjk: Added Index for Country Number in Nations table
		If dbEmpire.TableDefs("Nations").Indexes.Count = 0 Then
			idxNew = dbEmpire.TableDefs("Nations").CreateIndex("PrimaryKey")
			With idxNew
				.Primary = True
				.Unique = True
				'UPGRADE_WARNING: Couldn't resolve default property of object idxNew.Fields.Append. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Fields.Append(.CreateField("Number"))
			End With
			
			dbEmpire.TableDefs("Nations").Indexes.Append(idxNew)
			dbEmpire.TableDefs("Nations").Indexes.Refresh()
		End If
		
		'280506 rjk: Add timestamps to sector, ship, land, plane and nuke tables
		If dbEmpire.TableDefs("ship").Fields.Count = 36 Then
			dbEmpire.TableDefs("ship").Fields.Append(dbEmpire.TableDefs("ship").CreateField("timestamp", DAO.DataTypeEnum.dbLong))
			dbEmpire.TableDefs("ship").Fields("timestamp").DefaultValue = 0
			If dbEmpire.TableDefs("ship").RecordCount > 0 Then
				rsShip = dbEmpire.OpenRecordset("ship")
				rsShip.MoveFirst()
				While Not rsShip.EOF
					rsShip.Edit()
					rsShip.Fields("timestamp").Value = 0
					rsShip.Update()
					rsShip.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("land").Fields.Count = 43 Then
			dbEmpire.TableDefs("land").Fields.Append(dbEmpire.TableDefs("land").CreateField("timestamp", DAO.DataTypeEnum.dbLong))
			dbEmpire.TableDefs("land").Fields("timestamp").DefaultValue = 0
			If dbEmpire.TableDefs("land").RecordCount > 0 Then
				rsLand = dbEmpire.OpenRecordset("land")
				rsLand.MoveFirst()
				While Not rsLand.EOF
					rsLand.Edit()
					rsLand.Fields("timestamp").Value = 0
					rsLand.Update()
					rsLand.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("plane").Fields.Count = 23 Then
			dbEmpire.TableDefs("plane").Fields.Append(dbEmpire.TableDefs("plane").CreateField("timestamp", DAO.DataTypeEnum.dbLong))
			dbEmpire.TableDefs("plane").Fields("timestamp").DefaultValue = 0
			If dbEmpire.TableDefs("plane").RecordCount > 0 Then
				rsPlane = dbEmpire.OpenRecordset("plane")
				rsPlane.MoveFirst()
				While Not rsPlane.EOF
					rsPlane.Edit()
					rsPlane.Fields("timestamp").Value = 0
					rsPlane.Update()
					rsPlane.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("nuke").Fields.Count = 6 Then
			dbEmpire.TableDefs("nuke").Fields.Append(dbEmpire.TableDefs("nuke").CreateField("timestamp", DAO.DataTypeEnum.dbLong))
			dbEmpire.TableDefs("nuke").Fields("timestamp").DefaultValue = 0
			If dbEmpire.TableDefs("nuke").RecordCount > 0 Then
				rsNuke = dbEmpire.OpenRecordset("nuke")
				rsNuke.MoveFirst()
				While Not rsNuke.EOF
					rsNuke.Edit()
					rsNuke.Fields("timestamp").Value = 0
					rsNuke.Update()
					rsNuke.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("Sectors").Fields.Count = 89 Then
			dbEmpire.TableDefs("Sectors").Fields.Append(dbEmpire.TableDefs("Sectors").CreateField("timestamp", DAO.DataTypeEnum.dbLong))
			dbEmpire.TableDefs("Sectors").Fields("timestamp").DefaultValue = 0
			If dbEmpire.TableDefs("Sectors").RecordCount > 0 Then
				rsSectors = dbEmpire.OpenRecordset("Sectors")
				rsSectors.MoveFirst()
				While Not rsSectors.EOF
					rsSectors.Edit()
					rsSectors.Fields("timestamp").Value = 0
					rsSectors.Update()
					rsSectors.MoveNext()
				End While
			End If
		End If
		
		'250606 rjk: Add off to ship, land, plane and nuke tables
		If dbEmpire.TableDefs("ship").Fields.Count = 37 Then
			dbEmpire.TableDefs("ship").Fields.Append(dbEmpire.TableDefs("ship").CreateField("off", DAO.DataTypeEnum.dbBoolean))
			dbEmpire.TableDefs("ship").Fields("off").DefaultValue = 0
			If dbEmpire.TableDefs("ship").RecordCount > 0 Then
				rsShip = dbEmpire.OpenRecordset("ship")
				rsShip.MoveFirst()
				While Not rsShip.EOF
					rsShip.Edit()
					rsShip.Fields("off").Value = 0
					rsShip.Update()
					rsShip.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("land").Fields.Count = 44 Then
			dbEmpire.TableDefs("land").Fields.Append(dbEmpire.TableDefs("land").CreateField("off", DAO.DataTypeEnum.dbBoolean))
			dbEmpire.TableDefs("land").Fields("off").DefaultValue = 0
			If dbEmpire.TableDefs("land").RecordCount > 0 Then
				rsLand = dbEmpire.OpenRecordset("land")
				rsLand.MoveFirst()
				While Not rsLand.EOF
					rsLand.Edit()
					rsLand.Fields("off").Value = 0
					rsLand.Update()
					rsLand.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("plane").Fields.Count = 24 Then
			dbEmpire.TableDefs("plane").Fields.Append(dbEmpire.TableDefs("plane").CreateField("off", DAO.DataTypeEnum.dbBoolean))
			dbEmpire.TableDefs("plane").Fields("off").DefaultValue = 0
			If dbEmpire.TableDefs("plane").RecordCount > 0 Then
				rsPlane = dbEmpire.OpenRecordset("plane")
				rsPlane.MoveFirst()
				While Not rsPlane.EOF
					rsPlane.Edit()
					rsPlane.Fields("off").Value = 0
					rsPlane.Update()
					rsPlane.MoveNext()
				End While
			End If
		End If
		If dbEmpire.TableDefs("nuke").Fields.Count = 7 Then
			dbEmpire.TableDefs("nuke").Fields.Append(dbEmpire.TableDefs("nuke").CreateField("off", DAO.DataTypeEnum.dbBoolean))
			dbEmpire.TableDefs("nuke").Fields("off").DefaultValue = 0
			If dbEmpire.TableDefs("nuke").RecordCount > 0 Then
				rsNuke = dbEmpire.OpenRecordset("nuke")
				rsNuke.MoveFirst()
				While Not rsNuke.EOF
					rsNuke.Edit()
					rsNuke.Fields("off").Value = 0
					rsNuke.Update()
					rsNuke.MoveNext()
				End While
			End If
		End If
		
		'180606 rjk: Add id index for nuke Table if it is missing
		If dbEmpire.TableDefs("nuke").Indexes.Count = 2 Then
			idxNew = dbEmpire.TableDefs("nuke").CreateIndex("id")
			With idxNew
				.Primary = False
				.Unique = False
				'UPGRADE_WARNING: Couldn't resolve default property of object idxNew.Fields.Append. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Fields.Append(.CreateField("id"))
			End With
			
			dbEmpire.TableDefs("nuke").Indexes.Append(idxNew)
			dbEmpire.TableDefs("nuke").Indexes.Refresh()
		End If
		
		'230806 rjk: Added timestamps for oil and fert.
		If dbEmpire.TableDefs("Sea sectors").Fields.Count = 4 Then
			dbEmpire.TableDefs("Sea sectors").Fields.Append(dbEmpire.TableDefs("Sea sectors").CreateField("ftimestamp", DAO.DataTypeEnum.dbDate))
			dbEmpire.TableDefs("Sea sectors").Fields("ftimestamp").DefaultValue = vbNullString
			dbEmpire.TableDefs("Sea sectors").Fields.Append(dbEmpire.TableDefs("Sea sectors").CreateField("otimestamp", DAO.DataTypeEnum.dbDate))
			dbEmpire.TableDefs("Sea sectors").Fields("otimestamp").DefaultValue = vbNullString
			If dbEmpire.TableDefs("Sea sectors").RecordCount > 0 Then
				rsSea = dbEmpire.OpenRecordset("Sea sectors")
				rsSea.MoveFirst()
				While Not rsSea.EOF
					rsSea.Edit()
					rsSea.Fields("ftimestamp").Value = GetWinACEUTC
					rsSea.Fields("otimestamp").Value = GetWinACEUTC
					rsSea.Update()
					rsSea.MoveNext()
				End While
			End If
		End If
		
		'230806 rjk: Added end cargo for orders table.
		If dbEmpire.TableDefs("Ship orders").Fields.Count = 25 Then
			For iLoop = 1 To 6
				dbEmpire.TableDefs("Ship orders").Fields.Append(dbEmpire.TableDefs("Ship orders").CreateField("end cargo " & CStr(iLoop), DAO.DataTypeEnum.dbText, 50))
				dbEmpire.TableDefs("Ship orders").Fields("end cargo " & CStr(iLoop)).Properties.Append(dbEmpire.TableDefs("Ship orders").CreateProperty("UnicodeCompression", DAO.DataTypeEnum.dbBoolean, True))
				dbEmpire.TableDefs("Ship orders").Fields("end cargo " & CStr(iLoop)).Properties("AllowZeroLength").Value = True
				dbEmpire.TableDefs("Ship orders").Fields("end cargo " & CStr(iLoop)).DefaultValue = vbNullString
				dbEmpire.TableDefs("Ship orders").Fields("cargo " & CStr(iLoop)).Name = "start cargo " & CStr(iLoop)
				'End If
				
				If dbEmpire.TableDefs("Ship orders").RecordCount > 0 Then
					rsShipOrders = dbEmpire.OpenRecordset("Ship orders")
					rsShipOrders.MoveFirst()
					While Not rsShipOrders.EOF
						rsShipOrders.Edit()
						rsShipOrders.Fields("end cargo " & CStr(iLoop)).Value = rsShipOrders.Fields("start cargo " & CStr(iLoop)).Value
						rsShipOrders.Update()
						rsShipOrders.MoveNext()
					End While
					rsShipOrders.Close()
					'UPGRADE_NOTE: Object rsShipOrders may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsShipOrders = Nothing
				End If
			Next 
		End If
		
		'030108 rjk: Added elevation to the sector table
		If dbEmpire.TableDefs("Sectors").Fields.Count = 90 Then
			dbEmpire.TableDefs("Sectors").Fields.Append(dbEmpire.TableDefs("Sectors").CreateField("elev", DAO.DataTypeEnum.dbInteger))
			dbEmpire.TableDefs("Sectors").Fields("elev").DefaultValue = 0
			If dbEmpire.TableDefs("Sectors").RecordCount > 0 Then
				rsSectors = dbEmpire.OpenRecordset("Sectors")
				rsSectors.MoveFirst()
				While Not rsSectors.EOF
					rsSectors.Edit()
					rsSectors.Fields("elev").Value = 0
					rsSectors.Update()
					rsSectors.MoveNext()
				End While
			End If
		End If
		
		'Check nuke Table, added country field if necessary
		'If dbEmpire.TableDefs("nuke").Fields.Count = 5 Then
		'    dbEmpire.TableDefs("nuke").Fields.Append dbEmpire.TableDefs("nuke").CreateField("country", dbInteger)
		'End If
		'010204 rjk: Add PrimaryKey (id) for nuke Table if it is missing
		'If dbEmpire.TableDefs("nuke").Indexes.Count = 1 Then
		'    Set idxNew = dbEmpire.TableDefs("nuke").CreateIndex("PrimaryKey")
		'    With idxNew
		'        .Primary = True
		'        .Unique = True
		'        .Fields.Append .CreateField("id")
		'    End With
		'
		'    dbEmpire.TableDefs("nuke").Indexes.Append idxNew
		'    dbEmpire.TableDefs("nuke").Indexes.Refresh
		'End If
		
		'Check land Table, added civ and uw fields if necessary
		'If dbEmpire.TableDefs("land").Fields.Count = 41 Then
		'    dbEmpire.TableDefs("land").Fields.Append dbEmpire.TableDefs("land").CreateField("civ", dbInteger)
		'    dbEmpire.TableDefs("land").Fields.Append dbEmpire.TableDefs("land").CreateField("uw", dbInteger)
		'End If
		
		'If dbEmpire.TableDefs("Version Info").Fields(3).Name = "ETU per undate" Then
		'    dbEmpire.TableDefs("Version Info").Fields(3).Name = "ETU per update"
		'End If
		
		'120703 rjk: Added Items table for storing the items used in the game
		'If dbEmpire.TableDefs.Count = 28 Then 'add Items table
		'    Dim tdfNew As TableDef
		'
		'    Set tdfNew = dbEmpire.CreateTableDef("Items")
		'    With tdfNew
		'        .Fields.Append .CreateField("letter", dbText)
		'        .Fields.Append .CreateField("value", dbInteger)
		'        .Fields.Append .CreateField("sell", dbBoolean)
		'        .Fields.Append .CreateField("lbs", dbInteger)
		'        .Fields.Append .CreateField("pack_rg", dbInteger)
		'        .Fields.Append .CreateField("pack_wh", dbInteger)
		'        .Fields.Append .CreateField("pack_ur", dbInteger)
		'        .Fields.Append .CreateField("pack_bk", dbInteger)
		'        .Fields.Append .CreateField("name", dbText)
		'        .Fields.Append .CreateField("p_sname", dbText)
		'        .Fields.Append .CreateField("cond_name", dbText)
		'        .Fields.Append .CreateField("db_name", dbText)
		'        .Fields.Append .CreateField("sector_panel_label", dbText)
		'        .Fields.Append .CreateField("distribution_panel_label", dbText)
		'        .Fields.Append .CreateField("form_name", dbText)
		'        .Fields.Append .CreateField("intelligence_names", dbText)
		'        .Fields.Append .CreateField("form_label", dbText)
		'
		'        Set idxNew = .CreateIndex("PrimaryKey")
		'        With idxNew
		'                .Fields.Append .CreateField("letter")
		'        End With
		'        .Indexes.Append idxNew
		'
		'        .Indexes.Refresh
		'    End With
		'    dbEmpire.TableDefs.Append tdfNew
		'End If
	End Sub
	
	'060304 rjk: Changed the Enemy timestamp to UTC
	'210404 rjk: Remove the local copy of tz to use the global tz
	Private Sub UpdateEnemyTimestamp()
		
		'update the timestamp if is not UTC format
		If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
			rsEnemySect.MoveFirst()
			While Not rsEnemySect.EOF
				If Len(rsEnemySect.Fields("timestamp").Value) > 0 Then
					If Mid(rsEnemySect.Fields("timestamp").Value, 5, 1) <> "/" Then
						rsEnemySect.Edit()
						rsEnemySect.Fields("timestamp").Value = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, tz.Bias, EmpireTimeString(rsEnemySect.Fields("timestamp").Value)), "yyyy/mm/dd hh:mm:ss")
						rsEnemySect.Update()
					End If
				End If
				rsEnemySect.MoveNext()
			End While
		End If
		
		If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
			rsEnemyUnit.MoveFirst()
			While Not rsEnemyUnit.EOF
				If Len(rsEnemyUnit.Fields("timestamp").Value) > 0 Then
					If Mid(rsEnemyUnit.Fields("timestamp").Value, 5, 1) <> "/" Then
						rsEnemyUnit.Edit()
						rsEnemyUnit.Fields("timestamp").Value = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, tz.Bias, EmpireTimeString(rsEnemyUnit.Fields("timestamp").Value)), "yyyy/mm/dd hh:mm:ss")
						rsEnemyUnit.Update()
					End If
				End If
				rsEnemyUnit.MoveNext()
			End While
		End If
	End Sub
	
	Public Function OpenEmpireDB() As Boolean
		'Set Database name
		dbsName = My.Application.Info.DirectoryPath & "\Games\" & GameName & ".mdb"
		
		'Set data base files
		dbEmpire = DAODBEngine_definst.OpenDatabase(dbsName, False, False)
		CheckDBFields(dbEmpire)
		rsSectors = dbEmpire.OpenRecordset("Sectors")
		rsBmap = dbEmpire.OpenRecordset("Bmap")
		rsShip = dbEmpire.OpenRecordset("ship")
		rsShipList = dbEmpire.OpenRecordset("shiplist")
		rsLand = dbEmpire.OpenRecordset("land")
		rsNuke = dbEmpire.OpenRecordset("nuke")
		rsLandList = dbEmpire.OpenRecordset("landlist")
		rsPlaneList = dbEmpire.OpenRecordset("planelist")
		rsPlane = dbEmpire.OpenRecordset("plane")
		rsTeleHead = dbEmpire.OpenRecordset("Telegram Header")
		rsTeleBody = dbEmpire.OpenRecordset("Telegram Body")
		rsEnemySect = dbEmpire.OpenRecordset("Enemy Sect")
		rsEnemyUnit = dbEmpire.OpenRecordset("Enemy Units")
		UpdateEnemyTimestamp()
		rsVersion = dbEmpire.OpenRecordset("Version Info")
		rsBuild = dbEmpire.OpenRecordset("Build Info")
		rsSea = dbEmpire.OpenRecordset("Sea sectors")
		rsAirCombat = dbEmpire.OpenRecordset("Air Combat")
		rsShipOrders = dbEmpire.OpenRecordset("Ship orders")
		rsMissions = dbEmpire.OpenRecordset("Missions")
		rsSectorType = dbEmpire.OpenRecordset("Sector type")
		rsShowText = dbEmpire.OpenRecordset("Show Text")
		rsNations = dbEmpire.OpenRecordset("Nations")
		rsNation = dbEmpire.OpenRecordset("Nation")
		rsItems = dbEmpire.OpenRecordset("Items")
		Items = New EmpItems
		
		rsSectors.Index = "PrimaryKey"
		rsBmap.Index = "PrimaryKey"
		rsShip.Index = "location"
		rsLand.Index = "location"
		rsPlane.Index = "location"
		rsNuke.Index = "location"
		rsTeleHead.Index = "PrimaryKey"
		rsTeleBody.Index = "PrimaryKey"
		rsEnemySect.Index = "PrimaryKey"
		rsEnemyUnit.Index = "PrimaryKey"
		rsBuild.Index = "PrimaryKey"
		rsSea.Index = "PrimaryKey"
		rsAirCombat.Index = "PrimaryKey"
		rsShipOrders.Index = "PrimaryKey"
		rsMissions.Index = "PrimaryKey"
		rsSectorType.Index = "PrimaryKey"
		rsShowText.Index = "PrimaryKey"
		rsItems.Index = "PrimaryKey"
		rsNations.Index = "PrimaryKey"
		
		OpenEmpireDB = True
		IsDBOpen = True
		
		'Set the nation variables
		GetNationVariables()
	End Function
	
	Sub CloseEmpireDB()
		
		If IsDBOpen Then
			'Set data base files
			'UPGRADE_NOTE: Object rsSectors may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsSectors = Nothing
			'UPGRADE_NOTE: Object rsBmap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsBmap = Nothing
			'UPGRADE_NOTE: Object rsShip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsShip = Nothing
			'UPGRADE_NOTE: Object rsLand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsLand = Nothing
			'UPGRADE_NOTE: Object rsNuke may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsNuke = Nothing
			'UPGRADE_NOTE: Object rsShipList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsShipList = Nothing
			'UPGRADE_NOTE: Object rsLandList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsLandList = Nothing
			'UPGRADE_NOTE: Object rsPlane may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsPlane = Nothing
			'UPGRADE_NOTE: Object rsTeleHead may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTeleHead = Nothing
			'UPGRADE_NOTE: Object rsTeleBody may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTeleBody = Nothing
			'UPGRADE_NOTE: Object rsEnemySect may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsEnemySect = Nothing
			'UPGRADE_NOTE: Object rsEnemyUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsEnemyUnit = Nothing
			'UPGRADE_NOTE: Object rsVersion may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsVersion = Nothing
			'UPGRADE_NOTE: Object rsNations may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsNations = Nothing
			'UPGRADE_NOTE: Object rsNation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsNation = Nothing
			'UPGRADE_NOTE: Object rsBuild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsBuild = Nothing
			'UPGRADE_NOTE: Object rsSea may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsSea = Nothing
			'UPGRADE_NOTE: Object rsAirCombat may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsAirCombat = Nothing
			'UPGRADE_NOTE: Object rsShipOrders may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsShipOrders = Nothing
			'UPGRADE_NOTE: Object rsMissions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsMissions = Nothing
			
			dbEmpire.Close()
			IsDBOpen = False
		End If
		
	End Sub
	
	Public Sub ClearUnitsFiles()
		DeleteAllRecords(rsPlane)
		DeleteAllRecords(rsLand)
		DeleteAllRecords(rsShip)
	End Sub
	
	Public Sub ResetDataBase()
		
		On Error Resume Next 'MoveFirst will give error on empty dataset
		'Clear current sector File
		DeleteAllRecords(rsSectors)
		
		'Clear Plane, ship, and land file
		ClearUnitsFiles()
		
		DeleteAllRecords(rsNuke)
		
		'Clear Bmap file
		'don't clear the bmap file if we are suppressing the bmap update
		'120303 rjk: Switched options to frmOptions and boolean options
		If frmOptions.bSuppressBmapRefresh Then
			Exit Sub
		End If
		
		DeleteAllRecords(rsBmap)
	End Sub
	
	Public Sub DeleteAllRecords(ByRef rsDelete As DAO.Recordset)
		With rsDelete
			If .BOF And .EOF Then Exit Sub
			
			.MoveFirst()
			Do While Not .EOF
				.Delete()
				.MoveNext()
			Loop 
		End With
	End Sub
	
	Public Sub DeleteAllRecordsOlderThen(ByRef rsDelete As DAO.Recordset, ByRef dExpiry As Date)
		Dim dRecord As Date
		
		With rsDelete
			If .BOF And .EOF Then Exit Sub
			
			.MoveFirst()
			Do While Not .EOF
				dRecord = ConvertToLocalTimeFromWinACEUTC((.Fields("timestamp").Value))
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If DateDiff(Microsoft.VisualBasic.DateInterval.Day, dExpiry, dRecord) <= 0 Then
					.Delete()
				End If
				.MoveNext()
			Loop 
		End With
	End Sub
	
	Public Sub DeleteAllAirCombatRecordsOlderThen(ByRef dExpiry As Date)
		Dim dRecord As Date
		
		With rsAirCombat
			If .BOF And .EOF Then Exit Sub
			
			.MoveFirst()
			Do While Not .EOF
				dRecord = ConvertToLocalTimeFromWinACEUTC((.Fields("nextupdate").Value))
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If DateDiff(Microsoft.VisualBasic.DateInterval.Day, dExpiry, dRecord) <= 0 Then
					.Delete()
				End If
				.MoveNext()
			Loop 
		End With
	End Sub
	
	Public Sub DeleteBuildRecords(ByRef strType As String)
		With rsBuild
			If .BOF And .EOF Then Exit Sub
			
			.MoveFirst()
			Do While Not .EOF
				If .Fields("type").Value = strType Then
					.Delete()
				End If
				.MoveNext()
			Loop 
		End With
	End Sub
	
	Public Sub ShiftOrigin(ByRef sx As Short, ByRef sy As Short)
		
		Dim hldIndex As Object
		Dim holdx As Short
		Dim holdy As Short
		Dim n As Short
		
		On Error Resume Next 'MoveFirst will give error on empty dataset
		
		'Programmer notes
		'There is a problem with just updating the x,y values directly in the
		'files keyed by location (Sector, bmap, Enemy Sect)  Since files are
		'processed in key order, the danger is that records will be processed
		'over and ower, as increasing the x,y values will put them later in the
		'key order. The + - 2000 techique used to avoid this insures that after
		'the first read/update, the record WILL go later in the key order, where
		'the second read/update will correct it and push it back earlier in the key
		'order so it won't be read again
		
		'Update current sector File
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsSectors.Index
		rsSectors.Index = "Primary Key"
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			If rsSectors.Fields("x").Value < 1000 Then
				rsSectors.Edit()
				rsSectors.Fields("x").Value = rsSectors.Fields("x").Value - sx + 2000
				rsSectors.Fields("y").Value = rsSectors.Fields("y").Value - sy + 2000
				holdx = rsSectors.Fields("dist_x").Value - sx
				holdy = rsSectors.Fields("dist_y").Value - sy
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsSectors.Fields("dist_x").Value = holdx
				rsSectors.Fields("dist_y").Value = holdy
				rsSectors.Update()
			End If
			rsSectors.MoveNext()
		End While
		
		'Second Pass
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			If rsSectors.Fields("x").Value > 1000 Then
				rsSectors.Edit()
				holdx = rsSectors.Fields("x").Value - 2000
				holdy = rsSectors.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsSectors.Fields("x").Value = holdx
				rsSectors.Fields("y").Value = holdy
				rsSectors.Update()
			End If
			rsSectors.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsSectors.Index = hldIndex
		
		'Update Sea Information file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsSea.Index
		rsSea.Index = "PrimaryKey"
		rsSea.MoveFirst()
		While Not rsSea.EOF
			If rsSea.Fields("x").Value < 1000 Then
				rsSea.Edit()
				rsSea.Fields("x").Value = rsSea.Fields("x").Value - sx + 2000
				rsSea.Fields("y").Value = rsSea.Fields("y").Value - sy + 2000
				rsSea.Update()
			End If
			rsSea.MoveNext()
		End While
		
		'Second Pass
		rsSea.MoveFirst()
		While Not rsSea.EOF
			If rsSea.Fields("x").Value > 1000 Then
				rsSea.Edit()
				holdx = rsSea.Fields("x").Value - 2000
				holdy = rsSea.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsSea.Fields("x").Value = holdx
				rsSea.Fields("y").Value = holdy
				rsSea.Update()
			End If
			rsSea.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsSea.Index = hldIndex
		
		'Update Bmap file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsBmap.Index
		rsBmap.Index = "Primary Key"
		rsBmap.MoveFirst()
		While Not rsBmap.EOF
			If rsBmap.Fields("x").Value < 1000 Then
				rsBmap.Edit()
				rsBmap.Fields("x").Value = rsBmap.Fields("x").Value - sx + 2000
				rsBmap.Fields("y").Value = rsBmap.Fields("y").Value - sy + 2000
				rsBmap.Update()
			End If
			rsBmap.MoveNext()
		End While
		
		'Second Pass
		rsBmap.MoveFirst()
		While Not rsBmap.EOF
			If rsBmap.Fields("x").Value > 1000 Then
				rsBmap.Edit()
				holdx = rsBmap.Fields("x").Value - 2000
				holdy = rsBmap.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsBmap.Fields("x").Value = holdx
				rsBmap.Fields("y").Value = holdy
				rsBmap.Update()
			End If
			rsBmap.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsBmap.Index = hldIndex
		
		
		'Update Enemy Sector file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsEnemySect.Index
		rsEnemySect.Index = "Primary Key"
		rsEnemySect.MoveFirst()
		While Not rsEnemySect.EOF
			If rsEnemySect.Fields("x").Value < 1000 Then
				rsEnemySect.Edit()
				rsEnemySect.Fields("x").Value = rsEnemySect.Fields("x").Value - sx + 2000
				rsEnemySect.Fields("y").Value = rsEnemySect.Fields("y").Value - sy + 2000
				rsEnemySect.Update()
			End If
			rsEnemySect.MoveNext()
		End While
		
		rsEnemySect.MoveFirst()
		While Not rsEnemySect.EOF
			If rsEnemySect.Fields("x").Value > 1000 Then
				rsEnemySect.Edit()
				holdx = rsEnemySect.Fields("x").Value - 2000
				holdy = rsEnemySect.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsEnemySect.Fields("x").Value = holdx
				rsEnemySect.Fields("y").Value = holdy
				rsEnemySect.Update()
			End If
			rsEnemySect.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsEnemySect.Index = hldIndex
		
		'Update Plane file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsPlane.Index
		rsPlane.Index = "Primary Key"
		rsPlane.MoveFirst()
		While Not rsPlane.EOF
			If rsPlane.Fields("x").Value < 1000 Then
				rsPlane.Edit()
				rsPlane.Fields("x").Value = rsPlane.Fields("x").Value - sx + 2000
				rsPlane.Fields("y").Value = rsPlane.Fields("y").Value - sy + 2000
				rsPlane.Update()
			End If
			rsPlane.MoveNext()
		End While
		
		rsPlane.MoveFirst()
		While Not rsPlane.EOF
			If rsPlane.Fields("x").Value > 1000 Then
				rsPlane.Edit()
				holdx = rsPlane.Fields("x").Value - 2000
				holdy = rsPlane.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsPlane.Fields("x").Value = holdx
				rsPlane.Fields("y").Value = holdy
				rsPlane.Update()
			End If
			rsPlane.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPlane.Index = hldIndex
		
		'Update Land file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsLand.Index
		rsLand.Index = "Primary Key"
		rsLand.MoveFirst()
		While Not rsLand.EOF
			If rsLand.Fields("x").Value < 1000 Then
				rsLand.Edit()
				rsLand.Fields("x").Value = rsLand.Fields("x").Value - sx + 2000
				rsLand.Fields("y").Value = rsLand.Fields("y").Value - sy + 2000
				rsLand.Update()
			End If
			rsLand.MoveNext()
		End While
		rsLand.MoveFirst()
		While Not rsLand.EOF
			If rsLand.Fields("x").Value > 1000 Then
				rsLand.Edit()
				holdx = rsLand.Fields("x").Value - 2000
				holdy = rsLand.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsLand.Fields("x").Value = holdx
				rsLand.Fields("y").Value = holdy
				rsLand.Update()
			End If
			rsLand.MoveNext()
		End While
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsLand.Index = hldIndex
		
		'Update Ship file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsShip.Index
		rsShip.Index = "Primary Key"
		rsShip.MoveFirst()
		While Not rsShip.EOF
			If rsShip.Fields("x").Value < 1000 Then
				rsShip.Edit()
				rsShip.Fields("x").Value = rsShip.Fields("x").Value - sx + 2000
				rsShip.Fields("y").Value = rsShip.Fields("y").Value - sy + 2000
				rsShip.Update()
			End If
			rsShip.MoveNext()
		End While
		rsShip.MoveFirst()
		While Not rsShip.EOF
			If rsShip.Fields("x").Value > 1000 Then
				rsShip.Edit()
				holdx = rsShip.Fields("x").Value - 2000
				holdy = rsShip.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsShip.Fields("x").Value = holdx
				rsShip.Fields("y").Value = holdy
				rsShip.Update()
			End If
			rsShip.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsShip.Index = hldIndex
		
		'Update Nuke file
		rsNuke.MoveFirst()
		While Not rsNuke.EOF
			If rsNuke.Fields("x").Value < 1000 Then
				rsNuke.Edit()
				rsNuke.Fields("x").Value = rsNuke.Fields("x").Value - sx + 2000
				rsNuke.Fields("y").Value = rsNuke.Fields("y").Value - sy + 2000
				rsNuke.Update()
			End If
			rsNuke.MoveNext()
		End While
		
		rsNuke.MoveFirst()
		While Not rsNuke.EOF
			If rsNuke.Fields("x").Value > 1000 Then
				rsNuke.Edit()
				holdx = rsNuke.Fields("x").Value - 2000
				holdy = rsNuke.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsNuke.Fields("x").Value = holdx
				rsNuke.Fields("y").Value = holdy
				rsNuke.Update()
			End If
			rsNuke.MoveNext()
		End While
		
		'Update Enemy Unit file
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsEnemyUnit.Index
		rsEnemyUnit.Index = "Primary Key"
		rsEnemyUnit.MoveFirst()
		While Not rsEnemyUnit.EOF
			rsEnemyUnit.Edit()
			If rsEnemyUnit.Fields("x").Value < 1000 Then
				rsEnemyUnit.Fields("x").Value = rsEnemyUnit.Fields("x").Value - sx + 2000
				rsEnemyUnit.Fields("y").Value = rsEnemyUnit.Fields("y").Value - sy + 2000
				rsEnemyUnit.Update()
			End If
			rsEnemyUnit.MoveNext()
		End While
		
		rsEnemyUnit.MoveFirst()
		While Not rsEnemyUnit.EOF
			rsEnemyUnit.Edit()
			If rsEnemyUnit.Fields("x").Value > 1000 Then
				holdx = rsEnemyUnit.Fields("x").Value - 2000
				holdy = rsEnemyUnit.Fields("y").Value - 2000
				OffsetSectorLocation(holdx, holdy, 0, 0)
				rsEnemyUnit.Fields("x").Value = holdx
				rsEnemyUnit.Fields("y").Value = holdy
				rsEnemyUnit.Update()
			End If
			rsEnemyUnit.MoveNext()
		End While
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsEnemyUnit.Index = hldIndex
		
		'Move jump points
		rsNation.MoveFirst()
		rsNation.Edit()
		For n = 1 To 5
			If ParseSectors(holdx, holdy, (rsNation.Fields("JumpPoint" & CStr(n)).Value)) Then
				OffsetSectorLocation(holdx, holdy, -1 * sx, -1 * sy)
				rsNation.Fields("JumpPoint" & CStr(n)).Value = SectorString(holdx, holdy)
			End If
		Next n
		rsNation.Update()
		
	End Sub
	
	Public Sub GetNationVariables()
		On Error GoTo Empire_Error
		
		etsTable = New EmpTables
		estsTable = New EmpSymbolTables
		
		If rsNation.BOF And rsNation.EOF Then
			Exit Sub
		End If
		
		rsNation.MoveFirst()
		'Fill the nation variables from the database
		Education = rsNation.Fields("Education").Value
		TechLevel = rsNation.Fields("TechLevel").Value
		Happiness = rsNation.Fields("Happiness").Value
		Research = rsNation.Fields("Research").Value
		CapSect = rsNation.Fields("CapSect").Value
		Maxpop = rsNation.Fields("Maxpop").Value
		MaxSafeCivs = rsNation.Fields("MaxSafeCivs").Value
		MaxSafeUws = rsNation.Fields("MaxSafeUws").Value
		HapNeeded = rsNation.Fields("HapNeeded").Value
		MilReserves = rsNation.Fields("MilReserves").Value
		HarborAvail = rsNation.Fields("HarborAvail").Value
		AirportAvail = rsNation.Fields("AirportAvail").Value
		FortAvail = rsNation.Fields("FortAvail").Value
		HQAvail = rsNation.Fields("HqAvail").Value
		
		Exit Sub
Empire_Error: 
		EmpireError("GetNationVariables", vbNullString, vbNullString)
	End Sub
	
	Public Sub SaveNationVariables()
		On Error GoTo Empire_Error
		
		If rsNation.BOF And rsNation.EOF Then
			rsNation.AddNew()
		Else
			rsNation.MoveFirst()
			rsNation.Edit()
		End If
		
		'Fill the nation variables from the database
		rsNation.Fields("Education").Value = Education
		rsNation.Fields("TechLevel").Value = TechLevel
		rsNation.Fields("Happiness").Value = Happiness
		rsNation.Fields("Research").Value = Research
		rsNation.Fields("CapSect").Value = CapSect
		rsNation.Fields("Maxpop").Value = Maxpop
		rsNation.Fields("MaxSafeCivs").Value = MaxSafeCivs
		rsNation.Fields("MaxSafeUws").Value = MaxSafeUws
		rsNation.Fields("HapNeeded").Value = HapNeeded
		rsNation.Fields("MilReserves").Value = MilReserves
		rsNation.Fields("HarborAvail").Value = HarborAvail
		rsNation.Fields("AirportAvail").Value = AirportAvail
		rsNation.Fields("FortAvail").Value = FortAvail
		rsNation.Fields("HqAvail").Value = HQAvail
		rsNation.Update()
		
		Exit Sub
Empire_Error: 
		EmpireError("SaveNationVariables", vbNullString, vbNullString)
	End Sub
	
	Function SetAccessProperty(ByRef obj As Object, ByRef strName As String, ByRef intType As Short, ByRef varSetting As Object) As Boolean
		Dim prp As DAO.Property
		Const conPropNotFound As Short = 3270
		
		On Error GoTo ErrorSetAccessProperty
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.Properties. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object varSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.Properties(strName) = varSetting
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.Properties. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.Properties.Refresh()
		SetAccessProperty = True
		Exit Function
		
ErrorSetAccessProperty: 
		If Err.Number = conPropNotFound Then
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.CreateProperty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			prp = obj.CreateProperty(strName, intType, varSetting)
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.Properties. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj.Properties.Append(prp)
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.Properties. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj.Properties.Refresh()
			SetAccessProperty = True
		Else
			MsgBox(Err.Number & ": " & vbCrLf & Err.Description)
			SetAccessProperty = False
		End If
	End Function
End Module