Attribute VB_Name = "XdumpParse"
Option Explicit

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
Private tmStamp As Long
Private iRowCount As Integer

Public Sub ParseXdump(strLine As String, bUpdateTimeStamp As Boolean)
Dim strParams() As String

'check for an announcement or tele message in the middle of a dump
If InStr(strLine, "You have") > 0 Then
    Exit Sub
End If

If Left$(strLine, 6) = "XDUMP " Then
    xdType = XD_UNKNOWN
    tmStamp = 0
    strParams = Split(strLine, " ")
    iRowCount = 0
    If UBound(strParams) = 3 Then
        ParseXDMetaHeader (strParams(2))
        xdMeta = True
        tmStamp = CLng(strParams(3))
    ElseIf UBound(strParams) = 2 Then
        xdType = ParseXDTableName(strParams(1))
        If xdType <> XD_UNKNOWN Then
            tmStamp = CLng(strParams(2))
            If bUpdateTimeStamp Then
                Select Case xdType
                Case XD_LAND
                    tsLand = tmStamp
                Case XD_NUKE
                    tsNuke = tmStamp
                Case XD_PLANE
                    tsPlane = tmStamp
                Case XD_SECTOR
                    tsSectors = tmStamp
                Case XD_SHIP
                    tsShip = tmStamp
                Case XD_LOST
                    tsLost = tmStamp
                Case XD_LAND_CHR
                    DeleteBuildRecords "l"
                Case XD_PLANE_CHR
                    DeleteBuildRecords "p"
                Case XD_SHIP_CHR
                    DeleteBuildRecords "s"
                Case XD_SECTOR_CHR
                    DeleteBuildRecords "i"
                Case XD_COUNTRIES
                    DeleteAllRecords rsNations
                End Select
            End If
            If frmEmpCmd.ForceUpdates Then
                Select Case xdType
                Case XD_LAND
                    DeleteAllRecords rsLand
                Case XD_NUKE
                    DeleteAllRecords rsNuke
                Case XD_PLANE
                    DeleteAllRecords rsPlane
                Case XD_SECTOR
                    DeleteAllRecords rsSectors
                Case XD_SHIP
                    DeleteAllRecords rsShip
                End Select
            End If
            Select Case xdType
            Case XD_LAND_CHR
                UpdateTechLevelForShow "l", "b"
                UpdateTechLevelForShow "l", "c"
                UpdateTechLevelForShow "l", "s"
            Case XD_NUKE_CHR
                UpdateTechLevelForShow "n", "b"
                UpdateTechLevelForShow "n", "s"
            Case XD_PLANE_CHR
                UpdateTechLevelForShow "p", "b"
                UpdateTechLevelForShow "p", "c"
                UpdateTechLevelForShow "p", "s"
            Case XD_SECTOR_CHR
                UpdateTechLevelForShow "c", "b"
                UpdateTechLevelForShow "c", "h"
                UpdateTechLevelForShow "c", "s"
            Case XD_SHIP_CHR
                UpdateTechLevelForShow "s", "b"
                UpdateTechLevelForShow "s", "c"
                UpdateTechLevelForShow "s", "s"
            Case XD_VERSION
                UpdateTechLevelForShow "b", "b"
                UpdateTechLevelForShow "t", "b"
            Case XD_LEVEL, XD_META_TYPE, XD_MISSIONS, XD_SECTOR_NAVIGATION, _
                XD_SHIP_CHR_FLAGS
                ParseXDSymbolTableHeader xdType, strParams(1)
            End Select
        Else
            Debug.Print "ParseXdump:Invalid xdump type:" + CStr(xdType) + ", " + strLine
        End If
    Else
        EmpireError "ParseXdump", "Invalid xdump type:" + CStr(xdType), strLine
    End If
ElseIf Left$(strLine, 12) = "Usage: xdump" Then
    Exit Sub
ElseIf Left$(strLine, 1) = "/" Then
    Select Case xdType
    Case XD_COUNTRIES
        frmChat.ControlFlashForm
    Case XD_SHIP, XD_SECTOR, XD_LAND, XD_NUKE, XD_PLANE
        frmDrawMap.Refresh
        If rsVersion.Fields("RAILWAYS") And frmEmpCmd.ForceUpdates Then
            UpdateRailTrack
        End If
    Case XD_VERSION, XD_SHIP_CHR, XD_SECTOR_CHR, XD_LAND_CHR, XD_PLANE_CHR, _
        XD_PRODUCT_CHR, XD_ITEM_CHR, XD_INFRA_CHR
        If xdType = XD_ITEM_CHR Then
            Items.UpdateDatabase
        End If
        If frmToolShow.Visible Then
            frmToolShow.FillGrid
        End If
        
    End Select
    xdType = XD_UNKNOWN
    xdMeta = False
ElseIf xdMeta Then
    ParseXDMeta xdMetaType, strLine
Else
    Select Case xdType
    Case XD_COUNTRIES
        ParseXDCountries strLine
    Case XD_GAME
        ParseXDGame strLine
    Case XD_ITEM_CHR
        ParseXDItemChr strLine
    Case XD_INFRA_CHR
        ParseXDInfraChr strLine
    Case XD_LAND
        ParseXDLand strLine
    Case XD_LOST
        ParseXDLost strLine
    Case XD_LAND_CHR
        ParseXDLandChr strLine
    Case XD_NATION
        ParseXDNation strLine
    Case XD_NATION_RELATIONSHIPS
        ParseXDNationRelationships strLine
    Case XD_NUKE
        ParseXDNuke strLine
    Case XD_NUKE_CHR
        ParseXDNukeChr strLine
    Case XD_PLANE
        ParseXDPlane strLine
    Case XD_PLANE_CHR
        ParseXDPlaneChr strLine
    Case XD_SECTOR
        ParseXDSector strLine
    Case XD_SECTOR_CHR
        ParseXDSectorChr strLine
    Case XD_SHIP
        ParseXDShip strLine
    Case XD_SHIP_CHR
        ParseXDShipChr strLine
    Case XD_PRODUCT_CHR
        ParseXDProductChr strLine
    Case XD_UPDATES
        ParseXDUpdates strLine
    Case XD_VERSION
        ParseXDVersion strLine
    Case XD_LEVEL, XD_META_TYPE, XD_MISSIONS, XD_SECTOR_NAVIGATION, XD_SHIP_CHR_FLAGS
        ParseXDSymbolTable strLine
    Case XD_UNKNOWN
'        EmpireError "ParseXdump", "Invalid xdump unknown type", strLine
    Case Else
        EmpireError "ParseXdump", "No Parser type:" + CStr(xdType), strLine
    End Select
End If
End Sub

Private Sub ParseXDSector(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iOwnerIndex As Integer
Dim iXlocIndex As Integer
Dim iYlocIndex As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable

'Save and restore the index on entry and exit
hldIndex = rsSectors.Index
rsSectors.Index = "PrimaryKey"

On Error GoTo ParseXDSector_Exit

Set etTable = etsTable(XD_SECTOR)
If etTable Is Nothing Then
    EmpireError "ParseXDSector", _
        "The meta data was not found for sector", strLine
    rsSectors.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDSector", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsSectors.Index = hldIndex
    Exit Sub
End If

'Get the Sector information to get a record.
iOwnerIndex = etTable.GetZeroIndex("owner")
iXlocIndex = etTable.GetZeroIndex("xloc")
iYlocIndex = etTable.GetZeroIndex("yloc")

If iXlocIndex = -1 Or iYlocIndex = -1 Then
    EmpireError "ParseXDSector", _
        "The record identifier fields not present", _
         CStr(iXlocIndex) + ":" + CStr(iYlocIndex)
    rsSectors.Index = hldIndex
    Exit Sub
End If

rsSectors.Seek "=", strParams(iXlocIndex), strParams(iYlocIndex)
If rsSectors.NoMatch Then
    rsSectors.AddNew
Else
    rsSectors.Edit
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "owner"
        rsSectors.Fields("Country") = strParams(iIndex)
    Case "xloc"
        rsSectors.Fields("x") = strParams(iIndex)
    Case "yloc"
        rsSectors.Fields("y") = strParams(iIndex)
    Case "des"
        rsSectors.Fields("des") = GetSectorDes(Val(strParams(iIndex)))
    Case "newdes"
        rsSectors.Fields("sdes") = GetSectorDes(Val(strParams(iIndex)))
        If etTable.GetZeroIndex("des") <> -1 Then
            If strParams(etTable.GetZeroIndex("des")) = strParams(iIndex) Then
                rsSectors.Fields("sdes") = " "
            End If
        End If
    Case "effic"
        rsSectors.Fields("eff") = strParams(iIndex)
    Case "mobil"
        rsSectors.Fields("mob") = strParams(iIndex)
    Case "oldown"
        If strParams(iIndex) = strParams(iOwnerIndex) Then
            rsSectors.Fields("*") = " "
        Else
            rsSectors.Fields("*") = "*"
        End If
    Case "off", "min", "gold", "fert", "ocontent", "uran", "work", "avail", "terr", _
        "terr1", "terr2", "terr3", "uw", "food", "shell", "gun", "iron", "dust", "bar", _
        "oil", "lcm", "hcm", "rad", "road", "rail", "fallout", _
        "c_dist", "m_dist", "u_dist", "f_dist", "s_dist", _
        "g_dist", "p_dist", "i_dist", "d_dist", "b_dist", "o_dist", "l_dist", "h_dist", _
        "r_dist", "timestamp", "elev"
        rsSectors.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "terr0"
        rsSectors.Fields("terr") = strParams(iIndex)
    Case "petrol"
        rsSectors.Fields("pet") = strParams(iIndex)
    Case "civil"
        rsSectors.Fields("civ") = strParams(iIndex)
    Case "milit"
        rsSectors.Fields("mil") = strParams(iIndex)
    Case "dfense"
        rsSectors.Fields("defense") = strParams(iIndex)
    Case "coastal"
        rsSectors.Fields("coast") = strParams(iIndex)
    Case "xdist"
        rsSectors.Fields("dist_x") = strParams(iIndex)
    Case "ydist"
        rsSectors.Fields("dist_y") = strParams(iIndex)
    Case "mines"
        rsSectors.Fields("lmines") = strParams(iIndex)
    Case "c_del", "m_del", "u_del", "f_del", "s_del", "g_del", "p_del", _
        "i_del", "d_del", "b_del", "o_del", "l_del", "h_del", "r_del"
        rsSectors.Fields(etTable.Column(iHeaderIndex).Name) = DirectionString(Val(strParams(iIndex)) And 7)
        rsSectors.Fields(Left$(etTable.Column(iHeaderIndex).Name, 2) + "cut") = Val(strParams(iIndex)) And &HFFF8
    Case "pstage", "ptime", "che", "che_target", "access", "loyal"
    'South Pacific: The Search For Atlantis
    Case "frt"
        If frmOptions.bSPAtlantis Then
            rsSectors("fert") = strParams(iIndex)
        Else
            EmpireError "ParseXDSector", "Field not understood:" + strErrorField, strLine
        End If
    Case "runway"
        If frmOptions.bSPAtlantis Then
            rsSectors.Fields("min") = strParams(iIndex)
        Else
            EmpireError "ParseXDSector", "Field not understood:" + strErrorField, strLine
        End If
    Case "rdar"
        If frmOptions.bSPAtlantis Then
            rsSectors.Fields("gold") = strParams(iIndex)
        Else
            EmpireError "ParseXDSector", "Field not understood:" + strErrorField, strLine
        End If
    Case "nvigate"
        If frmOptions.bSPAtlantis Then
            rsSectors.Fields("ocontent") = strParams(iIndex)
        Else
            EmpireError "ParseXDSector", "Field not understood:" + strErrorField, strLine
        End If
    Case "dterr"
        If bDeity And VersionCheck(4, 3, 6) <> VER_LESS Then
            'In the future stored the deity territory
        Else
            EmpireError "ParseXDSector", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDSector", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
If iOwnerIndex = -1 Then
    rsSectors.Fields("country") = CountryNumber
End If

rsSectors.Update
rsSectors.Index = hldIndex

If Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY Then
    frmDrawMap.FillSectorBox Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex))
    If Not frmDrawMap.IsAnyTabActive Then
        frmDrawMap.DisplayFirstSectorPanel
    End If
End If

If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
    hldIndex = rsEnemySect.Index
    rsEnemySect.Index = "PrimaryKey"
    rsEnemySect.Seek "=", Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex))
    If Not rsEnemySect.NoMatch Then
        rsEnemySect.Delete
    End If
    rsEnemySect.Index = hldIndex
End If

Exit Sub

'Error handling routine
ParseXDSector_Exit:
EmpireError "ParseXDSector", strErrorField, strLine
rsSectors.Index = hldIndex
End Sub

Private Function GetSectorDes(iDchrIndex As Integer) As String
If Not (rsSectorType.BOF And rsSectorType.EOF) Then
    rsSectorType.MoveFirst
    While Not rsSectorType.EOF
        If rsSectorType.Fields("dchr_i") = iDchrIndex Then
            GetSectorDes = rsSectorType.Fields("des")
            Exit Function
        End If
        rsSectorType.MoveNext
    Wend
End If
GetSectorDes = "?"
End Function

Private Sub ParseXDSectorChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iDesIndex As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim hldBuildIndex As String
Dim etTable As EmpTable
Dim iCost As Integer
Dim iBuild As Integer
Dim iLcms As Integer
Dim iHcms As Integer

'Save and restore the index on entry and exit
hldIndex = rsSectorType.Index
rsSectorType.Index = "PrimaryKey"

On Error GoTo ParseXDSectorChr_Exit

Set etTable = etsTable(XD_SECTOR_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDSectorChr", _
        "The meta data was not found for sect-chr", strLine
    rsSectorType.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDSectorChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsSectorType.Index = hldIndex
    Exit Sub
End If

'Get the Sector information to get a record.
iDesIndex = etTable.GetZeroIndex("mnem")

If iDesIndex = -1 Then
    EmpireError "ParseXDSectorChr", _
        "The record identifier fields not present", _
         CStr(iDesIndex)
    rsSectorType.Index = hldIndex
    Exit Sub
End If
strParams(iDesIndex) = RemoveEscapeSequencesAndQuotes(strParams(iDesIndex))

rsSectorType.Seek "=", strParams(iDesIndex)
If rsSectorType.NoMatch Then
    rsSectorType.AddNew
Else
    rsSectorType.Edit
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "name"
        rsSectorType.Fields("desc") = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
    Case "mnem"
        rsSectorType.Fields("des") = strParams(iIndex)
    Case "mcst"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            EmpireError "ParseXDSectorChr", "Field not understood:" + strErrorField, strLine
        Else
            rsSectorType.Fields("mcost") = strParams(iIndex)
        End If
    Case "ostr"
        rsSectorType.Fields("offense") = Val(strParams(iIndex))
    Case "dstr"
        rsSectorType.Fields("defense") = Val(strParams(iIndex))
    Case "flg", "value", "uid", "maint"
    Case "pkg"
        rsSectorType.Fields("pack_civ") = Items.FindByLetter("c").Packing(Val(strParams(iIndex)))
        rsSectorType.Fields("pack_mil") = Items.FindByLetter("m").Packing(Val(strParams(iIndex)))
        rsSectorType.Fields("pack_uw") = Items.FindByLetter("u").Packing(Val(strParams(iIndex)))
        rsSectorType.Fields("pack_bar") = Items.FindByLetter("b").Packing(Val(strParams(iIndex)))
        rsSectorType.Fields("pack_other") = Items.FindByLetter("l").Packing(Val(strParams(iIndex)))
        If strParams(iDesIndex) = "c" Then
            rsVersion.MoveFirst
            rsVersion.Edit
            If Val(strParams(iIndex)) = PACKING_URBAN Then
                rsVersion.Fields("BIG_CITY") = True
            Else
                rsVersion.Fields("BIG_CITY") = False
            End If
            rsVersion.Update
        End If
        rsSectorType.Fields("pack_type") = Val(strParams(iIndex))
    Case "cost"
        iCost = strParams(iIndex)
    Case "lcms"
        iLcms = strParams(iIndex)
    Case "hcms"
        iHcms = strParams(iIndex)
    Case "build"
        iBuild = strParams(iIndex)
'        If strParams(iIndex) >= 0 Then
'        End If
'               dchr[x].d_mnem, dchr[x].d_cost, dchr[x].d_build,
'           dchr[x].d_lcms, dchr[x].d_hcms);
    Case "peffic"
        If VersionCheck(4, 3, 6) = VER_LESS Then
            EmpireError "ParseXDSectorChr", "Field not understood:" + strErrorField, strLine
        Else
            rsSectorType.Fields("eff") = strParams(iIndex)
        End If
    Case "prd"
        rsSectorType.Fields("pchr_i") = strParams(iIndex)
    Case "maxpop", "nav"
        rsSectorType.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "terrain"
        rsSectorType.Fields(etTable.Column(iHeaderIndex).Name) = Val(strParams(iIndex))
    Case "mob0"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            ' do nothing for the moment unable I figure this field does
            rsSectorType.Fields("mcost") = Val(strParams(iIndex))
        Else
            EmpireError "ParseXDSectorChr", "Field not understood:" + strErrorField, strLine
        End If
    Case "mob1"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            ' do nothing for the moment unable I figure this field does
            rsSectorType.Fields("mcost100") = Val(strParams(iIndex))
        Else
            EmpireError "ParseXDSectorChr", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDSectorChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

If iCost >= 0 Then
    If iCost > 0 Or iLcms > 0 Or iHcms > 0 Or iBuild <> 1 Then
        strErrorField = "Build Infrastructure"
        hldBuildIndex = rsBuild.Index
        rsBuild.Index = "PrimaryKey"
    
        If rsBuild.BOF And rsBuild.EOF Then
            rsBuild.AddNew
            rsBuild.Fields("type") = "i"
        Else
            rsBuild.Seek "=", "i", strParams(iDesIndex)
            If rsBuild.NoMatch Then
                rsBuild.AddNew
                rsBuild.Fields("type") = "i"
            Else
                rsBuild.Edit
            End If
        End If
        rsBuild.Fields("id") = strParams(iDesIndex)
        rsBuild.Fields("cost") = iCost
        rsBuild.Fields("lcm") = iLcms
        rsBuild.Fields("hcm") = iHcms
        rsBuild.Fields("stat1") = iBuild
        rsBuild.Update
        rsBuild.Index = hldBuildIndex
    End If
End If

rsSectorType.Fields("dchr_i") = iRowCount
rsSectorType.Fields("pcode") = ""
rsSectorType.Update
rsSectorType.Index = hldIndex
iRowCount = iRowCount + 1
Exit Sub

'Error handling routine
ParseXDSectorChr_Exit:
EmpireError "ParseXDSectorChr", strErrorField, strLine
rsSectorType.Index = hldIndex
End Sub

Private Sub ParseXDPlane(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iOwnerIndex As Integer
Dim iIDIndex As Integer
Dim iXlocIndex As Integer
Dim iYlocIndex As Integer
Dim iOpxIndex As Integer
Dim iOpyIndex As Integer
Dim iRadiusIndex As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim hldChrIndex As String
Dim etTable As EmpTable
Dim iPrevXloc As Integer
Dim iPrevYloc As Integer

'Save and restore the index on entry and exit
hldIndex = rsPlane.Index
rsPlane.Index = "PrimaryKey"

On Error GoTo ParseXDPlane_Exit

Set etTable = etsTable(XD_PLANE)
If etTable Is Nothing Then
    EmpireError "ParseXDPlane", _
        "The meta data was not found for plane", strLine
    rsPlane.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDPlane", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
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

If iIDIndex = -1 Or iOwnerIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or _
    iRadiusIndex = -1 Or iOpxIndex = -1 Or iOpxIndex = -1 Then
    EmpireError "ParseXDPlane", _
        "One of the following fields is missing: owner, id, xloc, yloc, opx, opy, rad", _
         ""
    rsPlane.Index = hldIndex
    Exit Sub
End If

If rsPlane.BOF And rsPlane.EOF Then
    rsPlane.AddNew
Else
    rsPlane.Seek "=", strParams(iIDIndex)
    If rsPlane.NoMatch Then
        rsPlane.AddNew
    Else
        rsPlane.Edit
    End If
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "owner"
        rsPlane.Fields("Country") = strParams(iIndex)
    Case "uid"
        rsPlane.Fields("id") = strParams(iIndex)
    Case "type"
        rsPlane.Fields("type") = GetUnitType(Val(strParams(iIndex)), "p")
    Case "wing"
        If estsTable(XD_META_TYPE)(etTable.Column(iHeaderIndex).ColType).Name = "c" Then
            strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
            If strParams(iIndex) = "" Or strParams(iIndex) = " " Then 'replace space with ~
                rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = "~"
            Else
                rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = Left$(strParams(iIndex), 1)
            End If
        Else
            If strParams(iIndex) = 32 Then 'replace space with ~
                rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = "~"
            Else
                rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = Chr$(strParams(iIndex))
            End If
        End If
    Case "xloc"
        If IsNull(rsPlane.Fields("x")) Then
            iPrevXloc = -1
        Else
            iPrevXloc = rsPlane.Fields("x")
        End If
        rsPlane.Fields("x") = strParams(iIndex)
    Case "yloc"
        If IsNull(rsPlane.Fields("y")) Then
            iPrevYloc = -1
        Else
            iPrevYloc = rsPlane.Fields("y")
        End If
        rsPlane.Fields("y") = strParams(iIndex)
    Case "effic"
        rsPlane.Fields("eff") = strParams(iIndex)
    Case "mobil"
        rsPlane.Fields("mob") = strParams(iIndex)
    Case "harden"
        rsPlane.Fields("hard") = strParams(iIndex)
    Case "range"
        rsPlane.Fields("react") = strParams(iIndex)
    Case "tech"
        rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        hldChrIndex = rsBuild.Index
        rsBuild.Index = "PrimaryKey"

        rsBuild.Seek "=", "p", rsPlane.Fields("type")
        If Not rsBuild.NoMatch Then
            rsPlane.Fields("att") = Fix(ComputePlaneAttDef(rsBuild.Fields("base att"), _
                                        rsBuild.Fields("tech"), rsPlane.Fields("tech")))
            rsPlane.Fields("def") = Fix(ComputePlaneAttDef(rsBuild.Fields("base def"), _
                                        rsBuild.Fields("tech"), rsPlane.Fields("tech")))
            rsPlane.Fields("acc") = Fix(ComputePlaneAccuracy(rsBuild.Fields("base acc"), _
                                        rsBuild.Fields("tech"), rsPlane.Fields("tech")))
            rsPlane.Fields("load") = Fix(ComputePlaneLoad(rsBuild.Fields("base load"), _
                                        rsBuild.Fields("tech"), rsPlane.Fields("tech")))
'            'need to fix for limited range planes
            rsPlane.Fields("range") = Fix(ComputePlaneRange(rsBuild.Fields("base spd"), _
                                        rsBuild.Fields("tech"), rsPlane.Fields("tech")))
            rsPlane.Fields("fuel") = rsBuild.Fields("stat6")
'        Else
'            EmpireError "ParseXDPlane", "Plane Characteristics not found:" + rsPlane.Fields("type"), strLine
        End If
        rsBuild.Index = hldChrIndex
    Case "ship", "land", "timestamp"
        rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "nuketype"
        rsPlane.Fields("nuke") = GetUnitType(Val(strParams(iIndex)), "n")
    Case "opx", "opy", "radius", "access", "theta"
    Case "mission"
        AddMission Val(strParams(iIndex)), Val(strParams(iIDIndex)), _
            Val(strParams(iOwnerIndex)), "p", _
            SectorString(Val(strParams(iOpxIndex)), Val(strParams(iOpyIndex))), _
            Val(strParams(iRadiusIndex))
    Case "flags"
        If (strParams(iIndex) And &H1) <> 0 Then
            rsPlane.Fields("laun") = "Y"
        Else
            rsPlane.Fields("laun") = "N"
        End If
        If (strParams(iIndex) And &H2) <> 0 Then
            rsPlane.Fields("orb") = "Y"
        Else
            rsPlane.Fields("orb") = "N"
        End If
        If (strParams(iIndex) And &H4) <> 0 Then
            rsPlane.Fields("grd") = "A"
        Else
            rsPlane.Fields("grd") = "G"
        End If
    Case "off"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            rsPlane.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        Else
            EmpireError "ParseXDPlane", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDPlane", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
If iOwnerIndex = -1 Then
    rsPlane.Fields("country") = CountryNumber
End If

rsPlane.Update
rsPlane.Index = hldIndex

If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or _
    (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
    frmDrawMap.FillGrid
End If

CheckForOldEnemyUnit CInt(strParams(iIDIndex)), "p"
Exit Sub

'Error handling routine
ParseXDPlane_Exit:
EmpireError "ParseXDPlane", strErrorField, strLine
rsPlane.Index = hldIndex
End Sub

Private Sub ParseXDPlaneChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iCapIndex As Integer
Dim iIDIndex As Integer
Dim iTechIndex As Integer
Dim iPos As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable
Dim lFlags As Long

'Save and restore the index on entry and exit
hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

On Error GoTo ParseXDPlaneChr_Exit

Set etTable = etsTable(XD_PLANE_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDPlaneChr", _
        "The meta data was not found for plane-chr", strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDPlaneChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

'Get the Index information to get a record.
iIDIndex = etTable.GetZeroIndex("name")
strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
iTechIndex = etTable.GetZeroIndex("tech")

If iIDIndex = -1 Then
    EmpireError "ParseXDPlaneChr", _
        "The record identifier fields not present", _
         CStr(iIDIndex)
    rsBuild.Index = hldIndex
    Exit Sub
End If

iPos = InStr(strParams(iIDIndex), " ")

If rsBuild.BOF And rsBuild.EOF Then
    rsBuild.AddNew
    rsBuild.Fields("type") = "p"
Else
    rsBuild.Seek "=", "p", Left$(strParams(iIDIndex), iPos - 1)
    If rsBuild.NoMatch Then
        rsBuild.AddNew
        rsBuild.Fields("type") = "p"
    Else
        rsBuild.Edit
    End If
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "name"
        If iPos > 0 Then
            rsBuild.Fields("id") = Left$(strParams(iIndex), iPos - 1)
            rsBuild.Fields("desc") = LTrim$(Mid$(strParams(iIndex), iPos + 1, Len(strParams(iIndex)) - iPos))
        Else
            EmpireError "ParseXDPlaneChr", "Incorrect format of name", strLine
        End If
    Case "cost", "tech"
        rsBuild.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "l_build":
        rsBuild.Fields("lcm") = strParams(iIndex)
    Case "h_build":
        rsBuild.Fields("hcm") = strParams(iIndex)
    Case "crew"
        rsBuild.Fields("mil") = strParams(iIndex)
    Case "flags"
        iCapIndex = 1
        lFlags = CLng(strParams(iIndex))
        If (lFlags And &H1) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "bomber"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H2) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "tactical"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H4) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "intercept"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H8) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "cargo"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H10) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "VTOL"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H20) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "missile"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H40) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "light"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H80) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "spy"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H100) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "image"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H200) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "satellite"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H400) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "stealth"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H800) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "SDI"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H1000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "half-stealth"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H2000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "x-light"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H4000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "helo"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H8000&) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "ASW"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H10000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "para"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H20000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "escort"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H40000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "mine"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H80000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "sweep"
            iCapIndex = iCapIndex + 1
        End If
        If (lFlags And &H100000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "marine"
            iCapIndex = iCapIndex + 1
        End If
        rsBuild.Fields("cap" + CStr(iCapIndex)) = ""
    Case "acc"
        rsBuild.Fields("base acc") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
            rsBuild.Fields("stat1") = Fix(ComputePlaneAccuracy(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "load"
        rsBuild.Fields("base load") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
            rsBuild.Fields("stat2") = Fix(ComputePlaneLoad(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "att"
        rsBuild.Fields("base att") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
            rsBuild.Fields("stat3") = Fix(ComputePlaneAttDef(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "def"
        rsBuild.Fields("base def") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
            rsBuild.Fields("stat4") = Fix(ComputePlaneAttDef(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "range"
        rsBuild.Fields("base spd") = CInt(strParams(iIndex))
        If CheckIndex(iTechIndex, "ParseXDPlaneChr", "tech") Then
            rsBuild.Fields("stat5") = Fix(ComputePlaneRange(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "fuel"
        rsBuild.Fields("stat6") = strParams(iIndex)
    Case "stealth"
        rsBuild.Fields("stat7") = strParams(iIndex)
    Case "type"
        rsBuild.Fields("chr_i") = strParams(iIndex)
    Case Else
        EmpireError "ParseXDPlaneChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

rsBuild.Fields("avail") = CalculateAvail(CInt(rsBuild.Fields("lcm")), CInt(rsBuild.Fields("hcm")))

rsBuild.Update
rsBuild.Index = hldIndex
Exit Sub

'Error handling routine
ParseXDPlaneChr_Exit:
EmpireError "ParseXDPlaneChr", strErrorField, strLine
rsBuild.Index = hldIndex
End Sub

Private Sub ParseXDShip(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iArrayIndex As Integer
Dim iHeaderIndex As Integer
Dim iOwnerIndex As Integer
Dim iIDIndex As Integer
Dim iXlocIndex As Integer
Dim iYlocIndex As Integer
Dim iOpxIndex As Integer
Dim iOpyIndex As Integer
Dim iRadiusIndex As Integer
Dim strErrorField As String
Dim hldShipIndex As String
Dim hldChrIndex As String
Dim iXstartIndex As Integer
Dim iYstartIndex As Integer
Dim iXendIndex As Integer
Dim iYendIndex As Integer
Dim iTypeIndex As Integer
Dim strStartCargo(6) As String
Dim strEndCargo(6) As String
Dim iStart(6) As Integer
Dim iEnd(6) As Integer
Dim etTable As EmpTable
Dim bAutoNav As Boolean
Dim iPrevXloc As Integer
Dim iPrevYloc As Integer

'Save and restore the index on entry and exit
hldShipIndex = rsShip.Index
rsShip.Index = "PrimaryKey"

On Error GoTo ParseXDShip_Exit

Set etTable = etsTable(XD_SHIP)
If etTable Is Nothing Then
    EmpireError "ParseXDShip", _
        "The meta data was not found for ship", strLine
    rsShip.Index = hldShipIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDShip", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
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

If iOwnerIndex = -1 Or iIDIndex = -1 Or iOpxIndex = -1 Or iOpyIndex = -1 Or _
    iRadiusIndex = -1 Or iXlocIndex = -11 Or iYlocIndex = -1 Or iTypeIndex = -1 Or _
    iXstartIndex = -1 Or iYstartIndex = -1 Or iXendIndex = -1 Or iYendIndex = -1 _
    Then
    EmpireError "ParseXDShip", _
        "One of the following fields is missing: owner, id, xloc, yloc, opx, opy, rad, type" + _
        ", startx, starty, endx or endy", _
         CStr(iIDIndex)
    rsShip.Index = hldShipIndex
    Exit Sub
End If

If rsShip.BOF And rsShip.EOF Then
    rsShip.AddNew
Else
    rsShip.Seek "=", strParams(iIDIndex)
    If rsShip.NoMatch Then
        rsShip.AddNew
    Else
        rsShip.Edit
    End If
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "owner"
        rsShip.Fields("Country") = strParams(iIndex)
    Case "uid"
        rsShip.Fields("id") = strParams(iIndex)
    Case "type"
        rsShip.Fields("type") = GetUnitType(Val(strParams(iIndex)), "s")
    Case "fleet"
        If estsTable(XD_META_TYPE)(etTable.Column(iHeaderIndex).ColType).Name = "c" Then
            strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
            If strParams(iIndex) = "" Or strParams(iIndex) = " " Then 'replace space with ~
                rsShip.Fields("flt") = "~"
            Else
                rsShip.Fields("flt") = Left$(strParams(iIndex), 1)
            End If
        Else
            If strParams(iIndex) = 32 Then 'replace space with ~
                rsShip.Fields("flt") = "~"
            Else
                rsShip.Fields("flt") = Chr$(strParams(iIndex))
            End If
        End If
    Case "xloc"
        If IsNull(rsShip.Fields("x")) Then
            iPrevXloc = -1
        Else
            iPrevXloc = rsShip.Fields("x")
        End If
        rsShip.Fields("x") = strParams(iIndex)
    Case "yloc"
        If IsNull(rsShip.Fields("y")) Then
            iPrevYloc = 0
        Else
            iPrevYloc = rsShip.Fields("y")
        End If
        rsShip.Fields("y") = strParams(iIndex)
    Case "effic"
        rsShip.Fields("eff") = strParams(iIndex)
    Case "mobil"
        rsShip.Fields("mob") = strParams(iIndex)
    Case "nplane"
        rsShip.Fields("pln") = strParams(iIndex)
    Case "nland"
        rsShip.Fields("land") = strParams(iIndex)
    Case "nchoppers"
        rsShip.Fields("he") = strParams(iIndex)
    Case "nxlight"
        rsShip.Fields("xl") = strParams(iIndex)
    Case "civil"
        rsShip.Fields("civ") = strParams(iIndex)
    Case "milit"
        rsShip.Fields("mil") = strParams(iIndex)
    Case "xbuilt"
        rsShip.Fields("origx") = strParams(iIndex)
    Case "ybuilt"
        rsShip.Fields("origy") = strParams(iIndex)
    Case "name"
        rsShip.Fields(etTable.Column(iHeaderIndex).Name) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
    Case "tech"
        rsShip.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        hldChrIndex = rsBuild.Index
        rsBuild.Index = "PrimaryKey"

        rsBuild.Seek "=", "s", GetUnitType(Val(strParams(iTypeIndex)), "s")
        If Not rsBuild.NoMatch Then
            rsShip.Fields("def") = Fix(ComputeShipDefense(rsBuild.Fields("base def"), _
                                        rsBuild.Fields("tech"), rsShip.Fields("tech")))
            rsShip.Fields("spd") = Fix(ComputeShipSpeed(rsBuild.Fields("base spd"), _
                                        rsBuild.Fields("tech"), rsShip.Fields("tech")))
            rsShip.Fields("vis") = Fix(ComputeShipVisibility(rsBuild.Fields("base vul"), _
                                        rsBuild.Fields("tech"), rsShip.Fields("tech")))
            rsShip.Fields("rng") = Fix(ComputeShipRange(rsBuild.Fields("base frnge"), _
                                        rsBuild.Fields("tech"), rsShip.Fields("tech")))
            rsShip.Fields("fir") = Fix(ComputeShipFiringRange(rsBuild.Fields("base load"), _
                                        rsBuild.Fields("tech"), rsShip.Fields("tech")))
'        Else
'            EmpireError "ParseXDShip", "Ship Characteristics not found:" + rsShip.Fields("type"), strLine
        End If
        rsBuild.Index = hldChrIndex
    Case "timestamp", "uw", "food", "fuel", "shell", "gun", "petrol", "iron", "dust", "bar", _
         "oil", "lcm", "hcm", "rad"
        rsShip.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "opx", "opy", "radius", "flags", "pstage", "ptime", "access", _
         "builder", _
        "mquota", "path", "follow", "rflags", "rpath"
    Case "autonav"
        If Val(strParams(iIndex)) > 0 Then
            bAutoNav = True
        Else
            bAutoNav = False
        End If
    Case "mission"
        AddMission Val(strParams(iIndex)), Val(strParams(iIDIndex)), _
            Val(strParams(iOwnerIndex)), "s", _
            SectorString(Val(strParams(iOpxIndex)), Val(strParams(iOpyIndex))), _
            Val(strParams(iRadiusIndex))
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
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            rsShip.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        Else
            EmpireError "ParseXDShip", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDShip", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
If iOwnerIndex = -1 Then
    rsShip.Fields("country") = CountryNumber
End If

rsShip.Update
rsShip.Index = hldShipIndex
       
AddOrders bAutoNav, Val(strParams(iIDIndex)), Val(strParams(iOwnerIndex)), _
    0, 0, _
    SectorString(Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex))), _
    SectorString(Val(strParams(iXstartIndex)), Val(strParams(iYstartIndex))), _
    SectorString(Val(strParams(iXendIndex)), Val(strParams(iYendIndex))), _
    strStartCargo(), strEndCargo(), iStart(), iEnd()

If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or _
    (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
    frmDrawMap.FillGrid
End If

CheckForOldEnemyUnit CInt(strParams(iIDIndex)), "s"
Exit Sub

'Error handling routine
ParseXDShip_Exit:
EmpireError "ParseXDShip", strErrorField, strLine
rsShip.Index = hldShipIndex
End Sub

Private Sub ParseXDShipChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iCapIndex As Integer
Dim iIDIndex As Integer
Dim iTechIndex As Integer
Dim iTypeIndex As Integer
Dim iPos As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable

'Save and restore the index on entry and exit
hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

On Error GoTo ParseXDShipChr_Exit

Set etTable = etsTable(XD_SHIP_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDShipChr", _
        "The meta data was not found for ship-chr", strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDShipChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

'Get the Index information to get a record.
iIDIndex = etTable.GetZeroIndex("name")
strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
iTechIndex = etTable.GetZeroIndex("tech")

If iIDIndex = -1 Then
    EmpireError "ParseXDShipChr", _
        "The record identifier field not present", _
         CStr(iIDIndex)
    rsBuild.Index = hldIndex
    Exit Sub
End If
If iTechIndex = -1 Then
    EmpireError "ParseXDShipChr", _
        "The tech field not present", _
         CStr(iTechIndex)
    rsBuild.Index = hldIndex
    Exit Sub
End If

iPos = InStr(strParams(iIDIndex), " ")
If rsBuild.BOF And rsBuild.EOF Then
    rsBuild.AddNew
    rsBuild.Fields("type") = "s"
Else
    rsBuild.Seek "=", "s", Left$(strParams(iIDIndex), iPos - 1)
    If rsBuild.NoMatch Then
        rsBuild.AddNew
        rsBuild.Fields("type") = "s"
    Else
        rsBuild.Edit
    End If
End If
iCapIndex = 1
iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "name"
        If iPos > 0 Then
            rsBuild.Fields("id") = Left$(strParams(iIndex), iPos - 1)
            rsBuild.Fields("desc") = LTrim$(Mid$(strParams(iIndex), iPos + 1, Len(strParams(iIndex)) - iPos))
        Else
            EmpireError "ParseXDShipChr", "Incorrect format of name", strLine
        End If
    Case "cost", "tech"
        rsBuild.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "l_build":
        rsBuild.Fields("lcm") = strParams(iIndex)
    Case "h_build":
        rsBuild.Fields("hcm") = strParams(iIndex)
    Case "civil"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "c"
            iCapIndex = iCapIndex + 1
        End If
    Case "milit"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "m"
            iCapIndex = iCapIndex + 1
        End If
    Case "shell"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "s"
            iCapIndex = iCapIndex + 1
        End If
    Case "gun"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "g"
            iCapIndex = iCapIndex + 1
        End If
    Case "petrol"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "p"
            iCapIndex = iCapIndex + 1
        End If
    Case "iron"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "i"
            iCapIndex = iCapIndex + 1
        End If
    Case "dust"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "d"
            iCapIndex = iCapIndex + 1
        End If
    Case "bar"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "b"
            iCapIndex = iCapIndex + 1
        End If
    Case "food"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "f"
            iCapIndex = iCapIndex + 1
        End If
    Case "oil"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "o"
            iCapIndex = iCapIndex + 1
        End If
    Case "lcm"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "l"
            iCapIndex = iCapIndex + 1
        End If
    Case "hcm"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "h"
            iCapIndex = iCapIndex + 1
        End If
    Case "uw"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "u"
            iCapIndex = iCapIndex + 1
        End If
    Case "rad"
        If Len(strParams(iIndex)) > 0 And strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "r"
            iCapIndex = iCapIndex + 1
        End If
    Case "flags"
        If (strParams(iIndex) And &H1) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "fish"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H2) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "torp"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H4) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "dchrg"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H8) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "plane"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H10) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "miss"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H20) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "oil"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H40) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "sonar"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H80) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "mine"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H100) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "sweep"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H200) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "sub"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H400) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "spy"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H800) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "land"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H1000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "sub-torp"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H2000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "trade"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H4000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "semi-land"
            iCapIndex = iCapIndex + 1
        End If
'        If (strParams(iIndex) And &H8000&) <> 0 Then
'            '#define M_XLIGHT    bit(15) /* can hold xlight planes */
'        End If
'        If (strParams(iIndex) And &H10000) <> 0 Then
'            '#define M_CHOPPER   bit(16) /* can hold choppers */
'        End If
        If (strParams(iIndex) And &H20000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "oiler"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H40000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "supply"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H80000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "canal"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H100000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "anti-missile"
            iCapIndex = iCapIndex + 1
        End If
    Case "armor"
        rsBuild.Fields("base def") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
            rsBuild.Fields("stat1") = Fix(ComputeShipDefense(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "speed"
        rsBuild.Fields("base spd") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
            rsBuild.Fields("stat2") = Fix(ComputeShipSpeed(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "visib"
        rsBuild.Fields("base vul") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
            rsBuild.Fields("stat3") = Fix(ComputeShipVisibility(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "vrnge"
            rsBuild.Fields("stat4") = strParams(iIndex)
    Case "frnge"
        rsBuild.Fields("base frnge") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
            rsBuild.Fields("stat5") = Fix(ComputeShipRange(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "glim"
        rsBuild.Fields("base load") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDShipChr", "tech") Then
            rsBuild.Fields("stat6") = Fix(ComputeShipFiringRange(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "nland"
        rsBuild.Fields("stat7") = strParams(iIndex)
    Case "nplanes"
        rsBuild.Fields("stat8") = strParams(iIndex)
    Case "nchoppers"
        rsBuild.Fields("stat9") = strParams(iIndex)
    Case "nxlight"
        rsBuild.Fields("stat10") = strParams(iIndex)
    Case "fuelc"
        rsBuild.Fields("stat11") = strParams(iIndex)
    Case "fuelu"
        rsBuild.Fields("stat12") = strParams(iIndex)
    Case "type"
        rsBuild.Fields("chr_i") = strParams(iIndex)
    Case Else
        EmpireError "ParseXDShipChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

rsBuild.Fields("avail") = CalculateAvail(CInt(rsBuild.Fields("lcm")), CInt(rsBuild.Fields("hcm")))

rsBuild.Update
rsBuild.Index = hldIndex
Exit Sub

'Error handling routine
ParseXDShipChr_Exit:
EmpireError "ParseXDShipChr", strErrorField, strLine
rsBuild.Index = hldIndex
End Sub

Private Sub ParseXDLand(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iOwnerIndex As Integer
Dim iIDIndex As Integer
Dim iXlocIndex As Integer
Dim iYlocIndex As Integer
Dim iOpxIndex As Integer
Dim iOpyIndex As Integer
Dim iRadiusIndex As Integer
Dim iTypeIndex As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim hldChrIndex As String
Dim etTable As EmpTable
Dim iPrevXloc As Integer
Dim iPrevYloc As Integer

'Save and restore the index on entry and exit
hldIndex = rsLand.Index
rsLand.Index = "PrimaryKey"

On Error GoTo ParseXDLand_Exit

Set etTable = etsTable(XD_LAND)
If etTable Is Nothing Then
    EmpireError "ParseXDLand", _
        "The meta data was not found for land", strLine
    rsLand.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDLand", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
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

If iIDIndex = -1 Or iOwnerIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or _
    iOpxIndex = -1 Or iOpyIndex = -1 Or iRadiusIndex = -1 Or iTypeIndex = -1 Then
    EmpireError "ParseXDLand", _
        "One of the following fields is missing: owner, id, xloc, yloc, opx, opy, rad, tech", _
        ""
    rsLand.Index = hldIndex
    Exit Sub
End If

If rsLand.BOF And rsLand.EOF Then
    rsLand.AddNew
Else
    rsLand.Seek "=", strParams(iIDIndex)
    If rsLand.NoMatch Then
        rsLand.AddNew
    Else
        rsLand.Edit
    End If
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "owner"
        rsLand.Fields("Country") = strParams(iIndex)
    Case "uid"
        rsLand.Fields("id") = strParams(iIndex)
    Case "type"
        rsLand.Fields("type") = GetUnitType(Val(strParams(iIndex)), "l")
    Case "army"
        If estsTable(XD_META_TYPE)(etTable.Column(iHeaderIndex).ColType).Name = "c" Then
            strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
            If strParams(iIndex) = "" Or strParams(iIndex) = " " Then 'replace space with ~
                rsLand.Fields(etTable.Column(iHeaderIndex).Name) = "~"
            Else
                rsLand.Fields(etTable.Column(iHeaderIndex).Name) = Left$(strParams(iIndex), 1)
            End If
        Else
            If strParams(iIndex) = 32 Then 'replace space with ~
                rsLand.Fields(etTable.Column(iHeaderIndex).Name) = "~"
            Else
                rsLand.Fields(etTable.Column(iHeaderIndex).Name) = Chr$(strParams(iIndex))
            End If
        End If
    Case "xloc"
        If IsNull(rsLand.Fields("x")) Then
            iPrevXloc = -1
        Else
            iPrevXloc = rsLand.Fields("x")
        End If
        rsLand.Fields("x") = strParams(iIndex)
    Case "yloc"
        If IsNull(rsLand.Fields("y")) Then
            iPrevYloc = -1
        Else
            iPrevYloc = rsLand.Fields("y")
        End If
        rsLand.Fields("y") = strParams(iIndex)
    Case "effic"
        rsLand.Fields("eff") = strParams(iIndex)
    Case "mobil"
        rsLand.Fields("mob") = strParams(iIndex)
    Case "ship"
        rsLand.Fields("ship") = strParams(iIndex)
    Case "harden"
        rsLand.Fields("fort") = strParams(iIndex)
    Case "nxlight"
        rsLand.Fields("xl") = strParams(iIndex)
    Case "civil"
        rsLand.Fields("civ") = strParams(iIndex)
    Case "milit"
        rsLand.Fields("mil") = strParams(iIndex)
    Case "nxlight"
        rsLand.Fields("xl") = strParams(iIndex)
    Case "retreat"
        rsLand.Fields("retr") = strParams(iIndex)
    Case "name"
        rsLand.Fields(etTable.Column(iHeaderIndex).Name) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
    Case "tech"
        rsLand.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        hldChrIndex = rsBuild.Index
        rsBuild.Index = "PrimaryKey"

        rsBuild.Seek "=", "l", GetUnitType(Val(strParams(iTypeIndex)), "l")
        If Not rsBuild.NoMatch Then
            rsLand.Fields("att") = Round(ComputeLandAttDef(rsBuild.Fields("base att"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")), 1)
            rsLand.Fields("def") = Round(ComputeLandAttDef(rsBuild.Fields("base def"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")), 1)
            rsLand.Fields("vul") = Fix(ComputeLandVulerability(rsBuild.Fields("base vul"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
            rsLand.Fields("spd") = Fix(ComputeLandSpeed(rsBuild.Fields("base spd"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
            rsLand.Fields("vis") = rsBuild.Fields("stat5")
            rsLand.Fields("spy") = rsBuild.Fields("stat6")
            rsLand.Fields("radius") = rsBuild.Fields("stat7")
            rsLand.Fields("frg") = Fix(ComputeLandFiringRange(rsBuild.Fields("base frnge"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
            rsLand.Fields("acc") = Fix(ComputeLandAccuracy(rsBuild.Fields("base acc"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
            rsLand.Fields("dam") = Fix(ComputeLandDamage(rsBuild.Fields("base dam"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
            rsLand.Fields("amm") = Fix(ComputeLandAmmo(rsBuild.Fields("base load"), _
                                        rsBuild.Fields("base dam"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
            rsLand.Fields("aaf") = Fix(ComputeLandAAF(rsBuild.Fields("base aaf"), _
                                        rsBuild.Fields("tech"), rsLand.Fields("tech")))
'    lp->lnd_fuelc = (int)LND_FC(lcp->l_fuelc, tech_diff);
'    lp->lnd_fuelu = (int)LND_FU(lcp->l_fuelu, tech_diff);
'    lp->lnd_maxlight = (int)LND_XPL(lcp->l_nxlight, tech_diff);
'    lp->lnd_maxland = (int)LND_MXL(lcp->l_mxland, tech_diff);
'        Else
'            EmpireError "ParseXDLand", "Land Characteristics not found:" + rsLand.Fields("type"), strLine
        End If
        rsBuild.Index = hldChrIndex
    Case "timestamp", "react", "uw", "food", "fuel", "shell", "gun", "petrol", "iron", "dust", "bar", _
         "oil", "lcm", "hcm", "rad", "nland", "land"
        rsLand.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "opx", "opy", "radius", "flags", "pstage", "ptime", "access", _
         "rpath", "rflags"
    Case "mission"
        AddMission Val(strParams(iIndex)), Val(strParams(iIDIndex)), _
            Val(strParams(iOwnerIndex)), "l", _
            SectorString(Val(strParams(iOpxIndex)), Val(strParams(iOpyIndex))), _
            Val(strParams(iRadiusIndex))
    Case "off"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            rsLand.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        Else
            EmpireError "ParseXDLand", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDLand", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
If iOwnerIndex = -1 Then
    rsLand.Fields("country") = CountryNumber
End If

rsLand.Update
rsLand.Index = hldIndex

If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or _
    (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
    frmDrawMap.FillGrid
End If

CheckForOldEnemyUnit CInt(strParams(iIDIndex)), "l"
Exit Sub

'Error handling routine
ParseXDLand_Exit:
EmpireError "ParseXDLand", strErrorField, strLine
rsLand.Index = hldIndex
End Sub

Private Sub ParseXDLandChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iCapIndex As Integer
Dim iIDIndex As Integer
Dim iTechIndex As Integer
Dim iDamageIndex As Integer
Dim iPos As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable

'Save and restore the index on entry and exit
hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

On Error GoTo ParseXDLandChr_Exit

Set etTable = etsTable(XD_LAND_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDLandChr", _
        "The meta data was not found for item", strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDLandChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

'Get the ID information to get a record.
iIDIndex = etTable.GetZeroIndex("name")
strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
iTechIndex = etTable.GetZeroIndex("tech")
iDamageIndex = etTable.GetZeroIndex("dam")

If iIDIndex = -1 Then
    EmpireError "ParseXDLandChr", _
        "The record identifier field not present", _
         CStr(iIDIndex)
    rsBuild.Index = hldIndex
    Exit Sub
End If

iPos = InStr(strParams(iIDIndex), " ")
'
If rsBuild.BOF And rsBuild.EOF Then
    rsBuild.AddNew
    rsBuild.Fields("type") = "l"
Else
    rsBuild.Seek "=", "l", Left$(strParams(iIDIndex), iPos - 1)
    If rsBuild.NoMatch Then
        rsBuild.AddNew
        rsBuild.Fields("type") = "l"
    Else
        rsBuild.Edit
    End If
End If

rsBuild.Fields("gun") = 0

iHeaderIndex = 1
iCapIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "name"
        If iPos > 0 Then
            rsBuild.Fields("id") = Left$(strParams(iIndex), iPos - 1)
            rsBuild.Fields("desc") = LTrim$(Right$(strParams(iIndex), Len(strParams(iIndex)) - iPos))
        Else
            EmpireError "ParseXDLandChr", "Incorrect format of name", strLine
        End If
    Case "cost", "tech"
        rsBuild.Fields(etTable.Column(iHeaderIndex).Name) = Val(strParams(iIndex))
    Case "l_build":
        rsBuild.Fields("lcm") = Val(strParams(iIndex))
    Case "h_build":
        rsBuild.Fields("hcm") = Val(strParams(iIndex))
    Case "g_build":
        rsBuild.Fields("gun") = Val(strParams(iIndex))
    Case "s_build":
        'shells not used yet
    Case "civil"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "c"
            iCapIndex = iCapIndex + 1
        End If
    Case "milit"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "m"
            iCapIndex = iCapIndex + 1
        End If
    Case "shell"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "s"
            iCapIndex = iCapIndex + 1
        End If
    Case "gun"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "g"
            iCapIndex = iCapIndex + 1
        End If
    Case "petrol"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "p"
            iCapIndex = iCapIndex + 1
        End If
    Case "iron"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "i"
            iCapIndex = iCapIndex + 1
        End If
    Case "dust"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "d"
            iCapIndex = iCapIndex + 1
        End If
    Case "bar"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "b"
            iCapIndex = iCapIndex + 1
        End If
    Case "food"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "f"
            iCapIndex = iCapIndex + 1
        End If
    Case "oil"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "o"
            iCapIndex = iCapIndex + 1
        End If
    Case "lcm"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "l"
            iCapIndex = iCapIndex + 1
        End If
    Case "hcm"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "h"
            iCapIndex = iCapIndex + 1
        End If
    Case "uw"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "u"
            iCapIndex = iCapIndex + 1
        End If
    Case "rad"
        If strParams(iIndex) > 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = strParams(iIndex) + "r"
            iCapIndex = iCapIndex + 1
        End If
    Case "flags"
        If (strParams(iIndex) And &H1) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "xlight"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H2) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "engineer"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H4) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "supply"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H8) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "security"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H10) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "light"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H20) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "marine"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H40) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "recon"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H80) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "radar"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H100) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "assault"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H200) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "flak"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H400) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "spy"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H800) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "train"
            iCapIndex = iCapIndex + 1
        End If
        If (strParams(iIndex) And &H1000) <> 0 Then
            rsBuild.Fields("cap" + CStr(iCapIndex)) = "heavy"
            iCapIndex = iCapIndex + 1
        End If
    Case "att"
        rsBuild.Fields("base att") = Val(strParams(iIndex))
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat1") = Round(ComputeLandAttDef(Val(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()), 1)
'            rsBuild.Fields("stat1") = Fix(ComputeLandAttDef(Val(strParams(iIndex)), CInt(strParams(iTechIndex)), CInt(rsNation.Fields("TechLevel"))) * 10) / 10
        End If
    Case "def"
        rsBuild.Fields("base def") = Val(strParams(iIndex))
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat2") = Round(ComputeLandAttDef(Val(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()), 1)
'            rsBuild.Fields("stat2") = Fix(ComputeLandAttDef(Val(strParams(iIndex)), CInt(strParams(iTechIndex)), CInt(rsNation.Fields("TechLevel"))) * 10) / 10
        End If
    Case "vul"
        rsBuild.Fields("base vul") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat3") = Fix(ComputeLandVulerability(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "spd"
        rsBuild.Fields("base spd") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat4") = Fix(ComputeLandSpeed(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "vis"
        rsBuild.Fields("stat5") = strParams(iIndex)
    Case "spy"
        rsBuild.Fields("stat6") = strParams(iIndex)
    Case "rmax"
        rsBuild.Fields("stat7") = strParams(iIndex)
    Case "frg"
        rsBuild.Fields("base frnge") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat8") = Fix(ComputeLandFiringRange(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "acc"
        rsBuild.Fields("base acc") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat9") = Fix(ComputeLandAccuracy(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "dam"
        rsBuild.Fields("base dam") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat10") = Fix(ComputeLandDamage(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "ammo"
        rsBuild.Fields("base load") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") And _
           CheckIndex(iDamageIndex, "ParseXDLandChr", "damage") Then
            rsBuild.Fields("stat11") = Fix(ComputeLandAmmo(CInt(strParams(iIndex)), CInt(strParams(iDamageIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "aaf"
        rsBuild.Fields("base aaf") = strParams(iIndex)
        If CheckIndex(iTechIndex, "ParseXDLandChr", "tech") Then
            rsBuild.Fields("stat12") = Fix(ComputeLandAAF(CInt(strParams(iIndex)), CInt(strParams(iTechIndex)), GetTechLevelForShow()))
        End If
    Case "fuelc"
        rsBuild.Fields("stat13") = strParams(iIndex)
    Case "fuelu"
        rsBuild.Fields("stat14") = strParams(iIndex)
    Case "nxlight"
        rsBuild.Fields("stat15") = strParams(iIndex)
    Case "nland"
        rsBuild.Fields("stat16") = strParams(iIndex)
    Case "type"
        rsBuild.Fields("chr_i") = strParams(iIndex)
    Case Else
        EmpireError "ParseXDLandChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

rsBuild.Fields("avail") = CalculateAvail(CInt(rsBuild.Fields("lcm")), CInt(rsBuild.Fields("hcm")))

rsBuild.Update
rsBuild.Index = hldIndex
Exit Sub

'Error handling routine
ParseXDLandChr_Exit:
EmpireError "ParseXDLandChr", strErrorField, strLine
rsBuild.Index = hldIndex
End Sub

Private Sub ParseXDNuke(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iOwnerIndex As Integer
Dim iIDIndex As Integer
Dim iTypeIndex As Integer
Dim iXlocIndex As Integer
Dim iYlocIndex As Integer
Dim iOffIndex As Integer
Dim iTypeLoop As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable
Dim iPrevXloc As Integer
Dim iPrevYloc As Integer

hldIndex = rsNuke.Index
rsNuke.Index = "PrimaryKey"

On Error GoTo ParseXDNuke_Exit

Set etTable = etsTable(XD_NUKE)
If etTable Is Nothing Then
    EmpireError "ParseXDNuke", _
        "The meta data was not found for nuke", strLine
    rsNuke.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDNuke", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsNuke.Index = hldIndex
    Exit Sub
End If

'Get the Sector information to get a record.
iOwnerIndex = etTable.GetZeroIndex("owner")
iIDIndex = etTable.GetZeroIndex("uid")
If VersionCheck(4, 3, 3) <> VER_LESS Then
    iTypeIndex = etTable.GetZeroIndex("type")
Else
    iTypeIndex = etTable.GetZeroIndex("types")
End If
iXlocIndex = etTable.GetZeroIndex("xloc")
iYlocIndex = etTable.GetZeroIndex("yloc")
iOffIndex = etTable.GetZeroIndex("off")

If iIDIndex = -1 Or iTypeIndex = -1 Or iOwnerIndex = -1 Or iXlocIndex = -1 _
    Or iYlocIndex = -1 Or (iOffIndex = -1 And VersionCheck(4, 3, 6) <> VER_LESS) Then
    EmpireError "ParseXDNuke", _
        "One of the following fields is missing: ", _
         "owner, uid, type(s), xloc, yloc"
    rsNuke.Index = hldIndex
    Exit Sub
End If
If VersionCheck(4, 3, 3) <> VER_LESS Then
    rsNuke.Seek "=", strParams(iIDIndex)
    If rsNuke.NoMatch Then
        rsNuke.AddNew
    Else
        rsNuke.Edit
    End If
    strErrorField = "owner"
    rsNuke.Fields("country") = strParams(iOwnerIndex)
    strErrorField = "id"
    rsNuke.Fields("id") = strParams(iIDIndex)
    strErrorField = "x"
    If IsNull(rsNuke.Fields("x")) Then
        iPrevXloc = -1
    Else
        iPrevXloc = rsNuke.Fields("x")
    End If
    rsNuke.Fields("x") = strParams(iXlocIndex)
    strErrorField = "y"
    If IsNull(rsNuke.Fields("y")) Then
        iPrevYloc = -1
    Else
        iPrevYloc = rsNuke.Fields("y")
    End If
    rsNuke.Fields("y") = strParams(iYlocIndex)
    strErrorField = "type"
    rsNuke.Fields("type") = GetUnitType(CInt(strParams(iTypeIndex)), "n")
    rsNuke.Fields("num") = 1
    If VersionCheck(4, 3, 6) <> VER_LESS Then
        rsNuke.Fields("off") = strParams(iOffIndex)
    End If
    rsNuke.Update
Else
    For iTypeLoop = 0 To etTable.Column("types").Length - 1
        If strParams(iTypeIndex + iTypeLoop) > 0 Then
            'Save and restore the index on entry and exit
            rsNuke.Seek "=", strParams(iIDIndex), GetUnitType(iTypeLoop, "n")
            If rsNuke.NoMatch Then
                rsNuke.AddNew
            Else
                rsNuke.Edit
            End If
            strErrorField = GetUnitType(iTypeLoop, "n")
            rsNuke.Fields("country") = strParams(iOwnerIndex)
            rsNuke.Fields("id") = strParams(iIDIndex)
            rsNuke.Fields("x") = strParams(iXlocIndex)
            rsNuke.Fields("y") = strParams(iYlocIndex)
            rsNuke.Fields("type") = GetUnitType(iTypeLoop, "n")
            rsNuke.Fields("num") = strParams(iTypeIndex + iTypeLoop)
            rsNuke.Update
        Else
            rsNuke.Seek "=", strParams(iIDIndex), GetUnitType(iTypeLoop, "n")
            If Not rsNuke.NoMatch Then
                rsNuke.Delete
            End If
        End If
    Next
End If
rsNuke.Index = hldIndex

If (Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY) Or _
    (iPrevXloc = frmDrawMap.SelX And iPrevYloc = frmDrawMap.SelY) Then
    frmDrawMap.FillGrid
End If

CheckForOldEnemyUnit CInt(strParams(iIDIndex)), "n"
Exit Sub

'Error handling routine
ParseXDNuke_Exit:
EmpireError "ParseXDNuke", strErrorField, strLine
rsNuke.Index = hldIndex
End Sub

Private Sub ParseXDNukeChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iIDIndex As Integer
Dim iTechIndex As Integer
Dim iPos As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable
Dim sDRNukeConst As Single

'Save and restore the index on entry and exit
hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

On Error GoTo ParseXDNukeChr_Exit

Set etTable = etsTable(XD_NUKE_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDNukeChr", _
        "The meta data was not found for nuke-chr", strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDNukeChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

'Get the ID information to get a record.
iIDIndex = etTable.GetZeroIndex("name")
strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
iTechIndex = etTable.GetZeroIndex("tech")

If iIDIndex = -1 Then
    EmpireError "ParseXDNukeChr", _
        "The record identifier field not present", _
         CStr(iIDIndex)
    rsBuild.Index = hldIndex
    Exit Sub
End If

iPos = InStr(strParams(iIDIndex), " ")
'
If rsBuild.BOF And rsBuild.EOF Then
    rsBuild.AddNew
    rsBuild.Fields("type") = "n"
Else
    rsBuild.Seek "=", "n", Left$(Left$(strParams(iIDIndex), iPos - 1), 5)
    If rsBuild.NoMatch Then
        rsBuild.AddNew
        rsBuild.Fields("type") = "n"
    Else
        rsBuild.Edit
    End If
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "name"
        If iPos > 0 Then
            rsBuild.Fields("id") = Left$(Left$(strParams(iIndex), iPos - 1), 5)
            rsBuild.Fields("desc") = LTrim$(Mid$(strParams(iIndex), iPos + 1, Len(strParams(iIndex)) - iPos))
        Else
            EmpireError "ParseXDNukeChr", "Incorrect format of name", strLine
        End If
    Case "cost", "avail"
        rsBuild.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
    Case "l_build":
        rsBuild.Fields("lcm") = strParams(iIndex)
    Case "h_build":
        rsBuild.Fields("hcm") = strParams(iIndex)
    Case "o_build":
        rsBuild.Fields("oil") = strParams(iIndex)
    Case "r_build":
        rsBuild.Fields("rad") = strParams(iIndex)
    Case "flags"
        If (strParams(iIndex) And &H1) <> 0 Then
            rsBuild.Fields("cap1") = "neutron"
        Else
            rsBuild.Fields("cap1") = ""
        End If
    Case "blast"
        rsBuild.Fields("stat1") = strParams(iIndex)
    Case "dam"
        rsBuild.Fields("stat2") = strParams(iIndex)
    Case "weight"
        rsBuild.Fields("stat3") = strParams(iIndex)
    Case "tech"
        rsBuild.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
        rsBuild.Fields("stat4") = strParams(iIndex)
        If rsVersion.BOF And rsVersion.EOF Then
            sDRNukeConst = 0#
        Else
            sDRNukeConst = rsVersion.Fields("DRNUKE Const")
        End If
        If sDRNukeConst > 0# Then
            rsBuild.Fields("stat5") = Int((strParams(iIndex) * sDRNukeConst)) + 1
        Else
            rsBuild.Fields("stat5") = 0
        End If
    Case "type"
        rsBuild.Fields("chr_i") = strParams(iIndex)
    Case Else
        EmpireError "ParseXDNukeChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

rsBuild.Update
rsBuild.Index = hldIndex
Exit Sub

'Error handling routine
ParseXDNukeChr_Exit:
EmpireError "ParseXDNukeChr", strErrorField, strLine
rsBuild.Index = hldIndex
End Sub

Private Sub ParseXDItemChr(strLine As String)
Dim strParams() As String
Dim iIDIndex As Integer
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim strErrorField As String
Dim eiChr As EmpItem
Dim bNew As Boolean
Dim etTable As EmpTable

On Error GoTo ParseXDItemChr_Exit

Set etTable = etsTable(XD_ITEM_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDItemChr", _
        "The meta data was not found for item-chr", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDItemChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

'Get the ID information to get a record.
iIDIndex = etTable.GetZeroIndex("mnem")

If iIDIndex = -1 Then
    EmpireError "ParseXDItemChr", _
        "The record identifier field not present", _
         CStr(iIDIndex)
    Exit Sub
End If

strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))

If strParams(iIDIndex) = "0" Then 'skip "unused" row
    Exit Sub
End If
    
Set eiChr = Items.FindByLetter(strParams(iIDIndex))
If eiChr Is Nothing Then
    Set eiChr = New EmpItem
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
        eiChr.Value = strParams(iIndex)
    Case "sell"
        If strParams(iIndex) = 0 Then
            eiChr.Sellable = False
        Else
            eiChr.Sellable = True
        End If
    Case "lbs"
        eiChr.Weight = strParams(iIndex)
    Case "uid"
        eiChr.ChrIndex = strParams(iIndex)
    Case "mnem"
        eiChr.Letter = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
    Case "pkg"
        eiChr.Packing(PACKING_INEFFICIENT) = strParams(iIndex)
        eiChr.Packing(PACKING_REGULAR) = strParams(iIndex + 1)
        eiChr.Packing(PACKING_WAREHOUSE) = strParams(iIndex + 2)
        eiChr.Packing(PACKING_URBAN) = strParams(iIndex + 3)
        eiChr.Packing(PACKING_BANK) = strParams(iIndex + 4)
        iIndex = iIndex + 4
    Case "melt_denom"
    Case Else
        EmpireError "ParseXDItemChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

eiChr.ChrIndex = iRowCount
If bNew Then
    eiChr.IntelligenceNames = eiChr.ItemName
    eiChr.ProductName = eiChr.ItemName
    eiChr.FormName = eiChr.ItemName
    eiChr.ConditionalName = Left$(eiChr.ItemName, 3)
    eiChr.DatabaseName = Left$(eiChr.ItemName, 3)
    eiChr.SectorPanelLabel = UCase(Left$(eiChr.ItemName, 1)) + Mid$(eiChr.ItemName, 2, 2) + ":"
    eiChr.DistributionPanelLabel = UCase(Left$(eiChr.ItemName, 1)) + Mid$(eiChr.ItemName, 2, 2)
    Items.AddItem (eiChr)
End If

iRowCount = iRowCount + 1
Exit Sub

'Error handling routine
ParseXDItemChr_Exit:
EmpireError "ParseXDItemChr", strErrorField, strLine
End Sub

Private Sub ParseXDInfraChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iIDIndex As Integer
Dim iPos As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable
Dim bEnable As Boolean

'Save and restore the index on entry and exit
hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

On Error GoTo ParseXDInfraChr_Exit

Set etTable = etsTable(XD_INFRA_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDInfraChr", _
        "The meta data was not found for infra", strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDInfraChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsBuild.Index = hldIndex
    Exit Sub
End If

'Get the ID information to get a record.
iIDIndex = etTable.GetZeroIndex("name")
If iIDIndex = -1 Then
    EmpireError "ParseXDInfraChr", _
        "The record identifier field not present", _
         CStr(iIDIndex)
    rsBuild.Index = hldIndex
    Exit Sub
End If

strParams(iIDIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIDIndex))
iPos = InStr(strParams(iIDIndex), " ")
If iPos > 0 Then
    strParams(iIDIndex) = Left$(strParams(iIDIndex), iPos - 1)
End If
If Len(strParams(iIDIndex)) > 5 Then
    strParams(iIDIndex) = Left$(strParams(iIDIndex), 5)
End If

If rsBuild.BOF And rsBuild.EOF Then
    rsBuild.AddNew
    rsBuild.Fields("type") = "i"
Else
    rsBuild.Seek "=", "i", strParams(iIDIndex)
    If rsBuild.NoMatch Then
        rsBuild.AddNew
        rsBuild.Fields("type") = "i"
    Else
        rsBuild.Edit
    End If
End If

iHeaderIndex = 1
bEnable = True
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "name"
        rsBuild.Fields("id") = strParams(iIndex)
    Case "dcost"
        rsBuild.Fields("cost") = strParams(iIndex)
    Case "lcms"
        rsBuild.Fields("lcm") = strParams(iIndex)
    Case "hcms"
        rsBuild.Fields("hcm") = strParams(iIndex)
    Case "mcost"
        rsBuild.Fields("stat2") = strParams(iIndex)
    Case "enable"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            If strParams(iIndex) = "0" Then
                If Not rsBuild.NoMatch Then
                    rsBuild.Delete
                Else
                    rsBuild.CancelUpdate
                End If
                Exit Sub
            End If
        Else
            EmpireError "ParseXDLandChr", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDLandChr", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex

rsBuild.Update
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
EmpireError "ParseXDInfraChr", strErrorField, strLine
rsBuild.Index = hldIndex
End Sub

Private Sub ParseXDProductChr(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable

'Save and restore the index on entry and exit
hldIndex = rsSectorType.Index
rsSectorType.Index = "PrimaryKey"

On Error GoTo ParseXDProductChr_Exit

Set etTable = etsTable(XD_PRODUCT_CHR)
If etTable Is Nothing Then
    EmpireError "ParseXDProductChr", _
        "The meta data was not found for product-chr", strLine
    rsSectorType.Index = hldIndex
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDProductChr", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    rsSectorType.Index = hldIndex
    Exit Sub
End If

rsSectorType.MoveFirst
While Not rsSectorType.EOF And iRowCount <> 0
    If iRowCount = rsSectorType.Fields("pchr_i") Then
        rsSectorType.Edit
        iHeaderIndex = 1
        For iIndex = 0 To UBound(strParams)
            strErrorField = etTable.Column(iHeaderIndex).Name
            Select Case etTable.Column(iHeaderIndex).Name
            Case "sname"
                rsSectorType.Fields("product") = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
            Case "name"
            Case "ctype"
                rsSectorType.Fields("use1s") = GetType(Val(strParams(iIndex)))
                iIndex = iIndex + 1
                rsSectorType.Fields("use2s") = GetType(Val(strParams(iIndex)))
                iIndex = iIndex + 1
                rsSectorType.Fields("use3s") = GetType(Val(strParams(iIndex)))
            Case "camt"
                rsSectorType.Fields("use1n") = strParams(iIndex)
                iIndex = iIndex + 1
                rsSectorType.Fields("use2n") = strParams(iIndex)
                iIndex = iIndex + 1
                rsSectorType.Fields("use3n") = strParams(iIndex)
            Case "nlndx"
                rsSectorType.Fields("level") = GetNationType(Val(strParams(iIndex)))
            Case "cost"
                rsSectorType.Fields(etTable.Column(iHeaderIndex).Name) = strParams(iIndex)
            Case "nrndx", "level", "uid"
            Case "type"
                rsSectorType.Fields("pcode") = GetType(Val(strParams(iIndex)))
            Case "nlmin"
                rsSectorType.Fields("min") = strParams(iIndex)
            Case "effic"
                If VersionCheck(4, 3, 6) = VER_LESS Then
                    rsSectorType.Fields("eff") = strParams(iIndex)
                Else
                    EmpireError "ParseXDProductChr", "Field not understood:" + strErrorField, strLine
                End If
            Case "nllag"
                rsSectorType.Fields("lag") = strParams(iIndex)
            Case "nrdep"
                rsSectorType.Fields("dep") = strParams(iIndex)
            Case Else
                EmpireError "ParseXDProductChr", "Field not understood:" + strErrorField, strLine
            End Select
            iHeaderIndex = iHeaderIndex + 1
        Next iIndex
        rsSectorType.Update
    End If
    rsSectorType.MoveNext
Wend
rsSectorType.Index = hldIndex
iRowCount = iRowCount + 1
Exit Sub

'Error handling routine
ParseXDProductChr_Exit:
EmpireError "ParseXDProductChr", strErrorField, strLine
rsSectorType.Index = hldIndex
End Sub

Private Sub ParseXDGame(strLine As String)
Dim strParams() As String
Dim etTable As EmpTable
Dim strErrorField As String
Dim iIndex As Integer
Static iTurn As Integer

On Error GoTo ParseXDGame_Exit

Set etTable = etsTable(XD_GAME)
If etTable Is Nothing Then
    EmpireError "ParseXDGame", _
        "The meta data was not found for game", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDGame", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iIndex + 1).Name
    Select Case etTable.Column(iIndex + 1).Name
    Case "turn"
        If iTurn = 0 Then
            iTurn = CInt(strParams(iIndex))
        ElseIf iTurn <> CInt(strParams(iIndex)) And frmEmpCmd.bAutoUpdate Then
            iTurn = CInt(strParams(iIndex))
            UpdateCommands
        End If
    Case "upd_disable"
        frmEmpCmd.bUpdateEnabled = Not CBool(strParams(iIndex))
    Case "tick", "rt", "down"
        'ignore
    Case Else
        EmpireError "ParseXDGame", "Field not understood:" + strErrorField, strLine
    End Select
Next
Exit Sub

ParseXDGame_Exit:
EmpireError "ParseXDGame", strErrorField, strLine
End Sub


Private Sub ParseXDUpdates(strLine As String)
Dim strParams() As String
Dim etTable As EmpTable
Dim strErrorField As String
Dim iIndex As Integer
Dim lTimeUpdate As Long
Dim lTimeCurrent As Long

If Not frmEmpCmd.bSetUpdateTime Then
    Exit Sub
End If
frmEmpCmd.bSetUpdateTime = False

On Error GoTo ParseXDUpdates_Exit

Set etTable = etsTable(XD_UPDATES)
If etTable Is Nothing Then
    EmpireError "ParseXDUpdates", _
        "The meta data was not found for update", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDUpdates", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iIndex + 1).Name
    Select Case etTable.Column(iIndex + 1).Name
    Case "time"
            lTimeUpdate = CLng(strParams(iIndex))
            If lTimeUpdate > 0 Then
                lTimeCurrent = CLng(tmStamp)
                frmDrawMap.TimerAtUpdate = Timer
                frmDrawMap.SecondsToUpdate = lTimeUpdate - lTimeCurrent
                frmDrawMap.UpdateTimer.Interval = 1000
                frmDrawMap.UpdateTimer.Enabled = True
            Else
                frmDrawMap.TimerAtUpdate = 0
                frmDrawMap.SecondsToUpdate = 0
            End If
    Case Else
        EmpireError "ParseXDUpdates", "Field not understood:" + strErrorField, strLine
    End Select
Next
Exit Sub

ParseXDUpdates_Exit:
EmpireError "ParseXDUpdates", strErrorField, strLine
End Sub

Private Sub ParseXDVersion(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iETUupdate As Integer
Dim iBuildBridgeHcm As Integer
Dim iBuildTowerHcm As Integer
Dim strErrorField As String
Dim etTable As EmpTable
Dim bXDumpServer As Boolean

On Error GoTo ParseXDVersion_Exit

Set etTable = etsTable(XD_VERSION)
If etTable Is Nothing Then
    EmpireError "ParseXDVersion", _
        "The meta data was not found for version", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDVersion", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

iETUupdate = etTable.GetZeroIndex("etu_per_update")
iBuildBridgeHcm = etTable.GetZeroIndex("buil_bh")
iBuildTowerHcm = etTable.GetZeroIndex("buil_tower_bh")

If VersionCheck(4, 3, 0) <> VER_LESS Then
    bXDumpServer = True
Else
    bXDumpServer = False
End If

If rsVersion.RecordCount = 0 Then
    rsVersion.AddNew
Else
    rsVersion.MoveFirst
    rsVersion.Edit
End If
iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "privname", "privlog", "update_policy", "update_window", "update_times", _
         "blitz_time", "update_demandpolicy", "update_wantmin", "update_missed", _
         "update_demandtimes", "game_days", "game_hours", _
         "ALL_BLEED", "AUTO_POWER", "BLITZ", "BRIDGETOWERS", "DEFENSE_INFRA", _
         "DEMANDUPDATE", "DRNUKE", "EASY_BRIDGES", "FALLOUT", "FUEL", "GODNEWS", _
         "GUINEA_PIGS", "HIDDEN", _
         "INTERDICT_ATT", "LANDSPIES", "LOANS", "LOSE_CONTACT", "MARKET", "MOB_ACCESS", _
         "NOMOBCOST", "NO_FORT_FIRE", "NO_PLAGUE", _
         "PINPOINTMISSILE", _
         "RES_POP", "SAIL", "SECT_MOB_ACCESS", "SHOWPLANE", "SLOW_WAR", _
         "SUPER_BARS", "TECH_POP", "TRADESHIPS", "TREATIES", "UNIT_MOB_ACCESS", _
         "m_m_p_d", "max_btus", _
         "max_idle", "players_at_00", _
         "war_cost", "ally_factor", _
         "money_land", "money_plane", "money_ship", _
         "torpedo_damage", "fort_max_interdiction_range", "land_max_interdiction_range", _
         "ship_max_interdiction_range", "flakscale", "combat_mob", "people_damage", _
         "unit_damage", "collateral_dam", "assault_penalty", _
         "sect_mob_neg_factor", "unit_mob_neg_factor", _
         "mission_mob_cost", "decay_per_etu", "fallout_spread", _
         "MARK_DELAY", "TRADE_DELAY", "buytax", "tradetax", "trade_1_dist", _
         "trade_2_dist", "trade_3_dist", "trade_1", "trade_2", "trade_3", _
         "trade_ally_bonus", "trade_ally_cut", "fuel_mult", "update_demand", _
         "maxnoc", "max_idle_visitor", "login_grace_time"
    Case "version"
        Dim strVersions() As String
        Dim iPos As Integer
        strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
        iPos = InStrRev(strParams(iIndex), " ")
        If Len(strParams(iIndex)) > 1 And iPos > 0 Then
            strVersions = Split(Right$(strParams(iIndex), Len(strParams(iIndex)) - iPos + 1), ".")
            If UBound(strVersions) = 2 Then
                rsVersion.Fields("Major") = strVersions(0)
                rsVersion.Fields("Minor") = strVersions(1)
                rsVersion.Fields("Revision") = strVersions(2)
            Else
                EmpireError "ParseXDVersion", "Wrong number of fields in version:" + strErrorField, strLine
            End If
        Else
            EmpireError "ParseXDVersion", "Unexpected format in version:" + strErrorField, strLine
        End If
    Case "WORLD_X"
        rsVersion.Fields("World X") = strParams(iIndex)
    Case "WORLD_Y"
        rsVersion.Fields("World Y") = strParams(iIndex)
    Case "s_p_etu"
        rsVersion.Fields("Empire Time Unit") = strParams(iIndex)
    Case "etu_per_update"
        rsVersion.Fields("ETU per update") = strParams(iIndex)
    Case "btu_build_rate"
        rsVersion.Fields("Civs for BTUs") = (1# / (Val(strParams(iIndex)) * 100#))
    Case "fgrate"
        rsVersion.Fields("Food grow") = Val(strParams(iIndex)) * 100#
    Case "fcrate"
        rsVersion.Fields("Food harvest") = Val(strParams(iIndex)) * 1000#
    Case "obrate"
        rsVersion.Fields("Civ Birthrate") = Val(strParams(iIndex)) * 1000#
    Case "uwbrate"
        rsVersion.Fields("Uw Birthrate") = Val(strParams(iIndex)) * 1000#
    Case "eatrate"
        rsVersion.Fields("Adult food") = Val(strParams(iIndex)) * 1000#
    Case "babyeat"
        rsVersion.Fields("Baby food") = Val(strParams(iIndex)) * 1000#
    Case "bankint"
        rsVersion.Fields("Bank Interest") = Val(strParams(iIndex)) * 1000#
    Case "money_civ"
        rsVersion.Fields("Civ Taxes") = Val(strParams(iIndex)) * 1000#
    Case "money_uw"
        rsVersion.Fields("UW Taxes") = Val(strParams(iIndex)) * 1000#
    Case "money_mil"
        rsVersion.Fields("Mil Salary") = -Val(strParams(iIndex)) * 1000#
    Case "money_res"
        rsVersion.Fields("Reserve Salary") = -Val(strParams(iIndex)) * 1000#
    Case "rollover_avail_max"
        rsVersion.Fields("Rollover Avail") = strParams(iIndex)
    Case "hap_cons"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            rsVersion.Fields("Civs per stroller") = Val(strParams(iIndex)) / Val(strParams(iETUupdate))
        End If
    Case "edu_cons"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            rsVersion.Fields("Civs per graduate") = Val(strParams(iIndex)) / Val(strParams(iETUupdate))
        End If
    Case "hap_avg"
        rsVersion.Fields("Happy average") = Val(strParams(iIndex))
    Case "edu_avg"
        rsVersion.Fields("Edu average") = Val(strParams(iIndex))
    Case "tech_log_base"
        rsVersion.Fields("Tech base") = Val(strParams(iIndex))
    Case "easy_tech"
        rsVersion.Fields("Easy tech") = Val(strParams(iIndex))
    Case "level_age_rate"
        rsVersion.Fields("Tech Decay Rate") = 1
        rsVersion.Fields("Tech Decay Time") = Val(strParams(iIndex))
    Case "fire_range_factor"
        rsVersion.Fields("Fire range") = strParams(iIndex)
    Case "sect_mob_max"
        rsVersion.Fields("Max mob - Sector") = strParams(iIndex)
    Case "ship_mob_max"
        rsVersion.Fields("Max mob - Ship") = strParams(iIndex)
    Case "plane_mob_max"
        rsVersion.Fields("Max mob - Plane") = strParams(iIndex)
    Case "land_mob_max"
        rsVersion.Fields("Max mob - Land") = strParams(iIndex)
    Case "sect_mob_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            rsVersion.Fields("Mob gain - Sector") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
        End If
    Case "ship_mob_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            rsVersion.Fields("Mob gain - Ship") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
        End If
    Case "plane_mob_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            rsVersion.Fields("Mob gain - Plane") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
        End If
    Case "land_mob_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            rsVersion.Fields("Mob gain - Land") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
        End If
    Case "ship_grow_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            If Val(strParams(iIndex)) * Val(strParams(iETUupdate)) > 100 Then
                rsVersion.Fields("Eff gain - Ship") = 100#
            Else
                rsVersion.Fields("Eff gain - Ship") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
            End If
        End If
    Case "plane_grow_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            If Val(strParams(iIndex)) * Val(strParams(iETUupdate)) > 100 Then
                rsVersion.Fields("Eff gain - Plane") = 100#
            Else
                rsVersion.Fields("Eff gain - Plane") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
            End If
        End If
    Case "land_grow_scale"
        If CheckIndex(iETUupdate, "ParseXDVersion", "ETU per update") Then
            If Val(strParams(iIndex)) * Val(strParams(iETUupdate)) > 100 Then
                rsVersion.Fields("Eff gain - Land") = 100#
            Else
                rsVersion.Fields("Eff gain - Land") = Val(strParams(iIndex)) * Val(strParams(iETUupdate))
            End If
        End If
    Case "buil_bh", "buil_tower_bh"
    Case "buil_bt"
    'pr("Bridges require %g tech,", buil_bt);
    'pr(" %d hcm,", buil_bh);
    'pr(" %d workers,\n", 0);
        ParseShowText "Bridges require " + CStr(Int(Val(strParams(iIndex)))) + " tech, " + _
            CStr(strParams(iBuildBridgeHcm)) + " hcm, 0 workers", "b", "b", 2
    Case "buil_bc"
    'pr("%d available workforce, and cost $%g\n",
    '  (SCT_BLD_WORK(0, buil_bh) * SCT_MINEFF + 99) / 100,
    '  buil_bc);
        ParseShowText CStr(Int((SectorBuildWork(0, CInt(strParams(iBuildBridgeHcm))) * 20 + 99) / 100)) + _
            " available workforce, and cost $" + CStr(Int(Val(strParams(iIndex)))), _
            "b", "b", 3
    Case "buil_tower_bt"
    'pr("Bridge Towers require %g tech,", buil_tower_bt);
    'pr(" %d hcm,", buil_tower_bh);
    'pr(" %d workers,\n", 0);
        ParseShowText "Bridge Towers require " + CStr(Int(Val(strParams(iIndex)))) + " tech, " + _
            CStr(strParams(iBuildTowerHcm)) + " hcm, 0 workers", "t", "b", 2
    Case "buil_tower_bc"
    'pr("%d available workforce, and cost $%g\n",
    '   (SCT_BLD_WORK(0, buil_tower_bh) * SCT_MINEFF + 99) / 100,
    '   buil_tower_bc);
        ParseShowText CStr(Int((SectorBuildWork(0, CInt(strParams(iBuildTowerHcm))) * 20 + 99) / 100)) + _
            " available workforce, and cost $" + CStr(Int(Val(strParams(iIndex)))), _
            "t", "b", 3
    Case "NOFOOD"
        If strParams(iIndex) = 1 Then
            rsVersion.Fields("Food Needed") = "N"
        Else
            rsVersion.Fields("Food Needed") = "Y"
        End If
    Case "BIG_CITY", "GO_RENEW", "RAILWAYS"
        If strParams(iIndex) = 1 Then
            rsVersion.Fields(etTable.Column(iHeaderIndex).Name) = True
        Else
            rsVersion.Fields(etTable.Column(iHeaderIndex).Name) = False
        End If
    Case "drnuke_const"
            rsVersion.Fields("DRNUKE Const") = Val(strParams(iIndex))
    Case "LIMBER"
        If frmOptions.bHeavyMetal Then
            'Assume the bHeavyMetal has set both of this options for the time being
        Else
            EmpireError "ParseXDVersion", "Field not understood:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDVersion", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
rsVersion.Update
If Not bXDumpServer And VersionCheck(4, 3, 0) <> VER_LESS Then
    RequestMetaTables
End If
Exit Sub

'Error handling routine
ParseXDVersion_Exit:
EmpireError "ParseXDVersion", strErrorField, strLine
End Sub

Public Function GetIndexStringArray(strArray() As String, strMatch As String) As Integer
Dim iIndex As Integer
Dim iAdjustedIndex As Integer

If IsNull(strArray) Then
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

Private Function GetUnitType(eiChrIndex As Integer, strUnit As String) As String
If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("chr_i") = eiChrIndex And rsBuild.Fields("type") = strUnit Then
            GetUnitType = rsBuild.Fields("id")
            Exit Function
        End If
        rsBuild.MoveNext
    Wend
End If
GetUnitType = "????"
End Function

Public Function RemoveEscapeSequencesAndQuotes(strParam As String) As String
Dim iStartPos As Integer
Dim iFoundPos As Integer
Dim strOutput As String

iFoundPos = InStr(strParam, "\")

If iFoundPos = 0 Then
    If InStr(strParam, Chr$(34)) = 1 Then
        strParam = Mid$(strParam, 2, Len(strParam) - 2)
    End If
    RemoveEscapeSequencesAndQuotes = strParam
    Exit Function
End If

iStartPos = 1
While iFoundPos <> 0
    strOutput = strOutput + Mid$(strParam, iStartPos, iFoundPos - iStartPos)
    strOutput = strOutput + Chr$("&" + Mid$(strParam, iFoundPos + 1, 3))
    iStartPos = iFoundPos + 4
    iFoundPos = InStr(iStartPos, strParam, "\")
Wend
strParam = strOutput + Mid$(strParam, iStartPos)
If InStr(strParam, Chr$(34)) = 1 Then
    strParam = Mid$(strParam, 2, Len(strParam) - 2)
End If
RemoveEscapeSequencesAndQuotes = strParam
End Function

Private Function CheckIndex(iIndex As Integer, strFunction As String, strField As String)
If iIndex < 0 Then
    EmpireError strFunction, "The " + strField + " field not present", CStr(iIndex)
    CheckIndex = False
Else
    CheckIndex = True
End If
End Function

Private Sub ParseXDMetaHeader(strType As String)
Dim xdType As enumXdumpType

xdType = ParseXDTableName(strType)
etsTable.AddUpdate xdType, strType
xdMetaType = xdType
End Sub

Private Sub ParseXDSymbolTableHeader(xdSymbolTable As enumXdumpType, strType As String)
estsTable.AddUpdate xdSymbolTable, strType
End Sub

Private Sub ParseXDNation(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iCnumIndex As Integer
Dim iXcapIndex As Integer
Dim strErrorField As String
Dim etTable As EmpTable

On Error GoTo ParseXDNation_Exit

Set etTable = etsTable(XD_NATION)
If etTable Is Nothing Then
    EmpireError "ParseXDNation", _
        "The meta data was not found for nat", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDNation", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

iCnumIndex = etTable.GetZeroIndex("cnum")
iXcapIndex = etTable.GetZeroIndex("xcap")

If iCnumIndex = -1 Or iXcapIndex = -1 Then
    EmpireError "ParseXDNation", _
        "One of the following is missing: country number or X location for the capital", ""
    Exit Sub
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "cnum"
        CountryNumber = Val(strParams(iIndex))
    Case "cname"
        If VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
            frmEmpCmd.CountryName = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
            Nations.AddNation Val(strParams(iCnumIndex)), frmEmpCmd.CountryName
            If rsNations.BOF And rsNations.EOF Then
                rsNations.AddNew
            Else
                rsNations.Seek "=", Val(strParams(iCnumIndex))
                If rsNations.NoMatch Then
                    rsNations.AddNew
                Else
                    rsNations.Edit
                End If
            End If
            rsNations.Fields("Name") = frmEmpCmd.CountryName
            rsNations.Fields("Number") = Val(strParams(iCnumIndex))
            rsNations.Update
        Else
            EmpireError "ParseXDNation", "Field not understood:" + strErrorField, strLine
        End If
    Case "xcap"
    Case "ycap"
        CapSect = SectorString(Val(strParams(iXcapIndex)), Val(strParams(iIndex)))
    Case "access"
        If VersionCheck(4, 3, 10) = VER_LESS Then
            EmpireError "ParseXDCountries", "Access field invalid:" + strErrorField, strLine
        End If
    Case "milreserve":
        MilReserves = Val(strParams(iIndex))
    Case "education":
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
        MaxSafeCivs = Int(CSng(Maxpop) / (1# + _
            (Val(rsVersion.Fields("Civ Birthrate")) / 1000#) * _
            Val(rsVersion.Fields("ETU per update"))))
        If MaxSafeCivs > Maxpop Then
           MaxSafeCivs = Maxpop
        End If
        MaxSafeUws = Int(CSng(Maxpop) / (1# + _
            (Val(rsVersion.Fields("Uw Birthrate")) / 1000#) * _
            Val(rsVersion.Fields("ETU per update"))))
        If MaxSafeUws > Maxpop Then
           MaxSafeUws = Maxpop
        End If
    Case "happiness"
        Happiness = Val(strParams(iIndex))
    Case "max"
    Case "stat", "passwd", "ip", "hostname", "userid", _
         "xorg", "yorg", "dayno", "update", "missed", _
        "tgms", "ann", "minused", "timeused", "btu", "money", "login", _
        "logout", "newstim", "annotim", _
         "flags"
    Case "relations", "contacts", "rejects"
        If VersionCheck(4, 3, 4) <> VER_LESS Or frmOptions.bSPAtlantis Then
            EmpireError "ParseXDNation", "Field not understood:" + strErrorField, strLine
        End If
        iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
    Case Else
        EmpireError "ParseXDNation", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
SaveNationVariables
Exit Sub

'Error handling routine
ParseXDNation_Exit:
EmpireError "ParseXDNation", strErrorField, strLine
End Sub

Private Sub ParseXDCountries(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iHeaderIndex As Integer
Dim iCnumIndex As Integer
Dim iStatusIndex As Integer
Dim iRelIndex As Integer
Dim iStatus As Integer
Dim strErrorField As String
Dim etTable As EmpTable

On Error GoTo ParseXDCountries_Exit

Set etTable = etsTable(XD_COUNTRIES)
If etTable Is Nothing Then
    EmpireError "ParseXDCountries", _
        "The meta data was not found for cou", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDCountries", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

iCnumIndex = etTable.GetZeroIndex("cnum")
iStatusIndex = etTable.GetZeroIndex("stat")

If iCnumIndex = -1 And iStatusIndex = -1 Then
    EmpireError "ParseXDCountries", _
        "One of the following is missing: country number or status", ""
    Exit Sub
End If

iHeaderIndex = 1
For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iHeaderIndex).Name
    Select Case etTable.Column(iHeaderIndex).Name
    Case "cname"
        iStatus = Val(strParams(iStatusIndex))
        Select Case iStatus
        Case STAT_SANCT, STAT_ACTIVE, STAT_GOD
            strParams(iIndex) = RemoveEscapeSequencesAndQuotes(strParams(iIndex))
            Nations.AddNation Val(strParams(iCnumIndex)), strParams(iIndex)
            If rsNations.BOF And rsNations.EOF Then
                rsNations.AddNew
            Else
                rsNations.Seek "=", Val(strParams(iCnumIndex))
                If rsNations.NoMatch Then
                    rsNations.AddNew
                Else
                    rsNations.Edit
                End If
            End If
            rsNations.Fields("Name") = strParams(iIndex)
            rsNations.Fields("Number") = Val(strParams(iCnumIndex))
            rsNations.Update
        End Select
    Case "cnum", "stat", "cname", "passwd", "ip", "hostname", "userid", _
        "xcap", "ycap", "xorg", "yorg", "dayno", "update", "missed", _
        "tgms", "ann", "minused", "timeused", "btu", "milreserve", "money", "login", _
        "logout", "newstim", "annotim", "tech", "research", "education", _
        "happiness", "flags"
    Case "relations" ' 7 0 99 38
        If VersionCheck(4, 3, 4) <> VER_LESS Then
            For iRelIndex = 0 To etTable.Column(iHeaderIndex).Length - 1
                If Nations.NationName(iRelIndex) <> vbNullString Then
                    Nations.Relation(Val(strParams(iCnumIndex)), iRelIndex) = strParams(iIndex + iRelIndex)
                End If
            Next
        End If
        iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
    Case "contacts" ' 6 1 99 -1
        iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
    Case "rejects" ' 6 1 99 -1
        iIndex = iIndex + etTable.Column(iHeaderIndex).Length - 1
    Case "access" ' 6 0 0 -1
        If VersionCheck(4, 3, 10) = VER_LESS Or Not bDeity Then
            EmpireError "ParseXDCountries", "Access field invalid:" + strErrorField, strLine
        End If
    Case Else
        EmpireError "ParseXDCountries", "Field not understood:" + strErrorField, strLine
    End Select
    iHeaderIndex = iHeaderIndex + 1
Next iIndex
Exit Sub

'Error handling routine
ParseXDCountries_Exit:
EmpireError "ParseXDCountries", strErrorField, strLine
End Sub

Private Sub ParseXDNationRelationships(strLine As String)
Dim strParams() As String
Dim etTable As EmpTable
Dim strErrorField As String
Dim iIndex As Integer
Dim iValue As Integer

On Error GoTo ParseXDNationRelationships_Exit

Set etTable = etsTable(XD_NATION_RELATIONSHIPS)
If etTable Is Nothing Then
    EmpireError "ParseXDNationRelationships", _
        "The meta data was not found for nation relationships", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDNationRelationships", _
        "The number of fields does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iIndex + 1).Name
    Select Case etTable.Column(iIndex + 1).Name
    Case "value"
        iValue = CInt(strParams(iIndex))
    Case "name"
        Select Case RemoveEscapeSequencesAndQuotes(strParams(iIndex))
        Case "unknown"
            iREL_AT_WAR = -1
            iREL_SITZKRIEG = -1   'SLOW_WAR only
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
            EmpireError "ParseXDNationRelationships", "Expected Relationship: " + strErrorField, strLine
        End Select
    Case Else
        EmpireError "ParseXDNationRelationships", "Field not understood:" + strErrorField, strLine
    End Select
Next
Exit Sub

ParseXDNationRelationships_Exit:
EmpireError "ParseXDNationRelationships", strErrorField, strLine
End Sub

Private Sub ParseXDLost(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iTypeIndex As Integer
Dim iIDIndex As Integer
Dim iXlocIndex As Integer
Dim iYlocIndex As Integer
Dim iTimestamp As Long
Dim strErrorField As String
Dim hldIndex As String
Dim etTable As EmpTable
Dim rsLost As Recordset

hldIndex = ""

On Error GoTo ParseXDLost_Exit

Set etTable = etsTable(XD_LOST)
If etTable Is Nothing Then
    EmpireError "ParseXDLost", _
        "The meta data was not found for lost", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDLost", _
        "The number of fields does not match the header", _
        CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

iTypeIndex = etTable.GetZeroIndex("type")
iIDIndex = etTable.GetZeroIndex("id")
iXlocIndex = etTable.GetZeroIndex("x")
iYlocIndex = etTable.GetZeroIndex("y")
iTimestamp = etTable.GetZeroIndex("timestamp")

If iTypeIndex = -1 Or iIDIndex = -1 Or iXlocIndex = -1 Or iYlocIndex = -1 Or _
    iTimestamp = -1 Then
    EmpireError "ParseXDLost", _
        "One of the following: type, id, x, y or timestamp is not present", ""
    Exit Sub
End If

Select Case strParams(iTypeIndex)
Case LT_SECTOR
    Set rsLost = rsSectors
    strErrorField = "Invalid Sector"
    hldIndex = rsLost.Index
    rsLost.Index = "PrimaryKey"
    rsLost.Seek "=", CInt(strParams(iXlocIndex)), CInt(strParams(iYlocIndex))
    If Not rsLost.NoMatch Then
        If CLng(strParams(iTimestamp)) > rsLost.Fields("timestamp") Then
            rsLost.Delete
        End If
    End If
Case LT_SHIP
    Set rsLost = rsShip
    strErrorField = "Invalid Ship"
    hldIndex = rsLost.Index
    rsLost.Index = "PrimaryKey"
    rsLost.Seek "=", CInt(strParams(iIDIndex))
    If Not rsLost.NoMatch Then
        If CLng(strParams(iTimestamp)) > rsLost.Fields("timestamp") Then
            rsLost.Delete
        End If
    End If
Case LT_PLANE
    Set rsLost = rsPlane
    strErrorField = "Invalid Plane"
    hldIndex = rsLost.Index
    rsLost.Index = "PrimaryKey"
    rsLost.Seek "=", CInt(strParams(iIDIndex))
    If Not rsLost.NoMatch Then
        If CLng(strParams(iTimestamp)) > rsLost.Fields("timestamp") Then
            rsLost.Delete
        End If
    End If
Case LT_LAND
    Set rsLost = rsLand
    strErrorField = "Invalid Land"
    hldIndex = rsLost.Index
    rsLost.Index = "PrimaryKey"
    rsLost.Seek "=", CInt(strParams(iIDIndex))
    If Not rsLost.NoMatch Then
        If CLng(strParams(iTimestamp)) > rsLost.Fields("timestamp") Then
            rsLost.Delete
        End If
    End If
Case LT_NUKE
    Set rsLost = rsNuke
    strErrorField = "Invalid Nuke"
    hldIndex = rsLost.Index
    rsLost.Index = "PrimaryKey"
    rsLost.Seek "=", CInt(strParams(iIDIndex))
    While Not rsLost.NoMatch
        rsLost.Delete
        rsLost.Seek "=", CInt(strParams(iIDIndex))
    Wend
End Select

rsLost.Index = hldIndex

If Val(strParams(iXlocIndex)) = frmDrawMap.SelX And Val(strParams(iYlocIndex)) = frmDrawMap.SelY Then
    frmDrawMap.FillSectorBox Val(strParams(iXlocIndex)), Val(strParams(iYlocIndex))
    If Not frmDrawMap.IsAnyTabActive Then
        frmDrawMap.DisplayFirstSectorPanel
    End If
End If

Exit Sub

'Error handling routine
ParseXDLost_Exit:
EmpireError "ParseXDLost", strErrorField, strLine
If Len(hldIndex) > 0 Then
    rsLost.Index = hldIndex
End If
End Sub

Private Sub ParseXDMeta(xdTable As enumXdumpType, strLine As String)
Dim etTable As EmpTable

Set etTable = etsTable.Table(xdTable)
If Not etTable Is Nothing Then
    etTable.AddUpdateColumn strLine
End If
End Sub

Private Sub ParseXDSymbolTable(strLine As String)
Dim strParams() As String
Dim iIndex As Integer
Dim iValueIndex As Integer
Dim strErrorField As String
Dim etTable As EmpTable

On Error GoTo ParseXDSymbolTable_Exit

Set etTable = etsTable(xdType)
If etTable Is Nothing Then
    EmpireError "ParseXDSymbolTable", _
        "The meta data was not found for table (" + CStr(xdType) + ")", strLine
    Exit Sub
End If

strParams = Split(strLine, " ")

If UBound(strParams) + 1 <> etTable.ParameterCount Then
    EmpireError "ParseXDSymbolTable", _
        "The number of fields for table " + etTable.Name + " does not match the header", CStr(etTable.ParameterCount) + ":" + strLine
    Exit Sub
End If

iValueIndex = etTable.GetZeroIndex("value")

If iValueIndex = -1 Then
    EmpireError "ParseXDSymbolTable", _
        "The record identifier field not present", _
         CStr(iValueIndex)
    Exit Sub
End If

For iIndex = 0 To UBound(strParams)
    strErrorField = etTable.Column(iIndex + 1).Name
    Select Case etTable.Column(iIndex + 1).Name
    Case "name"
        estsTable(xdType).AddUpdateSymbolEntry CLng(strParams(iValueIndex)), _
            RemoveEscapeSequencesAndQuotes(strParams(iIndex))
    Case "value"
    End Select
Next
Exit Sub

'Error handling routine
ParseXDSymbolTable_Exit:
EmpireError "ParseXDSymbolTable", strErrorField, strLine
End Sub

Private Function ParseXDTableName(strName As String)
Select Case strName
Case "country"
    If VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
        ParseXDTableName = XD_COUNTRIES
    Else
        ParseXDTableName = XD_NATION
    End If
Case "game"
    ParseXDTableName = XD_GAME
Case "item"
    ParseXDTableName = XD_ITEM_CHR
Case "infrastructure"
    ParseXDTableName = XD_INFRA_CHR
Case "land"
    ParseXDTableName = XD_LAND
Case "land-chr"
    ParseXDTableName = XD_LAND_CHR
Case "level"
    ParseXDTableName = XD_LEVEL
Case "lost"
    ParseXDTableName = XD_LOST
Case "meta"
    ParseXDTableName = XD_META
Case "missions"
    ParseXDTableName = XD_MISSIONS
Case "nat"
    If VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
        ParseXDTableName = XD_NATION
    Else
        ParseXDTableName = XD_COUNTRIES
    End If
Case "nation-relationships"
    ParseXDTableName = XD_NATION_RELATIONSHIPS
Case "nuke"
    ParseXDTableName = XD_NUKE
Case "nuke-chr"
    ParseXDTableName = XD_NUKE_CHR
Case "sect"
    ParseXDTableName = XD_SECTOR
Case "sect-chr"
    ParseXDTableName = XD_SECTOR_CHR
Case "sector-navigation"
    ParseXDTableName = XD_SECTOR_NAVIGATION
Case "ship"
    ParseXDTableName = XD_SHIP
Case "ship-chr"
    ParseXDTableName = XD_SHIP_CHR
Case "ship-chr-flags"
    ParseXDTableName = XD_SHIP_CHR_FLAGS
Case "plane"
    ParseXDTableName = XD_PLANE
Case "plane-chr"
    ParseXDTableName = XD_PLANE_CHR
Case "product"
    ParseXDTableName = XD_PRODUCT_CHR
Case "updates"
    ParseXDTableName = XD_UPDATES
Case "version"
    ParseXDTableName = XD_VERSION
Case "meta-type"
    ParseXDTableName = XD_META_TYPE
Case Else
    ParseXDTableName = XD_UNKNOWN
End Select
End Function

Public Sub RequestMetaTables()
If VersionCheck(4, 3, 0) = VER_LESS Then
    Exit Sub
End If
frmEmpCmd.SubmitEmpireCommand "xdump meta meta", False
frmEmpCmd.SubmitEmpireCommand "xdump meta ver", False
frmEmpCmd.SubmitEmpireCommand "xdump meta item", False
frmEmpCmd.SubmitEmpireCommand "xdump meta infra", False
frmEmpCmd.SubmitEmpireCommand "xdump meta product", False
frmEmpCmd.SubmitEmpireCommand "xdump meta sect-chr", False
frmEmpCmd.SubmitEmpireCommand "xdump meta land-chr", False
frmEmpCmd.SubmitEmpireCommand "xdump meta ship-chr", False
frmEmpCmd.SubmitEmpireCommand "xdump meta nuke-chr", False
frmEmpCmd.SubmitEmpireCommand "xdump meta plane-chr", False
frmEmpCmd.SubmitEmpireCommand "xdump meta sect", False
frmEmpCmd.SubmitEmpireCommand "xdump meta land", False
frmEmpCmd.SubmitEmpireCommand "xdump meta ship", False
frmEmpCmd.SubmitEmpireCommand "xdump meta nuke", False
frmEmpCmd.SubmitEmpireCommand "xdump meta plane", False
frmEmpCmd.SubmitEmpireCommand "xdump meta nat", False
frmEmpCmd.SubmitEmpireCommand "xdump meta lost", False
frmEmpCmd.SubmitEmpireCommand "xdump meta cou", False
frmEmpCmd.SubmitEmpireCommand "xdump meta meta-type", False
frmEmpCmd.SubmitEmpireCommand "xdump meta missions", False
frmEmpCmd.SubmitEmpireCommand "xdump meta level", False
frmEmpCmd.SubmitEmpireCommand "xdump meta sector-navigation", False
frmEmpCmd.SubmitEmpireCommand "xdump meta ship-chr-flags", False
If VersionCheck(4, 3, 10) <> VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "xdump meta updates", False
    frmEmpCmd.SubmitEmpireCommand "xdump meta game", False
End If
frmEmpCmd.SubmitEmpireCommand "xdump meta nation-relationships", False
frmEmpCmd.SubmitEmpireCommand "xdump meta-type *", False
frmEmpCmd.SubmitEmpireCommand "xdump missions *", False
frmEmpCmd.SubmitEmpireCommand "xdump level *", False
frmEmpCmd.SubmitEmpireCommand "xdump sector-navigation *", False
frmEmpCmd.SubmitEmpireCommand "xdump ship-chr-flags *", False
frmEmpCmd.SubmitEmpireCommand "xdump nation-relationships *", False
End Sub

Public Sub RequestConfigurationTables()
frmEmpCmd.SubmitEmpireCommand "xdump ver *", False
GetNation
frmEmpCmd.SubmitEmpireCommand "xdump item *", False
frmEmpCmd.SubmitEmpireCommand "xdump sect-chr *", False
frmEmpCmd.SubmitEmpireCommand "xdump product *", False
frmEmpCmd.SubmitEmpireCommand "xdump infra *", False
RequestUnitConfigurationTables
End Sub

Public Sub RequestUnitConfigurationTables()
frmEmpCmd.SubmitEmpireCommand "xdump land-chr *", False
frmEmpCmd.SubmitEmpireCommand "xdump ship-chr *", False
frmEmpCmd.SubmitEmpireCommand "xdump nuke-chr *", False
frmEmpCmd.SubmitEmpireCommand "xdump plane-chr *", False
End Sub

Private Function GetTechLevelForShow() As Integer
If Len(frmToolShow.txtTech) > 0 Then
    GetTechLevelForShow = CInt(frmToolShow.txtTech)
ElseIf bDeity Then
    GetTechLevelForShow = 1000
Else
    If rsNation.BOF And rsNation.EOF Then
        GetTechLevelForShow = 0
    Else
        rsNation.MoveFirst
        GetTechLevelForShow = CInt(rsNation.Fields("TechLevel"))
    End If
End If
End Function

Private Sub UpdateTechLevelForShow(strUnit As String, strReport As String)
Dim strText As String
    
strText = "Printing for tech level '" + CStr(GetTechLevelForShow) + "'"
ParseShowText strText, strUnit, strReport, 1
End Sub

Private Function GetType(iIndex As Integer) As String
Dim eiItem As EmpItem

If iIndex = -1 Then
    GetType = ""
    Exit Function
End If

Set eiItem = Items.FindByChrIndex(iIndex)

If eiItem Is Nothing Then
    GetType = ""
Else
    GetType = eiItem.Letter
End If
End Function

Private Function GetNationType(lIndex As Long) As String
Select Case Left$(estsTable(XD_LEVEL)(lIndex).Name, 3)
Case "non"
    GetNationType = ""
Case "tec"
    GetNationType = "tech"
Case Else
    GetNationType = Left$(estsTable(XD_LEVEL)(lIndex).Name, 3)
End Select
End Function

Public Function MaxPopulation(Optional strDes As String) As Integer
If Not IsMissing(strDes) Then
    rsSectorType.Seek "=", strDes
ElseIf frmOptions.bSPAtlantis Then
    rsSectorType.Seek "=", "i"
Else
    rsSectorType.Seek "=", "m"
End If
If Not rsSectorType.NoMatch Then
    MaxPopulation = rsSectorType.Fields("maxpop")
Else
    MaxPopulation = 999
End If
End Function

Private Function DirectionString(iValue As Integer) As String
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

Private Sub AddMission(lMission As Long, iID As Integer, iOwner As Integer, strType As String, strSector As String, iRadius As Integer)
If lMission > 0 Then
    If rsMissions.BOF And rsMissions.EOF Then
        rsMissions.AddNew
    Else
        rsMissions.Seek "=", strType, iID
        If Not rsMissions.NoMatch Then
            rsMissions.Edit
        Else
            rsMissions.AddNew
        End If
    End If
    rsMissions.Fields("id") = iID
    rsMissions.Fields("owner") = iOwner
    rsMissions.Fields("type") = strType
    rsMissions.Fields("op sector") = strSector
    rsMissions.Fields("mission") = estsTable(XD_MISSIONS)(lMission).Name
    rsMissions.Fields("radius") = iRadius
    rsMissions.Update
ElseIf Not (rsMissions.BOF And rsMissions.EOF) Then
    rsMissions.Seek "=", strType, iID
    If Not rsMissions.NoMatch Then
        rsMissions.Delete
    End If
End If
End Sub

Private Sub AddOrders(bAutoNav As Boolean, iID As Integer, iOwner As Integer, _
    iLength As Integer, iETA As Integer, _
    strCurSector As String, strStartSector As String, strEndSector As String, _
    strStartCargo() As String, strEndCargo() As String, iStart() As Integer, iEnd() As Integer)
Dim iIndex As Integer

If bAutoNav Then
    If rsShipOrders.BOF And rsShipOrders.EOF Then
        rsShipOrders.AddNew
    Else
        rsShipOrders.Seek "=", iID
        If Not rsShipOrders.NoMatch Then
            rsShipOrders.Edit
        Else
            rsShipOrders.AddNew
        End If
    End If
    rsShipOrders.Fields("id") = iID
    rsShipOrders.Fields("curr sector") = strCurSector
    rsShipOrders.Fields("start sector") = strEndSector
    rsShipOrders.Fields("end sector") = strStartSector
    rsShipOrders.Fields("length") = iLength
    rsShipOrders.Fields("eta") = iETA
    rsShipOrders.Fields("owner") = iOwner
    For iIndex = 1 To 6
        rsShipOrders.Fields("start cargo " + CStr(iIndex)) = strStartCargo(iIndex)
        rsShipOrders.Fields("end cargo " + CStr(iIndex)) = strEndCargo(iIndex)
        rsShipOrders.Fields("start level " + CStr(iIndex)) = iStart(iIndex)
        rsShipOrders.Fields("end level " + CStr(iIndex)) = iEnd(iIndex)
    Next
    rsShipOrders.Update
ElseIf Not (rsShipOrders.BOF And rsShipOrders.EOF) Then
    rsShipOrders.Seek "=", iID
    If Not rsShipOrders.NoMatch Then
        rsShipOrders.Delete
    End If
End If
If frmDrawMap.PromptUp And frmDrawMap.PromptForm Is frmPromptOrder Then
    frmPromptOrder.OrderLoadData
End If
End Sub

Private Function CalculateAvail(iLCM As Integer, iHCM As Integer) As Integer
CalculateAvail = 20 + iLCM + (2 * iHCM)
End Function

'HeavyMetal modification
Private Sub UpdateRailTrack()
Dim xa, ya, xas, yas As Variant
Dim iIndex As Integer
Dim strSect() As String
Dim iCount As Integer
Dim vSect As Variant
Dim iSecX As Integer
Dim iSecY As Integer

xa = Array(-2, -1, 1, 2, 1, -1)
ya = Array(0, -1, -1, 0, 1, 1)
xas = Array(1, 2, 1, -1, -2, -1)
yas = Array(-1, 0, 1, 1, 0, -1)

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
rsSectors.MoveFirst
While Not rsSectors.EOF
    rsSectorType.Seek "=", rsSectors.Fields("des")
    If rsSectorType.NoMatch Then
        Exit Sub
    End If
    rsSectors.Edit
    If rsSectorType.Fields("mcost100") = 0 And rsSectors.Fields("eff") >= 5 Then
        rsSectors.Fields("rail") = 1
        iCount = iCount + 1
        ReDim Preserve strSect(1 To iCount)
        strSect(iCount) = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
    Else
        rsSectors.Fields("rail") = 0
    End If
    rsSectors.Update
    rsSectors.MoveNext
Wend
If iCount = 0 Then
    Exit Sub
End If

For Each vSect In strSect
    For iIndex = 0 To 5
        If ParseSectors(iSecX, iSecY, CStr(vSect)) Then
            rsSectors.Seek "=", iSecX + CInt(xa(iIndex)), iSecY + CInt(ya(iIndex))
            If Not rsSectors.NoMatch Then
                rsSectorType.Seek "=", rsSectors.Fields("des")
                If rsSectorType.NoMatch Then
                    Exit Sub
                End If
                If rsSectors.Fields("eff") >= 60 Then
                    rsSectors.Edit
                    rsSectors.Fields("rail") = rsSectors.Fields("rail") + 1
                    rsSectors.Update
                End If
            End If
        End If
    Next iIndex
Next

End Sub
