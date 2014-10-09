Attribute VB_Name = "EmpDataBase"
Option Explicit

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
Public dbsName As String            'Hold empire database name
Public GameName As String
Public dbEmpire As Database
Public rsSectors As Recordset
Public rsBmap As Recordset
Public rsShip As Recordset
Public rsShipList As Recordset
Public rsLandList As Recordset
Public rsPlaneList As Recordset
Public rsLand As Recordset
Public rsPlane As Recordset
Public rsNuke As Recordset
Public rsTeleHead As Recordset
Public rsTeleBody As Recordset
Public rsEnemySect As Recordset
Public rsEnemyUnit As Recordset
Public rsVersion As Recordset
Public rsNation As Recordset
Public rsNations As Recordset
Public rsBuild As Recordset
Public rsSea As Recordset
Public rsAirCombat As Recordset
Public rsShipOrders As Recordset
Public rsMissions As Recordset
Public rsSectorType As Recordset
Public rsShowText As Recordset
Public rsItems As Recordset '120703 rjk: Added to support different commodities (theme games)

Public IsDBOpen As Boolean

'Nation Variables
Public Education As Single
Public TechLevel As Single
Public Happiness As Single
Public Research As Single
Public CapSect As String
Public Maxpop As Long
Public MaxSafeCivs As Long
Public MaxSafeUws As Long
Public HapNeeded As Single
Public MilReserves As Long
Public HarborAvail As Boolean
Public AirportAvail As Boolean
Public FortAvail As Boolean
Public HQAvail As Boolean

Private Sub CheckDBFields(dbEmpire As Database)
Dim idxNew As Index
Dim iLoop As Integer

'270304 rjk: Added 'Rollover Avail' to Version table
If dbEmpire.TableDefs("Version Info").Fields.Count = 39 Then
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("Rollover Avail", dbInteger)
    dbEmpire.TableDefs("Version Info").Fields("Rollover Avail").DefaultValue = 0
    If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
        Set rsVersion = dbEmpire.OpenRecordset("Version Info")
        rsVersion.MoveFirst
        While Not rsVersion.EOF
            rsVersion.Edit
            rsVersion.Fields("Rollover Avail") = 0
            rsVersion.Update
            rsVersion.MoveNext
        Wend
    End If
End If
'220305 rjk: Added 'Version' to Version table
If dbEmpire.TableDefs("Version Info").Fields.Count = 40 Then
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("Major", dbInteger)
    dbEmpire.TableDefs("Version Info").Fields("Major").DefaultValue = 0
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("Minor", dbInteger)
    dbEmpire.TableDefs("Version Info").Fields("Minor").DefaultValue = 0
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("Revision", dbInteger)
    dbEmpire.TableDefs("Version Info").Fields("Revision").DefaultValue = 0
    If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
        Set rsVersion = dbEmpire.OpenRecordset("Version Info")
        rsVersion.MoveFirst
        While Not rsVersion.EOF
            rsVersion.Edit
            rsVersion.Fields("Major") = 0
            rsVersion.Fields("Minor") = 0
            rsVersion.Fields("Revision") = 0
            rsVersion.Update
            rsVersion.MoveNext
        Wend
    End If
End If
'181205 rjk: Added 'Eff gain - Sect' to Version table
If dbEmpire.TableDefs("Version Info").Fields.Count = 43 Then
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("Eff gain - Sect", dbInteger)
    dbEmpire.TableDefs("Version Info").Fields("Eff gain - Sect").DefaultValue = 100
    If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
        Set rsVersion = dbEmpire.OpenRecordset("Version Info")
        rsVersion.MoveFirst
        While Not rsVersion.EOF
            rsVersion.Edit
            rsVersion.Fields("Eff gain - Sect") = 100
            rsVersion.Update
            rsVersion.MoveNext
        Wend
    End If
End If
'110206 rjk: Added 'drnuke_const' to Version table
If dbEmpire.TableDefs("Version Info").Fields.Count = 44 Then
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("DRNUKE Const", dbSingle)
    dbEmpire.TableDefs("Version Info").Fields("DRNUKE Const").DefaultValue = 0#
    If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
        Set rsVersion = dbEmpire.OpenRecordset("Version Info")
        rsVersion.MoveFirst
        While Not rsVersion.EOF
            rsVersion.Edit
            rsVersion.Fields("DRNUKE Const") = 0#
            rsVersion.Update
            rsVersion.MoveNext
        Wend
    End If
End If
'020907 rjk: Added 'timezoneadj' to Version table - difference between local time
'            and server time
If dbEmpire.TableDefs("Version Info").Fields.Count = 45 Then
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("Time Zone Adj", dbLong)
    dbEmpire.TableDefs("Version Info").Fields("Time Zone Adj").DefaultValue = 0
    If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
        Set rsVersion = dbEmpire.OpenRecordset("Version Info")
        rsVersion.MoveFirst
        While Not rsVersion.EOF
            rsVersion.Edit
            rsVersion.Fields("Time Zone Adj") = 0
            rsVersion.Update
            rsVersion.MoveNext
        Wend
    End If
End If
'160209 rjk: Added RAILWAYS options, originally part of the Heavy Metal/Plastic
If dbEmpire.TableDefs("Version Info").Fields.Count = 46 Then
    dbEmpire.TableDefs("Version Info").Fields.Append dbEmpire.TableDefs("Version Info").CreateField("RAILWAYS", dbBoolean)
    dbEmpire.TableDefs("Version Info").Fields("RAILWAYS").DefaultValue = False
    If dbEmpire.TableDefs("Version Info").RecordCount > 0 Then
        Set rsVersion = dbEmpire.OpenRecordset("Version Info")
        rsVersion.MoveFirst
        While Not rsVersion.EOF
            rsVersion.Edit
            rsVersion.Fields("RAILWAYS") = False
            rsVersion.Update
            rsVersion.MoveNext
        Wend
    End If
End If
'270804 rjk: Added Index to Sector Type table
If dbEmpire.TableDefs("Sector Type").Fields.Count = 25 Then
    dbEmpire.TableDefs("Sector Type").Fields.Append dbEmpire.TableDefs("Sector Type").CreateField("dchr_i", dbInteger)
    dbEmpire.TableDefs("Sector Type").Fields("dchr_i").DefaultValue = 0
    If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
        Set rsSectorType = dbEmpire.OpenRecordset("Sector Type")
        rsSectorType.MoveFirst
        While Not rsSectorType.EOF
            rsSectorType.Edit
            rsSectorType.Fields("dchr_i") = 0
            rsSectorType.Update
            rsSectorType.MoveNext
        Wend
    End If
End If
'100904 rjk: Added Product Index to Sector Type table
If dbEmpire.TableDefs("Sector Type").Fields.Count = 26 Then
    dbEmpire.TableDefs("Sector Type").Fields.Append dbEmpire.TableDefs("Sector Type").CreateField("pchr_i", dbInteger)
    dbEmpire.TableDefs("Sector Type").Fields("pchr_i").DefaultValue = 0
    If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
        Set rsSectorType = dbEmpire.OpenRecordset("Sector Type")
        rsSectorType.MoveFirst
        While Not rsSectorType.EOF
            rsSectorType.Edit
            rsSectorType.Fields("pchr_i") = 0
            rsSectorType.Update
            rsSectorType.MoveNext
        Wend
    End If
End If

'180606 rjk: Added mcost at 100% to the Sector Type table for 4.3.6 servers
If dbEmpire.TableDefs("Sector Type").Fields.Count = 27 Then
    dbEmpire.TableDefs("Sector Type").Fields.Append dbEmpire.TableDefs("Sector Type").CreateField("mcost100", dbSingle)
    dbEmpire.TableDefs("Sector Type").Fields("mcost100").DefaultValue = -2#
    If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
        Set rsSectorType = dbEmpire.OpenRecordset("Sector Type")
        rsSectorType.MoveFirst
        While Not rsSectorType.EOF
            rsSectorType.Edit
            rsSectorType.Fields("mcost100") = -2#
            rsSectorType.Update
            rsSectorType.MoveNext
        Wend
    End If
End If

'190606 rjk: Added nav
If dbEmpire.TableDefs("Sector Type").Fields.Count = 28 Then
    dbEmpire.TableDefs("Sector Type").Fields.Append dbEmpire.TableDefs("Sector Type").CreateField("nav", dbInteger)
    dbEmpire.TableDefs("Sector Type").Fields("nav").DefaultValue = 0
    If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
        Set rsSectorType = dbEmpire.OpenRecordset("Sector Type")
        rsSectorType.MoveFirst
        While Not rsSectorType.EOF
            rsSectorType.Edit
            rsSectorType.Fields("nav") = 0
            rsSectorType.Update
            rsSectorType.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("Sector Type").Fields.Count = 29 Then
    dbEmpire.TableDefs("Sector Type").Fields.Append dbEmpire.TableDefs("Sector Type").CreateField("pack_type", dbInteger)
    dbEmpire.TableDefs("Sector Type").Fields("pack_type").DefaultValue = 1
    If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
        Set rsSectorType = dbEmpire.OpenRecordset("Sector Type")
        rsSectorType.MoveFirst
        While Not rsSectorType.EOF
            rsSectorType.Edit
            rsSectorType.Fields("pack_type") = 1
            rsSectorType.Update
            rsSectorType.MoveNext
        Wend
    End If
End If
'160209 rjk: Added Terrain to the sector type table
If dbEmpire.TableDefs("Sector Type").Fields.Count = 30 Then
    dbEmpire.TableDefs("Sector Type").Fields.Append dbEmpire.TableDefs("Sector Type").CreateField("terrain", dbInteger)
    dbEmpire.TableDefs("Sector Type").Fields("terrain").DefaultValue = 0
    If dbEmpire.TableDefs("Sector Type").RecordCount > 0 Then
        Set rsSectorType = dbEmpire.OpenRecordset("Sector Type")
        rsSectorType.MoveFirst
        While Not rsSectorType.EOF
            rsSectorType.Edit
            rsSectorType.Fields("terrain") = 0
            rsSectorType.Update
            rsSectorType.MoveNext
        Wend
    End If
End If

'280804 rjk: Added Index to Build Info table
If dbEmpire.TableDefs("Build Info").Fields.Count = 52 Then
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("chr_i", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("chr_i").DefaultValue = 0
    If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
        Set rsBuild = dbEmpire.OpenRecordset("Build Info")
        rsBuild.MoveFirst
        While Not rsBuild.EOF
            rsBuild.Edit
            rsBuild.Fields("chr_i") = 0
            rsBuild.Update
            rsBuild.MoveNext
        Wend
    End If
End If
'120306 rjk: Added Base Speed and Base Firing Range to Build Info table
If dbEmpire.TableDefs("Build Info").Fields.Count = 53 Then
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base spd", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base spd").DefaultValue = 0
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base frnge", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base frnge").DefaultValue = 0
    If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
        Set rsBuild = dbEmpire.OpenRecordset("Build Info")
        rsBuild.MoveFirst
        While Not rsBuild.EOF
            rsBuild.Edit
            rsBuild.Fields("base spd") = 0
            rsBuild.Fields("base frnge") = 0
            rsBuild.Update
            rsBuild.MoveNext
        Wend
    End If
End If
'230306 rjk: Added Base Attack and Base Defense to Build Info table
If dbEmpire.TableDefs("Build Info").Fields.Count = 55 Then
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base att", dbSingle)
    dbEmpire.TableDefs("Build Info").Fields("base att").DefaultValue = 0#
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base def", dbSingle)
    dbEmpire.TableDefs("Build Info").Fields("base def").DefaultValue = 0#
    If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
        Set rsBuild = dbEmpire.OpenRecordset("Build Info")
        rsBuild.MoveFirst
        While Not rsBuild.EOF
            rsBuild.Edit
            rsBuild.Fields("base att") = 0#
            rsBuild.Fields("base def") = 0#
            rsBuild.Update
            rsBuild.MoveNext
        Wend
    End If
End If
'210406 rjk: Added Base Accuracy to Build Info table
If dbEmpire.TableDefs("Build Info").Fields.Count = 57 Then
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base acc", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base acc").DefaultValue = 0#
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base load", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base load").DefaultValue = 0#
    If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
        Set rsBuild = dbEmpire.OpenRecordset("Build Info")
        rsBuild.MoveFirst
        While Not rsBuild.EOF
            rsBuild.Edit
            rsBuild.Fields("base acc") = 0#
            rsBuild.Fields("base load") = 0#
            rsBuild.Update
            rsBuild.MoveNext
        Wend
    End If
End If
'220406 rjk: Added Base Accuracy to Build Info table
If dbEmpire.TableDefs("Build Info").Fields.Count = 59 Then
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base vul", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base vul").DefaultValue = 0#
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base dam", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base dam").DefaultValue = 0#
    dbEmpire.TableDefs("Build Info").Fields.Append dbEmpire.TableDefs("Build Info").CreateField("base aaf", dbInteger)
    dbEmpire.TableDefs("Build Info").Fields("base aaf").DefaultValue = 0#
    If dbEmpire.TableDefs("Build Info").RecordCount > 0 Then
        Set rsBuild = dbEmpire.OpenRecordset("Build Info")
        rsBuild.MoveFirst
        While Not rsBuild.EOF
            rsBuild.Edit
            rsBuild.Fields("base vul") = 0#
            rsBuild.Fields("base dam") = 0#
            rsBuild.Fields("base aaf") = 0#
            rsBuild.Update
            rsBuild.MoveNext
        Wend
    End If
End If
'300804 rjk: Add Index for Title to the Telegram Header table
If dbEmpire.TableDefs("Telegram Header").Indexes.Count = 3 Then
    Set idxNew = dbEmpire.TableDefs("Telegram Header").CreateIndex("Title")
    With idxNew
        .Primary = False
        .Unique = False
        .Fields.Append .CreateField("Title")
    End With

    dbEmpire.TableDefs("Telegram Header").Indexes.Append idxNew
    dbEmpire.TableDefs("Telegram Header").Indexes.Refresh
End If
'050904 rjk: Added Index to Items table
If dbEmpire.TableDefs("Items").Fields.Count = 17 Then
    dbEmpire.TableDefs("Items").Fields.Append dbEmpire.TableDefs("Items").CreateField("pack_ie", dbInteger)
    dbEmpire.TableDefs("Items").Fields("pack_ie").DefaultValue = 1
    dbEmpire.TableDefs("Items").Fields.Append dbEmpire.TableDefs("Items").CreateField("chr_i", dbInteger)
    dbEmpire.TableDefs("Items").Fields("chr_i").DefaultValue = 0
    If dbEmpire.TableDefs("Items").RecordCount > 0 Then
        Set rsItems = dbEmpire.OpenRecordset("Items")
        rsItems.MoveFirst
        While Not rsItems.EOF
            rsItems.Edit
            rsItems.Fields("pack_ie") = 1
            rsItems.Fields("chr_i") = 0
            rsItems.Update
            rsItems.MoveNext
        Wend
    End If
End If
'110605 rjk: Added location Index to Mission table for reacting planes.
If dbEmpire.TableDefs("Missions").Indexes.Count = 1 Then
    Set idxNew = dbEmpire.TableDefs("Missions").CreateIndex("Mission")
    With idxNew
        .Primary = False
        .Unique = False
        .Fields.Append .CreateField("type")
        .Fields.Append .CreateField("mission")
    End With

    dbEmpire.TableDefs("Missions").Indexes.Append idxNew
    dbEmpire.TableDefs("Missions").Indexes.Refresh
End If
'170206 rjk: Added Index for Country Number in Nations table
If dbEmpire.TableDefs("Nations").Indexes.Count = 0 Then
    Set idxNew = dbEmpire.TableDefs("Nations").CreateIndex("PrimaryKey")
    With idxNew
        .Primary = True
        .Unique = True
        .Fields.Append .CreateField("Number")
    End With

    dbEmpire.TableDefs("Nations").Indexes.Append idxNew
    dbEmpire.TableDefs("Nations").Indexes.Refresh
End If

'280506 rjk: Add timestamps to sector, ship, land, plane and nuke tables
If dbEmpire.TableDefs("ship").Fields.Count = 36 Then
    dbEmpire.TableDefs("ship").Fields.Append dbEmpire.TableDefs("ship").CreateField("timestamp", dbLong)
    dbEmpire.TableDefs("ship").Fields("timestamp").DefaultValue = 0
    If dbEmpire.TableDefs("ship").RecordCount > 0 Then
        Set rsShip = dbEmpire.OpenRecordset("ship")
        rsShip.MoveFirst
        While Not rsShip.EOF
            rsShip.Edit
            rsShip.Fields("timestamp") = 0
            rsShip.Update
            rsShip.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("land").Fields.Count = 43 Then
    dbEmpire.TableDefs("land").Fields.Append dbEmpire.TableDefs("land").CreateField("timestamp", dbLong)
    dbEmpire.TableDefs("land").Fields("timestamp").DefaultValue = 0
    If dbEmpire.TableDefs("land").RecordCount > 0 Then
        Set rsLand = dbEmpire.OpenRecordset("land")
        rsLand.MoveFirst
        While Not rsLand.EOF
            rsLand.Edit
            rsLand.Fields("timestamp") = 0
            rsLand.Update
            rsLand.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("plane").Fields.Count = 23 Then
    dbEmpire.TableDefs("plane").Fields.Append dbEmpire.TableDefs("plane").CreateField("timestamp", dbLong)
    dbEmpire.TableDefs("plane").Fields("timestamp").DefaultValue = 0
    If dbEmpire.TableDefs("plane").RecordCount > 0 Then
        Set rsPlane = dbEmpire.OpenRecordset("plane")
        rsPlane.MoveFirst
        While Not rsPlane.EOF
            rsPlane.Edit
            rsPlane.Fields("timestamp") = 0
            rsPlane.Update
            rsPlane.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("nuke").Fields.Count = 6 Then
    dbEmpire.TableDefs("nuke").Fields.Append dbEmpire.TableDefs("nuke").CreateField("timestamp", dbLong)
    dbEmpire.TableDefs("nuke").Fields("timestamp").DefaultValue = 0
    If dbEmpire.TableDefs("nuke").RecordCount > 0 Then
        Set rsNuke = dbEmpire.OpenRecordset("nuke")
        rsNuke.MoveFirst
        While Not rsNuke.EOF
            rsNuke.Edit
            rsNuke.Fields("timestamp") = 0
            rsNuke.Update
            rsNuke.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("Sectors").Fields.Count = 89 Then
    dbEmpire.TableDefs("Sectors").Fields.Append dbEmpire.TableDefs("Sectors").CreateField("timestamp", dbLong)
    dbEmpire.TableDefs("Sectors").Fields("timestamp").DefaultValue = 0
    If dbEmpire.TableDefs("Sectors").RecordCount > 0 Then
        Set rsSectors = dbEmpire.OpenRecordset("Sectors")
        rsSectors.MoveFirst
        While Not rsSectors.EOF
            rsSectors.Edit
            rsSectors.Fields("timestamp") = 0
            rsSectors.Update
            rsSectors.MoveNext
        Wend
    End If
End If

'250606 rjk: Add off to ship, land, plane and nuke tables
If dbEmpire.TableDefs("ship").Fields.Count = 37 Then
    dbEmpire.TableDefs("ship").Fields.Append dbEmpire.TableDefs("ship").CreateField("off", dbBoolean)
    dbEmpire.TableDefs("ship").Fields("off").DefaultValue = 0
    If dbEmpire.TableDefs("ship").RecordCount > 0 Then
        Set rsShip = dbEmpire.OpenRecordset("ship")
        rsShip.MoveFirst
        While Not rsShip.EOF
            rsShip.Edit
            rsShip.Fields("off") = 0
            rsShip.Update
            rsShip.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("land").Fields.Count = 44 Then
    dbEmpire.TableDefs("land").Fields.Append dbEmpire.TableDefs("land").CreateField("off", dbBoolean)
    dbEmpire.TableDefs("land").Fields("off").DefaultValue = 0
    If dbEmpire.TableDefs("land").RecordCount > 0 Then
        Set rsLand = dbEmpire.OpenRecordset("land")
        rsLand.MoveFirst
        While Not rsLand.EOF
            rsLand.Edit
            rsLand.Fields("off") = 0
            rsLand.Update
            rsLand.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("plane").Fields.Count = 24 Then
    dbEmpire.TableDefs("plane").Fields.Append dbEmpire.TableDefs("plane").CreateField("off", dbBoolean)
    dbEmpire.TableDefs("plane").Fields("off").DefaultValue = 0
    If dbEmpire.TableDefs("plane").RecordCount > 0 Then
        Set rsPlane = dbEmpire.OpenRecordset("plane")
        rsPlane.MoveFirst
        While Not rsPlane.EOF
            rsPlane.Edit
            rsPlane.Fields("off") = 0
            rsPlane.Update
            rsPlane.MoveNext
        Wend
    End If
End If
If dbEmpire.TableDefs("nuke").Fields.Count = 7 Then
    dbEmpire.TableDefs("nuke").Fields.Append dbEmpire.TableDefs("nuke").CreateField("off", dbBoolean)
    dbEmpire.TableDefs("nuke").Fields("off").DefaultValue = 0
    If dbEmpire.TableDefs("nuke").RecordCount > 0 Then
        Set rsNuke = dbEmpire.OpenRecordset("nuke")
        rsNuke.MoveFirst
        While Not rsNuke.EOF
            rsNuke.Edit
            rsNuke.Fields("off") = 0
            rsNuke.Update
            rsNuke.MoveNext
        Wend
    End If
End If

'180606 rjk: Add id index for nuke Table if it is missing
If dbEmpire.TableDefs("nuke").Indexes.Count = 2 Then
    Set idxNew = dbEmpire.TableDefs("nuke").CreateIndex("id")
    With idxNew
        .Primary = False
        .Unique = False
        .Fields.Append .CreateField("id")
    End With

    dbEmpire.TableDefs("nuke").Indexes.Append idxNew
    dbEmpire.TableDefs("nuke").Indexes.Refresh
End If

'230806 rjk: Added timestamps for oil and fert.
If dbEmpire.TableDefs("Sea sectors").Fields.Count = 4 Then
    dbEmpire.TableDefs("Sea sectors").Fields.Append dbEmpire.TableDefs("Sea sectors").CreateField("ftimestamp", dbDate)
    dbEmpire.TableDefs("Sea sectors").Fields("ftimestamp").DefaultValue = vbNullString
    dbEmpire.TableDefs("Sea sectors").Fields.Append dbEmpire.TableDefs("Sea sectors").CreateField("otimestamp", dbDate)
    dbEmpire.TableDefs("Sea sectors").Fields("otimestamp").DefaultValue = vbNullString
    If dbEmpire.TableDefs("Sea sectors").RecordCount > 0 Then
        Set rsSea = dbEmpire.OpenRecordset("Sea sectors")
        rsSea.MoveFirst
        While Not rsSea.EOF
            rsSea.Edit
            rsSea.Fields("ftimestamp") = GetWinACEUTC
            rsSea.Fields("otimestamp") = GetWinACEUTC
            rsSea.Update
            rsSea.MoveNext
        Wend
    End If
End If

'230806 rjk: Added end cargo for orders table.
If dbEmpire.TableDefs("Ship orders").Fields.Count = 25 Then
    For iLoop = 1 To 6
        dbEmpire.TableDefs("Ship orders").Fields.Append dbEmpire.TableDefs("Ship orders").CreateField("end cargo " & CStr(iLoop), dbText, 50)
        dbEmpire.TableDefs("Ship orders").Fields("end cargo " & CStr(iLoop)).Properties.Append dbEmpire.TableDefs("Ship orders").CreateProperty("UnicodeCompression", dbBoolean, True)
        dbEmpire.TableDefs("Ship orders").Fields("end cargo " & CStr(iLoop)).Properties("AllowZeroLength").Value = True
        dbEmpire.TableDefs("Ship orders").Fields("end cargo " & CStr(iLoop)).DefaultValue = vbNullString
        dbEmpire.TableDefs("Ship orders").Fields("cargo " & CStr(iLoop)).Name = "start cargo " & CStr(iLoop)
'End If

        If dbEmpire.TableDefs("Ship orders").RecordCount > 0 Then
            Set rsShipOrders = dbEmpire.OpenRecordset("Ship orders")
            rsShipOrders.MoveFirst
            While Not rsShipOrders.EOF
                rsShipOrders.Edit
                rsShipOrders.Fields("end cargo " & CStr(iLoop)) = rsShipOrders.Fields("start cargo " & CStr(iLoop))
                rsShipOrders.Update
                rsShipOrders.MoveNext
            Wend
            rsShipOrders.Close
            Set rsShipOrders = Nothing
        End If
    Next
End If

'030108 rjk: Added elevation to the sector table
If dbEmpire.TableDefs("Sectors").Fields.Count = 90 Then
    dbEmpire.TableDefs("Sectors").Fields.Append dbEmpire.TableDefs("Sectors").CreateField("elev", dbInteger)
    dbEmpire.TableDefs("Sectors").Fields("elev").DefaultValue = 0
    If dbEmpire.TableDefs("Sectors").RecordCount > 0 Then
        Set rsSectors = dbEmpire.OpenRecordset("Sectors")
        rsSectors.MoveFirst
        While Not rsSectors.EOF
            rsSectors.Edit
            rsSectors.Fields("elev") = 0
            rsSectors.Update
            rsSectors.MoveNext
        Wend
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
    rsEnemySect.MoveFirst
    While Not rsEnemySect.EOF
        If Len(rsEnemySect.Fields("timestamp")) > 0 Then
            If Mid$(rsEnemySect.Fields("timestamp"), 5, 1) <> "/" Then
                rsEnemySect.Edit
                rsEnemySect.Fields("timestamp") = Format(DateAdd("n", tz.Bias, EmpireTimeString(rsEnemySect.Fields("timestamp"))), "yyyy/mm/dd hh:mm:ss")
                rsEnemySect.Update
            End If
        End If
        rsEnemySect.MoveNext
    Wend
End If

If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
    rsEnemyUnit.MoveFirst
    While Not rsEnemyUnit.EOF
        If Len(rsEnemyUnit.Fields("timestamp")) > 0 Then
            If Mid$(rsEnemyUnit.Fields("timestamp"), 5, 1) <> "/" Then
                rsEnemyUnit.Edit
                rsEnemyUnit.Fields("timestamp") = Format(DateAdd("n", tz.Bias, EmpireTimeString(rsEnemyUnit.Fields("timestamp"))), "yyyy/mm/dd hh:mm:ss")
                rsEnemyUnit.Update
            End If
        End If
        rsEnemyUnit.MoveNext
    Wend
End If
End Sub

Public Function OpenEmpireDB() As Boolean
'Set Database name
dbsName = App.path + "\Games\" + GameName + ".mdb"

'Set data base files
Set dbEmpire = OpenDatabase(dbsName, False, False)
CheckDBFields dbEmpire
Set rsSectors = dbEmpire.OpenRecordset("Sectors")
Set rsBmap = dbEmpire.OpenRecordset("Bmap")
Set rsShip = dbEmpire.OpenRecordset("ship")
Set rsShipList = dbEmpire.OpenRecordset("shiplist")
Set rsLand = dbEmpire.OpenRecordset("land")
Set rsNuke = dbEmpire.OpenRecordset("nuke")
Set rsLandList = dbEmpire.OpenRecordset("landlist")
Set rsPlaneList = dbEmpire.OpenRecordset("planelist")
Set rsPlane = dbEmpire.OpenRecordset("plane")
Set rsTeleHead = dbEmpire.OpenRecordset("Telegram Header")
Set rsTeleBody = dbEmpire.OpenRecordset("Telegram Body")
Set rsEnemySect = dbEmpire.OpenRecordset("Enemy Sect")
Set rsEnemyUnit = dbEmpire.OpenRecordset("Enemy Units")
UpdateEnemyTimestamp
Set rsVersion = dbEmpire.OpenRecordset("Version Info")
Set rsBuild = dbEmpire.OpenRecordset("Build Info")
Set rsSea = dbEmpire.OpenRecordset("Sea sectors")
Set rsAirCombat = dbEmpire.OpenRecordset("Air Combat")
Set rsShipOrders = dbEmpire.OpenRecordset("Ship orders")
Set rsMissions = dbEmpire.OpenRecordset("Missions")
Set rsSectorType = dbEmpire.OpenRecordset("Sector type")
Set rsShowText = dbEmpire.OpenRecordset("Show Text")
Set rsNations = dbEmpire.OpenRecordset("Nations")
Set rsNation = dbEmpire.OpenRecordset("Nation")
Set rsItems = dbEmpire.OpenRecordset("Items")
Set Items = New EmpItems

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
GetNationVariables
End Function

Sub CloseEmpireDB()

If IsDBOpen Then
    'Set data base files
    Set rsSectors = Nothing
    Set rsBmap = Nothing
    Set rsShip = Nothing
    Set rsLand = Nothing
    Set rsNuke = Nothing
    Set rsShipList = Nothing
    Set rsLandList = Nothing
    Set rsPlane = Nothing
    Set rsTeleHead = Nothing
    Set rsTeleBody = Nothing
    Set rsEnemySect = Nothing
    Set rsEnemyUnit = Nothing
    Set rsVersion = Nothing
    Set rsNations = Nothing
    Set rsNation = Nothing
    Set rsBuild = Nothing
    Set rsSea = Nothing
    Set rsAirCombat = Nothing
    Set rsShipOrders = Nothing
    Set rsMissions = Nothing
    
    dbEmpire.Close
    IsDBOpen = False
End If

End Sub

Public Sub ClearUnitsFiles()
DeleteAllRecords rsPlane
DeleteAllRecords rsLand
DeleteAllRecords rsShip
End Sub

Public Sub ResetDataBase()

On Error Resume Next        'MoveFirst will give error on empty dataset
'Clear current sector File
DeleteAllRecords rsSectors

'Clear Plane, ship, and land file
ClearUnitsFiles

DeleteAllRecords rsNuke

'Clear Bmap file
'don't clear the bmap file if we are suppressing the bmap update
'120303 rjk: Switched options to frmOptions and boolean options
If frmOptions.bSuppressBmapRefresh Then
    Exit Sub
End If

DeleteAllRecords rsBmap
End Sub

Public Sub DeleteAllRecords(rsDelete As Recordset)
With rsDelete
   If .BOF And .EOF Then Exit Sub

   .MoveFirst
   Do While Not .EOF
      .Delete
      .MoveNext
   Loop
End With
End Sub

Public Sub DeleteAllRecordsOlderThen(rsDelete As Recordset, dExpiry As Date)
Dim dRecord As Date

With rsDelete
    If .BOF And .EOF Then Exit Sub
    
    .MoveFirst
    Do While Not .EOF
        dRecord = ConvertToLocalTimeFromWinACEUTC(.Fields("timestamp").Value)
        If DateDiff("d", dExpiry, dRecord) <= 0 Then
            .Delete
        End If
        .MoveNext
    Loop
End With
End Sub

Public Sub DeleteAllAirCombatRecordsOlderThen(dExpiry As Date)
Dim dRecord As Date

With rsAirCombat
    If .BOF And .EOF Then Exit Sub
    
    .MoveFirst
    Do While Not .EOF
        dRecord = ConvertToLocalTimeFromWinACEUTC(.Fields("nextupdate").Value)
        If DateDiff("d", dExpiry, dRecord) <= 0 Then
            .Delete
        End If
        .MoveNext
    Loop
End With
End Sub

Public Sub DeleteBuildRecords(strType As String)
With rsBuild
    If .BOF And .EOF Then Exit Sub

    .MoveFirst
    Do While Not .EOF
        If .Fields("type") = strType Then
            .Delete
        End If
        .MoveNext
    Loop
End With
End Sub

Public Sub ShiftOrigin(sx As Integer, sy As Integer)

Dim hldIndex As Variant
Dim holdx As Integer
Dim holdy As Integer
Dim n As Integer

On Error Resume Next        'MoveFirst will give error on empty dataset

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
hldIndex = rsSectors.Index
rsSectors.Index = "Primary Key"
rsSectors.MoveFirst
While Not rsSectors.EOF
    If rsSectors.Fields("x") < 1000 Then
        rsSectors.Edit
        rsSectors.Fields("x") = rsSectors.Fields("x") - sx + 2000
        rsSectors.Fields("y") = rsSectors.Fields("y") - sy + 2000
        holdx = rsSectors.Fields("dist_x") - sx
        holdy = rsSectors.Fields("dist_y") - sy
        OffsetSectorLocation holdx, holdy, 0, 0
        rsSectors.Fields("dist_x") = holdx
        rsSectors.Fields("dist_y") = holdy
        rsSectors.Update
    End If
    rsSectors.MoveNext
Wend

'Second Pass
rsSectors.MoveFirst
While Not rsSectors.EOF
    If rsSectors.Fields("x") > 1000 Then
        rsSectors.Edit
        holdx = rsSectors.Fields("x") - 2000
        holdy = rsSectors.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsSectors.Fields("x") = holdx
        rsSectors.Fields("y") = holdy
        rsSectors.Update
    End If
    rsSectors.MoveNext
Wend
rsSectors.Index = hldIndex

'Update Sea Information file
hldIndex = rsSea.Index
rsSea.Index = "PrimaryKey"
rsSea.MoveFirst
While Not rsSea.EOF
    If rsSea.Fields("x") < 1000 Then
        rsSea.Edit
        rsSea.Fields("x") = rsSea.Fields("x") - sx + 2000
        rsSea.Fields("y") = rsSea.Fields("y") - sy + 2000
        rsSea.Update
    End If
    rsSea.MoveNext
Wend

'Second Pass
rsSea.MoveFirst
While Not rsSea.EOF
    If rsSea.Fields("x") > 1000 Then
        rsSea.Edit
        holdx = rsSea.Fields("x") - 2000
        holdy = rsSea.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsSea.Fields("x") = holdx
        rsSea.Fields("y") = holdy
        rsSea.Update
    End If
    rsSea.MoveNext
Wend
rsSea.Index = hldIndex

'Update Bmap file
hldIndex = rsBmap.Index
rsBmap.Index = "Primary Key"
rsBmap.MoveFirst
While Not rsBmap.EOF
    If rsBmap.Fields("x") < 1000 Then
        rsBmap.Edit
        rsBmap.Fields("x") = rsBmap.Fields("x") - sx + 2000
        rsBmap.Fields("y") = rsBmap.Fields("y") - sy + 2000
        rsBmap.Update
    End If
    rsBmap.MoveNext
Wend

'Second Pass
rsBmap.MoveFirst
While Not rsBmap.EOF
    If rsBmap.Fields("x") > 1000 Then
        rsBmap.Edit
        holdx = rsBmap.Fields("x") - 2000
        holdy = rsBmap.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsBmap.Fields("x") = holdx
        rsBmap.Fields("y") = holdy
        rsBmap.Update
    End If
    rsBmap.MoveNext
Wend
rsBmap.Index = hldIndex


'Update Enemy Sector file
hldIndex = rsEnemySect.Index
rsEnemySect.Index = "Primary Key"
rsEnemySect.MoveFirst
While Not rsEnemySect.EOF
    If rsEnemySect.Fields("x") < 1000 Then
        rsEnemySect.Edit
        rsEnemySect.Fields("x") = rsEnemySect.Fields("x") - sx + 2000
        rsEnemySect.Fields("y") = rsEnemySect.Fields("y") - sy + 2000
        rsEnemySect.Update
    End If
    rsEnemySect.MoveNext
Wend

rsEnemySect.MoveFirst
While Not rsEnemySect.EOF
    If rsEnemySect.Fields("x") > 1000 Then
        rsEnemySect.Edit
        holdx = rsEnemySect.Fields("x") - 2000
        holdy = rsEnemySect.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsEnemySect.Fields("x") = holdx
        rsEnemySect.Fields("y") = holdy
        rsEnemySect.Update
    End If
    rsEnemySect.MoveNext
Wend
rsEnemySect.Index = hldIndex

'Update Plane file
hldIndex = rsPlane.Index
rsPlane.Index = "Primary Key"
rsPlane.MoveFirst
While Not rsPlane.EOF
    If rsPlane.Fields("x") < 1000 Then
        rsPlane.Edit
        rsPlane.Fields("x") = rsPlane.Fields("x") - sx + 2000
        rsPlane.Fields("y") = rsPlane.Fields("y") - sy + 2000
        rsPlane.Update
    End If
    rsPlane.MoveNext
Wend

rsPlane.MoveFirst
While Not rsPlane.EOF
    If rsPlane.Fields("x") > 1000 Then
        rsPlane.Edit
        holdx = rsPlane.Fields("x") - 2000
        holdy = rsPlane.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsPlane.Fields("x") = holdx
        rsPlane.Fields("y") = holdy
        rsPlane.Update
    End If
    rsPlane.MoveNext
Wend
rsPlane.Index = hldIndex

'Update Land file
hldIndex = rsLand.Index
rsLand.Index = "Primary Key"
rsLand.MoveFirst
While Not rsLand.EOF
    If rsLand.Fields("x") < 1000 Then
        rsLand.Edit
        rsLand.Fields("x") = rsLand.Fields("x") - sx + 2000
        rsLand.Fields("y") = rsLand.Fields("y") - sy + 2000
        rsLand.Update
    End If
    rsLand.MoveNext
Wend
rsLand.MoveFirst
While Not rsLand.EOF
    If rsLand.Fields("x") > 1000 Then
        rsLand.Edit
        holdx = rsLand.Fields("x") - 2000
        holdy = rsLand.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsLand.Fields("x") = holdx
        rsLand.Fields("y") = holdy
        rsLand.Update
    End If
    rsLand.MoveNext
Wend

rsLand.Index = hldIndex

'Update Ship file
hldIndex = rsShip.Index
rsShip.Index = "Primary Key"
rsShip.MoveFirst
While Not rsShip.EOF
    If rsShip.Fields("x") < 1000 Then
        rsShip.Edit
        rsShip.Fields("x") = rsShip.Fields("x") - sx + 2000
        rsShip.Fields("y") = rsShip.Fields("y") - sy + 2000
        rsShip.Update
    End If
    rsShip.MoveNext
Wend
rsShip.MoveFirst
While Not rsShip.EOF
    If rsShip.Fields("x") > 1000 Then
        rsShip.Edit
        holdx = rsShip.Fields("x") - 2000
        holdy = rsShip.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsShip.Fields("x") = holdx
        rsShip.Fields("y") = holdy
        rsShip.Update
    End If
    rsShip.MoveNext
Wend
rsShip.Index = hldIndex

'Update Nuke file
rsNuke.MoveFirst
While Not rsNuke.EOF
    If rsNuke.Fields("x") < 1000 Then
        rsNuke.Edit
        rsNuke.Fields("x") = rsNuke.Fields("x") - sx + 2000
        rsNuke.Fields("y") = rsNuke.Fields("y") - sy + 2000
        rsNuke.Update
    End If
    rsNuke.MoveNext
Wend

rsNuke.MoveFirst
While Not rsNuke.EOF
    If rsNuke.Fields("x") > 1000 Then
        rsNuke.Edit
        holdx = rsNuke.Fields("x") - 2000
        holdy = rsNuke.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsNuke.Fields("x") = holdx
        rsNuke.Fields("y") = holdy
        rsNuke.Update
    End If
    rsNuke.MoveNext
Wend

'Update Enemy Unit file
hldIndex = rsEnemyUnit.Index
rsEnemyUnit.Index = "Primary Key"
rsEnemyUnit.MoveFirst
While Not rsEnemyUnit.EOF
    rsEnemyUnit.Edit
    If rsEnemyUnit.Fields("x") < 1000 Then
        rsEnemyUnit.Fields("x") = rsEnemyUnit.Fields("x") - sx + 2000
        rsEnemyUnit.Fields("y") = rsEnemyUnit.Fields("y") - sy + 2000
        rsEnemyUnit.Update
    End If
    rsEnemyUnit.MoveNext
Wend

rsEnemyUnit.MoveFirst
While Not rsEnemyUnit.EOF
    rsEnemyUnit.Edit
    If rsEnemyUnit.Fields("x") > 1000 Then
        holdx = rsEnemyUnit.Fields("x") - 2000
        holdy = rsEnemyUnit.Fields("y") - 2000
        OffsetSectorLocation holdx, holdy, 0, 0
        rsEnemyUnit.Fields("x") = holdx
        rsEnemyUnit.Fields("y") = holdy
        rsEnemyUnit.Update
    End If
    rsEnemyUnit.MoveNext
Wend
rsEnemyUnit.Index = hldIndex

'Move jump points
rsNation.MoveFirst
rsNation.Edit
For n = 1 To 5
    If ParseSectors(holdx, holdy, rsNation.Fields("JumpPoint" + CStr(n)).Value) Then
        OffsetSectorLocation holdx, holdy, -1 * sx, -1 * sy
        rsNation.Fields("JumpPoint" + CStr(n)).Value = _
        SectorString(holdx, holdy)
    End If
Next n
rsNation.Update

End Sub

Public Sub GetNationVariables()
On Error GoTo Empire_Error

Set etsTable = New EmpTables
Set estsTable = New EmpSymbolTables

If rsNation.BOF And rsNation.EOF Then
    Exit Sub
End If

rsNation.MoveFirst
'Fill the nation variables from the database
Education = rsNation.Fields("Education")
TechLevel = rsNation.Fields("TechLevel")
Happiness = rsNation.Fields("Happiness")
Research = rsNation.Fields("Research")
CapSect = rsNation.Fields("CapSect")
Maxpop = rsNation.Fields("Maxpop")
MaxSafeCivs = rsNation.Fields("MaxSafeCivs")
MaxSafeUws = rsNation.Fields("MaxSafeUws")
HapNeeded = rsNation.Fields("HapNeeded")
MilReserves = rsNation.Fields("MilReserves")
HarborAvail = rsNation.Fields("HarborAvail")
AirportAvail = rsNation.Fields("AirportAvail")
FortAvail = rsNation.Fields("FortAvail")
HQAvail = rsNation.Fields("HqAvail")

Exit Sub
Empire_Error:
EmpireError "GetNationVariables", vbNullString, vbNullString
End Sub

Public Sub SaveNationVariables()
On Error GoTo Empire_Error

If rsNation.BOF And rsNation.EOF Then
    rsNation.AddNew
Else
    rsNation.MoveFirst
    rsNation.Edit
End If

'Fill the nation variables from the database
rsNation.Fields("Education") = Education
rsNation.Fields("TechLevel") = TechLevel
rsNation.Fields("Happiness") = Happiness
rsNation.Fields("Research") = Research
rsNation.Fields("CapSect") = CapSect
rsNation.Fields("Maxpop") = Maxpop
rsNation.Fields("MaxSafeCivs") = MaxSafeCivs
rsNation.Fields("MaxSafeUws") = MaxSafeUws
rsNation.Fields("HapNeeded") = HapNeeded
rsNation.Fields("MilReserves") = MilReserves
rsNation.Fields("HarborAvail") = HarborAvail
rsNation.Fields("AirportAvail") = AirportAvail
rsNation.Fields("FortAvail") = FortAvail
rsNation.Fields("HqAvail") = HQAvail
rsNation.Update

Exit Sub
Empire_Error:
EmpireError "SaveNationVariables", vbNullString, vbNullString
End Sub

Function SetAccessProperty(obj As Object, strName As String, intType As Integer, varSetting As Variant) As Boolean
Dim prp As Property
Const conPropNotFound As Integer = 3270

On Error GoTo ErrorSetAccessProperty
obj.Properties(strName) = varSetting
obj.Properties.Refresh
SetAccessProperty = True
Exit Function
  
ErrorSetAccessProperty:
If Err = conPropNotFound Then
    Set prp = obj.CreateProperty(strName, intType, varSetting)
    obj.Properties.Append prp
    obj.Properties.Refresh
    SetAccessProperty = True
Else
    MsgBox Err & ": " & vbCrLf & Err.Description
    SetAccessProperty = False
End If
End Function


