Attribute VB_Name = "EmpCommon"
Option Explicit

'Change Log:
'081103 efj: removal of dead variables and code
'090703 rjk: Corrected SectorDistance to consider the world wraparound when determining the shortest path
'090703 rjk: EmpireError Added CStr to Application versions and removed semicolon
'092803 rjk: Added ndump timestamp.
'093003 rjk: Expand the check for Fleet identifiers to include A-Z and ~ as well for GetUnitList.
'100203 rjk: When firing on a Sea sector the damage goes to 100% to 0%.
'            The Defense strength for the sea is zero which causes a divide by zero error
'            in SectorEffAfterDamage so just return 0 when defense strength is zero.
'100503 rjk: Removed ndump timestamp
'101403 rjk: Added ndump timestamp (try #2)
'101903 rjk: Added check for Multiple Sectors in ParseSectors, return false if multiple sectors are present
'102003 rjk: Fixed leap year bug in SecondsDifference
'102503 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
'112903 rjk: Updated Error routine to point to sourceforge for assistance.
'120303 rjk: Switched options to frmOptions and boolean options
'120703 rjk: Changed hldindex to hldIndex.
'121203 rjk: Switched Items and Item objects.
'122903 rjk: Removed chance calculation from PlagueRisk to the display time to allow display options.
'260104 rjk: Added y sector for St@r W@rs.
'270104 rjk: Adjusted mobility calculations for St@r W@rs ships (star ships)
'280304 rjk: Adjusted the Production for ROLLOVER_AVAIL - retro
'080404 rjk: Removed the 999 limit on worker avail to match the server in the production routine
'130604 rjk: Added the ability to display path from sector to sector for flying
'200604 rjk: Fixed ParseSectors to verify valid address (0,-1 not valid) while parsing
'210704 rjk: Added "~" to production for StarWars II
'210704 rjk: Switched to mcost check instead of designation check for mobility cost
'260804 rjk: Fixed civ at 5000 for production calculation for 2K4.
'270904 rjk: Added routines to support xdump parsing.
'011204 rjk: Fixed routines to support xdump parsing, added CalculateTechDiff.
'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
'            ships with SWEEP capabilities.
'070605 rjk: Changed PlagueRisk to use the maxpop sector field instead of hard coding to
'            999 or 9999 for Cities.
'230705 rjk: Added EE9 Theme game option to suppress the production limits.
'281205 rjk: Added back the material limit from the resource depletion.
'281205 rjk: Added max. eff. gain limits for sect.
'281205 rjk: Use infrastructure information in the production calculations.
'180206 rjk: Added SectorBuildWork function.
'120306 rjk: Fixed Firing Range and Mobility Costs.  Add Function to Compute Unit Counts
'            for ships.
'230406 rjk: Added Function to Compute Unit Counts for Land Units.
'140506 rjk: Removed max production limits for South Pacific Atlantis
'170506 rjk: Removed the unused commodities for SP: Atlantis.
'200606 rjk: Added Nuke support to GetUnitList.
'            Incorporate mobility cost changes for 4.3.6 servers.
'250606 rjk: Do not enter minefields when calculating mobility for pre 4.3.6 servers.
'            Check nav infrastructure for SP: Atlantis.
'110706 rjk: Accomodate the change to show sect s for 4.3.6 servers.
'160706 rjk: Added AutoView for navigate and march.
'091206 rjk: Fixed the Fort fire range to apply the scaling before adding the
'            efficency bonus.
'221206 rjk: Fixed ship mobility cost for minesweeps for 4.3.6 servers.
'240307 rjk: Generalize oil production to check sector type.
'090607 rjk: Compute whether production is material limited or now.
'            Fixed new efficiency calculation to deal negatives during
'            tear down.
'            Removed the RolloverAvail scaling by sector efficiency.
'120607 rjk: Modified SecondsDifference to accomodate RFC 2822 date change
'            for server 4.3.10 -- scheduler updated
'301007 rjk: Added option to allow configuration of the maximum path length
'260408 rjk: Added 'q' for Heavy Metal II biofuel plant to Production
'120508 rjk: Allow ~ to food produce as well as oil
'170508 rjk: Fix production to use the depletion value in sect-chr
'220508 rjk: Fix production to need max civ when not producing any product.
'240508 rjk: Modified Production to use the 1% cost from the sect-chr.
'310508 rjk: Allow plains designation for Heavy Metal II
'070608 rjk: Fix the packing factor for mobility calculations.
'270608 rjk: Added navigation of ships under bridges that you do not own.
'160208 rjk: Change LoadTypeBox to use terrain check to determine if a sector
'            type can be redesignated.
'250409 rjk: Fix production so it will show production of military for enlistment
'            center when stopped.  It was incorrectly looking at the stopped state
'            instead of the not owned state.
'090311 kab: Modified the date parses to deal more regional effects.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

'EmpCmd Constants for ENEMY SECTORS
Public Const ES_DES = 1
Public Const ES_X = 2
Public Const ES_Y = 3
Public Const ES_OWNER = 4
Public Const ES_OLDOWNER = 5
Public Const ES_EFFICIENCY = 6
Public Const ES_ROAD = 7
Public Const ES_RAIL = 8
Public Const ES_DEFENSE = 9
Public Const ES_CIV = 10
Public Const ES_MIL = 11
Public Const ES_SHELL = 12
Public Const ES_GUN = 13
Public Const ES_PETROL = 14
Public Const ES_FOOD = 15
Public Const ES_IRON = 16
Public Const ES_BARS = 17
Public Const ES_NUM = 17

'EmpCmd Constants for ENEMY UNITS
Public Const EU_ID = 1
Public Const EU_X = 2
Public Const EU_Y = 3
Public Const EU_OWNER = 4
Public Const EU_TYPE = 5
' Public Const EU_CLASS = 6   removed efj 8/2003
Public Const EU_MIL = 7
Public Const EU_TECH = 8
Public Const EU_EFF = 9
' Public Const EU_WAKE = 10   removed efj 8/2003
' Public Const EU_NUM = 10   removed efj 8/2003

'EmpCmd Constants for LOST ITEMS
Public Const LS_TYPE = 0
Public Const LS_ID = 1
Public Const LS_X = 2
Public Const LS_Y = 3
' Public Const LS_TIMESTAMP = 4   removed efj 8/2003

'EmpCmd Constants for LOST TYPES
Public Const LT_SECTOR = 0
Public Const LT_SHIP = 1
Public Const LT_PLANE = 2
Public Const LT_LAND = 3
Public Const LT_NUKE = 4 '101403 rjk: Added back,  removed efj 8/2003

'Class EmpNationList Nation Relationships codes
Public iREL_AT_WAR As Integer
Public iREL_SITZKRIEG As Integer
Public iREL_MOBILIZING As Integer
Public iREL_HOSTILE As Integer
Public iREL_NEUTRAL As Integer
Public iREL_FRIENDLY As Integer
Public iREL_ALLIED As Integer


Public Const HEX_DIR = "yujnbg"

'Mobility Constants
Public Const MB_NORMAL = 1
'Public Const MB_DIST = 10   removed efj 8/2003
'Public Const MB_DELIVER = 4   removed efj 8/2003

'Public Const PB_NONE = 0   removed efj 8/2003
'Public Const PB_CIVS = 1   removed efj 8/2003
'Public Const PB_WAREHOUSE = 2   removed efj 8/2003
'Public Const PB_BANK = 3   removed efj 8/2003

Enum enumMobType  'drk 5/25/03. constants suck
    MT_COMMODITY = 1
    MT_LAND = 2
    MT_SHIP = 3
    MT_PLANE = 4 '111803 rjk: Added for DisplayPath in recon/fly/etc.
    MT_RAIL = 5
    MT_SMALL_SHIP = 6 'a ship small enough navigate a canal
End Enum

'Application name
Public Const APPNAME = "WinACE"
' Global Const LISTVIEW_BUTTON = 11   removed efj 8/2003

'Timestamps
Public tsSectors As Long
Public tsShip As Long
Public tsLand As Long
Public tsPlane As Long
Public tsLost As Long
Public tsNuke As Long '101403 rjk: Added timestamp to ndump
Public timNextUpdate As Date

'the production function fills in this data type.  before 2.2.8 this was an array of variants,
'which was a horrible programming practice (sorry guys).  There is absolutely no reason to
'use an array of variants when user defined types are available in the language!
'fixme: get some more specific data types for most of these
Public Type productionDataType
    NewEff As Integer        '1
    ProdAmount As Variant    '2
    MaxProdAmount As Variant '3
    ExcessCiv As Integer     '4
    newdes As Variant        '5
    unitsConsumed As Variant '6
    ItemProduced As String   '7
    BuildAvailCost As Variant '8
    ResourceEff As Variant   '9
    LevelEff As Variant      '10
    ProductionAvail As Variant
    prodValidFlag As Boolean 'if true, then production function completed successfully, otherwise it failed.
    prodMaxMaterialLimit As Boolean 'true if Max is material limited
End Type

'Error Variables
'Public errNumber As Long   removed efj 8/2003
'Public errDesc As String   removed efj 8/2003
'Public errIndex As Integer   removed efj 8/2003
'Public errLine As String   removed efj 8/2003

Public colDes As New Collection
Private arySectDes() As String

Public Function Max(ByVal a As Double, ByVal b As Double) As Double
    If (a > b) Then
        Max = a
    Else
        Max = b
    End If
End Function

Public Function Min(ByVal a As Double, ByVal b As Double) As Double
    If (a < b) Then
        Min = a
    Else
        Min = b
    End If
End Function


Sub LoadSectorDesc()

Dim n As Integer

On Error GoTo Empire_Error

While colDes.Count > 0
    colDes.Remove 1
Wend

ReDim arySectDes(1)

If rsSectorType.BOF And rsSectorType.EOF Then
    Exit Sub
End If

rsSectorType.MoveFirst

'Load sector descriptions from file
While Not rsSectorType.EOF
    colDes.Add rsSectorType.Fields("desc").Value, rsSectorType.Fields("des").Value
    n = n + 1
    ReDim Preserve arySectDes(n)
    arySectDes(n) = rsSectorType.Fields("des").Value + " [" + rsSectorType.Fields("desc").Value + "]"
    rsSectorType.MoveNext
Wend


Exit Sub
Empire_Error:
EmpireError "EmpCommon - LoadSectorDesc", vbNullString, vbNullString
End Sub
Function SectDesString(n As Integer) As String
On Error Resume Next
SectDesString = vbNullString
SectDesString = arySectDes(n)
End Function

Public Sub LoadCommodityBox(lb As Object)
Dim Index As Integer

lb.Clear
If frmOptions.bSPAtlantis Then
    For Index = 1 To 3
        lb.AddItem Items(Index).FormName
    Next Index
Else
    For Index = 1 To Items.Count
        lb.AddItem Items(Index).FormName
    Next Index
End If
End Sub

Public Function LoadTypebox(lb As Object) As String
Dim n As Integer
On Error GoTo Empire_Error

n = 0
lb.Clear
LoadTypebox = vbNullString

If rsSectorType.BOF And rsSectorType.EOF Then
    'Exit Sub
    Exit Function
End If

rsSectorType.MoveFirst

'Load sector descriptions from file
While Not rsSectorType.EOF
    If (Not bDeity And rsSectorType.Fields("dchr_i") <> rsSectorType.Fields("terrain")) Or _
        bDeity Then
        n = n + 1
        lb.AddItem CStr(rsSectorType.Fields("desc").Value)
        lb.ItemData(lb.NewIndex) = n
        LoadTypebox = LoadTypebox + rsSectorType.Fields("des").Value
    End If
    rsSectorType.MoveNext
Wend

Exit Function
Empire_Error:
EmpireError "LoadTypebox", CStr(n), vbNullString
End Function


Public Sub LoadRetreatBox(lb As Object)

'          I         Retreat when the ship is injured,
'                    i.e. whenever the ship is damaged
'                    by fire, bombs, or torping.
'          T         Retreat when a sub torpedos or tries
'                    to torpedo the ship.
'          B         Retreat when a plane bombs or tries
'                    to bomb the ship.
'          S         Retreat when the ship detects a sonar
'                    ping.
'          D         Retreat when the ship is depth-charged.
'          H         Retreat when helpless. A ship is helpless
'                    when it is fired upon, and no friendly
'                    ships/sectors (including the ship itself)
'                    are able to fire back at the aggressor.
'          U         Retreat upon a failed boarding attempt.
'          C         Clear the flags

lb.Clear
lb.AddItem "Injured"
lb.AddItem "Torpedoed"
lb.AddItem "Bombed"
lb.AddItem "Sonar"
lb.AddItem "Depth-Charge"
lb.AddItem "Helpless"
lb.AddItem "Unsuc. Board"
lb.AddItem "Clear Retreat"
End Sub


'Routine to compute the airports a set of planes are flying from
Public Function ComputeAirports(strPlanes As String) As Variant

Dim CurrUnit As Variant
Dim Airports() As String
Dim n As Integer
Dim i As Integer
Dim strLoc As String
Dim strLast As String
Dim found As Boolean
Dim First As Boolean
On Error GoTo Empire_Error

'Get current units
If Len(strPlanes) > 0 Then
    CurrUnit = GetUnitList(strPlanes, "p")
Else
    ReDim CurrUnit(0)
    CurrUnit(0) = vbNullString
    ComputeAirports = CurrUnit
    Exit Function
End If

rsPlane.Index = "PrimaryKey"
First = True

For n = LBound(CurrUnit) + 1 To UBound(CurrUnit)
    rsPlane.Seek "=", CInt(CurrUnit(n))
    If Not rsPlane.NoMatch Then
        strLoc = SectorString(rsPlane.Fields("x"), rsPlane.Fields("y"))
        'Always add first airport
        If First Then
            ReDim Preserve Airports(1)
            Airports(1) = strLoc
            First = False
        'If same airport as last plane (common) then skip
        ElseIf strLoc = strLast Then
        
        'Add to list if not already there
        Else
            'First search
            found = False
            For i = LBound(Airports) To UBound(Airports)
                If Airports(i) = strLoc Then
                    found = True
                End If
            Next i
            'then add if necessary
            If Not found Then
                i = UBound(Airports)
                ReDim Preserve Airports(i + 1)
                Airports(i + 1) = strLoc
            End If
        End If
        'Save last airport, as planes usually come out in order
        strLast = strLoc
    End If
Next n
        
ComputeAirports = Airports

Exit Function
Empire_Error:
EmpireError "ComputeAirports", vbNullString, vbNullString
End Function
Public Sub ComputeDestination(x1 As Integer, y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer, strPath)

Dim X As Integer
Dim p1 As Integer
On Error GoTo Empire_Error

X2 = x1
Y2 = y1
p1 = Len(strPath)
For X = 1 To p1
    Select Case Mid$(strPath, X, 1)
        Case "g"
            X2 = X2 - 2
        Case "y"
            X2 = X2 - 1
            Y2 = Y2 - 1
        Case "u"
            X2 = X2 + 1
            Y2 = Y2 - 1
        Case "j"
            X2 = X2 + 2
        Case "n"
            X2 = X2 + 1
            Y2 = Y2 + 1
        Case "b"
            X2 = X2 - 1
            Y2 = Y2 + 1
    End Select
Next X

'Handle worldwrap
If X2 > MapSizeX / 2 Then X2 = X2 - MapSizeX
If X2 <= -1 * MapSizeX / 2 Then X2 = X2 + MapSizeX
If Y2 > MapSizeY / 2 Then Y2 = Y2 - MapSizeY
If Y2 <= -1 * MapSizeY / 2 Then Y2 = Y2 + MapSizeY
    

Exit Sub
Empire_Error:
EmpireError "ComputeDestination", vbNullString, CStr(x1) + " " + CStr(y1) + " " + strPath
End Sub

Public Function SectorDistance(x1 As Integer, y1 As Integer, X2 As Integer, Y2 As Integer) As Integer
'returns distance as the plane flies
On Error GoTo Empire_Error

Dim dy As Integer
Dim dx As Integer

'090703 rjk: Change to select the shortest path, i.e. half the map size in the y directon
dy = Abs(y1 - Y2)
If dy > (MapSizeY / 2#) Then
    dy = MapSizeY - dy
End If

'090703 rjk: Change to select the shortest path, i.e. 1/4 the map size in the x directon
'dx = ((CInt(Abs(X2 - x1) - dy) / 2#))
dx = CInt(Abs(X2 - x1))
If dx > (MapSizeX / 2#) Then
    dx = MapSizeX - dx
End If
dx = (dx - dy) / 2#
If dx < 0 Then
    dx = 0
End If

SectorDistance = dx + dy
Exit Function
Empire_Error:
EmpireError "SectorDistance", vbNullString, CStr(x1) + ":" + CStr(y1) + ":" + CStr(X2) + ":" + CStr(Y2)
End Function

       
Public Function EmpirePathDirection(ByVal dx As Integer, ByVal dy As Integer) As String
On Error GoTo Empire_Error

Dim xdiff As Integer

'Check for valid numbers
If (dx + dy) Mod 2 <> 0 Then
    EmpirePathDirection = "h"
    Exit Function
End If

'If there is no change, then it is done
If dx = 0 And dy = 0 Then
    EmpirePathDirection = vbNullString
    Exit Function
End If

xdiff = Abs(dx) - Abs(dy)

'if there is a greater xdiff then move left or right first
If xdiff > 1 Then
    If dx > 0 Then
        EmpirePathDirection = "j" + EmpirePathDirection(dx - 2, dy)
    Else
        EmpirePathDirection = "g" + EmpirePathDirection(dx + 2, dy)
    End If
    Exit Function
End If

'head in a diagonal direction
If dy < 0 Then
    If dx >= 0 Then
        EmpirePathDirection = "u" + EmpirePathDirection(dx - 1, dy + 1)
    Else
        EmpirePathDirection = "y" + EmpirePathDirection(dx + 1, dy + 1)
    End If
Else
    If dx >= 0 Then
        EmpirePathDirection = "n" + EmpirePathDirection(dx - 1, dy - 1)
    Else
        EmpirePathDirection = "b" + EmpirePathDirection(dx + 1, dy - 1)
    End If
End If

Exit Function
Empire_Error:
EmpireError "EmpirePathDirection", vbNullString, CStr(dx) + ":" + CStr(dy)
End Function

Public Sub CopyListBoxToClipboard(lb As ListBox)
On Error GoTo Empire_Error

' Copy Selected text into a sring
Dim X As Integer
Dim txtstring As String

Clipboard.Clear  'Clear Clipboard.
For X = 0 To (lb.ListCount - 1)
    If lb.Selected(X) Then
        txtstring = txtstring + lb.List(X) + vbCr + vbLf
    End If
Next X

'If nothing is selected, copy the whole thing
If Len(txtstring) = 0 Then
    For X = 0 To (lb.ListCount - 1)
        txtstring = txtstring + lb.List(X) + vbCr + vbLf
    Next X
End If

Clipboard.SetText txtstring    ' Put text on Clipboard.
Exit Sub
Empire_Error:
EmpireError "CopyListBoxToClipboard", vbNullString, lb.Name
End Sub

Public Function ShipMoveCost(speed As Integer, Eff As Integer, Tech As Integer, strType As String) As Single
On Error GoTo Empire_Error

'The mobility cost for a ship to navigate or retreat is:
'  (sectors travelled) * 480 / (ship speed)
'Where
'  ship speed = (base speed) * (1 + (tech factor))
'  base speed = Max(0.01, efficiency * speed)
'  tech factor = (50 + tech) / (200 + tech)

Dim tech_factor As Single
Dim base_speed As Single
Dim ship_speed As Single

tech_factor = (50 + Tech) / (200 + Tech)

'base speed = Max(0.01, efficiency * speed)
base_speed = (Eff * speed) / 100#
If base_speed < 0.01 Then
    base_speed = 0.01
End If

ship_speed = (base_speed) * (1 + (tech_factor))
ShipMoveCost = 480 / ship_speed

If VersionCheck(4, 3, 6) = VER_LESS Then
    If frmDrawMap.UnitCharacteristicCheck(TYPE_S_SWEEP, strType) Then
        ShipMoveCost = ShipMoveCost * 2
    End If
End If

Exit Function
Empire_Error:
EmpireError "ShipMoveCost", vbNullString, CStr(speed) + ":" + CStr(Eff) + ":" + CStr(Tech)
End Function

Function PlaneFlyCost(range As Integer, Eff As Integer, distance As Integer) As Single
On Error GoTo Empire_Error

PlaneFlyCost = 5 + ((20 * 100 / Eff) * distance / range)

Exit Function
Empire_Error:
EmpireError "PlaneFlyCost", vbNullString, CStr(range) + ":" + CStr(Eff) + ":" + CStr(distance)
End Function


'Take a string with sector numbers ("x,y"), and break them into two integers
Public Function ParseSectors(ByRef sx As Integer, ByRef sy As Integer, strSect As String) As Boolean

Dim p1 As Integer
On Error GoTo ParseSectors_Error

'101803 rjk: Added check for Multiple Sectors, return false if multiple sectors are present
If InStr(strSect, "\") > 0 Then
    ParseSectors = False
End If

p1 = InStr(strSect, ",")
If p1 > 0 And Len(strSect) > p1 Then
    sx = CInt(Left$(strSect, p1 - 1))
    sy = CInt(Mid$(strSect, p1 + 1))
    If ((sx + sy) Mod 2) = 0 Then
        ParseSectors = True
    Else
        ParseSectors = False
    End If
    Exit Function
End If

ParseSectors_Error:
ParseSectors = False
End Function

'Take a string with sector numbers ("x1:x2,y1:y2"), and break them into four integers
Public Function ParseSectorRange(ByRef sx1 As Integer, ByRef sx2 As Integer, ByRef sy1 As Integer, ByRef sy2 As Integer, ByVal strSect As String) As Boolean

Dim p1 As Integer
On Error GoTo ParseSectorRange_Error

p1 = InStr(strSect, ":")
If p1 > 0 And Len(strSect) > p1 Then
    sx1 = CInt(Left$(strSect, p1 - 1))
    strSect = Mid$(strSect, p1 + 1)
    p1 = InStr(strSect, ",")
    If p1 > 0 And Len(strSect) > p1 Then
        sx2 = CInt(Left$(strSect, p1 - 1))
        strSect = Mid$(strSect, p1 + 1)
        p1 = InStr(strSect, ":")
        If p1 > 0 And Len(strSect) > p1 Then
            sy1 = CInt(Left$(strSect, p1 - 1))
            sy2 = CInt(Mid$(strSect, p1 + 1))
            ParseSectorRange = True
            Exit Function
        End If
    End If
End If

ParseSectorRange_Error:
ParseSectorRange = False
End Function


'Parse a string of units into a individual unit array ("xxx/yyy/zzz")
Public Function ParseUnitString(ByVal strUnit As String) As Variant
On Error GoTo Empire_Error

Dim UnitList() As String
Dim strItem As String
Dim n As Integer

'Parse the string, seperating a slashes
While Len(strUnit) > 0
    n = InStr(strUnit, "/")
    If n = 1 Then
        strUnit = Mid$(strUnit, 2)
    ElseIf n = 0 Then
        strItem = strUnit
        strUnit = vbNullString
    Else
        strItem = Left$(strUnit, n - 1)
        strUnit = Mid$(strUnit, n + 1)
    End If
    
    'Add Item to list
    If Len(strItem) > 0 Then
        On Error Resume Next
        n = 0   'Set first in case of error
        n = UBound(UnitList)
        ReDim Preserve UnitList(n + 1)
        UnitList(n + 1) = strItem
        On Error GoTo 0
    End If
Wend
        
ParseUnitString = UnitList()
Exit Function

Empire_Error:
EmpireError "ParseUnitString", vbNullString, strUnit
End Function

Public Function SecondsDifference(strTime1 As String, strTime2 As String)
'computes the seconds between two empire time strings
On Error GoTo Empire_Error

Dim strT1 As String
Dim strT2 As String
Dim dd As Long
Dim mm As Long
Dim mn As Long
Dim hh As Long

strTime1 = Trim$(strTime1)
strTime2 = Trim$(strTime2)
If Asc(Left$(strTime1, 1)) >= Asc("0") And Asc(Left$(strTime1, 1)) <= Asc("9") Then
    dd = -Trim$(Left$(strTime1, 2)) + Trim$(Left$(strTime2, 2))
    
    mm = 0 'month
    If Mid$(strTime1, 4, 3) <> Mid$(strTime2, 4, 3) Then
        mn = InStr("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", Mid$(strTime1, 4, 3))
        mm = Int(Mid$("031 028 031 030 031 030 031 031 030 031 030 031", mn, 3))
    End If
    
    'leap year adjustment
    '102003 rjk: Fixed leap year - mn is the position in the string, not the actually month
    '            so switched mn = 2 to mn = 5
    'If mn = 2 And (Year(Now) Mod 4) = 0 Then
    If mn = 5 And (Year(Now) Mod 4) = 0 Then
        mm = 29
    End If
    
    dd = dd + mm
    strT1 = Trim$(Mid$(strTime1, 12))
    strT2 = Trim$(Mid$(strTime2, 12))
Else
    mm = 0 'month
    If Left$(strTime1, 3) <> Left$(strTime2, 3) Then
        mn = InStr("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", Left$(strTime1, 3))
        mm = Int(Mid$("031 028 031 030 031 030 031 031 030 031 030 031", mn, 3))
    End If
    
    'leap year adjustment
    '102003 rjk: Fixed leap year - mn is the position in the string, not the actually month
    '            so switched mn = 2 to mn = 5
    'If mn = 2 And (Year(Now) Mod 4) = 0 Then
    If mn = 5 And (Year(Now) Mod 4) = 0 Then
        mm = 29
    End If
    
    dd = mm - Trim$(Mid$(strTime1, 5, 2)) + Trim$(Mid$(strTime2, 5, 2))
    strT1 = Trim$(Mid$(strTime1, 7))
    strT2 = Trim$(Mid$(strTime2, 7))
End If
hh = CInt(Left$(strT2, 2)) - CInt(Left$(strT1, 2)) + 24 * dd
mm = CInt(Mid$(strT2, 4, 2)) - CInt(Mid$(strT1, 4, 2)) + 60 * hh
SecondsDifference = CInt(Mid$(strT2, 7, 2)) - CInt(Mid$(strT1, 7, 2)) + 60 * mm

Exit Function

Empire_Error:
EmpireError "SecondsDifference", vbNullString, strTime1 + ":" + strTime2
End Function

'Changed by KAB 20110309
'strTime looked like this: "on mar 09 09:11:35 2011" ' notice the lowercase mar and the "on", old index values indicated that the day string "on" has been 3 chars long previous/with other regional settings
'problems was suspected to be caused by regional settings
Public Function EmpireTimeString(strTime As String) As Date
On Error GoTo Empire_Error

'Dim strT1 As String   removed efj 8/2003
'Dim strT2 As String   removed efj 8/2003
Dim dd As Integer
Dim mm As Integer
Dim nn As Integer
Dim hh As Integer
Dim ss As Integer
Dim yy As Integer

Dim MonthOfYear: MonthOfYear = "jan feb mar apr may jun jul aug sep oct nov dec"
Dim a_Cnt As Long
Dim a_CntValid As Long
Dim a_TimeStr
Dim a_hhnnssStr
Dim a_TimeStrValid
ReDim a_TimeStrValid(6) As String

a_TimeStr = Split(strTime, " ")

If Not IsArray(a_TimeStr) Then
    Err.Raise 13, , "a_TimeStr not an array"
    Exit Function
End If

a_CntValid = 0
For a_Cnt = 0 To UBound(a_TimeStr)
    If (Trim$(a_TimeStr(a_Cnt)) <> "") Then
        a_TimeStrValid(a_CntValid) = a_TimeStr(a_Cnt)
        a_CntValid = a_CntValid + 1
    End If
Next

a_hhnnssStr = Split(a_TimeStrValid(3), ":")
If Not IsArray(a_hhnnssStr) Then
    Err.Raise 14, , "a_hhnnssStr not an array"
    Exit Function
End If

mm = InStr(MonthOfYear, LCase(a_TimeStrValid(1)))
If mm = 0 Then
    Exit Function
End If

mm = ((mm - 1) / 4) + 1
dd = Int(a_TimeStrValid(2))
hh = Int(a_hhnnssStr(0))
nn = Int(a_hhnnssStr(1))
ss = Int(a_hhnnssStr(2))
yy = Int(a_TimeStrValid(4))
EmpireTimeString = DateSerial(yy, mm, dd) + TimeSerial(hh, nn, ss)

Exit Function

Empire_Error:
EmpireError "EmpireTimeString", vbNullString, strTime
End Function

'This function returns the string between two other strings
'used for extracting sector numbers, etc.
Function StringBetween(strSource As String, strLeft As String, Optional strRight As String) As String
On Error GoTo Empire_Error
Dim n As Integer
Dim n2 As Integer

'Can't do anything with a null source string
If Len(strSource) = 0 Then
    StringBetween = vbNullString
    Exit Function
End If

'if the left string is empty, then start at the beginning
If Len(strLeft) = 0 Then
    n = 1
Else
    'otherwise, find it and start afterwards
    n = InStr(strSource, strLeft)
    If n = 0 Then
        StringBetween = vbNullString
        Exit Function
    End If
    n = n + Len(strLeft)
End If

'if the right string is empty, run to the end of the source
If Len(strRight) = 0 Then
    StringBetween = Trim$(Mid$(strSource, n))
Else
    'pull the area between
    If n = 0 Then
        n2 = InStr(strSource, strRight)
        StringBetween = Trim$(Left$(strSource, n2 - 1))
        Exit Function
    End If
    n2 = InStr(n, strSource, strRight)
    StringBetween = Trim$(Mid$(strSource, n, n2 - n))
End If

Exit Function

Empire_Error:
EmpireError "StringBetween", vbNullString, strSource
End Function

Function SectorMobCost(sx As Integer, sy As Integer, ByVal MobType As enumMobType, Optional PostUpdate As Variant) As Single
If IsMissing(PostUpdate) Then
    PostUpdate = False
End If

On Error GoTo Empire_Error

If VersionCheck(4, 3, 6) <> VER_LESS Then

'    base = dchr[sp->sct_type].d_mob0;
'    if (base < 0)
'    return -1.0;
'
'    /* linear function in eff, d_mcst at 0%, d_emcst at 100% */
'    base += (dchr[sp->sct_type].d_mob1 - base) * sp->sct_effic / 100;
'    if (CANT_HAPPEN(base < 0))
'    base = 0;
'
'    if (mobtype == MOB_MOVE || mobtype == MOB_MARCH) {
'    /* linear function in road, base at 0%, base/10 at 100% */
'    cost = base;
'    if (intrchr[INT_ROAD].in_enable)
'        cost -= base * 0.009 * sp->sct_road;
'    } else if (mobtype == MOB_RAIL) {
'    if (!intrchr[INT_RAIL].in_enable || sp->sct_rail <= 0)
'        return -1.0;
'    /* linear function in rail, base at 0%, base/100 at 100% */
'    cost = base - base * 0.0099 * sp->sct_rail;
'    } else {
'    CANT_REACH();
'    cost = base;
'    }
'    if (CANT_HAPPEN(cost < 0))
'    cost = 0;
'
'    if (mobtype == MOB_MOVE)
'    return MAX(cost, 0.001);
'    if (sp->sct_own != sp->sct_oldown && sp->sct_mobil <= 0)
'    /* slow down land units in newly taken sectors */
'    return cost + 0.2;
'    return MAX(cost, 0.02);
    Dim sBase As Single
    Dim sCost As Single
    Dim strDes As String
    
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        strDes = rsSectors.Fields("des")
    Else
        If MobType = MT_SHIP Or MobType = MT_SMALL_SHIP Then
            rsBmap.Seek "=", sx, sy
            If rsBmap.NoMatch Then
                strDes = "."   'if not found, assume sea
            Else
                strDes = rsBmap.Fields("des")
            End If
        Else
            SectorMobCost = 9999
            Exit Function
        End If
        'Deal with sea sectors.
    End If

    rsSectorType.Seek "=", strDes
    If rsSectorType.NoMatch Then
        SectorMobCost = 9999
        Exit Function
    End If
    
    If MobType = MT_SHIP Or MobType = MT_SMALL_SHIP Then
        Select Case estsTable(XD_SECTOR_NAVIGATION)(rsSectorType.Fields("nav")).Name
        Case "sea"
            SectorMobCost = 1
        Case "harbor"
            SectorMobCost = 9999
            If Not rsSectors.NoMatch Then
                If rsSectors.Fields("eff") >= 2 Then
                    SectorMobCost = 1
                End If
            End If
        Case "canal"
            SectorMobCost = 9999
            If MobType = MT_SMALL_SHIP Then
                If Not rsSectors.NoMatch Then
                    If rsSectors.Fields("eff") >= 2 Then
                        SectorMobCost = 1
                    End If
                End If
            End If
        Case "bridge"
            SectorMobCost = 9999
            If Not rsSectors.NoMatch Then
                If rsSectors.Fields("eff") >= 60 Then
                    SectorMobCost = 1
                End If
            Else
                Dim hldIndex As String
                
                hldIndex = rsEnemySect.Index
                rsEnemySect.Index = "PrimaryKey"
                rsEnemySect.Seek "=", sx, sy
                If Not rsEnemySect.NoMatch Then
                    If rsEnemySect.Fields("eff") >= 60 Then
                        SectorMobCost = 1
                    End If
                End If
                rsEnemySect.Index = hldIndex
            End If
        Case Else
            SectorMobCost = 9999
        End Select
        Exit Function
    End If
    
    sBase = rsSectorType.Fields("mcost")
    If sBase < 0 Then
        SectorMobCost = 9999
        Exit Function
    End If
    
    sBase = sBase + ((CSng(rsSectorType.Fields("mcost100")) - sBase) * rsSectors.Fields("eff") / 100#)
    
    If sBase <= 0 Then
        'Error should not happen
        SectorMobCost = 9999
        Exit Function
    End If
    
    If MobType = MT_COMMODITY Or MobType = MT_LAND Then
        sCost = sBase
        rsBuild.Seek "=", "i", "road"
        If Not rsSectorType.NoMatch Then
            sCost = sCost - (sBase * 0.009 * rsSectors.Fields("road"))
        End If
    ElseIf MobType = MT_RAIL Then
            rsBuild.Seek "=", "i", "rail"
            If rsSectorType.NoMatch Or rsSectors.Fields("rail") <= 0 Then
                SectorMobCost = 9999
                Exit Function
            End If
            sCost = sBase - (sBase * 0.0099 * rsSectors.Fields("rail"))
    Else
        'Error should not happen
        sCost = sBase
    End If
    
    If sCost < 0 Then
        'Error should not happen
        sCost = 0
    End If
    
    If MobType = MT_COMMODITY Then
        SectorMobCost = Max(sCost, 0.001)
        Exit Function
    End If
    
    If rsSectors.Fields("*") = "*" And rsSectors.Fields("mob") <= 0 Then
        SectorMobCost = sCost + 0.2
        Exit Function
    End If
    SectorMobCost = Max(sCost, 0.02)
    
Else ' before 4.3.6 calculations
    If MobType = MT_PLANE Then
        SectorMobCost = 1
        Exit Function
    End If
    
    Dim Des As String
    Dim Eff As Integer
    
    'get sector info
    rsSectors.Seek "=", sx, sy
    If rsSectors.NoMatch Then
        If MobType = MT_SHIP Then
            rsBmap.Seek "=", sx, sy
            If rsBmap.NoMatch Then
                Des = "."   'if not found, assume sea
            Else
                Des = rsBmap.Fields("des")
            End If
            Eff = 100
        Else
            SectorMobCost = 999#
            Exit Function
        End If
    Else
        Des = rsSectors.Fields("des")
        Eff = rsSectors.Fields("eff")
        
        'if this is post-update, then find out sector des and efficency
        'after the update
        If CBool(PostUpdate) Then
            Dim v As productionDataType
            v = Production(rsSectors)
            If v.prodValidFlag Then
                Des = CStr(v.newdes)
                Eff = CInt(v.NewEff)
            End If
        End If
    End If
    
    'passable sea hexes have mob of 1, all others 999
    If MobType = MT_SHIP Then
        If frmOptions.bStarWars Then '260104 rjk: Added for St@r W@rs, modified for StarWars II
            If Des = "\" Then
                SectorMobCost = 999
            Else
                SectorMobCost = 1
            End If
        Else
            If Des = "." Or _
                Des = "h" And (Eff > 1) Or _
                Des = "=" And (Eff > 59) Or _
                Des = "c" And (Eff > 59) And rsVersion.Fields("BIG_CITY") Then
                SectorMobCost = 1
            ElseIf frmOptions.bSPAtlantis And Not rsSectors.NoMatch Then
                If rsSectors.Fields("ocontent") > 59 Then
                    SectorMobCost = 1
                Else
                    SectorMobCost = 999
                End If
            Else
                SectorMobCost = 999
            End If
        End If
        Exit Function
    End If
    
    rsSectorType.Seek "=", Des
    If rsSectorType.NoMatch Then
        SectorMobCost = 999#
        Exit Function
    Else
        SectorMobCost = rsSectorType.Fields("mcost")
        If SectorMobCost = 0 Then
            SectorMobCost = 999#
            Exit Function
        End If
    End If
    
    If (rsSectors.Fields("*") = "*") And _
        (rsSectors.Fields("mob") <= 0) Then 'if newly occupied territory
        If SectorMobCost < 2 Then           'negate road sec bonus
            SectorMobCost = 2#
        End If
    Else                                    'else give
                                            'road infrastructure bonus
        SectorMobCost = SectorMobCost / (1# + (rsSectors.Fields("road") / 122#))
    End If
    
    If SectorMobCost < 1 Then               'can't do better than 1
        SectorMobCost = 1                   'at this point
    End If
    
    If rsSectorType.Fields("mcost") >= 25 Then '210704 rjk: Switched mcost instead of designation check
        SectorMobCost = ((SectorMobCost * 10#) - Eff) / 115#
    Else
        SectorMobCost = ((SectorMobCost * 100#) - Eff) / 500#
    End If
    
    If MobType = MT_LAND Then
        If SectorMobCost < 0.01 Then            'must alway be at least
            SectorMobCost = 0.01                'minimum cost for land
        End If
    Else
        If SectorMobCost < 0.001 Then           'must alway be at least
            SectorMobCost = 0.001               'minimum cost
        End If
    End If
End If
Exit Function

Empire_Error:
EmpireError "SectorMobCost", vbNullString, CStr(sx) + ":" + CStr(sy)
End Function


Function BestPath(strSect As String, strDest As String, ByVal MobType As enumMobType, Optional PostUpdate As Variant) As Variant
Dim MaxRadius As Integer

MaxRadius = frmOptions.iMaximumPathLength

If IsMissing(PostUpdate) Then
    PostUpdate = False
End If

ReDim range(0 To MaxRadius * 4 + 5, 0 To MaxRadius * 2 + 3) As Integer
ReDim paths(0 To MaxRadius * 4 + 5, 0 To MaxRadius * 2 + 3) As String
ReDim mcost(0 To MaxRadius * 4 + 5, 0 To MaxRadius * 2 + 3) As Single
ReDim pcost(0 To MaxRadius * 4 + 5, 0 To MaxRadius * 2 + 3) As Single
Dim StartX As Integer
Dim StartY As Integer
Dim EndX As Integer
Dim EndY As Integer
Dim X As Integer
Dim Y As Integer
Dim sx As Integer
Dim sy As Integer
Dim tx As Integer
Dim ty As Integer
Dim n As Integer
Dim offy As Integer
Dim offx As Integer
Dim holdy As Integer
Dim holdx As Integer
Dim v(1 To 2) As Variant
v(1) = vbNullString
v(2) = 0

Const UNCHECKED_SECTOR = 0
Const CHECKED_SECTOR = 1
Const DESTINATION_SECTOR = 2
Const ADJACENT_SECTOR = 3
Const UNPASSABLE_SECTOR = 4
Const ORIGIN_SECTOR = 5

'Setup directional arrays for the six directions
Dim xa, ya, pa As Variant
xa = Array(-2, -1, 1, 2, 1, -1)
ya = Array(0, -1, -1, 0, 1, 1)
pa = Array("g", "y", "u", "j", "n", "b")

On Error Resume Next

'verify start sector and set offsets
If ParseSectors(StartX, StartY, strSect) Then
    offx = StartX - (MaxRadius + 2) * 2
    offy = StartY - MaxRadius
    If MobType <> MT_SHIP And MobType <> MT_SMALL_SHIP Then
        rsSectors.Seek "=", StartX, StartY
        If rsSectors.NoMatch Then
            BestPath = v
            Exit Function
        End If
    End If
Else
    BestPath = v
    Exit Function
End If

'verify endpoint sector and set offsets
If ParseSectors(EndX, EndY, strDest) Then
    If MobType <> MT_SHIP And MobType <> MT_PLANE And MobType <> MT_SMALL_SHIP Then
        rsSectors.Seek "=", EndX, EndY
        If rsSectors.NoMatch Then
            BestPath = v
            Exit Function
        End If
    End If
Else
    BestPath = v
    Exit Function
End If

'if start and finish hex are the same, then no cost vv
If StartX = EndX And StartY = EndY Then
    BestPath = v
    Exit Function
End If

Dim colx As New Collection
Dim coly As New Collection
Dim mincost As Single

'The x and y collections hold the sectors to be seached
'start with the initial sector
colx.Add (MaxRadius + 2) * 2
coly.Add MaxRadius
range((MaxRadius + 2) * 2, MaxRadius) = ORIGIN_SECTOR

Dim MaxX As Integer
Dim MaxY As Integer
'Adjust x for world wrap
MaxX = rsVersion.Fields("World X") / 2
MaxY = rsVersion.Fields("World Y") / 2

'set the destination sector - check for world wrap.
sx = EndX
sy = EndY
If (StartX > 0) And (EndX < 0) Or (StartX < 0 And EndX > 0) Then
    If (Abs(StartX) + Abs(EndX)) > MaxX Then
        If StartX > 0 Then
            sx = EndX + rsVersion.Fields("World X")
        Else
            sx = EndX - rsVersion.Fields("World X")
        End If
    End If
End If
If (StartY > 0) And (EndY < 0) Or (StartY < 0 And EndY > 0) Then
    If (Abs(StartY) + Abs(EndY)) > MaxY Then
        If StartY > 0 Then
            sy = EndY + rsVersion.Fields("World Y")
        Else
            sy = EndY - rsVersion.Fields("World Y")
        End If
    End If
End If

'if the distance is only one hex, then short path is cost
If (Abs(StartX - sx) + Abs(StartY - sy) = 2) And _
    (Abs(StartY - sy) <= 1) Then 'note - 2 in the y direction are not adjacent
    v(1) = EmpirePathDirection(sx - StartX, sy - StartY)
    v(2) = SectorMobCost(EndX, EndY, MobType, PostUpdate)
    BestPath = v
    Exit Function
End If

range(sx - offx, sy - offy) = DESTINATION_SECTOR
mcost(sx - offx, sy - offy) = SectorMobCost(EndX, EndY, MobType, PostUpdate)
pcost(sx - offx, sy - offy) = 9999
mincost = 9999

'setup the six adjacent sectors
For n = 0 To 5
    holdx = EndX + xa(n)
    holdy = EndY + ya(n)
    OffsetSectorLocation holdx, holdy, 0, 0
    If MobType = MT_SHIP Or MobType = MT_SMALL_SHIP Or MobType = MT_PLANE Then
        range(sx - offx + xa(n), sy - offy + ya(n)) = ADJACENT_SECTOR
        mcost(sx - offx + xa(n), sy - offy + ya(n)) = SectorMobCost(holdx, holdy, MobType, PostUpdate)
    Else
        rsSectors.Seek "=", holdx, holdy
        'if its not in the database - its unpassable
        If rsSectors.NoMatch Then
            range(sx - offx + xa(n), sy - offy + ya(n)) = UNPASSABLE_SECTOR
        Else
            range(sx - offx + xa(n), sy - offy + ya(n)) = ADJACENT_SECTOR
            mcost(sx - offx + xa(n), sy - offy + ya(n)) = SectorMobCost(holdx, holdy, MobType, PostUpdate)
        End If
    End If
Next n

'Keep going thru the grids until you have explored all
While colx.Count > 0

    X = colx.Item(1)
    Y = coly.Item(1)
    colx.Remove 1
    coly.Remove 1
    holdx = X + offx
    holdy = Y + offy
    OffsetSectorLocation holdx, holdy, 0, 0
    If range(X, Y) = UNCHECKED_SECTOR Then
        If MobType = MT_SHIP Or MobType = MT_SMALL_SHIP Or MobType = MT_PLANE Then
            range(X, Y) = CHECKED_SECTOR
        Else
            rsSectors.Seek "=", holdx, holdy
            'if its not in the database - its unpassable
            If rsSectors.NoMatch Then
                range(X, Y) = UNPASSABLE_SECTOR
            Else
                range(X, Y) = CHECKED_SECTOR
            End If
        End If
    End If
    
    Select Case range(X, Y)
        'if its unpassable, then we are done
        Case UNPASSABLE_SECTOR
        
        'if its adjacent to the destination sector, only look that way
        Case ADJACENT_SECTOR
            For n = 0 To 5
                If range(X + xa(n), Y + ya(n)) = DESTINATION_SECTOR And _
                (pcost(X + xa(n), Y + ya(n)) > (pcost(X, Y) + mcost(X + xa(n), Y + ya(n)))) Then
                    paths(X + xa(n), Y + ya(n)) = paths(X, Y) + pa(n)
                    pcost(X + xa(n), Y + ya(n)) = pcost(X, Y) + mcost(X + xa(n), Y + ya(n))
                    mincost = pcost(X, Y)
                End If
            Next
        
        'otherwise, explore out in every direction
        Case Else
            If range(X, Y) = ORIGIN_SECTOR Then
                range(X, Y) = UNPASSABLE_SECTOR
            End If

            For n = 0 To 5
                'first make sure the new address is in the grid boundries
                If (X + xa(n) >= LBound(paths, 1)) And (X + xa(n) <= UBound(paths, 1)) _
                    And (Y + ya(n) >= LBound(paths, 2)) And (Y + ya(n) <= UBound(paths, 2)) Then
                    
                    'compute the mobility cost if it has not already been done
                    If mcost(X + xa(n), Y + ya(n)) = 0 Then
                        tx = holdx + xa(n)
                        ty = holdy + ya(n)
                        OffsetSectorLocation tx, ty, 0, 0
                        mcost(X + xa(n), Y + ya(n)) = SectorMobCost(tx, ty, MobType, PostUpdate)
                    End If
                    
                    'if we haven't visted this sector, or if this is a cheaper path,
                    If (Len(paths(X + xa(n), Y + ya(n))) = 0) Or (pcost(X + xa(n), Y + ya(n)) > (pcost(X, Y) + mcost(X + xa(n), Y + ya(n)))) _
                            And range(X + xa(n), Y + ya(n)) <> UNPASSABLE_SECTOR Then
                        
                        'update the best path and total cost
                        paths(X + xa(n), Y + ya(n)) = paths(X, Y) + pa(n)
                        pcost(X + xa(n), Y + ya(n)) = pcost(X, Y) + mcost(X + xa(n), Y + ya(n))
                        
                        'if the cost to this hex is less than the mininum cost
                        'found so far, scan out from this new hex as well.
                        If pcost(X + xa(n), Y + ya(n)) < mincost Then
                            colx.Add X + xa(n)
                            coly.Add Y + ya(n)
                        End If
                    End If
                End If
            Next n
    End Select
Wend

'now find the path at the end

BestPath_Exit:

X = sx - offx
Y = sy - offy

If X < LBound(paths, 1) Or X > UBound(paths, 1) _
    Or Y < LBound(paths, 2) Or Y > UBound(paths, 2) Then
    v(1) = vbNullString
    v(2) = 0
Else
    v(1) = paths(X, Y)
    v(2) = pcost(X, Y)
End If
BestPath = v

End Function

Public Function PathCost(strSect As String, strPath As String, MobType As Integer) As Single

' Const MaxRadius = 20   removed efj 8/2003

Dim Cost As Single
Dim n As Integer
Dim X As Integer
Dim Y As Integer
Dim X2 As Integer
Dim Y2 As Integer

On Error Resume Next

'verify starting sector
If ParseSectors(X, Y, strSect) Then
    If MobType <> MT_SHIP And MobType <> MT_SMALL_SHIP Then
        rsSectors.Seek "=", X, Y
        If rsSectors.NoMatch Then
            PathCost = 0
            Exit Function
        End If
    End If
Else
    PathCost = 0
    Exit Function
End If

'Walk through the path computing path cost
For n = 1 To Len(strPath)
    If InStr("gyujnb", Mid$(strPath, n, 1)) > 0 Then
        ComputeDestination X, Y, X2, Y2, Mid$(strPath, n, 1)
        Cost = SectorMobCost(X2, Y2, MobType)
        If Cost = 0 Or Cost > 100 Then  'impassable sector
            PathCost = 0
            Exit Function
        End If
        PathCost = PathCost + Cost
        X = X2
        Y = Y2
    End If
Next n

End Function

Public Function MobilityWeight(strItem As String, ByVal Count As Long, Optional SourceDes As String, Optional MoveBonus As Integer) As Single

If IsMissing(MoveBonus) Then
    MoveBonus = MB_NORMAL
End If

If MoveBonus < MB_NORMAL Then
    MoveBonus = MB_NORMAL
End If

If IsMissing(SourceDes) Then
    SourceDes = vbNullString
End If

On Error GoTo Empire_Error
Dim w As Integer
Dim pb As Integer

'figure weight of items
Select Case strItem
    Case "b"
        w = 50
    Case "g"
        w = 10
    Case "r"
        w = 8
    Case "d"
        w = 5
    Case "u"
        w = 2
    Case Else
        w = 1
End Select

If SourceDes = vbNullString Then
    pb = 1
Else
    rsSectorType.Seek "=", SourceDes
    If rsSectorType.NoMatch Then
        pb = 1
    Else
        If Not (rsItems.BOF And rsItems.EOF) Then
            rsItems.Seek "=", strItem
            If Not rsItems.NoMatch Then
                Select Case rsSectorType.Fields("pack_type")
                Case 0: 'Inefficiency
                    pb = rsItems.Fields("pack_ie")
                Case 1: 'Normal
                    pb = rsItems.Fields("pack_rg")
                Case 2: 'Warehouse
                    pb = rsItems.Fields("pack_wh")
                Case 3: 'Urban
                    pb = rsItems.Fields("pack_ur")
                Case 4: 'Bank
                    pb = rsItems.Fields("pack_bk")
                End Select
            End If
        Else
            Select Case strItem
                Case "c"
                    pb = rsSectorType.Fields("pack_civ")
                Case "m"
                    pb = rsSectorType.Fields("pack_mil")
                Case "u"
                    pb = rsSectorType.Fields("pack_uw")
                Case "b"
                    pb = rsSectorType.Fields("pack_bar")
                Case Else
                    pb = rsSectorType.Fields("pack_other")
            End Select
        End If
    End If
End If

'The exact formula is:
'  mob cost = (amount) * (weight) * (path cost) / (source packing bonus)

MobilityWeight = CSng(Count * w) / CSng(pb * MoveBonus)
Exit Function

Empire_Error:
EmpireError "MobilityWeight", vbNullString, strItem

End Function


Public Function NewPop(rs As Recordset, Poptype As String) As Single

On Error GoTo Empire_Error
'if you haven't got a version, then quit
If rsVersion.BOF And rsVersion.EOF Then
    NewPop = -1
    Exit Function
End If

'must have nation vars to do right
If Maxpop <= 0 Then
    NewPop = -1
    Exit Function
End If

'calc the number of babies born
Dim babies As Single
Dim Maxbabies As Long
Dim MaxSectorPop As Long

'check max pop
MaxSectorPop = Maxpop
rsSectorType.Seek "=", rs.Fields("des")
If Not rsSectorType.NoMatch Then
    MaxSectorPop = rsSectorType.Fields("maxpop")
End If

Select Case Left$(Poptype, 1)
    Case "c"
        babies = rsVersion.Fields("ETU per update") * rsVersion.Fields("Civ Birthrate")
        babies = babies * (rs.Fields("civ") / 1000)
        Maxbabies = MaxSectorPop - rs.Fields("civ")
    
    Case "u"
        babies = rsVersion.Fields("ETU per update") * rsVersion.Fields("Uw Birthrate")
        babies = babies * (rs.Fields("uw") / 1000)
        Maxbabies = MaxSectorPop - rs.Fields("uw")
    
    Case Else
        babies = 0
        Maxbabies = 0
End Select

'make sure their aren't more than the sector can hold
If babies > Maxbabies Then
    babies = Maxbabies
End If
If babies < 0 Then
    babies = 0
End If
NewPop = babies

Exit Function

Empire_Error:
EmpireError "NewPop", vbNullString, Poptype
End Function

Public Function FoodRequired(rs As Recordset) As Integer
On Error GoTo Empire_Error

'if you haven't got a version, then quit
If rsVersion.BOF And rsVersion.EOF Then
    FoodRequired = -1
    Exit Function
End If

'If this is a blitz, there may not be any food required
If rsVersion.Fields("Food Needed") = "N" Then
    FoodRequired = 0
    Exit Function
End If

'If there are only a few mil in the sector, no food is needed
If rs.Fields("civ") = 0 And rs.Fields("uw") = 0 And rs.Fields("mil") < 11 Then
    FoodRequired = 0
    Exit Function
End If
    
Dim NewUW As Long
Dim NewCiv As Long

NewUW = Int(NewPop(rs, "u") + 0.499)
NewCiv = Int(NewPop(rs, "c") + 0.499)

'if we get an error in calculating new pop, start here
If NewUW < 0 Or NewCiv < 0 Then
    FoodRequired = -1
    Exit Function
End If


'calc the food needed to raise babies
FoodRequired = CInt(((CSng(NewUW + NewCiv) / 1000#) * rsVersion.Fields("Baby food")) + 0.499)

'add the food needed to feed the population
FoodRequired = FoodRequired + CInt((CSng((rs.Fields("civ") + rs.Fields("uw") + rs.Fields("mil") _
    + NewUW + NewCiv) / 1000#) * rsVersion.Fields("Adult food") * rsVersion.Fields("ETU per update")) + 0.499)

Exit Function

Empire_Error:
EmpireError "FoodRequired", vbNullString, CStr(rs.Fields("x")) + ":" + CStr(rs.Fields("y"))
End Function

Public Function Production(rs As Recordset) As productionDataType
On Error GoTo Empire_Error

Dim Prod As productionDataType

'Products:
'   Item      $    Lcm  Hcm Iron Dust  Oil  Rad   Tech   Production Eff.
'   Shells:   3     2    1    0    0    0    0     20    (tlev-20)/(tech-10)
'   Guns:    30     5    10   0    0    1    0     20    (tlev-20)/(tech-10)
'   Iron:     0     0    0    0    0    0    0      0    0
'   Dust:     0     0    0    0    0    0    0      0    0
'   Bars:    10     0    0    0    5    0    0      0    0
'   Food:     0     0    0    0    0    0    0      0    (tlev+10)/(tlev+20)
'   Oil:      0     0    0    0    0    0    0      0    (tlev+10)/(tlev+20)
'   Petrol    1     0    0    0    0    1    0     20    (tlev-20)/(tech-10)
'   Lcm:      0     0    0    1    0    0    0      0    (tlev+10)/(tlev+20)
'   Hcm:      0     0    0    2    0    0    0      0    (tlev+10)/(tlev+20)
'   Rad:      2     0    0    0    0    0    0     40    (tlev-40)/(tlev-30)
'   Educate:  9     1    0    0    0    0    0      0    0
'   Happy:    9     1    0    0    0    0    0      0    0
'   Tech:  5*ETUS  10    0    0    1    5    0      0    (edu-5)/(edu+5)
'   Resrch:  90    10    0    0    1    5    0      0    (edu-5)/(edu+5)

'if you haven't got a version, then quit
If rsVersion.BOF And rsVersion.EOF Then
    Prod.prodValidFlag = False: Production = Prod
    Exit Function
End If

Dim NewUW As Long
Dim NewCiv As Long
Dim NewEff As Integer
Dim newdes As String
Dim civs As Integer
Dim uws As Integer
Dim Mil As Integer
Dim buildavail As Long
Dim baseavail As Long
Dim maxavail As Long
Dim work_factor As Single
Dim ETU_factor As Single
Dim workcost As Integer
Dim buildcost As Long
Dim avail As Long
Dim um_avail As Long
Dim existingavail As Long

'+++++++++++++++++++++++++++++++++++++++++++++++
' Part 1 - calculate The sectors new efficency
'+++++++++++++++++++++++++++++++++++++++++++++++

NewUW = Int(NewPop(rs, "u"))
NewCiv = Int(NewPop(rs, "c"))

'if we get an error in calculating new pop, quit here
If NewUW < 0 Or NewCiv < 0 Then
    Prod.prodValidFlag = False: Production = Prod
    Exit Function
End If

'get the record for the sector type.  If we have an error, quit here
rsSectorType.Seek "=", rs.Fields("des")
If rsSectorType.NoMatch Then
    Prod.prodValidFlag = False: Production = Prod
    Exit Function
End If
    
'Compute safe civilians
Dim MaxSectorPop As Long
Dim SafeCivs As Long
MaxSectorPop = rsSectorType.Fields("maxpop")
SafeCivs = (MaxSafeCivs / Maxpop) * (rsSectorType.Fields("maxpop"))

If frmOptions.b2K4 Then
    civs = 5000
Else
    civs = rs.Fields("civ")
End If
If civs > MaxSectorPop Then
    civs = MaxSectorPop
End If
uws = rs.Fields("uw")
If uws > MaxSectorPop Then
    uws = MaxSectorPop
End If
Mil = rs.Fields("mil")
If Mil > MaxSectorPop Then
    Mil = MaxSectorPop
End If

If rs.Fields("work") = 100 Then
    work_factor = 1
Else
    work_factor = CSng(rs.Fields("work")) / 100#
End If

If Len(Trim$(rs.Fields("sdes"))) = 0 Then
    existingavail = rs.Fields("avail")
    If existingavail > rsVersion.Fields("Rollover Avail") Then
        existingavail = rsVersion.Fields("Rollover Avail")
    End If
Else
    existingavail = 0
End If

'Note - Ken went through avail routine in 4.2.8 and documented where rounding was done.
'changed routines here to simulate that rounding
'080404 rjk: Removed the 999 limit on worker avail to match the server
ETU_factor = rsVersion.Fields("ETU per update")
baseavail = Int(Int(((work_factor * civs) + uws + (Mil * 0.4)) * ETU_factor) / 100) + existingavail
avail = baseavail + Int(Int(((work_factor * NewCiv) + NewUW) * ETU_factor) / 100)
maxavail = baseavail + Int(Int(((work_factor * (MaxSectorPop - civs)) + NewUW) * ETU_factor) / 100)
um_avail = CLng((Int((uws + Mil * 0.4) * ETU_factor) + Int(NewUW * ETU_factor)) / 100)

'check the amount of avail used in building and reducing eff
'1 avail reduces eff by 4
buildavail = Int(avail / 2)
avail = avail - buildavail

'calc how much avail will be used reducing the old sector
If Len(Trim$(rs.Fields("sdes"))) > 0 And rs.Fields("eff") > 0 _
        And rs.Fields("off") = 0 Then
    workcost = Int((rs.Fields("eff") + 3) / 4)
    If workcost > buildavail Then
        workcost = buildavail
        newdes = rs.Fields("des")
    Else
        newdes = rs.Fields("sdes")
    End If
    buildavail = buildavail - workcost
    buildcost = workcost
    NewEff = rs.Fields("eff") - (workcost * 4)
    If NewEff < 0 Then
        NewEff = 0
    End If
Else
    NewEff = rs.Fields("eff")
    newdes = rs.Fields("des")
    buildcost = 0
End If

'calc how much avail will be used to increase the new sector
If rs.Fields("off") = 0 Then 'sector not stopped
    If Not (rsVersion.BOF And rsVersion.EOF) Then
        rsVersion.MoveFirst
        If (100 - NewEff) > rsVersion.Fields("Eff gain - Sect").Value Then
            workcost = rsVersion.Fields("Eff gain - Sect").Value
        Else
            workcost = 100 - NewEff
        End If
    Else
        workcost = 100 - NewEff
    End If
Else
    workcost = 0
End If

'if sector is becoming a fort, make sure there are hcm's in
'sector to bring it up.
If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.Seek "=", "i", newdes
    If Not rsBuild.NoMatch Then
        If rsBuild.Fields("lcm") > 0 Then
            If workcost > rs.Fields("lcm") / rsBuild.Fields("lcm") Then
                workcost = rs.Fields("lcm") / rsBuild.Fields("lcm")
            End If
        End If
        If rsBuild.Fields("hcm") > 0 Then
            If workcost > rs.Fields("hcm") / rsBuild.Fields("hcm") Then
                workcost = rs.Fields("hcm") / rsBuild.Fields("hcm")
            End If
        End If
        buildcost = buildcost + (workcost * rsBuild.Fields("stat1"))
    Else
        buildcost = buildcost + workcost
    End If
Else
    buildcost = buildcost + workcost
End If

If workcost > buildavail Then
    workcost = buildavail
End If
buildavail = buildavail - workcost
NewEff = NewEff + workcost
maxavail = maxavail - buildcost

'add left over buildavail back to avail
avail = avail + buildavail

'Set the production array with current information
Prod.NewEff = NewEff          'Prod(1) = NewEff    'new efficency
Prod.ProdAmount = 0           'Prod(2) = 0         'production amount
Prod.MaxProdAmount = 0        'Prod(3) = 0         'max production amount
Prod.ExcessCiv = 0            'Prod(4) = 0         'excess civ
Prod.newdes = newdes          'Prod(5) = newdes    'designation
Prod.unitsConsumed = 0        'Prod(6) = 0         'units consumed
Prod.ItemProduced = vbNullString        'Prod(7) = vbnullstring        'item produced
Prod.BuildAvailCost = buildcost 'Prod(8) = buildcost 'build cost in avail
Prod.ResourceEff = 0          'Prod(9) = 0         'resource efficency
Prod.LevelEff = 0             'Prod(10) = 0        'level efficency
Prod.ProductionAvail = avail  'Prod(11) = avail    'production avail
Prod.prodValidFlag = True

'if you haven't reached 60%, or if the sector is stopped or
'if  sector is enlistment and not owned
'then nothing is produced
If NewEff < 60 Or rs.Fields("off") = 1 Or _
    (newdes = "e" And rs.Fields("*") = "*") Then
    'check against the max safe civs - all others are excess
    If (rs.Fields("civ") - SafeCivs) > Prod.ExcessCiv Then
        Prod.ExcessCiv = rs.Fields("civ") - SafeCivs
    End If
    Production = Prod
    Exit Function
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++
' Part 2 - calculate what the sector will produce
'+++++++++++++++++++++++++++++++++++++++++++++++++

Dim pe As Single
Dim sector_pe As Single
Dim resource_pe As Single
Dim level_pe As Single
Dim unit_work As Integer
Dim unit_prod As Integer
Dim material_limit As Integer
Dim worker_limit As Integer
Dim max_limit As Integer
Dim max_produce As Integer
Dim n As Integer

'get the record for the new sector type.  If we have an error, quit here
rsSectorType.Seek "=", newdes
If rsSectorType.NoMatch Then
    Prod.prodValidFlag = False: Production = Prod
    Exit Function
End If

'initial production efficiency is the sector efficiency
sector_pe = NewEff / 100#
unit_prod = 100
max_produce = 999

'compute level production efficency
If Len(rsSectorType.Fields("level")) = 0 Then
    level_pe = 1
Else
    Select Case Left$(rsSectorType.Fields("level"), 3)
    Case "tec"
        level_pe = TechLevel
        
    Case "edu"
        level_pe = Education
        
    Case "hap"
        level_pe = Happiness
    
    Case "res"
        level_pe = Research

    End Select
        
    If level_pe >= rsSectorType.Fields("min") Then
        If level_pe - rsSectorType.Fields("min") + rsSectorType.Fields("lag") <> 0 Then
            level_pe = ((level_pe - rsSectorType.Fields("min")) / _
                (level_pe - rsSectorType.Fields("min") + rsSectorType.Fields("lag")))
        Else
            level_pe = 1
        End If
    Else
        level_pe = 0
    End If
End If

If level_pe < 0 Then
    level_pe = 0
End If


resource_pe = 1
Prod.ItemProduced = rsSectorType.Fields("pcode")
If rsSectorType.Fields("product") <> " " Then
    material_limit = 9999
Else
    material_limit = 0
End If
unit_prod = rsSectorType.Fields("eff")
unit_work = rsSectorType.Fields("use1n") + rsSectorType.Fields("use2n") + rsSectorType.Fields("use3n")
If unit_work = 0 Then
    unit_work = 1
Else
    Dim usen As Integer
    Dim uses As String
    Dim limit As Long
    material_limit = 9999
    For n = 1 To 3
        Select Case n
            Case 1
                usen = rsSectorType.Fields("use1n")
                uses = rsSectorType.Fields("use1s")
            Case 2
                usen = rsSectorType.Fields("use2n")
                uses = rsSectorType.Fields("use2s")
            Case 3
                usen = rsSectorType.Fields("use3n")
                uses = rsSectorType.Fields("use3s")
        End Select
        
        If usen > 0 And Len(uses) > 0 Then
            Select Case uses
                Case "b"
                    limit = rs.Fields("bar")
                Case "c"
                    limit = rs.Fields("civ")
                Case "d"
                    limit = rs.Fields("dust")
                Case "f"
                    limit = rs.Fields("food")
                Case "g"
                    limit = rs.Fields("gun")
                Case "h"
                    limit = rs.Fields("hcm")
                Case "i"
                    limit = rs.Fields("iron")
                Case "l"
                    limit = rs.Fields("lcm")
                Case "m"
                    limit = rs.Fields("mil")
                Case "o"
                    limit = rs.Fields("oil")
                Case "p"
                    limit = rs.Fields("pet")
                Case "r"
                    limit = rs.Fields("rad")
                Case "s"
                    limit = rs.Fields("shell")
                Case "u"
                    limit = rs.Fields("uw")
                Case Else
                    limit = 9999
            End Select
            limit = limit / usen
        
            If limit < material_limit Then
                material_limit = limit
            End If
        End If
    Next n
End If

'Handle the exceptions
Select Case newdes
    Case "m"    'iron
        resource_pe = (rs.Fields("min") / 100)
        material_limit = (rs.Fields("min") * 100)
    
    Case "g", "^"   'dust
        resource_pe = (rs.Fields("gold") / 100)
        material_limit = (rs.Fields("gold") * 100)
        If rsSectorType.Fields("dep") > 0 Then
            material_limit = material_limit / rsSectorType.Fields("dep")
        End If
  
  'oil, ":" is for IW5 only
    Case "o", ":"
        If rsSectorType.Fields("product") = "oil" Then
            resource_pe = (rs.Fields("ocontent") / 100)
            material_limit = (rs.Fields("ocontent") * 100)
            If rsSectorType.Fields("dep") > 0 Then
                material_limit = material_limit / rsSectorType.Fields("dep")
            End If
        Else
            If frmOptions.bMaxProd Then
                material_limit = 0
                level_pe = 0
            End If
        End If
        
    Case "~" '210704 rjk: ~ StarWars II, generalize uses sector type, 20080512 HeavyMetal as well
        If rsSectorType.Fields("product") = "oil" Then
            material_limit = (rs.Fields("ocontent") * 100)
            If rsSectorType.Fields("dep") > 0 Then
                material_limit = material_limit / rsSectorType.Fields("dep")
            End If
        ElseIf rsSectorType.Fields("product") = "food" Then
            resource_pe = (rs.Fields("fert") / 100)
            material_limit = (rs.Fields("fert") * 100)
        Else
            If frmOptions.bMaxProd Then
                material_limit = 0
                level_pe = 0
            End If
        End If
        
    Case "u"    'rads
        resource_pe = (rs.Fields("uran") / 100)
        material_limit = (rs.Fields("uran") * 100)
        If rsSectorType.Fields("dep") > 0 Then
            material_limit = material_limit / rsSectorType.Fields("dep")
        End If
  
    Case "a"    'food
        resource_pe = (rs.Fields("fert") / 100)
        
    
    Case "j"    'lcms
        'Pat sometimes run game with no iron - min content sectors
        'directly produce lcms and hcms
        If Not (rsNation.BOF And rsNation.EOF) Then
            rsNation.MoveFirst
            If rsNation.Fields("NoIron").Value Then
                pe = pe * (rs.Fields("min") / 100)
                material_limit = 9999
            End If
        End If
    
    Case "k"    'hcms
        'Pat sometimes run game with no iron - min content sectors
        'directly produce lcms and hcms
        If Not (rsNation.BOF And rsNation.EOF) Then
            rsNation.MoveFirst
            If rsNation.Fields("NoIron").Value Then
                resource_pe = (rs.Fields("min") / 100)
                material_limit = 9999
                unit_work = 1
            End If
       End If
    
    'drk@unxsoft.com 10/15/02: this is silly.  Why isn't this in the database?
    'why is this field even here?  It's not used anywhere?  Why just a single
    'letter?  Why does happiness and education have codes, but tech and res
    'don't?  I think I'll be changing this a bunch... FIXME.
    Case "l"    'education
        Prod.ItemProduced = "e"
    
    Case "p"   'happiness
        Prod.ItemProduced = "x"
    
    '"v" is for IW5 only!, 260104 rjk: "y" is for St@r W@rs only
    '"q" is for Heavy Metal II only
    Case "t", "r", "b", "%", "i", "d", ":", "v", "y", "q"  ' ':', 'v' for IW5 only!
        'nothing special needed
    
    Case "h", "!", "*", "n", "f" 'These produce avail, see below

    Case "e"    'enlistment - this is special
        Dim recruit_rate As Single
        
        recruit_rate = (rsVersion.Fields("ETU per update") / 20)
        Prod.ProdAmount = Int((Mil + 10) * recruit_rate)
        If Prod.ProdAmount > Int((civs + NewCiv) / 2) - Mil Then
            Prod.ProdAmount = Int((civs + NewCiv) / 2) - Mil
        End If
        
        'compute max amount
        Prod.MaxProdAmount = Int((MaxSectorPop / 2 - recruit_rate * 10) / (recruit_rate + 1))
        Prod.MaxProdAmount = Int((Prod.MaxProdAmount + 10) * recruit_rate)
        
        'compute excess civs
        Prod.ExcessCiv = (Int((Mil + 10) * recruit_rate) + Mil) * 2
        
        'Take into account the birth rate
        Prod.ExcessCiv = CInt((Prod.ExcessCiv / (1 + ((rsVersion.Fields("ETU per update") * rsVersion.Fields("Civ Birthrate")) / 1000))) + 0.499)
        Prod.ExcessCiv = civs - Prod.ExcessCiv

        'check against the max safe civs - all others are excess
        If (rs.Fields("civ") - SafeCivs) > Prod.ExcessCiv Then
            Prod.ExcessCiv = rs.Fields("civ") - SafeCivs
        End If

        Prod.ItemProduced = "m"
        Production = Prod
        Exit Function

    Case Else
        ' modification by thomas lullier
        ' max pop vs max prod
        'note by daniel: this case should only apply to sectors that don't produce
        'anything (e.g. highways, radar stations, etc.)  If a sector does produce
        'something, but doesn't need any special case, lump it in with t, r, and b
        'above. 10/15/02
        '120303 rjk: Switched options to frmOptions and boolean options
        If frmOptions.bMaxProd Then
            material_limit = 0
            level_pe = 0
        End If

End Select

Select Case newdes
    Case "h", "!", "*", "n", "f"
        'these 'produce' avail (well, everthing produces avail, but for these it gets
        'reported in the Pd: field, b/c that's the most interesting thing for them
        Prod.ProdAmount = avail
        Prod.MaxProdAmount = maxavail
        Prod.prodMaxMaterialLimit = False
    
    Case Else

        pe = sector_pe * resource_pe

        'determine if resources or workers are the limitation
        worker_limit = CInt(avail * pe / unit_work)
        If material_limit < worker_limit Then
            worker_limit = material_limit
        End If

        max_limit = CInt(maxavail * pe / unit_work)

        'tech and research keep decimal.  All others round off
        If newdes = "t" Or newdes = "r" Then
            Prod.ProdAmount = CSng(Format$(level_pe * worker_limit, "##0.0"))
            Prod.MaxProdAmount = CSng(Format$(level_pe * max_limit, "##0.0"))
        Else
            Prod.ProdAmount = CInt(level_pe * CInt(unit_prod * 0.01 * worker_limit))
            Prod.MaxProdAmount = CInt(level_pe * CInt(unit_prod * 0.01 * max_limit))
        End If
        Prod.unitsConsumed = worker_limit

        'Check against the max
        If Prod.MaxProdAmount > max_produce And Not frmOptions.bEE9 And Not frmOptions.bSPAtlantis Then
            Prod.MaxProdAmount = max_produce
        End If

        ' additionnal check against the max using material_limit
        ' by thomas lullier
        If Prod.MaxProdAmount > material_limit Then
            Prod.MaxProdAmount = material_limit
            Prod.prodMaxMaterialLimit = True
        Else
            Prod.prodMaxMaterialLimit = False
        End If

        'Check against the max
        If Prod.ProdAmount > max_produce And Not frmOptions.bEE9 And Not frmOptions.bSPAtlantis Then
            Prod.ProdAmount = max_produce
        End If

        'Check against the max
        If material_limit > 999 And level_pe > 0 And unit_prod > 0 And _
           Not frmOptions.bEE9 And Not frmOptions.bSPAtlantis Then
            material_limit = CInt(999 / (level_pe * unit_prod * 0.01))
        End If
    End Select

'+++++++++++++++++++++++++++++++++++++++++++++++++
' Part 3 - calculate extra civs or civs needed.
'+++++++++++++++++++++++++++++++++++++++++++++++++

Dim avail_needed As Long
Dim civ_needed As Long
'calculate the civs needed to produce everything you can
If pe > 0 Then
    avail_needed = CLng(((material_limit * unit_work) / pe) + 0.499)
Else
    avail_needed = 0
End If

'check and make sure you have at least 2 * buildcost
If avail_needed < buildcost Then
    avail_needed = buildcost
End If
avail_needed = avail_needed + buildcost

If avail_needed > 0 Then
    If avail_needed > existingavail Then
        avail_needed = avail_needed - existingavail
    Else
        avail_needed = 0
    End If
End If
    
'Now calculate the number of civs needed to generate that avail
civ_needed = CLng(((avail_needed - um_avail) / (rsVersion.Fields("ETU per update") / 100)) + 0.499)

'Take into account the birth rate
civ_needed = CLng((civ_needed / (1 + ((rsVersion.Fields("ETU per update") * rsVersion.Fields("Civ Birthrate")) / 1000))) + 0.499)

If (SafeCivs > 0) And (civ_needed > SafeCivs) Then
    civ_needed = SafeCivs
End If

If civ_needed < 1 Then
    If civs > 1 Then
        civ_needed = 1
    Else
        civ_needed = 0
    End If
End If

Prod.ExcessCiv = civs - civ_needed

'check against the max safe civs - all others are excess
If (rs.Fields("civ") - SafeCivs) > Prod.ExcessCiv Then
    Prod.ExcessCiv = rs.Fields("civ") - SafeCivs
End If

Prod.ResourceEff = resource_pe
Prod.LevelEff = level_pe
 
Production = Prod
Exit Function

Empire_Error:
EmpireError "Production", vbNullString, CStr(rs.Fields("x")) + ":" + CStr(rs.Fields("y"))
End Function

Public Function GetUnitList(strUnit As String, strType As String) As Variant
On Error GoTo Empire_Error

Dim sx As Integer
Dim sy As Integer
Dim n As Integer
Dim units As Variant
Dim nID As Integer
Dim hldIndex As Variant
Dim rs As Recordset
Dim Fleet As String

If Len(Trim$(strUnit)) = 0 Then
    GetUnitList = 0
    Exit Function
End If

'set indexs
Select Case strType
    Case "l"
        Set rs = rsLand
        Fleet = "army"
    Case "p"
        Set rs = rsPlane
        Fleet = "wing"
    Case "s"
        Set rs = rsShip
        Fleet = "flt"
    Case "n"
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            Set rs = rsNuke
            Fleet = ""
        Else
            GetUnitList = 0
            Exit Function
        End If
    Case Else
        GetUnitList = 0
        Exit Function
End Select

hldIndex = rs.Index
rs.Index = "PrimaryKey"

If InStr(strUnit, "/") > 0 Then
    'list of units
    GetUnitList = ParseUnitString(strUnit)

'if it has a comma, its probably a sector number
ElseIf InStr(strUnit, ",") > 0 Then
    If Not ParseSectors(sx, sy, strUnit) Then
        'who knows at this point - just quit
        GetUnitList = 0
        rs.Index = hldIndex
        Exit Function
    End If
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        ReDim units(0)
        n = 0
        While (Not rs.EOF)
            If sx = rs.Fields("x").Value And _
               sy = rs.Fields("y").Value Then
                n = n + 1
                ReDim Preserve units(n)
                units(n) = rs.Fields("id").Value
            End If
            rs.MoveNext
        Wend
        GetUnitList = units
    Else
        GetUnitList = 0
        rs.Index = hldIndex
        Exit Function
    End If
    
'if it's a single letter, its a fleet
'093003 rjk: Expand the check for Fleet identifiers to include A-Z and ~ as well.
ElseIf (Len(strUnit) = 1) And (((strUnit >= "a") And (strUnit <= "z")) Or _
       ((strUnit >= "A") And (strUnit <= "Z")) Or (strUnit = "~")) And Len(Fleet) > 0 Then
    rs.MoveFirst
    ReDim units(0)
    n = 0
    While (Not rs.EOF)
        If rs.Fields(Fleet) = strUnit Then
            n = n + 1
            ReDim Preserve units(n)
            units(n) = rs.Fields("id").Value
        End If
        rs.MoveNext
    Wend
    GetUnitList = units
Else
    'should just be a ship number
    nID = 9999
    On Error Resume Next
    nID = CInt(strUnit) 'in case not a valid number - fall through
    On Error GoTo Empire_Error
    rs.Seek "=", nID
    If Not rs.NoMatch Then
        ReDim units(1)
        units(1) = nID
        GetUnitList = units
    Else
        'who knows at this point - just quit
        GetUnitList = 0
    End If
End If
rs.Index = hldIndex
Exit Function

Empire_Error:
EmpireError "GetUnitList", vbNullString, strUnit
End Function

Public Function SectorDefenseStrength(sct_type As String) As Single
On Error GoTo Empire_Error

rsSectorType.Seek "=", sct_type
If rsSectorType.NoMatch Then
    SectorDefenseStrength = 1.5
Else
    SectorDefenseStrength = rsSectorType.Fields("defense")
End If

Exit Function
Empire_Error:
EmpireError "SectorDefenseStrength", vbNullString, sct_type
End Function

Public Function SectorEffAfterDamage(dam As Integer, sct_type As String, sct_eff As Integer) As Integer
On Error GoTo Empire_Error

'Start with the amount of damage to apply (dam).
'Need sector type to get defensive strength (dstr)
'If unknown, figure a sector is 100% sector efficiency to start (sct_effic)

Dim sector_strength As Single
Dim d_dstr As Single
Dim nDam As Long
Dim sct_effic As Single
Dim eff_lost As Single
Dim neff_lost As Integer

sct_effic = CSng(sct_eff)
sector_strength = 1#
If sct_type = "^" Then
   sector_strength = 2#
Else
    sector_strength = 1#
End If

d_dstr = SectorDefenseStrength(sct_type)

'100203 rjk: When firing on a Sea sector the damage goes to 100% to 0%.
'            The Defense strength for the sea is zero which causes a divide by zero error so
'            just return 0 to begin with.
If d_dstr = 0 Then
    SectorEffAfterDamage = 0
    Exit Function
End If
   
sector_strength = sector_strength + (d_dstr - sector_strength) * (sct_effic / 100#)

If (sector_strength > d_dstr) Then
   sector_strength = d_dstr
End If

nDam = CLng(dam * (1# / sector_strength))
eff_lost = (sct_effic * (CSng(nDam) / (nDam + 100)))
neff_lost = CInt(eff_lost)

'Make sure we always round down.
If neff_lost > eff_lost Then
    neff_lost = neff_lost - 1
End If

SectorEffAfterDamage = sct_eff - neff_lost
If SectorEffAfterDamage < 0 Then
    SectorEffAfterDamage = 0
End If
Exit Function

Empire_Error:
EmpireError "SectorEffAfterDamage", vbNullString, sct_type + ":" + CStr(dam) + ":" + CStr(sct_eff)
End Function

Function FortRange(Tech As Single, Optional Eff As Variant) As Single
On Error GoTo Empire_Error

'default is 100
If IsMissing(Eff) Then
    Eff = 100
End If

FortRange = 7 * (Tech + 50) / (Tech + 200)
'Adjust for firing range in version
If Not (rsVersion.BOF And rsVersion.EOF) Then
    rsVersion.MoveFirst
    FortRange = FortRange * rsVersion.Fields("Fire range")
End If

If Eff > 59 Then
    FortRange = FortRange + 1
End If

Exit Function

Empire_Error:
EmpireError "FortRange", vbNullString, CStr(Tech) + ":" + CStr(Eff)
End Function

Function UnitFireRange(Tech As Single, range As Integer) As Single
On Error GoTo Empire_Error
UnitFireRange = range * (Tech + 50) / (Tech + 200) / 2
Exit Function

Empire_Error:
EmpireError "UnitFireRange", vbNullString, CStr(Tech) + ":" + CStr(range)
End Function

Function RadarRange(Tech As Single) As Integer
On Error GoTo Empire_Error
RadarRange = Int(16 * (Tech + 50) / (Tech + 200))
Exit Function

Empire_Error:
EmpireError "RadarRange", vbNullString, CStr(Tech)
End Function
Function CoastwatchRange(Tech As Single) As Integer
On Error GoTo Empire_Error
CoastwatchRange = Int(14 * (Tech + 50) / (Tech + 200))
Exit Function

Empire_Error:
EmpireError "CoastwatchRange", vbNullString, CStr(Tech)
End Function

Function ComputeRangeStat(Tech As Single, Class As String, id As String) As Integer
On Error GoTo Empire_Error

Dim buildtech As Single
Dim frnge As Integer

rsBuild.Seek "=", Class, id
If rsBuild.NoMatch Then
    ComputeRangeStat = 0
    Exit Function
End If

buildtech = rsBuild.Fields("tech").Value

If rsBuild.Fields("base frnge") > 0 Then
    If Class = "s" Then
        ComputeRangeStat = Fix(ComputeShipRange(rsBuild.Fields("base frnge"), _
                CInt(buildtech), CInt(Tech)))
    Else
        ComputeRangeStat = Fix(ComputeLandFiringRange(rsBuild.Fields("base frnge"), _
                CInt(buildtech), CInt(Tech)))
    End If
Else
    If Class = "s" Then
        '0404040 rjk: Replace with show information with the calculation
        'techdiff = (int)(tlev - mp->m_tech);
        'SHP_RNG(mp->m_frnge, techdiff)
        '#define SHP_RNG(b, t) (t ? (b * (logx((double)t, (double)35.0) < 1.0 ? 1.0 : \
        '             logx((double)t, (double)35.0))) : b)
        'drk 9/7/03.  removed reference to ship stats table
        'fixme: really should have a routine to convert "range" to "stat5"
        'instead of hardcoding it.
        frnge = rsBuild.Fields("stat5")
    End If
    
    If Class = "l" Then
        'ditto
        frnge = rsBuild.Fields("stat8")
    End If
    
    If Tech <= (buildtech + 35) Then
        ComputeRangeStat = frnge
    Else
        ComputeRangeStat = Int(frnge * Log(Tech - buildtech) / Log(35))
    End If
End If
Exit Function

Empire_Error:
EmpireError "ComputeRangeStat", vbNullString, CStr(Tech) + ":" + Class + ":" + id
End Function


Public Function LogX(dLog As Double, dBase As Double) As Double
'LogX = Log10(dLog) / Log10(dBase)
LogX = Log(dLog) / Log(dBase)
End Function

Public Function Log10(dLog As Double) As Double
Log10 = Log(dLog) / Log(10#)
End Function

Public Function ComputeShipDefense(iArmor As Integer, iBaseTech As Integer, iTech As Integer) As Double
'(short)SHP_DEF(mp->m_armor, techdiff),
'#define SHP_DEF(b, t) (t ? (b * (logx((double)t, (double)40.0) < 1.0 ? 1.0 : \
'                logx((double)t, (double)40.0))) : b)
ComputeShipDefense = ComputeLogFunction(iArmor, iBaseTech, iTech, 40#)
End Function

Public Function ComputeShipSpeed(iSpeed As Integer, iBaseTech As Integer, iTech As Integer) As Double
'(short)SHP_SPD(mp->m_speed, techdiff),rsNation.Fields("TechLevel"))
'#define SHP_SPD(b, t) (t ? (b * (logx((double)t, (double)35.0) < 1.0 ? 1.0 : \
'                 logx((double)t, (double)35.0))) : b)
ComputeShipSpeed = ComputeLogFunction(iSpeed, iBaseTech, iTech, 35#)
End Function

Public Function ComputeShipRange(iFrnge As Integer, iBaseTech As Integer, iTech As Integer) As Double
'(short)SHP_RNG(mp->m_frnge, techdiff),
'#define SHP_RNG(b, t) (t ? (b * (logx((double)t, (double)35.0) < 1.0 ? 1.0 : \
'                 logx((double)t, (double)35.0))) : b)
ComputeShipRange = ComputeLogFunction(iFrnge, iBaseTech, iTech, 35#)
End Function

Public Function ComputeShipFiringRange(iGlim As Integer, iBaseTech As Integer, iTech As Integer) As Double
'(short)SHP_FIR(mp->m_glim, techdiff));
'#define SHP_FIR(b, t) (t ? (b * (logx((double)t, (double)60.0) < 1.0 ? 1.0 : \
'                 logx((double)t, (double)60.0))) : b)
ComputeShipFiringRange = ComputeLogFunction(iGlim, iBaseTech, iTech, 60#)
End Function

Public Function ComputeShipVisibility(iVisib As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'(short)SHP_VIS(mp->m_visib, techdiff),
'#define SHP_VIS(b, t) (b * (1 - (sqrt((double)t) / 50)))
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputeShipVisibility = iVisib * (1 - (Sqr(iTechDiff) / 50))
End Function
       
Public Function ComputePlaneAttDef(iAttDef As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'(int)PLN_ATTDEF(pp->pl_att, (int)(tlev - pp->pl_tech)),
'(int)PLN_ATTDEF(pp->pl_def, (int)(tlev - pp->pl_tech)),
'#define PLN_ATTDEF(b, t) (b + ((b?1:0) * ((t/20)>10?10:(t/20))))
iTechDiff = iTech - iBaseTech
ComputePlaneAttDef = iAttDef
If iAttDef <> 0 Then
    If (iTechDiff / 20) > 10 Then
        ComputePlaneAttDef = ComputePlaneAttDef + 10
    Else
        ComputePlaneAttDef = ComputePlaneAttDef + (iTechDiff / 20)
    End If
End If
End Function

Public Function ComputePlaneRange(iRange As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'(int)PLN_RAN(pp->pl_range, (int)(tlev - pp->pl_tech)),
'#define PLN_RAN(b, t) (t ? (b + (logx((double)t, (double)2.0))) : b)
iTechDiff = iTech - iBaseTech
ComputePlaneRange = iRange
If iTechDiff > 0 Then
    ComputePlaneRange = ComputePlaneRange + LogX(CDbl(iTechDiff), 2#)
End If
End Function

Public Function ComputePlaneAccuracy(iAccuracy As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'(int)PLN_ACC(pp->pl_acc, (int)(tlev - pp->pl_tech)),
'#define PLN_ACC(b, t) (b * (1.0 - (sqrt((double)t) / 50.)))
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputePlaneAccuracy = iAccuracy * (1# - (Sqr(iTechDiff) / 50#))
End Function

Public Function ComputePlaneLoad(iLoad As Integer, iBaseTech As Integer, iTech As Integer) As Double
'(int)PLN_LOAD(pp->pl_load, (int)(tlev - pp->pl_tech)),
'#define PLN_LOAD(b, t) (t ? (b * (logx((double)t, (double)50.0) < 1.0 ? 1.0 : \
'                  logx((double)t, (double)50.0))) : b)
ComputePlaneLoad = ComputeLogFunction(iLoad, iBaseTech, iTech, 50#)
End Function

Public Function ComputeLandAttDef(dAttDef As Double, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'#define LND_ATTDEF(b, t) (((b) * (1.0 + ((sqrt((double)(t)) / 100.0) * 4.0))) \
'                          > 127 ? 127 : \
'                          ((b) * (1.0 + ((sqrt((double)(t)) / 100.0) * 4.0))))
'(float)LND_ATTDEF(lcp->l_att, ourtlev)
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputeLandAttDef = dAttDef * (1# + ((Sqr(CDbl(iTechDiff)) / 100#) * 4#))
If ComputeLandAttDef > 127# Then
    ComputeLandAttDef = 127#
End If
End Function

Public Function ComputeLandVulerability(iVulerability As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'#define LND_VUL(b, t) ((b * (1.0 - ((sqrt((double)t) / 100.0) * 1.1))) < 0\
'               ? 0 : (b * (1.0 - ((sqrt((double)t) / 100.0) * 1.1))))
'(int)LND_VUL(lcp->l_vul, ourtlev)
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputeLandVulerability = CDbl(iVulerability) * (1# - ((Sqr(iTechDiff) / 100#) * 1.1))
If ComputeLandVulerability <= 0# Then
    ComputeLandVulerability = 0#
End If
End Function

Public Function ComputeLandSpeed(iSpeed As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'#define LND_SPD(b, t) ((b * (1.0 + ((sqrt((double)t) / 100.0) * 2.1))) > 127\
'               ? 127 : (b * (1.0 + ((sqrt((double)t) / 100.0) * 2.1))))
'(int)LND_SPD(lcp->l_spd, ourtlev)
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputeLandSpeed = CDbl(iSpeed) * (1# + ((Sqr(iTechDiff) / 100#) * 2.1))
If ComputeLandSpeed >= 127# Then
    ComputeLandSpeed = 127#
End If
End Function

Public Function ComputeLandFiringRange(iFrnge As Integer, iBaseTech As Integer, iTech As Integer) As Double
'#define LND_FRG(b, t) ((t) ? \
'                       ((b) * (logx((double)(t), (double)35.0) < 1.0 ? 1.0 : \
'                               logx((double)(t), (double)35.0))) : (b))
'(int)LND_FRG(lcp->l_frg, ourtlev)
ComputeLandFiringRange = ComputeLogFunction(iFrnge, iBaseTech, iTech, 35#)
End Function

Public Function ComputeLandAccuracy(iAccuracy As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'#define LND_ACC(b, t) ((b * (1.0 - ((sqrt((double)t) / 100.0) * 1.1))) < 0\
'               ? 0 : (b * (1.0 - ((sqrt((double)t) / 100.0) * 1.1))))
'(int)LND_ACC(lcp->l_acc, ourtlev)
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputeLandAccuracy = CDbl(iAccuracy) * (1# - ((Sqr(iTechDiff) / 100#) * 1.1))
If ComputeLandAccuracy <= 0# Then
    ComputeLandAccuracy = 0#
End If
End Function

Public Function ComputeLandDamage(iDamage As Integer, iBaseTech As Integer, iTech As Integer) As Double
'#define LND_DAM(b, t) ((t) ? \
'                       ((b) * (logx((double)(t), (double)60.0) < 1.0 ? 1.0 : \
'                               logx((double)(t), (double)60.0))) : (b))
'        '(int)LND_DAM(lcp->l_dam, ourtlev)
ComputeLandDamage = ComputeLogFunction(iDamage, iBaseTech, iTech, 60#)
End Function

Public Function ComputeLandAmmo(iAmmo As Integer, iDamage As Integer, iBaseTech As Integer, iTech As Integer) As Double
'#define LND_AMM(b, d, t) ((b) ? ((LND_DAM((d), (t)) / 2) + 1) : 0)
'(int)LND_AMM(lcp->l_ammo, lcp->l_dam, ourtlev)
If iAmmo <> 0 Then
    ComputeLandAmmo = (ComputeLandDamage(iDamage, iBaseTech, iTech) / 2#) + 1#
Else
    ComputeLandAmmo = 0#
End If
End Function
    
Public Function ComputeLandAAF(iAAF As Integer, iBaseTech As Integer, iTech As Integer) As Double
Dim iTechDiff As Integer
'#define LND_AAF(b, t) ((b * (1.0 + ((sqrt((double)t) / 100.0) * 3.0))) > 127\
'               ? 127 : (b * (1.0 + ((sqrt((double)t) / 100.0) * 3.0))))
'(int)LND_AAF(lcp->l_aaf, ourtlev)
iTechDiff = CalculateTechDiff(iTech, iBaseTech)
ComputeLandAAF = CDbl(iAAF) * (1# + ((Sqr(iTechDiff) / 100#) * 3#))
If ComputeLandAAF >= 127# Then
    ComputeLandAAF = 127#
End If
End Function

Private Function ComputeLogFunction(iBase As Integer, iBaseTech As Integer, iTech As Integer, dScale As Double) As Double
Dim iTechDiff As Integer
ComputeLogFunction = CDbl(iBase)
iTechDiff = iTech - iBaseTech
If iTechDiff <= 0 Then
    Exit Function
End If
If LogX(CDbl(iTechDiff), dScale) > 1# Then
    ComputeLogFunction = ComputeLogFunction * LogX(CDbl(iTechDiff), dScale)
End If
End Function

Private Function CalculateTechDiff(iNationTech, iBaseTech) As Integer
CalculateTechDiff = iNationTech - iBaseTech
If CalculateTechDiff < 0 Then
    CalculateTechDiff = 0
End If
End Function

Function ComputeSpeedStat(Tech As Single, Class As String, id As String) As Integer
On Error GoTo Empire_Error

Dim buildtech As Single
Dim speed As Integer

rsBuild.Seek "=", Class, id
If rsBuild.NoMatch Then
    ComputeSpeedStat = 0
    Exit Function
End If

buildtech = rsBuild.Fields("tech").Value

If rsBuild.Fields("base spd") > 0 Then
    If Class = "s" Then
        ComputeSpeedStat = Fix(ComputeShipSpeed(rsBuild.Fields("base spd"), rsBuild.Fields("tech"), _
            CInt(Tech)))
    Else
        ComputeSpeedStat = Fix(ComputeLandSpeed(rsBuild.Fields("base spd"), rsBuild.Fields("tech"), _
            CInt(Tech)))
    End If
Else
    If Class = "s" Then
        speed = rsBuild.Fields("stat2") 'see ComputeRangeStat for comment.
    End If
    
    If Tech <= (buildtech + 35) Then
        ComputeSpeedStat = speed
    Else
        ComputeSpeedStat = Int(speed * Log(Tech - buildtech) / Log(35))
    End If
End If
Exit Function
    
Empire_Error:
    EmpireError "ComputeSpeedStat", vbNullString, CStr(Tech) + ":" + Class + ":" + id
    
End Function

Public Function SectorBuildWork(iLCM As Integer, iHCM As Integer) As Integer
SectorBuildWork = iLCM + 2 * iHCM
End Function

'This routine builds the error message for all parsing errors
Public Sub EmpireError(Routine As String, var As String, Line As String)

Dim StrMsg As String
Dim strTemp As String
Dim filenumber As Integer
Dim FileName As String
Static HideMsg As Boolean

'open error log
VerifySubDirectory "Errors", True
FileName = App.path + "\Errors\" + "WinACE_error_log" + " " + CStr(Year(Now)) + "-" + Format$(Month(Now), "00") + "-" _
        + Format$(Day(Now), "00") + ".txt"
filenumber = FreeFile   ' Get unused file number.
Open FileName For Append As #filenumber

'drk 6/5/03 : added version to the log entry so that when people send me error logs I know what I'm dealing with
'090703 rjk: Added CStr to Application versions and removed semicolon
Print #filenumber, "Log Entry " + Format$(Now, "hh:mm:ss AMPM") + " " + Format$(Now, "dddd, mmm d yyyy") + " version: " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)

StrMsg = "WinACE encountered an error during processing." + vbNewLine
strTemp = "VB Error: "

If Err.Number > 0 Then
    strTemp = strTemp + CStr(Err.Number) + " "
End If

If Len(Err.Description) > 0 Then
    strTemp = strTemp + "- " + Err.Description + " "
End If

Print #filenumber, strTemp
StrMsg = StrMsg + strTemp

If Len(Routine) > 0 Then
    strTemp = "Routine: " + Routine
    
    If Len(var) > 0 Then
        strTemp = strTemp + "   Loc: " + var
    End If
    StrMsg = StrMsg + vbNewLine + strTemp
    Print #filenumber, strTemp
End If
    
If Len(Line) > 0 Then
    strTemp = "Input Line: " + Line
    StrMsg = StrMsg + vbNewLine + strTemp
    Print #filenumber, strTemp
End If

'112903 rjk: Updated to point to sourceforge for assistance.
StrMsg = StrMsg + vbNewLine + "For assistance - please submit a request to sourceforge.net/projects/winace with the file WinACE_error_log.txt" _
    + vbNewLine + vbNewLine + "Continue to show errors?"

Print #filenumber, vbNullString
Close #filenumber

If Not HideMsg Then
    If MsgBox(StrMsg, vbYesNo + vbCritical, "WinACE Error") = vbNo Then
        HideMsg = True
    End If
End If
End Sub

Function ReversePath(strPath) As String
Dim i As Integer
Dim ch As String

ReversePath = vbNullString
If Not (Len(strPath) > 0) Then
    Exit Function
End If

For i = Len(strPath) To 1 Step -1
    Select Case Mid$(strPath, i, 1)
        Case "y"
            ch = "n"
        Case "u"
            ch = "b"
        Case "j"
            ch = "g"
        Case "n"
            ch = "y"
        Case "b"
            ch = "u"
        Case "g"
            ch = "j"
        Case Else
            ch = vbNullString
    End Select
    ReversePath = ReversePath + ch
Next i

End Function

Public Function PlagueRisk(rs As Recordset) As Single
'From code
'    plg_num = ((vec[I_CIVIL] + vec[I_MILIT] + vec[I_UW]) / civvies) *
'        ((vec[I_IRON] + vec[I_OIL] + (vec[I_RAD] * 2)) / 10.0 +
'        np->nat_level[NAT_TLEV] + 100.0);
'    plg_denom = eff + mobil + 100 + np->nat_level[NAT_RLEV];
'    plg_chance = ((plg_num / plg_denom) - 1.0) * 0.01;

'the value civvies is maxpop based on the sector designation.
'122903 rjk: Moved chance calculation to the display time to allow display options.

Dim plg_num As Single
Dim plg_denom As Single
Dim civvies As Single

civvies = CSng(Maxpop)
rsSectorType.Seek "=", rs.Fields("des")
If Not rsSectorType.NoMatch Then
    civvies = CSng(rsSectorType.Fields("maxpop"))
End If

plg_num = (rs.Fields("civ") + rs.Fields("mil") + rs.Fields("uw")) / civvies

'drk@unxsoft.com 5/13/02.  I think this is correct now.  Someone please let me know if not.
plg_num = plg_num * ( _
                        ((rs.Fields("iron") + rs.Fields("oil") + (rs.Fields("rad") * 2)) / 10#) _
                        + rsNation.Fields("TechLevel") + 100 _
                    )
plg_denom = rs.Fields("eff") + rs.Fields("mob") + 100 + rsNation.Fields("Research")

'PlagueRisk = ((plg_num / plg_denom) - 1#)
PlagueRisk = (plg_num / plg_denom)
If PlagueRisk < 0 Then
    PlagueRisk = 0
End If
End Function

Public Function SectorString(sx As Integer, sy As Integer) As String
SectorString = CStr(sx) + "," + CStr(sy)
End Function

Public Sub ComputeUnitCountsForShips()
Dim hldIndex As String
Dim hldBldIndex As String
Dim iCapIndex As Integer
Dim bFound As Boolean

If rsShip.BOF And rsShip.EOF Then
    Exit Sub
End If
hldIndex = rsShip.Index
rsShip.Index = "PrimaryKey"
hldBldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"
rsShip.MoveFirst
Do
    rsShip.Edit
    rsShip.Fields("pln") = 0
    rsShip.Fields("he") = 0
    rsShip.Fields("xl") = 0
    rsShip.Fields("land") = 0
    rsShip.Update
    rsShip.MoveNext
Loop Until rsShip.EOF
If Not (rsLand.BOF And rsLand.EOF) Then
    rsLand.MoveFirst
    Do
        If rsLand.Fields("ship") <> -1 Then
            rsShip.Seek "=", rsLand.Fields("ship")
            If Not rsShip.NoMatch Then
                rsShip.Edit
                rsShip.Fields("land") = rsShip.Fields("land") + 1
                rsShip.Update
            End If
        End If
        rsLand.MoveNext
    Loop Until rsLand.EOF
End If
If Not (rsPlane.BOF And rsPlane.EOF) Then
    rsPlane.MoveFirst
    Do
        If rsPlane.Fields("ship") <> -1 Then
            rsShip.Seek "=", rsPlane.Fields("ship")
            If Not rsShip.NoMatch Then
                If Not (rsBuild.BOF And rsBuild.EOF) Then
                    rsBuild.Seek "=", "p", rsPlane.Fields("type")
                    If Not rsBuild.NoMatch Then
                        iCapIndex = 1
                        bFound = False
                        While rsBuild.Fields("cap" + CStr(iCapIndex)) <> "" And Not bFound
                            If rsBuild.Fields("cap" + CStr(iCapIndex)) = "x-light" Then
                                rsShip.Edit
                                rsShip.Fields("xl") = rsShip.Fields("xl") + 1
                                rsShip.Update
                                bFound = True
                            End If
                            iCapIndex = iCapIndex + 1
                        Wend
                        iCapIndex = 1
                        While rsBuild.Fields("cap" + CStr(iCapIndex)) <> "" And Not bFound
                            If rsBuild.Fields("cap" + CStr(iCapIndex)) = "helo" Then
                                rsShip.Edit
                                rsShip.Fields("he") = rsShip.Fields("he") + 1
                                rsShip.Update
                                bFound = True
                            End If
                            iCapIndex = iCapIndex + 1
                        Wend
                        iCapIndex = 1
                        While rsBuild.Fields("cap" + CStr(iCapIndex)) <> "" And Not bFound
                            If rsBuild.Fields("cap" + CStr(iCapIndex)) = "light" Then
                                rsShip.Edit
                                rsShip.Fields("pln") = rsShip.Fields("pln") + 1
                                rsShip.Update
                                bFound = True
                            End If
                            iCapIndex = iCapIndex + 1
                        Wend
                    End If
                End If
            End If
        End If
        rsPlane.MoveNext
    Loop Until rsPlane.EOF
End If
rsBuild.Index = hldBldIndex
rsShip.Index = hldIndex
End Sub


Public Sub ComputeUnitCountsForLandUnits()
Dim hldIndex As String
Dim hldBldIndex As String
Dim iCapIndex As Integer
Dim iCurIndex As Integer
Dim bFound As Boolean
Dim varBookmark As Variant

If rsLand.BOF And rsLand.EOF Then
    Exit Sub
End If

hldIndex = rsLand.Index
rsLand.Index = "PrimaryKey"
hldBldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

rsLand.MoveFirst
Do
    rsLand.Edit
    rsLand.Fields("xl") = 0
    rsLand.Fields("nland") = 0
    rsLand.Update
    rsLand.MoveNext
Loop Until rsLand.EOF
If Not (rsLand.BOF And rsLand.EOF) Then
    rsLand.MoveFirst
    Do
        If rsLand.Fields("land") <> -1 Then
            varBookmark = rsLand.Bookmark
            rsLand.Seek "=", rsLand.Fields("land")
            If Not rsLand.NoMatch Then
                rsLand.Edit
                rsLand.Fields("nland") = rsLand.Fields("nland") + 1
                rsLand.Update
            End If
            rsLand.Move 0, varBookmark
        End If
        rsLand.MoveNext
    Loop Until rsLand.EOF
End If
If Not (rsPlane.BOF And rsPlane.EOF) Then
    rsPlane.MoveFirst
    Do
        If rsPlane.Fields("land") <> -1 Then
            rsLand.Seek "=", rsPlane.Fields("land")
            If Not rsLand.NoMatch Then
                If Not (rsBuild.BOF And rsBuild.EOF) Then
                    rsBuild.Seek "=", "p", rsPlane.Fields("type")
                    If Not rsBuild.NoMatch Then
                        iCapIndex = 1
                        bFound = False
                        While rsBuild.Fields("cap" + CStr(iCapIndex)) <> "" And Not bFound
                            If rsBuild.Fields("cap" + CStr(iCapIndex)) = "x-light" Then
                                rsLand.Edit
                                rsLand.Fields("xl") = rsLand.Fields("xl") + 1
                                rsLand.Update
                                bFound = True
                            End If
                            iCapIndex = iCapIndex + 1
                        Wend
                    End If
                End If
            End If
        End If
        rsPlane.MoveNext
    Loop Until rsPlane.EOF
End If
rsBuild.Index = hldBldIndex
rsLand.Index = hldIndex
End Sub

