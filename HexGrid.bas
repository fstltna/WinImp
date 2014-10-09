Attribute VB_Name = "HexGrid"
Option Explicit

'Change Log:
'092403 rjk: Switched hard coding to use UnitCharacteristicCheck.
'100803 rjk: Change NUMBER_OF_BASIC_COLORS from 8 to 9 to add 'Not Yours' to NORMAL display
'            Added default for 'Not Yours' colors to be the same as "Owned"
'100903 rjk: Made ConditionMet Public for bulk move capability
'101403 rjk: Removed duplicate cases in ConditionMet with incorrect logic.
'101903 rjk: Added GetConditionalSectors() and made ConditionMet private again.
'102003 rjk: Added Delivery ColorScheme
'102303 rjk: Added CountMultipleSectors to count the number Multiple Sectors in the Sector field
'110103 rjk: Added FriendlyPlane GIF. Added food required to the graduated colorscheme
'            Fill in designation from rsEnemySector if bmap does not have a designation.
'110303 rjk: Added Neutral Plane and Land GIF.
'110403 rjk: Added a Null check to reading a designation from the rsEnemySect table in LoadScreen
'            Show both designations bmap and rsEnemySector if they are different.
'110503 rjk: Added food shortage to the graduated colorscheme
'111103 rjk: Added MOBILIZING and SITZKRIEG relations as Enemy for showing Units
'112003 rjk: Changed the check for valid enemy sector designation to be a length check instead of null check
'112203 rjk: Removed HG_BIG_PIC.  Added Distribution thresholds to F4 Distribution ColorScheme
'            Override not known or unknown on the bmap if enemyIntelligence has a designation.
'            Added displaying values for enemy sectors in graduated mode.
'            Added the ability to either show current or new designation and both in NORMAL/CONDITIONAL modes.
'112303 rjk: Changed the size and position of the units to allow the display of the designation.
'            Fixed the restriction code to only display three units.
'112403 rjk: Added the ability to show food shortages and food required (food excess)
'112503 rjk: Added the ability to select conditional display from rsEnemySector table.
'            Set default value when a field is missing or unknown for ConditionMet functions
'112903 rjk: Added ShareBmap checks to the determination used for displaying both bmap and enemySect designations
'112903 rjk: Removed ViewUnits and replaced with unitView to set Unit View Options. Also moved the pictures.
'120303 rjk: Moved the Unit Icon positions to make more visible
'120303 rjk: Added Food required and Food Excess for F4 Conditional
'121803 rjk: Added Plague Risk scaled by 10
'122903 rjk: Changed Plague Risk to show percentage chance
'150204 rjk: Added Background Image selection, Added Border Color selection
'180204 rjk: Changed bridges to be transparent as well.
'050304 rjk: Display ally/enemy bridges and bridgeheads from intelligence
'110304 rjk: Switched to IsNull check for EnemySector designation instead of Length Check.
'            Show "?" if the EnemySector designation is null
'140304 rjk: Fixed Invalid procedure when including enemy sectors on graduated display with an item
'            that enemy sector do not record.
'300304 rjk: Fixed to deal with strength fields being Null
'050404 rjk: Added excess civilians to the graduated and conditional color schemes
'160404 rjk: Fixed the displaying both designations when rsSector and rsEnemySector both present.
'            Fixed the displaying of commodities values in Distribution and Deliveries mode.
'030704 rjk: Fixed the deity mode to show countries in their colors in ColorScheme NORMAL.
'300704 rjk: Added Expired Units.
'080904 rjk: Fixed F4 Delivery ColorScheme to work with Star Wars radius mode.
'090904 rjk: Fixed the F4 Delivery ColorScheme to use the correct colors.
'011204 rjk: Generalize map printing function for use by PrintMap.
'140705 rjk: Fixed the limit check for MobilityCheck.
'160705 rjk: Added Reacting Planes.
'261105 rjk: Fixed the displaying enemy intelligence only sectors.
'171205 rjk: Fixed the displaying mines over enemy intelligence.
'181205 rjk: Fixed Reacting Planes to work with an empty missions table.
'290606 rjk: Added Our Nuke Units.
'090607 rjk: Added PD to show production.
'170508 rjk: Added FS (Food shortage) to the F4 Conditional.
'220508 rjk: Added Build Cost to F4 graduated and conditional.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

Public NbrHexWide As Integer
Public NbrHexTall As Integer
Private HexSideLength As Single
Private HexFontSize As Single
Public HSL866 As Single
Public HSL5 As Single
Public OriginX As Integer
Public OriginY As Integer
Public CurrentSectorX As Integer
Public CurrentSectorY As Integer
'Colors used in drawing the hex grid
Public Const NUMBER_OF_PLAYER_COLORS = 40
'100803 rjk: Change from 8 to 9 to add 'Not Yours'
'102003 rjk: from 9 to 11 for Delivery colors
Public Const NUMBER_OF_BASIC_COLORS = 11
Public PlayerColors(0 To NUMBER_OF_PLAYER_COLORS) As Long
Public PlayerText(0 To NUMBER_OF_PLAYER_COLORS) As Long
Public BasicColors(1 To NUMBER_OF_BASIC_COLORS) As Long
Public BasicText(1 To NUMBER_OF_BASIC_COLORS) As Long

Public Const vbMediumGreen = &HC000&
Public Const vbDarkGreen = &H8000&
Public Const vbDarkBlue = &H800000

'Color scheme used in drawing the hex grid
Public Const MAX_COLOR_CONDITIONS = 5
Public Const COLORSCHEME_NORMAL = 0
Public Const COLORSCHEME_CONDITIONAL = 1
Public Const COLORSCHEME_GRADUATED = 2
Public Const COLORSCHEME_DISTRIBUTION = 3
Public Const COLORSCHEME_DELIVERIES = 4 '102003 rjk: Added for Delivery ColorScheme
Public ColorScheme As Integer
Public GradColorItem As String
Public GradColorHigh As Long
Public GradColorLow As Long
Public GradColorDspVal As Boolean
Public CondColorParms() As Variant
Public CondColorParmNumber As Integer
Public strCommodity As String '102003 rjk: Added for Delivery ColorScheme
Public Enum enumDisplayDesignation '112203 rjk: Added for controlling designation display
    DD_CURRENT
    DD_NEW
    DD_BOTH
    DD_BOTH_CURRENT
    DD_BOTH_NEW
End Enum
Public dspDesignation As enumDisplayDesignation
Public bIncludeEnemySectorsF4Display As Boolean '112203 rjk: Added for controlling the including of enemysector data
'112503 rjk: Added for controlling the default for when a condition field does not exist in the EnemySect or Sectors table
Public bConditionalMissing As Boolean

'Hex Directions
Public Const HEX_TOPLEFT = 1
Public Const HEX_TOPRIGHT = 2
Public Const HEX_RIGHT = 3
Public Const HEX_BOTTOMRIGHT = 4
Public Const HEX_BOTTOMLEFT = 5
Public Const HEX_LEFT = 6

'Drawing Font
Const MAP_DRAWING_FONT = "Rockwell"
Const MAP_BACKUP_FONT = "Times New Roman"

'****************************************************************
'Windows API/Global Declarations for Draw Filled Objects
'****************************************************************
Type POINTAPI
    X As Long
    Y As Long
End Type

Declare Function Polygon Lib "GDI32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function SetPolyFillMode Lib "GDI32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long

'PolyFill Modes
Public Const ALTERNATE = 1
Public Const WINDING = 2

Dim Ply(10) As POINTAPI

Public Type tScreenInfo
    iUnitCounts(15) As Integer
    cDes As String * 1
    cSecondDes As String * 1
    iValue As Integer
    bDisplayColor As Byte
    bOwner As Boolean
    bNotYours As Boolean
End Type
Dim ScreenInfo() As tScreenInfo

Dim FullBmap() As String * 1
Dim TextSize(32 To 126, 1 To 2) As Single
Public bMapBitMapLoaded As Boolean

Public Sub DrawGrid(p As Object)
On Error GoTo Empire_Error
Dim OrgX As Integer
Dim OrgY As Integer
Dim X As Integer
Dim Y As Integer

p.DrawWidth = 2   ' Set DrawWidth.
'p.ForeColor = QBColor(0)
p.ForeColor = frmOptions.lBorderColor
p.FillStyle = 0 ' Sold Fill

'set PolyFillMode to fill whole shape
SetPolyFillMode p.hDC, WINDING

If HexSideLength = 0 Then
    Exit Sub
End If

NbrHexWide = Int(p.ScaleWidth / (HSL866 * 2))
NbrHexTall = Int(p.ScaleHeight / (HSL5 * 3))

OrgX = OriginX
OrgY = OriginY

If Not bMapBitMapLoaded Then
    Set frmDrawMap.picMap.Picture = Nothing
    If frmOptions.strImageFileName <> "" Then
        Set frmDrawMap.picMap.Picture = LoadPicture(frmOptions.strImageFileName)
    End If
    bMapBitMapLoaded = True
End If

LoadScreen

For Y = 1 To NbrHexTall
    For X = 1 To NbrHexWide
        Dim PosX As Single
        Dim PosY As Single
        
        'Get Center position and draw hex
        Coord OrgX, OrgY, PosX, PosY
        DrawSector p, PosX, PosY, OrgX, OrgY
        OrgX = OrgX + 2
        If OrgX >= (MapSizeX / 2) Then
            OrgX = OrgX - MapSizeX
        End If
                
    Next X
    OrgY = OrgY + 1
    If OrgY > MapSizeY / 2 Then
        OrgY = OrgY - MapSizeY
    End If
    OrgX = OriginX + Abs(OrgY Mod 2)
Next Y
Exit Sub
Empire_Error:
If Err.Number = 9 Then
    'Probably have version information off
    Dim nReply As Integer
    nReply = MsgBox("Error in Map routines, World size may have changes since last time you ran version." + _
        vbNewLine + "Press OK to update version informantion", vbExclamation + vbOKCancel)
    If nReply = vbOK Then
        frmEmpCmd.SubmitEmpireCommand "version", True
    End If
End If
EmpireError "DrawGrid", vbNullString, vbNullString

End Sub

Public Sub Coord(sectx As Integer, secty As Integer, ByRef PosX As Single, ByRef PosY As Single)

Dim offsetX As Integer
Dim offsetY As Integer

offsetX = sectx - OriginX
offsetY = secty - OriginY
If offsetX < 0 Then
    offsetX = offsetX + MapSizeX
End If
If offsetY < 0 Then
    offsetY = offsetY + MapSizeY
End If

PosX = (1 + offsetX) * HSL866
PosY = (1 + offsetY * 1.5) * HexSideLength
End Sub


'Private Sub DrawHexSide(f As Object, CenterX As Single, CenterY As Single, Side As Integer)    8/2003 efj  removed
'
'
'Select Case Side
'    Case HEX_TOPLEFT:
'        f.Line (CenterX, CenterY - HexSideLength)-(CenterX - HSL866, CenterY - HSL5)
'    Case HEX_TOPRIGHT:
'        f.Line (CenterX + HSL866, CenterY - HSL5)-(CenterX, CenterY - HexSideLength)
'    Case HEX_RIGHT:
'        f.Line (CenterX + HSL866, CenterY + HSL5)-(CenterX + HSL866, CenterY - HSL5)
'    Case HEX_BOTTOMRIGHT:
'        f.Line (CenterX, CenterY + HexSideLength)-(CenterX + HSL866, CenterY + HSL5)
'    Case HEX_BOTTOMLEFT:
'        f.Line (CenterX - HSL866, CenterY + HSL5)-(CenterX, CenterY + HexSideLength)
'    Case HEX_LEFT:
'        f.Line (CenterX - HSL866, CenterY - HSL5)-(CenterX - HSL866, CenterY + HSL5)
'End Select
'
'End Sub

Public Sub DrawHex(f As Object, CenterX As Single, CenterY As Single)

'HEX_TOPLEFT:
'f.Line (CenterX, CenterY - HexSideLength)-(CenterX - HSL866, CenterY - HSL5)
'HEX_TOPRIGHT:
'f.Line (CenterX + HSL866, CenterY - HSL5)-(CenterX, CenterY - HexSideLength)
'HEX_RIGHT:
'f.Line (CenterX + HSL866, CenterY + HSL5)-(CenterX + HSL866, CenterY - HSL5)
'HEX_BOTTOMRIGHT:
'f.Line (CenterX, CenterY + HexSideLength)-(CenterX + HSL866, CenterY + HSL5)
'HEX_BOTTOMLEFT:
'f.Line (CenterX - HSL866, CenterY + HSL5)-(CenterX, CenterY + HexSideLength)
'HEX_LEFT:
'f.Line (CenterX - HSL866, CenterY - HSL5)-(CenterX - HSL866, CenterY + HSL5)

'HEX_TOPLEFT:
 f.Line (CenterX, CenterY - HexSideLength)-(CenterX - HSL866, CenterY - HSL5)
'HEX_LEFT:
 f.Line -(CenterX - HSL866, CenterY + HSL5)
'HEX_BOTTOMLEFT:
 f.Line -(CenterX, CenterY + HexSideLength)
'HEX_BOTTOMRIGHT:
 f.Line -(CenterX + HSL866, CenterY + HSL5)
'HEX_RIGHT:
 f.Line -(CenterX + HSL866, CenterY - HSL5)
'HEX_TOPRIGHT:
 f.Line -(CenterX, CenterY - HexSideLength)

End Sub
Public Sub DrawHexCenterLine(f As Object, CenterX As Single, CenterY As Single, lenline As Single, Side As Integer)

Select Case Side
    Case HEX_TOPLEFT:
        f.Line (CenterX, CenterY)-(CenterX - 0.5 * lenline, CenterY - 0.8666 * lenline)
    Case HEX_TOPRIGHT:
        f.Line (CenterX, CenterY)-(CenterX + 0.5 * lenline, CenterY - 0.8666 * lenline)
    Case HEX_RIGHT:
        f.Line (CenterX, CenterY)-(CenterX + lenline, CenterY)
    Case HEX_BOTTOMRIGHT:
        f.Line (CenterX, CenterY)-(CenterX + 0.5 * lenline, CenterY + 0.8666 * lenline)
    Case HEX_BOTTOMLEFT:
        f.Line (CenterX, CenterY)-(CenterX - 0.5 * lenline, CenterY + 0.8666 * lenline)
    Case HEX_LEFT:
        f.Line (CenterX, CenterY)-(CenterX - lenline, CenterY)
End Select

End Sub

Sub DrawSolidHex(p As Object, CenterX As Single, CenterY As Single, FColor As Long)
Dim nPoints As Long

p.FillColor = FColor ' Fill Color
'build coordinates for shape
Ply(HEX_TOPLEFT).X = CenterX
Ply(HEX_TOPLEFT).Y = CenterY - HexSideLength

Ply(HEX_TOPRIGHT).X = CenterX + HSL866
Ply(HEX_TOPRIGHT).Y = CenterY - HSL5

Ply(HEX_RIGHT).X = CenterX + HSL866
Ply(HEX_RIGHT).Y = CenterY + HSL5

Ply(HEX_BOTTOMRIGHT).X = CenterX
Ply(HEX_BOTTOMRIGHT).Y = CenterY + HexSideLength

Ply(HEX_BOTTOMLEFT).X = CenterX - HSL866
Ply(HEX_BOTTOMLEFT).Y = CenterY + HSL5

Ply(HEX_LEFT).X = CenterX - HSL866
Ply(HEX_LEFT).Y = CenterY - HSL5
nPoints = 6

'draw Hex
Polygon p.hDC, Ply(1), nPoints
End Sub

Public Sub SetLetterSize(ByRef p As Object)
'Set Font Size
SetDrawingFont p

p.FontSize = 48   ' Set font size.
While p.TextHeight("l") > (HexSideLength * 1.25) And p.FontSize > 3
    p.FontSize = p.FontSize - 1
Wend
End Sub

'112903 rjk: Removed ViewUnits and replaced with unitView.VisibleUnit
Public Sub DrawSector(ByRef p As Object, CenterX As Single, CenterY As Single, sx As Integer, sy As Integer)
Dim ch As String
Dim ch2 As String
Dim n As Integer
Dim Percent As Single
Dim oldcolor
Dim oldfontsize
Dim ix As Integer
Dim iy As Integer

oldcolor = p.ForeColor
oldfontsize = p.Font.Size
ix = SIx(sx)
iy = SIy(sy)

'draw sectors then
If ScreenInfo(ix, iy).cDes = vbNullChar Then
    'uncharted territory
    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(4), BasicText(4))
    'dummy p.Print is necessary to get screen to update when
    'only one sector is updated.
    ch = " "
    p.Print ch
Else
    ch = ScreenInfo(ix, iy).cDes
    
    If ch = "." Then 'sea are drawn in blue
        If frmDrawMap.picMap.Picture <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
            p.FillStyle = vbFSTransparent
        End If
        DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2))
        p.FillStyle = vbFSSolid
        p.ForeColor = CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2))
        If ScreenInfo(ix, iy).cSecondDes <> vbNullChar Then
            ch = ScreenInfo(ix, iy).cSecondDes '100404 rjk: NOT WORKING
        Else
            ch = " "
        End If
        'if graduated on fertility or oil, and this is a sea hex,
        'we need to use the values from rsSea for display in the hex
        If ColorScheme = COLORSCHEME_GRADUATED And ScreenInfo(ix, iy).iValue >= 0 Then
            If ch = " " Then
                ch = CStr(ScreenInfo(ix, iy).iValue)
            Else
                ch2 = CStr(ScreenInfo(ix, iy).iValue)
                p.Font.Size = GetHexFontSize * 0.7
                p.CurrentX = CenterX - p.TextWidth(ch2) / 2
                p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
                p.Print ch2
            End If
        End If
        If frmUnitView.unitView.VisibleUnits(ScreenInfo(ix, iy).iUnitCounts) Or _
           ch2 <> "" Then
            p.Font.Size = GetHexFontSize * 0.7
            p.CurrentX = CenterX - p.TextWidth(ch) / 2
            p.CurrentY = CenterY
        Else
            p.CurrentX = CenterX - p.TextWidth(ch) / 2
            p.CurrentY = CenterY - p.TextHeight(ch) / 2
        End If
        p.Print ch

    ElseIf ch = "X" Then 'draw minefields in reverse colors
        DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2))
        p.ForeColor = CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2))
        p.CurrentX = CenterX - TextSize(Asc(ScreenInfo(ix, iy).cDes), 1)
        p.CurrentY = CenterY - TextSize(Asc(ScreenInfo(ix, iy).cDes), 2)
        p.Print ch
    Else
        If ScreenInfo(ix, iy).bOwner Then
            'owned sector
            Select Case ColorScheme
            Case COLORSCHEME_NORMAL
                If ch = "=" Then    'bridges over water are drawn in blue
                    If frmDrawMap.picMap.Picture <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
                        p.FillStyle = vbFSTransparent
                    End If
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2))
                    p.FillStyle = vbFSSolid
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2))
                ElseIf ScreenInfo(ix, iy).bNotYours Then
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(9), BasicText(9))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(9), BasicColors(9))
                Else
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(1), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(1), BasicColors(1))
                End If
                DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
        
            Case COLORSCHEME_CONDITIONAL
                'test if condition is true or false
                If ScreenInfo(ix, iy).iValue > 0 Then
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(5), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(5), BasicColors(1))
                Else
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(6), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(6), BasicColors(1))
                End If
                DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
            
            Case COLORSCHEME_GRADUATED
                If GradColorHigh = GradColorLow Then
                    Percent = 1
                Else
                    Percent = (ScreenInfo(ix, iy).iValue - GradColorLow) / (GradColorHigh - GradColorLow)
                End If
                DrawSolidHex p, CenterX, CenterY, _
                    CheckSelectedHex(sx, sy, BlendedColor(BasicColors(8), BasicColors(7), Percent), BasicText(1))
                If Percent > 0.5 Then
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(7), BasicColors(1))
                Else
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(8), BasicColors(1))
                End If
                If GradColorDspVal Then
                    ch2 = CStr(ScreenInfo(ix, iy).iValue)
                    p.Font.Size = GetHexFontSize * 0.7
                    p.CurrentX = CenterX - p.TextWidth(ch2) / 2
                    p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
                    p.Print ch2
                    p.CurrentX = CenterX - p.TextWidth(ch) / 2
                    p.CurrentY = CenterY
                    p.Print ch
                Else
                    DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
                End If
            
            Case COLORSCHEME_DISTRIBUTION
                '112203 rjk: Added displaying the threshold
                n = ScreenInfo(ix, iy).bDisplayColor
                n = Abs(n) Mod (NUMBER_OF_PLAYER_COLORS + 1)
                If ch = "=" Then
                    If frmDrawMap.picMap.Picture <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
                        p.FillStyle = vbFSTransparent
                    End If
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), PlayerColors(n))
                    p.FillStyle = vbFSSolid
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), BasicColors(2))
                Else
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerText(n), PlayerColors(n))
                End If
                If ScreenInfo(ix, iy).iValue = 0 Then  'no threshold
                    DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
                Else
                    ch2 = CStr(ScreenInfo(ix, iy).iValue)
                    p.Font.Size = GetHexFontSize * 0.7
                    p.CurrentX = CenterX - p.TextWidth(ch2) / 2
                    p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
                    p.Print ch2
                    p.CurrentX = CenterX - p.TextWidth(ch) / 2
                    p.CurrentY = CenterY
                    p.Print ch
                End If
            '102003 rjk: Added Delivery ColorScheme
            Case COLORSCHEME_DELIVERIES
                If ScreenInfo(ix, iy).iValue <> 0 Then
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(10), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(10), BasicColors(1))
                    Select Case ScreenInfo(ix, iy).iValue Mod 8
                    Case 1
                        p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - HSL866 + p.DrawWidth, CenterY + HSL5)
                    Case 2
                        p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    Case 3
                        p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    Case 4
                        p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + HSL866 - p.DrawWidth, CenterY + HSL5)
                    Case 5
                        p.Line (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    Case 6
                        p.Line (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    Case 7
                    End Select
                Else
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(11), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(11), BasicColors(1))
                End If
                ch2 = CStr(ScreenInfo(ix, iy).iValue \ 8)
'                If ch2 <> 0 Then
                If ScreenInfo(ix, iy).iValue <> 0 Then
                    p.Font.Size = GetHexFontSize * 0.7
                    p.CurrentX = CenterX - p.TextWidth(ch2) / 2
                    p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
                    p.Print ch2
                    p.CurrentX = CenterX - p.TextWidth(ch) / 2
                    p.CurrentY = CenterY
                    p.Print ch
                Else
                    DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
                End If
            End Select
        ElseIf ScreenInfo(ix, iy).bDisplayColor = 0 Then
            'bmap sector
            If ch = "=" Then
                If frmDrawMap.picMap.Picture <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
                    p.FillStyle = vbFSTransparent
                End If
                DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2))
                p.FillStyle = vbFSSolid
                p.ForeColor = CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2))
            ElseIf frmOptions.bStarWars And ch = "\" Then
                DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, RGB(255, 128, 64), BasicText(3))
                p.ForeColor = CheckSelectedHex(sx, sy, BasicText(3), RGB(255, 128, 64))
            Else
                DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(3), BasicText(3))
                p.ForeColor = CheckSelectedHex(sx, sy, BasicText(3), BasicColors(3))
            End If
            DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
                
        Else 'sector with enemy info
            n = ScreenInfo(ix, iy).bDisplayColor
            While n > NUMBER_OF_PLAYER_COLORS
                n = n - NUMBER_OF_PLAYER_COLORS
            Wend
            '112203 rjk: Added displaying values for enemy sectors in graduated mode
            If ColorScheme = COLORSCHEME_GRADUATED And bIncludeEnemySectorsF4Display And _
               GradColorDspVal And ScreenInfo(ix, iy).iValue <> -2 Then
                If GradColorHigh = GradColorLow Then
                    Percent = 1
                Else
                    Percent = (ScreenInfo(ix, iy).iValue - GradColorLow) / (GradColorHigh - GradColorLow)
                End If
                If ch = "=" Then
                    If frmDrawMap.picMap.Picture <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
                        p.FillStyle = vbFSTransparent
                    End If
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), PlayerColors(n))
                    p.FillStyle = vbFSSolid
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), BasicColors(2))
                Else
                    DrawSolidHex p, CenterX, CenterY, _
                        CheckSelectedHex(sx, sy, BlendedColor(BasicColors(8), BasicColors(7), Percent), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - HSL866 + p.DrawWidth, CenterY + HSL5)
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + HSL866 - p.DrawWidth, CenterY + HSL5)
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                End If
                If Percent > 0.5 Then
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(7), BasicColors(1))
                Else
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(8), BasicColors(1))
                End If
                If ScreenInfo(ix, iy).iValue = -1 Then
                    ch2 = "???"
                Else
                    ch2 = CStr(ScreenInfo(ix, iy).iValue)
                End If
                p.Font.Size = GetHexFontSize * 0.7
                p.CurrentX = CenterX - p.TextWidth(ch2) / 2
                p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
                p.Print ch2
                p.CurrentX = CenterX - p.TextWidth(ch) / 2
                p.CurrentY = CenterY
                p.Print ch
            '112503 rjk: Added the ability to select conditional display from rsEnemySector table.
            ElseIf ColorScheme = COLORSCHEME_CONDITIONAL And bIncludeEnemySectorsF4Display Then
                'test if condition is true or false
                If ScreenInfo(ix, iy).iValue > 0 Then
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(5), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - HSL866 + p.DrawWidth, CenterY + HSL5)
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + HSL866 - p.DrawWidth, CenterY + HSL5)
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(5), BasicColors(1))
                Else
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(6), BasicText(1))
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - HSL866 + p.DrawWidth, CenterY + HSL5)
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2))
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth)-(CenterX + HSL866 - p.DrawWidth, CenterY + HSL5)
                    p.Line (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    p.Line (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth)-(CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2))
                    p.ForeColor = CheckSelectedHex(sx, sy, BasicText(6), BasicColors(1))
                End If
                DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
            Else
                If ch = "=" Then
                    If frmDrawMap.picMap.Picture <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
                        p.FillStyle = vbFSTransparent
                    End If
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), PlayerColors(n))
                    p.FillStyle = vbFSSolid
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), BasicColors(2))
                Else
                    DrawSolidHex p, CenterX, CenterY, CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
                    p.ForeColor = CheckSelectedHex(sx, sy, PlayerText(n), PlayerColors(n))
                End If
                '110403 rjk: Show both designations bmap and rsEnemySector if they are different.
                DrawDesignation ScreenInfo, p, CenterX, CenterY, ix, iy, ch
            End If
        End If
    End If
End If
               
p.ForeColor = oldcolor
p.Font.Size = oldfontsize
DrawUnits p, CenterX, CenterY, sx, sy, HexSideLength
End Sub

Private Sub DrawDesignation(ScreenInfo() As tScreenInfo, p As Object, CenterX As Single, CenterY As Single, ix As Integer, iy As Integer, ch As String)
Dim ch2 As String

If frmUnitView.unitView.VisibleUnits(ScreenInfo(ix, iy).iUnitCounts) Then
    p.Font.Size = GetHexFontSize * 0.7
    p.CurrentX = CenterX - p.TextWidth(ch) / 2
    p.CurrentY = CenterY
Else
    If ScreenInfo(ix, iy).cSecondDes <> vbNullChar Then
        ch2 = ScreenInfo(ix, iy).cSecondDes
        p.Font.Size = GetHexFontSize * 0.7
        p.CurrentX = CenterX - p.TextWidth(ch2) / 2
        p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
        p.Print ch2
        p.CurrentX = CenterX - p.TextWidth(ch) / 2
        p.CurrentY = CenterY
    Else
        p.CurrentX = CenterX - TextSize(Asc(ScreenInfo(ix, iy).cDes), 1)
        p.CurrentY = CenterY - TextSize(Asc(ScreenInfo(ix, iy).cDes), 2)
    End If
End If
p.Print ch
End Sub

Private Function CheckSelectedHex(sx As Integer, sy As Integer, NormalColor As Long, SelectedColor As Long) As Long
If sx = CurrentSectorX And sy = CurrentSectorY Then
    CheckSelectedHex = SelectedColor
Else
    CheckSelectedHex = NormalColor
End If
End Function

'112303 rjk: Changed the size and position of the units to allow the display of the designation.
'            Fixed the restriction code to only display three units.
Public Sub DrawUnits(p As Object, CenterX As Single, CenterY As Single, PosX As Integer, PosY As Integer, lenSide As Single)
Dim n As Integer
Dim pic As Picture
Dim ix As Integer
Dim iy As Integer

ix = SIx(PosX)
iy = SIy(PosY)

For n = 1 To 3
    Set pic = frmUnitView.unitView.UnitPicture(ScreenInfo(ix, iy).iUnitCounts, (n = 1))
    If pic Is Nothing Then
        Exit Sub
    Else
        Select Case n '120303 rjk: Moved the positions to make more visible
        Case 1
            p.PaintPicture pic, CenterX - (0.8 * lenSide), CenterY - (0.5 * lenSide), 0.8 * lenSide, (pic.Height / pic.Width) * 0.8 * lenSide
        Case 2
            p.PaintPicture pic, CenterX, CenterY - (0.5 * lenSide), 0.8 * lenSide, (pic.Height / pic.Width) * 0.8 * lenSide
        Case 3
            p.PaintPicture pic, CenterX - (0.4 * lenSide), CenterY - (0.8 * lenSide), 0.8 * lenSide, (pic.Height / pic.Width) * 0.8 * lenSide
        End Select
    End If
Next n
End Sub

Public Sub SetDefaultPlayerColors()
PlayerColors(0) = vbWhite
PlayerColors(1) = 8421631
PlayerColors(2) = 65535
PlayerColors(3) = 33023
PlayerColors(4) = 16512
PlayerColors(5) = 32768
PlayerColors(6) = 16776960
PlayerColors(7) = 16744576
PlayerColors(8) = 16744703
PlayerColors(9) = 16711935
PlayerColors(10) = 16711808
PlayerColors(11) = 128
PlayerColors(12) = 32896
PlayerColors(13) = 8421440
PlayerColors(14) = 12615680
PlayerColors(15) = 12615935
PlayerColors(16) = 8388736
PlayerColors(17) = 8388672
PlayerColors(18) = 8454143
PlayerColors(19) = 11395647
PlayerColors(20) = 4830697

PlayerText(0) = vbBlack
PlayerText(1) = vbBlack
PlayerText(2) = vbBlack
PlayerText(3) = vbBlack
PlayerText(4) = vbWhite
PlayerText(5) = vbWhite
PlayerText(6) = vbBlack
PlayerText(7) = vbYellow
PlayerText(8) = vbBlack
PlayerText(9) = vbBlack
PlayerText(10) = vbWhite
PlayerText(11) = vbWhite
PlayerText(12) = vbWhite
PlayerText(13) = vbYellow
PlayerText(14) = vbYellow
PlayerText(15) = vbBlack
PlayerText(16) = vbWhite
PlayerText(17) = vbWhite
PlayerText(18) = vbBlack
PlayerText(19) = vbBlack
PlayerText(20) = vbBlack

Dim n As Integer
For n = 21 To NUMBER_OF_PLAYER_COLORS
    PlayerColors(n) = PlayerColors(n - 20)
    PlayerText(n) = PlayerText(n - 20)
Next n

End Sub

Public Sub SetDefaultBasicColors()
BasicColors(1) = RGB(0, 255, 135)
BasicColors(2) = RGB(0, 150, 255)
BasicColors(3) = RGB(188, 188, 188)
BasicColors(4) = vbBlack
BasicColors(9) = BasicColors(1) '100803 rjk: default 'Not Your' colors to normal color

BasicText(1) = vbBlack
BasicText(2) = vbWhite
BasicText(3) = vbBlack
BasicText(4) = vbWhite
BasicText(9) = BasicText(1) '100803 rjk: default 'Not Your' colors to normal color
End Sub

Public Sub SetDefaultComparisionColors()
BasicColors(5) = vbRed
BasicColors(6) = RGB(0, 255, 135)

BasicText(5) = vbYellow
BasicText(6) = vbBlack
End Sub

Public Sub SetDefaultGradColors()
BasicColors(7) = vbRed
BasicColors(8) = vbYellow

BasicText(7) = vbBlack
BasicText(8) = vbBlack
End Sub

Public Sub SetDefaultDeliveryColors()
BasicColors(10) = vbRed
BasicColors(11) = RGB(0, 255, 135)

BasicText(10) = vbYellow
BasicText(11) = vbBlack
End Sub

Public Function ColorComponent(Color As Long, component As Integer) As Long
Select Case component
    Case 1
        ColorComponent = Color And &HFF
    Case 2
        ColorComponent = ((Color And &HFF00) / &H100) And &HFF
    Case 3
        ColorComponent = (Color And &HFF0000) / &H10000
End Select
End Function

Public Function BlendedColor(Color1 As Long, Color2 As Long, Percent As Single) As Long

If Percent >= 1 Then
    BlendedColor = Color2
    Exit Function
End If

If Percent <= 0 Then
    BlendedColor = Color1
    Exit Function
End If

Dim r1 As Long
Dim r2 As Long
Dim g1 As Long
Dim g2 As Long
Dim b1 As Long
Dim b2 As Long
Dim r3 As Long
Dim g3 As Long
Dim b3 As Long

r2 = ColorComponent(Color2, 1)
g2 = ColorComponent(Color2, 2)
b2 = ColorComponent(Color2, 3)

r1 = ColorComponent(Color1, 1)
g1 = ColorComponent(Color1, 2)
b1 = ColorComponent(Color1, 3)

r3 = r1 + Percent * (r2 - r1)
g3 = g1 + Percent * (g2 - g1)
b3 = b1 + Percent * (b2 - b1)
BlendedColor = RGB(r3, g3, b3)
End Function

Private Function ConditionMet() As Boolean
Dim Cond As Boolean
Dim Done As Boolean
Dim n As Integer
Dim v As productionDataType

Cond = False
Done = False
n = 1
While n <= CondColorParmNumber And Not Done
    '120303 rjk: Added Food required and Food Excess for F4 Conditional
    Select Case CondColorParms(n, 1)
    Case "FR"
        Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(FoodRequired(rsSectors)), CLng(CondColorParms(n, 3)))
    Case "FS"
        If FoodRequired(rsSectors) - CLng(rsSectors.Fields("food").Value) > 0 Then
            Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), FoodRequired(rsSectors) - CLng(rsSectors.Fields("food").Value), CLng(CondColorParms(n, 3)))
        Else
            Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), 0, CLng(CondColorParms(n, 3)))
        End If
    Case "FE"
        Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(rsSectors.Fields("food").Value) - FoodRequired(rsSectors), CLng(CondColorParms(n, 3)))
    Case "EC" '050404 rjk: Added for Excess Civilians
        v = Production(rsSectors)
        If (v.prodValidFlag) Then
            Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(v.ExcessCiv), CLng(CondColorParms(n, 3)))
        Else
            Cond = bConditionalMissing
        End If
    Case "BC"
        v = Production(rsSectors)
        If (v.prodValidFlag) Then
            Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(v.BuildAvailCost), CLng(CondColorParms(n, 3)))
        Else
            Cond = bConditionalMissing
        End If
    Case "RP" '110605 rjk: Added RP - Reacting Planes
        Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(UpdateScreenInfoWithReactingPlanes(rsSectors.Fields("x").Value, rsSectors.Fields("y"))), CLng(CondColorParms(n, 3)))
    Case "PD"
        v = Production(rsSectors)
        If (v.prodValidFlag) Then
            Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(v.ProdAmount), CLng(CondColorParms(n, 3)))
        Else
            Cond = bConditionalMissing
        End If
    Case Else
        '112503 rjk: Check and make sure the strength fields not null
        If Not IsNull(rsSectors.Fields(CondColorParms(n, 1)).Value) Then
            If rsSectors.Fields(CondColorParms(n, 1)).Type = dbInteger Or _
                rsSectors.Fields(CondColorParms(n, 1)).Type = dbLong Then
                ' test numeric values
                Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CLng(rsSectors.Fields(CondColorParms(n, 1)).Value), CLng(CondColorParms(n, 3)))
            Else  ' test string values
                Select Case CondColorParms(n, 2)
                Case "="
                    Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value = CStr(CondColorParms(n, 3)))
                Case "#"
                    Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value <> CStr(CondColorParms(n, 3)))
                'Case "#" 101403 rjk: Removed duplicate case, also had wrong logic.
                '    Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value = CStr(CondColorParms(n, 3)))
                Case "<"
                    Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value < CStr(CondColorParms(n, 3)))
                Case ">"
                    Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value > CStr(CondColorParms(n, 3)))
                End Select
            End If
        Else
            Cond = bConditionalMissing
        End If
    End Select
    'we may not have to process any farther
    'if it is an "and" and the first value is false, or if it is an "or" and the first value is true
    If (CondColorParms(n, 4) = "&" And Not Cond) Or (CondColorParms(n, 4) = "|" And Cond) Then
        Done = True
    End If
    n = n + 1
Wend

ConditionMet = Cond
End Function

Private Function NumberConditionEvaluation(Operator As String, Value As Long, Condition As Long) As Boolean
Select Case Operator
Case "="
    NumberConditionEvaluation = (Value = Condition)
Case "#"
    NumberConditionEvaluation = (Value <> Condition)
Case "<"
    NumberConditionEvaluation = (Value < Condition)
Case ">"
    NumberConditionEvaluation = (Value > Condition)
End Select
End Function

'112503 rjk: Added for checking the conditional logic on EnemySector table.
Private Function ConditionMetEnemySector() As Boolean
Dim Cond As Boolean
Dim Done As Boolean
Dim n As Integer

Cond = False
Done = False
n = 1
While n <= CondColorParmNumber And Not Done
    Select Case CondColorParms(n, 1)
    Case "eff", "road", "rail", "defense", "civ", "mil", "shell", "gun", "pet", "food", "iron", "bar"
        ' test numeric values
        If rsEnemySect.Fields(CondColorParms(n, 1)).Value <> -1 Then
            Select Case CondColorParms(n, 2)
                Case "="
                    Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value = CLng(CondColorParms(n, 3)))
                Case "#"
                    Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value <> CLng(CondColorParms(n, 3)))
                Case "<"
                    Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value < CLng(CondColorParms(n, 3)))
                Case ">"
                    Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value > CLng(CondColorParms(n, 3)))
            End Select
        Else
            Cond = bConditionalMissing
        End If
    Case "country"
        If rsEnemySect.Fields(CondColorParms(n, 1)).Value <> -1 Then
            Select Case CondColorParms(n, 2)
                Case "="
                    Cond = (rsEnemySect.Fields("owner").Value = CLng(CondColorParms(n, 3)))
                Case "#"
                    Cond = (rsEnemySect.Fields("owner").Value <> CLng(CondColorParms(n, 3)))
                Case "<"
                    Cond = (rsEnemySect.Fields("owner").Value < CLng(CondColorParms(n, 3)))
                Case ">"
                    Cond = (rsEnemySect.Fields("owner").Value > CLng(CondColorParms(n, 3)))
            End Select
        Else
            Cond = bConditionalMissing
        End If
    Case Else
        Cond = bConditionalMissing
    End Select
    'we may not have to process any farther
    'if it is an "and" and the first value is false, or if it is an "or" and the first value is true
    If (CondColorParms(n, 4) = "&" And Not Cond) Or (CondColorParms(n, 4) = "|" And Cond) Then
        Done = True
    End If
    n = n + 1
Wend

ConditionMetEnemySector = Cond
End Function

Public Sub SetDrawingFont(ByRef obj As Object)

On Error Resume Next

'set to the backup first in case we have an error
'when we try to use the primary
obj.FontName = MAP_BACKUP_FONT

obj.FontName = MAP_DRAWING_FONT
obj.FontBold = True

End Sub

Public Sub CenterMap(p As Object, CenterX As Integer, CenterY As Integer)

NbrHexWide = Int(p.ScaleWidth / (HexSideLength * 0.866 * 2))
NbrHexTall = Int(p.ScaleHeight / (HexSideLength * 1.5))

OriginX = CenterX - NbrHexWide
OriginY = CenterY - NbrHexTall / 2

'check boundries
If OriginX > MapSizeX / 2 Then
    OriginX = OriginX = OriginX - MapSizeX
End If
    
If OriginX < MapSizeX / (-2) Then
    OriginX = OriginX + MapSizeX
End If

If OriginY > MapSizeY / 2 Then
    OriginY = OriginY - MapSizeY
End If

If OriginY < MapSizeY / (-2) Then
    OriginY = OriginY + MapSizeY
End If
    
'make sure its divisible by two
OriginX = Int(OriginX \ 2) * 2
OriginY = Int(OriginY \ 2) * 2
    
End Sub

Public Sub SetHexSideLength(p As Object, hlen As Single)
HexSideLength = hlen
HSL866 = HexSideLength * 0.866
HSL5 = HexSideLength * 0.5
SetLetterSize p
HexFontSize = p.FontSize

'load textsize array
Dim n As Integer
For n = 32 To 126
    TextSize(n, 1) = p.TextWidth(Chr(n)) / 2
    TextSize(n, 2) = p.TextHeight(Chr(n)) / 2
Next n

End Sub

Public Function GetHexSideLength() As Single
GetHexSideLength = HexSideLength
End Function
Public Function GetHexFontSize() As Single
GetHexFontSize = HexFontSize
End Function

Private Sub LoadScreen()
Dim X As Integer
Dim Y As Integer
Dim x1 As Integer
Dim y1 As Integer
Dim n As Integer
Dim topx As Integer
Dim topy As Integer
Static FullBmapLoaded As Boolean
Dim strKey As String
Dim DistColor As Integer
Dim DistPoint() As String
Dim DistCount() As Integer
Dim v As productionDataType

topx = OriginX + NbrHexWide * 2 + 2
topy = OriginY + NbrHexTall
ReDim ScreenInfo(OriginX To topx, OriginY To topy)

'112803 rjk: Moved Determination Unit code to EmpUnitView class
GetUnitCount ScreenInfo, topx, topy

'if supressbmap, the load static array and reload from that to speed
'up screen redraws, since we only have to load the bmap from the data
'base once.
'load static array if not already loaded.
If (Not FullBmapLoaded) Or (Not frmOptions.bSuppressBmapRefresh) Then
    ReDim FullBmap(-1 * MapSizeX / 2 To MapSizeX / 2, -1 * MapSizeY / 2 To MapSizeY / 2)
    If Not (rsBmap.BOF And rsBmap.EOF) Then
        rsBmap.MoveFirst
    End If
    While Not rsBmap.EOF
        FullBmap(rsBmap.Fields("x"), rsBmap.Fields("y")) = rsBmap.Fields("des")
        rsBmap.MoveNext
    Wend
    FullBmapLoaded = True
End If
'load bmap info
If Not (rsBmap.BOF And rsBmap.EOF) Then
    rsBmap.MoveFirst
End If
While Not rsBmap.EOF
    X = SIx(rsBmap.Fields("x"))
    Y = SIy(rsBmap.Fields("y"))
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
        ScreenInfo(X, Y).cDes = FullBmap(rsBmap.Fields("x"), rsBmap.Fields("y"))
        If ColorScheme = COLORSCHEME_GRADUATED Then
            ScreenInfo(X, Y).iValue = -2
        End If
    End If
    rsBmap.MoveNext
Wend
'load intellegence info
'110103 rjk: Modified using the imported intelligence information to fill in missing bmap info.
If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
    rsEnemySect.MoveFirst
End If
While Not rsEnemySect.EOF
    X = SIx(rsEnemySect.Fields("x"))
    Y = SIy(rsEnemySect.Fields("y"))
    rsSectors.Seek "=", X, Y
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy And _
       rsSectors.NoMatch Then
        '110304 rjk: Switched to IsNull check
        If IsNull(rsEnemySect.Fields("des")) Then
            If ScreenInfo(X, Y).cDes = "" Then
                ScreenInfo(X, Y).cDes = "?"
            End If
        Else
            '112203 rjk: Override not known or unknown 112903 rjk: Added ShareBmap checks as well
            If ScreenInfo(X, Y).cDes = "" Or ScreenInfo(X, Y).cDes = "?" Or _
               (ScreenInfo(X, Y).cDes >= "A" And ScreenInfo(X, Y).cDes <= "W") Or _
               (ScreenInfo(X, Y).cDes >= "Y" And ScreenInfo(X, Y).cDes <= "Z") Or _
               IsNumeric(ScreenInfo(X, Y).cDes) Or ScreenInfo(X, Y).cDes = vbNullChar _
               Then
                ScreenInfo(X, Y).cDes = rsEnemySect.Fields("des")
            ElseIf ScreenInfo(X, Y).cDes <> rsEnemySect.Fields("des") And _
                   rsEnemySect.Fields("des") <> "?" Then
                ScreenInfo(X, Y).cSecondDes = rsEnemySect.Fields("des")
            End If
        End If
        Select Case ColorScheme
        Case COLORSCHEME_NORMAL
            ScreenInfo(X, Y).bDisplayColor = rsEnemySect.Fields("owner")
        Case COLORSCHEME_GRADUATED '112203 rjk: Added the ability to show values from enemysectors in graduated mode
            If bIncludeEnemySectorsF4Display Then
                ScreenInfo(X, Y).bDisplayColor = rsEnemySect.Fields("owner")
                Select Case GradColorItem
                'x,y,des,owner,oldowner,eff,road,rail,defense,civ,mil,shell,gun,pet,food,iron,bar,timestamp"
                Case "eff", "road", "rail", "defense", "civ", "mil", "shell", "gun", "pet", "food", "iron", "bar"
                    ScreenInfo(X, Y).iValue = rsEnemySect.Fields(GradColorItem)
                Case Else
                    ScreenInfo(X, Y).iValue = -2
                End Select
            End If
        Case COLORSCHEME_CONDITIONAL '112503 rjk: Added
            If bIncludeEnemySectorsF4Display Then
                ScreenInfo(X, Y).bDisplayColor = rsEnemySect.Fields("owner")
                If ConditionMetEnemySector Then
                    ScreenInfo(X, Y).iValue = ScreenInfo(X, Y).iValue + 1
                End If
            End If
        End Select
    End If
    rsEnemySect.MoveNext
Wend
'load sector info
ReDim DistPoint(0)
ReDim DistCount(0)
If Not (rsSectors.BOF And rsSectors.EOF) Then
    rsSectors.MoveFirst
End If
While Not rsSectors.EOF
    X = SIx(rsSectors.Fields("x"))
    Y = SIy(rsSectors.Fields("y"))
    'for distribution based display, must build graduated array using all sectors,
    '   or colors will change based on map view
    If ColorScheme = COLORSCHEME_DISTRIBUTION Then
        x1 = rsSectors.Fields("dist_x")
        y1 = rsSectors.Fields("dist_y")
        DistColor = 0
        If Not (x1 = rsSectors.Fields("x") And y1 = rsSectors.Fields("y")) Then
            strKey = CStr(x1) + "," + CStr(y1)
            For n = 1 To UBound(DistPoint)
                If strKey = DistPoint(n) Then
                    DistColor = n
                    DistCount(n) = DistCount(n) + 1
                End If
            Next n
            If DistColor = 0 Then
                ReDim Preserve DistPoint(UBound(DistPoint) + 1)
                ReDim Preserve DistCount(UBound(DistCount) + 1)
                DistPoint(UBound(DistPoint)) = strKey
                DistCount(UBound(DistPoint)) = 1
                DistColor = UBound(DistPoint)
            End If
        End If
    End If
    
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
        If bDeity And rsSectors.Fields("country") <> 0 And ColorScheme = COLORSCHEME_NORMAL Then
            ScreenInfo(X, Y).bOwner = False
        Else
            ScreenInfo(X, Y).bOwner = True
        End If
        If dspDesignation = DD_NEW And rsSectors.Fields("sdes") <> " " Then '112203 Added the ability to show new designation
            ScreenInfo(X, Y).cDes = rsSectors.Fields("sdes")
        Else
            ScreenInfo(X, Y).cDes = rsSectors.Fields("des")
            If dspDesignation = DD_BOTH And rsSectors.Fields("sdes") <> " " Then
                ScreenInfo(X, Y).cSecondDes = rsSectors.Fields("sdes")
            End If
        End If
        If bDeity And ColorScheme = COLORSCHEME_NORMAL Then
            ScreenInfo(X, Y).bDisplayColor = rsSectors.Fields("country")
        ElseIf bDeity And ColorScheme = COLORSCHEME_DISTRIBUTION And rsSectors.Fields("country") = 0 Then
            ScreenInfo(X, Y).bDisplayColor = 0
        ElseIf rsSectors.Fields("*") = "*" And ColorScheme = COLORSCHEME_NORMAL Then '100803 rjk: Added 'Not Yours' identify to Map display
            ScreenInfo(X, Y).bNotYours = True
        Else
            ScreenInfo(X, Y).bNotYours = False
        End If
    
        Select Case ColorScheme
        Case COLORSCHEME_NORMAL
        Case COLORSCHEME_CONDITIONAL
            'test if condition is true or false
            If ConditionMet Then
                ScreenInfo(X, Y).iValue = 1
            Else
                ScreenInfo(X, Y).iValue = 0
            End If
        
        Case COLORSCHEME_GRADUATED
            Select Case GradColorItem
            Case "plague risk percentage chance"
                If PlagueRisk(rsSectors) > 1# Then '122903 rjk: Change to show percentage chance
                    ScreenInfo(X, Y).iValue = Int(PlagueRisk(rsSectors))
                Else
                    ScreenInfo(X, Y).iValue = 0
                End If
            Case "plague risk calculation X 10" '121803 rjk: added
                ScreenInfo(X, Y).iValue = CInt(PlagueRisk(rsSectors) * 10#)
            Case "mobility cost"
                Dim sMobCost As Single
                sMobCost = SectorMobCost(rsSectors.Fields("x"), rsSectors.Fields("y"), MT_COMMODITY)
                If sMobCost > 9# Then
                    ScreenInfo(X, Y).iValue = 999
                Else
                    ScreenInfo(X, Y).iValue = CInt(100# * sMobCost)
                End If
            Case "food required" '110103 rjk: Added food required to the graduated colorscheme
                ScreenInfo(X, Y).iValue = FoodRequired(rsSectors)
            Case "food shortage" '110503 rjk: Added food shortage to the graduated colorscheme
                If FoodRequired(rsSectors) > rsSectors.Fields("food") Then
                    ScreenInfo(X, Y).iValue = FoodRequired(rsSectors) - rsSectors.Fields("food")
                Else
                    ScreenInfo(X, Y).iValue = 0
                End If
            Case "food excess" '112403 rjk: Added food shortage to the graduated colorscheme
                ScreenInfo(X, Y).iValue = rsSectors.Fields("food") - FoodRequired(rsSectors)
            Case "excess civilians" '050404 rjk: Added for Excess civilians
                v = Production(rsSectors)
                If (v.prodValidFlag) Then
                    ScreenInfo(X, Y).iValue = CInt(v.ExcessCiv)
                Else
                    ScreenInfo(X, Y).iValue = 0
                End If
            Case "build cost"
                v = Production(rsSectors)
                If (v.prodValidFlag) Then
                    ScreenInfo(X, Y).iValue = CInt(v.BuildAvailCost)
                Else
                    ScreenInfo(X, Y).iValue = 0
                End If
            Case "reacting planes" '100605 rjk: Added for reacting planes
                ScreenInfo(X, Y).iValue = UpdateScreenInfoWithReactingPlanes(X, Y)
            Case "production"
                v = Production(rsSectors)
                If (v.prodValidFlag) Then
                    ScreenInfo(X, Y).iValue = CInt(v.ProdAmount)
                Else
                    ScreenInfo(X, Y).iValue = 0
                End If
            Case Else
                If IsNull(rsSectors.Fields(GradColorItem)) Then
                    ScreenInfo(X, Y).iValue = -1
                Else
                    ScreenInfo(X, Y).iValue = rsSectors.Fields(GradColorItem)
                End If
            End Select
        
        '102003 rjk: Modified to store delivery information as well
        Case COLORSCHEME_DELIVERIES
            Select Case rsSectors.Fields(strCommodity + "_del")
            Case "."
                ScreenInfo(X, Y).iValue = 0
            Case "g"
                ScreenInfo(X, Y).iValue = 1
            Case "y"
                ScreenInfo(X, Y).iValue = 2
            Case "u"
                ScreenInfo(X, Y).iValue = 3
            Case "j"
                ScreenInfo(X, Y).iValue = 4
            Case "n"
                ScreenInfo(X, Y).iValue = 5
            Case "b"
                ScreenInfo(X, Y).iValue = 6
            Case "0", "1", "2", "3", "4", "5", "6", "7"
                ScreenInfo(X, Y).iValue = 7
            End Select
            ScreenInfo(X, Y).iValue = ScreenInfo(X, Y).iValue + (rsSectors.Fields(strCommodity + "_cut") * 8#)
            
        Case COLORSCHEME_DISTRIBUTION
            '112203 rjk: Added displaying the threshold
            ScreenInfo(X, Y).bDisplayColor = DistColor Mod (NUMBER_OF_PLAYER_COLORS + 1)
            ScreenInfo(X, Y).iValue = rsSectors.Fields(strCommodity + "_dist")
            
        End Select
    End If
    rsSectors.MoveNext
Wend

If ColorScheme = COLORSCHEME_GRADUATED And GradColorDspVal Then
    If (GradColorItem = "fert" Or GradColorItem = "ocontent") Then
        If Not (rsSea.BOF And rsSea.EOF) Then
            rsSea.MoveFirst
        End If
        While Not rsSea.EOF
            X = SIx(rsSea.Fields("x"))
            Y = SIy(rsSea.Fields("y"))
            If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
                ScreenInfo(X, Y).iValue = rsSea.Fields(GradColorItem)
            End If
            rsSea.MoveNext
        Wend
    End If
End If

'Add event markers for distribution centers
If ColorScheme = COLORSCHEME_DISTRIBUTION Then
    EventMarkers.Clear
    For n = 1 To UBound(DistPoint)
        If ParseSectors(X, Y, DistPoint(n)) Then
            x1 = n
            While x1 > NUMBER_OF_PLAYER_COLORS
                x1 = x1 - NUMBER_OF_PLAYER_COLORS
            Wend
            EventMarkers.Add X, Y, x1, "DP for " + CStr(DistCount(n)) + " Hexs"
        End If
    Next n
End If
End Sub

Function SIx(X As Integer) As Integer
If X < OriginX Then
    SIx = X + MapSizeX
Else
    SIx = X
End If
End Function

Function SIy(Y As Integer) As Integer
If Y < OriginY Then
    SIy = Y + MapSizeY
Else
    SIy = Y
End If

End Function

Public Sub UpdateScreenInfo(sx As Integer, sy As Integer)

Dim ix As Integer
Dim iy As Integer
rsSectors.Seek "=", sx, sy
If Not rsSectors.NoMatch Then
    ix = SIx(sx)
    iy = SIy(sy)
    ScreenInfo(ix, iy).cDes = rsSectors.Fields("des")
End If
End Sub

Public Function GetConditionalSectors() As String
Dim strSector As String

If ColorScheme = COLORSCHEME_CONDITIONAL Then
    rsSectors.MoveFirst
    Do While rsSectors.EOF = False
        If ConditionMet Then
            If Len(strSector) = 0 Then
                strSector = SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            Else
                strSector = strSector + "\" + SectorString(rsSectors.Fields("x"), rsSectors.Fields("y"))
            End If
        End If
        rsSectors.MoveNext
    Loop
Else
    strSector = ""
End If
GetConditionalSectors = strSector
End Function

'102303 rjk: Added to Count the number Multiple Sectors in the Sector field
Public Function CountMultipleSectors(strSector As String) As Integer
Dim nCount As Integer
Dim pos As Integer

If Len(strSector) = 0 Then
    CountMultipleSectors = 0
    Exit Function
End If

pos = 0
nCount = 0
Do
    nCount = nCount + 1
    pos = InStr(pos + 1, strSector, "\")
Loop While (pos <> 0)

CountMultipleSectors = nCount
End Function

Public Sub GetUnitCount(ScreenInfo() As tScreenInfo, topx As Integer, topy As Integer)
Dim X As Integer
Dim Y As Integer
Dim dExpiry As Date
Dim dRecord As Date

dExpiry = DateAdd("d", -frmUnitView.unitView.iExpiry, Now)

'load ships onto screen
If Not (rsShip.BOF And rsShip.EOF) Then
    rsShip.MoveFirst
End If
While Not rsShip.EOF
    '092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding
    '(ShowOnlyWarships And (rsShip.Fields("fir") > 0 Or rsShip.Fields("type") = "ls")) Then
    If Not frmOptions.bShowOnlyWarships Or _
       (frmOptions.bShowOnlyWarships And (rsShip.Fields("fir") > 0 Or _
       frmDrawMap.UnitCharacteristicCheck(TYPE_S_LAND, rsShip.Fields("type")))) Then
        X = SIx(rsShip.Fields("x"))
        Y = SIy(rsShip.Fields("y"))
        If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
            ScreenInfo(X, Y).iUnitCounts(VU_OUR_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_OUR_SHIPS) + 1
        End If
    End If
    rsShip.MoveNext
Wend
'load lands onto screen
If Not (rsLand.BOF And rsLand.EOF) Then
    rsLand.MoveFirst
End If
While Not rsLand.EOF
    X = SIx(rsLand.Fields("x"))
    Y = SIy(rsLand.Fields("y"))
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
        ScreenInfo(X, Y).iUnitCounts(VU_OUR_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_OUR_LAND_UNITS) + 1
    End If
    rsLand.MoveNext
Wend
'load planes onto screen
If Not (rsPlane.BOF And rsPlane.EOF) Then
    rsPlane.MoveFirst
End If
While Not rsPlane.EOF
    X = SIx(rsPlane.Fields("x"))
    Y = SIy(rsPlane.Fields("y"))
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
        ScreenInfo(X, Y).iUnitCounts(VU_OUR_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_OUR_PLANES) + 1
    End If
    rsPlane.MoveNext
Wend
'load nukes onto screen
If Not (rsNuke.BOF And rsNuke.EOF) Then
    rsNuke.MoveFirst
End If
While Not rsNuke.EOF
    X = SIx(rsNuke.Fields("x"))
    Y = SIy(rsNuke.Fields("y"))
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
        ScreenInfo(X, Y).iUnitCounts(VU_OUR_NUKES) = ScreenInfo(X, Y).iUnitCounts(VU_OUR_NUKES) + 1
    End If
    rsNuke.MoveNext
Wend
'load enemy units onto screen
If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
    rsEnemyUnit.MoveFirst
End If
    
While Not rsEnemyUnit.EOF
    X = SIx(rsEnemyUnit.Fields("x"))
    Y = SIy(rsEnemyUnit.Fields("y"))
    If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
        dRecord = ConvertToLocalTimeFromWinACEUTC(rsEnemyUnit.Fields("timestamp").Value)
        If frmUnitView.unitView.iExpiry = 0 Or _
           DateDiff("d", dExpiry, dRecord) > 0 Then
            Select Case rsEnemyUnit.Fields("class").Value
            Case "l"
                Select Case Nations.YourRelation(rsEnemyUnit.Fields("owner"))
                Case iREL_AT_WAR, iREL_HOSTILE, iREL_MOBILIZING, iREL_SITZKRIEG
                    ScreenInfo(X, Y).iUnitCounts(VU_ENEMY_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_ENEMY_LAND_UNITS) + 1
                Case iREL_ALLIED
                    ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_LAND_UNITS) + 1
                Case iREL_FRIENDLY
                    If frmOptions.friendlyUnit = FRIEND_NEUTRAL Then
                        ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_LAND_UNITS) + 1
                    Else
                        ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_LAND_UNITS) + 1
                    End If
                Case Else 'iREL_NEUTRAL or unknown or offline 120103 rjk: Added offline
                    ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_LAND_UNITS) + 1
                End Select
            Case "s"
                Select Case Nations.YourRelation(rsEnemyUnit.Fields("owner"))
                Case iREL_AT_WAR, iREL_HOSTILE, iREL_MOBILIZING, iREL_SITZKRIEG
                    ScreenInfo(X, Y).iUnitCounts(VU_ENEMY_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_ENEMY_SHIPS) + 1
                Case iREL_ALLIED
                    ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_SHIPS) + 1
                Case iREL_FRIENDLY
                    If frmOptions.friendlyUnit = FRIEND_NEUTRAL Then
                        ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_SHIPS) + 1
                    Else
                        ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_SHIPS) + 1
                    End If
                Case Else 'iREL_NEUTRAL or unknown or offline 120103 rjk: Added offline
                    ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_SHIPS) + 1
                End Select
            Case "p"
                Select Case Nations.YourRelation(rsEnemyUnit.Fields("owner"))
                Case iREL_AT_WAR, iREL_HOSTILE, iREL_MOBILIZING, iREL_SITZKRIEG
                    ScreenInfo(X, Y).iUnitCounts(VU_ENEMY_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_ENEMY_PLANES) + 1
                Case iREL_ALLIED
                    ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_PLANES) + 1
                Case iREL_FRIENDLY
                    If frmOptions.friendlyUnit = FRIEND_NEUTRAL Then
                        ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_PLANES) + 1
                    Else
                        ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_ALLIED_PLANES) + 1
                    End If
                End Select
                Case Else 'iREL_NEUTRAL or unknown or offline 120103 rjk: Added offline
                    ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_NEUTRAL_PLANES) + 1
            End Select
        Else
            Select Case rsEnemyUnit.Fields("class").Value
            Case "l"
                ScreenInfo(X, Y).iUnitCounts(VU_EXPIRED_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(VU_EXPIRED_LAND_UNITS) + 1
            Case "s"
                ScreenInfo(X, Y).iUnitCounts(VU_EXPIRED_SHIPS) = ScreenInfo(X, Y).iUnitCounts(VU_EXPIRED_SHIPS) + 1
            Case "p"
                ScreenInfo(X, Y).iUnitCounts(VU_EXPIRED_PLANES) = ScreenInfo(X, Y).iUnitCounts(VU_EXPIRED_PLANES) + 1
            End Select
        End If
    End If
    rsEnemyUnit.MoveNext
Wend
End Sub

Public Function UpdateScreenInfoWithReactingPlanes(X As Integer, Y As Integer) As Integer
Dim hldIndex As String
Dim iOpSecX As Integer
Dim iOpSecY As Integer
Dim iCount As Integer

If rsBmap.BOF And rsBmap.EOF Then
    UpdateScreenInfoWithReactingPlanes = 0
    Exit Function
End If

iCount = 0
hldIndex = rsMissions.Index
If rsMissions.EOF And rsMissions.BOF Then
    UpdateScreenInfoWithReactingPlanes = 0
    Exit Function
End If
rsMissions.Index = "Mission"
rsMissions.MoveFirst
rsMissions.Seek "=", "p", "air defense"
If rsMissions.NoMatch Then
    UpdateScreenInfoWithReactingPlanes = 0
    rsMissions.Index = hldIndex
    Exit Function
End If
Do Until rsMissions.EOF
    If rsMissions.Fields("type") <> "p" Or rsMissions.Fields("mission") <> "air defense" Then
        Exit Do
    End If
    ParseSectors iOpSecX, iOpSecY, rsMissions.Fields("op sector")
    If SectorDistance(X, Y, iOpSecX, iOpSecY) <= rsMissions("radius") Then
        iCount = iCount + 1
    End If
    rsMissions.MoveNext
Loop
rsMissions.Index = hldIndex
UpdateScreenInfoWithReactingPlanes = iCount
End Function
