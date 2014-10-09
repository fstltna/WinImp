Option Strict Off
Option Explicit On
Module HexGrid
	
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
	
	Public NbrHexWide As Short
	Public NbrHexTall As Short
	Private HexSideLength As Single
	Private HexFontSize As Single
	Public HSL866 As Single
	Public HSL5 As Single
	Public OriginX As Short
	Public OriginY As Short
	Public CurrentSectorX As Short
	Public CurrentSectorY As Short
	'Colors used in drawing the hex grid
	Public Const NUMBER_OF_PLAYER_COLORS As Short = 40
	'100803 rjk: Change from 8 to 9 to add 'Not Yours'
	'102003 rjk: from 9 to 11 for Delivery colors
	Public Const NUMBER_OF_BASIC_COLORS As Short = 11
	Public PlayerColors(NUMBER_OF_PLAYER_COLORS) As Integer
	Public PlayerText(NUMBER_OF_PLAYER_COLORS) As Integer
	'UPGRADE_WARNING: Lower bound of array BasicColors was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public BasicColors(NUMBER_OF_BASIC_COLORS) As Integer
	'UPGRADE_WARNING: Lower bound of array BasicText was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public BasicText(NUMBER_OF_BASIC_COLORS) As Integer
	
	Public Const vbMediumGreen As Integer = &HC000
	Public Const vbDarkGreen As Integer = &H8000
	Public Const vbDarkBlue As Integer = &H800000
	
	'Color scheme used in drawing the hex grid
	Public Const MAX_COLOR_CONDITIONS As Short = 5
	Public Const COLORSCHEME_NORMAL As Short = 0
	Public Const COLORSCHEME_CONDITIONAL As Short = 1
	Public Const COLORSCHEME_GRADUATED As Short = 2
	Public Const COLORSCHEME_DISTRIBUTION As Short = 3
	Public Const COLORSCHEME_DELIVERIES As Short = 4 '102003 rjk: Added for Delivery ColorScheme
	Public ColorScheme As Short
	Public GradColorItem As String
	Public GradColorHigh As Integer
	Public GradColorLow As Integer
	Public GradColorDspVal As Boolean
	Public CondColorParms() As Object
	Public CondColorParmNumber As Short
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
	Public Const HEX_TOPLEFT As Short = 1
	Public Const HEX_TOPRIGHT As Short = 2
	Public Const HEX_RIGHT As Short = 3
	Public Const HEX_BOTTOMRIGHT As Short = 4
	Public Const HEX_BOTTOMLEFT As Short = 5
	Public Const HEX_LEFT As Short = 6
	
	'Drawing Font
	Const MAP_DRAWING_FONT As String = "Rockwell"
	Const MAP_BACKUP_FONT As String = "Times New Roman"
	
	'****************************************************************
	'Windows API/Global Declarations for Draw Filled Objects
	'****************************************************************
	Structure POINTAPI
		Dim X As Integer
		Dim Y As Integer
	End Structure
	
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function Polygon Lib "GDI32" (ByVal hDC As Integer, ByRef lpPoint As POINTAPI, ByVal nCount As Integer) As Integer
	Declare Function SetPolyFillMode Lib "GDI32" (ByVal hDC As Integer, ByVal nPolyFillMode As Integer) As Integer
	
	'PolyFill Modes
	Public Const ALTERNATE As Short = 1
	Public Const WINDING As Short = 2
	
	Dim Ply(10) As POINTAPI
	
	Public Structure tScreenInfo
		<VBFixedArray(15)> Dim iUnitCounts() As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public cDes() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=1)> Public cSecondDes() As Char
		Dim iValue As Short
		Dim bDisplayColor As Byte
		Dim bOwner As Boolean
		Dim bNotYours As Boolean
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim iUnitCounts(15)
		End Sub
	End Structure
	Dim ScreenInfo() As tScreenInfo
	
	'UPGRADE_ISSUE: Declaration type not supported: Array of fixed-length strings. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
	Dim FullBmap() As String*1
	'UPGRADE_WARNING: Lower bound of array TextSize was changed from 32,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim TextSize(126, 2) As Single
	Public bMapBitMapLoaded As Boolean
	
	Public Sub DrawGrid(ByRef p As Object)
		On Error GoTo Empire_Error
		Dim OrgX As Short
		Dim OrgY As Short
		Dim X As Short
		Dim Y As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.DrawWidth = 2 ' Set DrawWidth.
		'p.ForeColor = QBColor(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.ForeColor = frmOptions.lBorderColor
		'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.FillStyle = 0 ' Sold Fill
		
		'set PolyFillMode to fill whole shape
		'UPGRADE_WARNING: Couldn't resolve default property of object p.hDC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetPolyFillMode(p.hDC, WINDING)
		
		If HexSideLength = 0 Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ScaleWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NbrHexWide = Int(p.ScaleWidth / (HSL866 * 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ScaleHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NbrHexTall = Int(p.ScaleHeight / (HSL5 * 3))
		
		OrgX = OriginX
		OrgY = OriginY
		
		If Not bMapBitMapLoaded Then
			'UPGRADE_NOTE: Object frmDrawMap.picMap.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			frmDrawMap.picMap.Image = Nothing
			If frmOptions.strImageFileName <> "" Then
				frmDrawMap.picMap.Image = System.Drawing.Image.FromFile(frmOptions.strImageFileName)
			End If
			bMapBitMapLoaded = True
		End If
		
		LoadScreen()
		
		Dim PosX As Single
		Dim PosY As Single
		For Y = 1 To NbrHexTall
			For X = 1 To NbrHexWide
				
				'Get Center position and draw hex
				Coord(OrgX, OrgY, PosX, PosY)
				DrawSector(p, PosX, PosY, OrgX, OrgY)
				OrgX = OrgX + 2
				If OrgX >= (MapSizeX / 2) Then
					OrgX = OrgX - MapSizeX
				End If
				
			Next X
			OrgY = OrgY + 1
			If OrgY > MapSizeY / 2 Then
				OrgY = OrgY - MapSizeY
			End If
			OrgX = OriginX + System.Math.Abs(OrgY Mod 2)
		Next Y
		Exit Sub
Empire_Error: 
		Dim nReply As Short
		If Err.Number = 9 Then
			'Probably have version information off
			nReply = MsgBox("Error in Map routines, World size may have changes since last time you ran version." & vbNewLine & "Press OK to update version informantion", MsgBoxStyle.Exclamation + MsgBoxStyle.OKCancel)
			If nReply = MsgBoxResult.OK Then
				frmEmpCmd.SubmitEmpireCommand("version", True)
			End If
		End If
		EmpireError("DrawGrid", vbNullString, vbNullString)
		
	End Sub
	
	Public Sub Coord(ByRef sectx As Short, ByRef secty As Short, ByRef PosX As Single, ByRef PosY As Single)
		
		Dim offsetX As Short
		Dim offsetY As Short
		
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
	
	Public Sub DrawHex(ByRef f As Object, ByRef CenterX As Single, ByRef CenterY As Single)
		
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
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Line((CenterX, CenterY - HexSideLength) - (CenterX - HSL866, CenterY - HSL5))
		'HEX_LEFT:
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Line((CenterX - HSL866, CenterY + HSL5))
		'HEX_BOTTOMLEFT:
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Line((CenterX, CenterY + HexSideLength))
		'HEX_BOTTOMRIGHT:
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Line((CenterX + HSL866, CenterY + HSL5))
		'HEX_RIGHT:
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Line((CenterX + HSL866, CenterY - HSL5))
		'HEX_TOPRIGHT:
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Line((CenterX, CenterY - HexSideLength))
		
	End Sub
	Public Sub DrawHexCenterLine(ByRef f As Object, ByRef CenterX As Single, ByRef CenterY As Single, ByRef lenline As Single, ByRef Side As Short)
		
		Select Case Side
			Case HEX_TOPLEFT
				'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f.Line((CenterX, CenterY) - (CenterX - 0.5 * lenline, CenterY - 0.8666 * lenline))
			Case HEX_TOPRIGHT
				'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f.Line((CenterX, CenterY) - (CenterX + 0.5 * lenline, CenterY - 0.8666 * lenline))
			Case HEX_RIGHT
				'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f.Line((CenterX, CenterY) - (CenterX + lenline, CenterY))
			Case HEX_BOTTOMRIGHT
				'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f.Line((CenterX, CenterY) - (CenterX + 0.5 * lenline, CenterY + 0.8666 * lenline))
			Case HEX_BOTTOMLEFT
				'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f.Line((CenterX, CenterY) - (CenterX - 0.5 * lenline, CenterY + 0.8666 * lenline))
			Case HEX_LEFT
				'UPGRADE_WARNING: Couldn't resolve default property of object f.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f.Line((CenterX, CenterY) - (CenterX - lenline, CenterY))
		End Select
		
	End Sub
	
	Sub DrawSolidHex(ByRef p As Object, ByRef CenterX As Single, ByRef CenterY As Single, ByRef FColor As Integer)
		Dim nPoints As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.FillColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		'UPGRADE_WARNING: Couldn't resolve default property of object p.hDC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Polygon(p.hDC, Ply(1), nPoints)
	End Sub
	
	Public Sub SetLetterSize(ByRef p As Object)
		'Set Font Size
		SetDrawingFont(p)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.FontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.FontSize = 48 ' Set font size.
		'UPGRADE_WARNING: Couldn't resolve default property of object p.FontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		While p.TextHeight("l") > (HexSideLength * 1.25) And p.FontSize > 3
			'UPGRADE_WARNING: Couldn't resolve default property of object p.FontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			p.FontSize = p.FontSize - 1
		End While
	End Sub
	
	'112903 rjk: Removed ViewUnits and replaced with unitView.VisibleUnit
	Public Sub DrawSector(ByRef p As Object, ByRef CenterX As Single, ByRef CenterY As Single, ByRef sx As Short, ByRef sy As Short)
		Dim ch As String
		Dim ch2 As String
		Dim n As Short
		Dim Percent As Single
		Dim oldcolor As Object
		Dim oldfontsize As Object
		Dim ix As Short
		Dim iy As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldcolor = p.ForeColor
		'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldfontsize = p.Font.Size
		ix = SIx(sx)
		iy = SIy(sy)
		
		'draw sectors then
		If ScreenInfo(ix, iy).cDes = vbNullChar Then
			'uncharted territory
			DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(4), BasicText(4)))
			'dummy p.Print is necessary to get screen to update when
			'only one sector is updated.
			ch = " "
			'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			p.Print(ch)
		Else
			ch = ScreenInfo(ix, iy).cDes
			
			If ch = "." Then 'sea are drawn in blue
				If CDbl(CObj(frmDrawMap.picMap.Image)) <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_ISSUE: Constant vbFSTransparent was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
					p.FillStyle = vbFSTransparent
				End If
				DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2)))
				'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: Constant vbFSSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				p.FillStyle = vbFSSolid
				'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
						'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.Font.Size = GetHexFontSize * 0.7
						'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.CurrentX = CenterX - p.TextWidth(ch2) / 2
						'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
						'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.Print(ch2)
					End If
				End If
				If frmUnitView.unitView.VisibleUnits(ScreenInfo(ix, iy).iUnitCounts) Or ch2 <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					p.Font.Size = GetHexFontSize * 0.7
					'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					p.CurrentX = CenterX - p.TextWidth(ch) / 2
					'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					p.CurrentY = CenterY
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					p.CurrentX = CenterX - p.TextWidth(ch) / 2
					'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					p.CurrentY = CenterY - p.TextHeight(ch) / 2
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.Print(ch)
				
			ElseIf ch = "X" Then  'draw minefields in reverse colors
				DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2)))
				'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.ForeColor = CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2))
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentX = CenterX - TextSize(Asc(ScreenInfo(ix, iy).cDes), 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentY = CenterY - TextSize(Asc(ScreenInfo(ix, iy).cDes), 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.Print(ch)
			Else
				If ScreenInfo(ix, iy).bOwner Then
					'owned sector
					Select Case ColorScheme
						Case COLORSCHEME_NORMAL
							If ch = "=" Then 'bridges over water are drawn in blue
								If CDbl(CObj(frmDrawMap.picMap.Image)) <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: Constant vbFSTransparent was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
									p.FillStyle = vbFSTransparent
								End If
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_ISSUE: Constant vbFSSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
								p.FillStyle = vbFSSolid
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2))
							ElseIf ScreenInfo(ix, iy).bNotYours Then 
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(9), BasicText(9)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(9), BasicColors(9))
							Else
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(1), BasicText(1)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(1), BasicColors(1))
							End If
							DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
							
						Case COLORSCHEME_CONDITIONAL
							'test if condition is true or false
							If ScreenInfo(ix, iy).iValue > 0 Then
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(5), BasicText(1)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(5), BasicColors(1))
							Else
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(6), BasicText(1)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(6), BasicColors(1))
							End If
							DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
							
						Case COLORSCHEME_GRADUATED
							If GradColorHigh = GradColorLow Then
								Percent = 1
							Else
								Percent = (ScreenInfo(ix, iy).iValue - GradColorLow) / (GradColorHigh - GradColorLow)
							End If
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BlendedColor(BasicColors(8), BasicColors(7), Percent), BasicText(1)))
							If Percent > 0.5 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(7), BasicColors(1))
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(8), BasicColors(1))
							End If
							If GradColorDspVal Then
								ch2 = CStr(ScreenInfo(ix, iy).iValue)
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Font.Size = GetHexFontSize * 0.7
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentX = CenterX - p.TextWidth(ch2) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Print(ch2)
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentX = CenterX - p.TextWidth(ch) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentY = CenterY
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Print(ch)
							Else
								DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
							End If
							
						Case COLORSCHEME_DISTRIBUTION
							'112203 rjk: Added displaying the threshold
							n = ScreenInfo(ix, iy).bDisplayColor
							n = System.Math.Abs(n) Mod (NUMBER_OF_PLAYER_COLORS + 1)
							If ch = "=" Then
								If CDbl(CObj(frmDrawMap.picMap.Image)) <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: Constant vbFSTransparent was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
									p.FillStyle = vbFSTransparent
								End If
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), PlayerColors(n)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_ISSUE: Constant vbFSSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
								p.FillStyle = vbFSSolid
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), BasicColors(2))
							Else
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, PlayerText(n), PlayerColors(n))
							End If
							If ScreenInfo(ix, iy).iValue = 0 Then 'no threshold
								DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
							Else
								ch2 = CStr(ScreenInfo(ix, iy).iValue)
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Font.Size = GetHexFontSize * 0.7
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentX = CenterX - p.TextWidth(ch2) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Print(ch2)
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentX = CenterX - p.TextWidth(ch) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentY = CenterY
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Print(ch)
							End If
							'102003 rjk: Added Delivery ColorScheme
						Case COLORSCHEME_DELIVERIES
							If ScreenInfo(ix, iy).iValue <> 0 Then
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(10), BasicText(1)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(10), BasicColors(1))
								Select Case ScreenInfo(ix, iy).iValue Mod 8
									Case 1
										'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5))
									Case 2
										'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
									Case 3
										'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
									Case 4
										'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5))
									Case 5
										'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										p.Line((CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
									Case 6
										'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										p.Line((CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
									Case 7
								End Select
							Else
								DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(11), BasicText(1)))
								'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.ForeColor = CheckSelectedHex(sx, sy, BasicText(11), BasicColors(1))
							End If
							ch2 = CStr(ScreenInfo(ix, iy).iValue \ 8)
							'                If ch2 <> 0 Then
							If ScreenInfo(ix, iy).iValue <> 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Font.Size = GetHexFontSize * 0.7
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentX = CenterX - p.TextWidth(ch2) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Print(ch2)
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentX = CenterX - p.TextWidth(ch) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.CurrentY = CenterY
								'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								p.Print(ch)
							Else
								DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
							End If
					End Select
				ElseIf ScreenInfo(ix, iy).bDisplayColor = 0 Then 
					'bmap sector
					If ch = "=" Then
						If CDbl(CObj(frmDrawMap.picMap.Image)) <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_ISSUE: Constant vbFSTransparent was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
							p.FillStyle = vbFSTransparent
						End If
						DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), BasicText(2)))
						'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_ISSUE: Constant vbFSSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
						p.FillStyle = vbFSSolid
						'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.ForeColor = CheckSelectedHex(sx, sy, BasicText(2), BasicColors(2))
					ElseIf frmOptions.bStarWars And ch = "\" Then 
						DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, RGB(255, 128, 64), BasicText(3)))
						'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.ForeColor = CheckSelectedHex(sx, sy, BasicText(3), RGB(255, 128, 64))
					Else
						DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(3), BasicText(3)))
						'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.ForeColor = CheckSelectedHex(sx, sy, BasicText(3), BasicColors(3))
					End If
					DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
					
				Else 'sector with enemy info
					n = ScreenInfo(ix, iy).bDisplayColor
					While n > NUMBER_OF_PLAYER_COLORS
						n = n - NUMBER_OF_PLAYER_COLORS
					End While
					'112203 rjk: Added displaying values for enemy sectors in graduated mode
					If ColorScheme = COLORSCHEME_GRADUATED And bIncludeEnemySectorsF4Display And GradColorDspVal And ScreenInfo(ix, iy).iValue <> -2 Then
						If GradColorHigh = GradColorLow Then
							Percent = 1
						Else
							Percent = (ScreenInfo(ix, iy).iValue - GradColorLow) / (GradColorHigh - GradColorLow)
						End If
						If ch = "=" Then
							If CDbl(CObj(frmDrawMap.picMap.Image)) <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_ISSUE: Constant vbFSTransparent was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
								p.FillStyle = vbFSTransparent
							End If
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), PlayerColors(n)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_ISSUE: Constant vbFSSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
							p.FillStyle = vbFSSolid
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), BasicColors(2))
						Else
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BlendedColor(BasicColors(8), BasicColors(7), Percent), BasicText(1)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
						End If
						If Percent > 0.5 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, BasicText(7), BasicColors(1))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, BasicText(8), BasicColors(1))
						End If
						If ScreenInfo(ix, iy).iValue = -1 Then
							ch2 = "???"
						Else
							ch2 = CStr(ScreenInfo(ix, iy).iValue)
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.Font.Size = GetHexFontSize * 0.7
						'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.CurrentX = CenterX - p.TextWidth(ch2) / 2
						'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
						'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.Print(ch2)
						'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.CurrentX = CenterX - p.TextWidth(ch) / 2
						'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.CurrentY = CenterY
						'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.Print(ch)
						'112503 rjk: Added the ability to select conditional display from rsEnemySector table.
					ElseIf ColorScheme = COLORSCHEME_CONDITIONAL And bIncludeEnemySectorsF4Display Then 
						'test if condition is true or false
						If ScreenInfo(ix, iy).iValue > 0 Then
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(5), BasicText(1)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, BasicText(5), BasicColors(1))
						Else
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(6), BasicText(1)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - HSL866 + p.DrawWidth, CenterY + HSL5))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX - p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + p.DrawWidth, CenterY - HexSideLength + (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY - HSL5 + p.DrawWidth) - (CenterX + HSL866 - p.DrawWidth, CenterY + HSL5))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX + HSL866 - p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX + p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object p.Line. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.Line((CenterX - HSL866 + p.DrawWidth, CenterY + HSL5 - p.DrawWidth) - (CenterX - p.DrawWidth, CenterY + HexSideLength - (p.DrawWidth * 2)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, BasicText(6), BasicColors(1))
						End If
						DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
					Else
						If ch = "=" Then
							If CDbl(CObj(frmDrawMap.picMap.Image)) <> 0 And (sx <> CurrentSectorX Or sy <> CurrentSectorY) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_ISSUE: Constant vbFSTransparent was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
								p.FillStyle = vbFSTransparent
							End If
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, BasicColors(2), PlayerColors(n)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.FillStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_ISSUE: Constant vbFSSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
							p.FillStyle = vbFSSolid
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, PlayerColors(n), BasicColors(2))
						Else
							DrawSolidHex(p, CenterX, CenterY, CheckSelectedHex(sx, sy, PlayerColors(n), PlayerText(n)))
							'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							p.ForeColor = CheckSelectedHex(sx, sy, PlayerText(n), PlayerColors(n))
						End If
						'110403 rjk: Show both designations bmap and rsEnemySector if they are different.
						DrawDesignation(ScreenInfo, p, CenterX, CenterY, ix, iy, ch)
					End If
				End If
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.ForeColor = oldcolor
		'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.Font.Size = oldfontsize
		DrawUnits(p, CenterX, CenterY, sx, sy, HexSideLength)
	End Sub
	
	Private Sub DrawDesignation(ByRef ScreenInfo() As tScreenInfo, ByRef p As Object, ByRef CenterX As Single, ByRef CenterY As Single, ByRef ix As Short, ByRef iy As Short, ByRef ch As String)
		Dim ch2 As String
		
		If frmUnitView.unitView.VisibleUnits(ScreenInfo(ix, iy).iUnitCounts) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			p.Font.Size = GetHexFontSize * 0.7
			'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			p.CurrentX = CenterX - p.TextWidth(ch) / 2
			'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			p.CurrentY = CenterY
		Else
			If ScreenInfo(ix, iy).cSecondDes <> vbNullChar Then
				ch2 = ScreenInfo(ix, iy).cSecondDes
				'UPGRADE_WARNING: Couldn't resolve default property of object p.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.Font.Size = GetHexFontSize * 0.7
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentX = CenterX - p.TextWidth(ch2) / 2
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentY = CenterY - p.TextHeight(ch2) * 0.8
				'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.Print(ch2)
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentX = CenterX - p.TextWidth(ch) / 2
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentY = CenterY
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentX = CenterX - TextSize(Asc(ScreenInfo(ix, iy).cDes), 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object p.CurrentY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				p.CurrentY = CenterY - TextSize(Asc(ScreenInfo(ix, iy).cDes), 2)
			End If
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object p.Print. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		p.Print(ch)
	End Sub
	
	Private Function CheckSelectedHex(ByRef sx As Short, ByRef sy As Short, ByRef NormalColor As Integer, ByRef SelectedColor As Integer) As Integer
		If sx = CurrentSectorX And sy = CurrentSectorY Then
			CheckSelectedHex = SelectedColor
		Else
			CheckSelectedHex = NormalColor
		End If
	End Function
	
	'112303 rjk: Changed the size and position of the units to allow the display of the designation.
	'            Fixed the restriction code to only display three units.
	Public Sub DrawUnits(ByRef p As Object, ByRef CenterX As Single, ByRef CenterY As Single, ByRef PosX As Short, ByRef PosY As Short, ByRef lenSide As Single)
		Dim n As Short
		Dim pic As System.Drawing.Image
		Dim ix As Short
		Dim iy As Short
		
		ix = SIx(PosX)
		iy = SIy(PosY)
		
		For n = 1 To 3
			pic = frmUnitView.unitView.UnitPicture(ScreenInfo(ix, iy).iUnitCounts, n = 1)
			If pic Is Nothing Then
				Exit Sub
			Else
				Select Case n '120303 rjk: Moved the positions to make more visible
					Case 1
						'UPGRADE_ISSUE: Picture property pic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture property pic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.PaintPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.PaintPicture(pic, CenterX - (0.8 * lenSide), CenterY - (0.5 * lenSide), 0.8 * lenSide, (pic.Height / pic.Width) * 0.8 * lenSide)
					Case 2
						'UPGRADE_ISSUE: Picture property pic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture property pic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.PaintPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.PaintPicture(pic, CenterX, CenterY - (0.5 * lenSide), 0.8 * lenSide, (pic.Height / pic.Width) * 0.8 * lenSide)
					Case 3
						'UPGRADE_ISSUE: Picture property pic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture property pic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_WARNING: Couldn't resolve default property of object p.PaintPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						p.PaintPicture(pic, CenterX - (0.4 * lenSide), CenterY - (0.8 * lenSide), 0.8 * lenSide, (pic.Height / pic.Width) * 0.8 * lenSide)
				End Select
			End If
		Next n
	End Sub
	
	Public Sub SetDefaultPlayerColors()
		PlayerColors(0) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
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
		
		PlayerText(0) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(1) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(2) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(3) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(4) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(5) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(6) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(7) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		PlayerText(8) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(9) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(10) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(11) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(12) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(13) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		PlayerText(14) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		PlayerText(15) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(16) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(17) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		PlayerText(18) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(19) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		PlayerText(20) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		
		Dim n As Short
		For n = 21 To NUMBER_OF_PLAYER_COLORS
			PlayerColors(n) = PlayerColors(n - 20)
			PlayerText(n) = PlayerText(n - 20)
		Next n
		
	End Sub
	
	Public Sub SetDefaultBasicColors()
		BasicColors(1) = RGB(0, 255, 135)
		BasicColors(2) = RGB(0, 150, 255)
		BasicColors(3) = RGB(188, 188, 188)
		BasicColors(4) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		BasicColors(9) = BasicColors(1) '100803 rjk: default 'Not Your' colors to normal color
		
		BasicText(1) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		BasicText(2) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		BasicText(3) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		BasicText(4) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		BasicText(9) = BasicText(1) '100803 rjk: default 'Not Your' colors to normal color
	End Sub
	
	Public Sub SetDefaultComparisionColors()
		BasicColors(5) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		BasicColors(6) = RGB(0, 255, 135)
		
		BasicText(5) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		BasicText(6) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
	End Sub
	
	Public Sub SetDefaultGradColors()
		BasicColors(7) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		BasicColors(8) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		
		BasicText(7) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		BasicText(8) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
	End Sub
	
	Public Sub SetDefaultDeliveryColors()
		BasicColors(10) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		BasicColors(11) = RGB(0, 255, 135)
		
		BasicText(10) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		BasicText(11) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
	End Sub
	
	Public Function ColorComponent(ByRef Color As Integer, ByRef component As Short) As Integer
		Select Case component
			Case 1
				ColorComponent = Color And &HFF
			Case 2
				ColorComponent = (CShort(Color And &HFF00) / &H100) And &HFF
			Case 3
				ColorComponent = CShort(Color And &HFF0000) / &H10000
		End Select
	End Function
	
	Public Function BlendedColor(ByRef Color1 As Integer, ByRef Color2 As Integer, ByRef Percent As Single) As Integer
		
		If Percent >= 1 Then
			BlendedColor = Color2
			Exit Function
		End If
		
		If Percent <= 0 Then
			BlendedColor = Color1
			Exit Function
		End If
		
		Dim r1 As Integer
		Dim r2 As Integer
		Dim g1 As Integer
		Dim g2 As Integer
		Dim b1 As Integer
		Dim b2 As Integer
		Dim r3 As Integer
		Dim g3 As Integer
		Dim b3 As Integer
		
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
		Dim n As Short
		Dim v As productionDataType
		
		Cond = False
		Done = False
		n = 1
		While n <= CondColorParmNumber And Not Done
			'120303 rjk: Added Food required and Food Excess for F4 Conditional
			Select Case CondColorParms(n, 1)
				Case "FR"
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(FoodRequired(rsSectors)), CInt(CondColorParms(n, 3)))
				Case "FS"
					If FoodRequired(rsSectors) - CInt(rsSectors.Fields("food").Value) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), FoodRequired(rsSectors) - CInt(rsSectors.Fields("food").Value), CInt(CondColorParms(n, 3)))
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), 0, CInt(CondColorParms(n, 3)))
					End If
				Case "FE"
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(rsSectors.Fields("food").Value) - FoodRequired(rsSectors), CInt(CondColorParms(n, 3)))
				Case "EC" '050404 rjk: Added for Excess Civilians
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If (v.prodValidFlag) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(v.ExcessCiv), CInt(CondColorParms(n, 3)))
					Else
						Cond = bConditionalMissing
					End If
				Case "BC"
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If (v.prodValidFlag) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object v.BuildAvailCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(v.BuildAvailCost), CInt(CondColorParms(n, 3)))
					Else
						Cond = bConditionalMissing
					End If
				Case "RP" '110605 rjk: Added RP - Reacting Planes
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(UpdateScreenInfoWithReactingPlanes((rsSectors.Fields("x").Value), rsSectors.Fields("y").Value)), CInt(CondColorParms(n, 3)))
				Case "PD"
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If (v.prodValidFlag) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(v.ProdAmount), CInt(CondColorParms(n, 3)))
					Else
						Cond = bConditionalMissing
					End If
				Case Else
					'112503 rjk: Check and make sure the strength fields not null
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rsSectors.Fields(CondColorParms(n, 1)).Value) Then
						If rsSectors.Fields(CondColorParms(n, 1)).Type = DAO.DataTypeEnum.dbInteger Or rsSectors.Fields(CondColorParms(n, 1)).Type = DAO.DataTypeEnum.dbLong Then
							' test numeric values
							'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Cond = NumberConditionEvaluation(CStr(CondColorParms(n, 2)), CInt(rsSectors.Fields(CondColorParms(n, 1)).Value), CInt(CondColorParms(n, 3)))
						Else ' test string values
							Select Case CondColorParms(n, 2)
								Case "="
									'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value = CStr(CondColorParms(n, 3)))
								Case "#"
									'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value <> CStr(CondColorParms(n, 3)))
									'Case "#" 101403 rjk: Removed duplicate case, also had wrong logic.
									'    Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value = CStr(CondColorParms(n, 3)))
								Case "<"
									'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value < CStr(CondColorParms(n, 3)))
								Case ">"
									'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Cond = (rsSectors.Fields(CondColorParms(n, 1)).Value > CStr(CondColorParms(n, 3)))
							End Select
						End If
					Else
						Cond = bConditionalMissing
					End If
			End Select
			'we may not have to process any farther
			'if it is an "and" and the first value is false, or if it is an "or" and the first value is true
			'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(n, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (CondColorParms(n, 4) = "&" And Not Cond) Or (CondColorParms(n, 4) = "|" And Cond) Then
				Done = True
			End If
			n = n + 1
		End While
		
		ConditionMet = Cond
	End Function
	
	'UPGRADE_NOTE: Operator was upgraded to Operator_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function NumberConditionEvaluation(ByRef Operator_Renamed As String, ByRef Value As Integer, ByRef Condition As Integer) As Boolean
		Select Case Operator_Renamed
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
		Dim n As Short
		
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
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value = CInt(CondColorParms(n, 3)))
							Case "#"
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value <> CInt(CondColorParms(n, 3)))
							Case "<"
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value < CInt(CondColorParms(n, 3)))
							Case ">"
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields(CondColorParms(n, 1)).Value > CInt(CondColorParms(n, 3)))
						End Select
					Else
						Cond = bConditionalMissing
					End If
				Case "country"
					If rsEnemySect.Fields(CondColorParms(n, 1)).Value <> -1 Then
						Select Case CondColorParms(n, 2)
							Case "="
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields("owner").Value = CInt(CondColorParms(n, 3)))
							Case "#"
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields("owner").Value <> CInt(CondColorParms(n, 3)))
							Case "<"
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields("owner").Value < CInt(CondColorParms(n, 3)))
							Case ">"
								'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Cond = (rsEnemySect.Fields("owner").Value > CInt(CondColorParms(n, 3)))
						End Select
					Else
						Cond = bConditionalMissing
					End If
				Case Else
					Cond = bConditionalMissing
			End Select
			'we may not have to process any farther
			'if it is an "and" and the first value is false, or if it is an "or" and the first value is true
			'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(n, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (CondColorParms(n, 4) = "&" And Not Cond) Or (CondColorParms(n, 4) = "|" And Cond) Then
				Done = True
			End If
			n = n + 1
		End While
		
		ConditionMetEnemySector = Cond
	End Function
	
	Public Sub SetDrawingFont(ByRef obj As Object)
		
		On Error Resume Next
		
		'set to the backup first in case we have an error
		'when we try to use the primary
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.FontName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.FontName = MAP_BACKUP_FONT
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.FontName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.FontName = MAP_DRAWING_FONT
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.FontBold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.FontBold = True
		
	End Sub
	
	Public Sub CenterMap(ByRef p As Object, ByRef CenterX As Short, ByRef CenterY As Short)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ScaleWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NbrHexWide = Int(p.ScaleWidth / (HexSideLength * 0.866 * 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object p.ScaleHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	Public Sub SetHexSideLength(ByRef p As Object, ByRef hlen As Single)
		HexSideLength = hlen
		HSL866 = HexSideLength * 0.866
		HSL5 = HexSideLength * 0.5
		SetLetterSize(p)
		'UPGRADE_WARNING: Couldn't resolve default property of object p.FontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HexFontSize = p.FontSize
		
		'load textsize array
		Dim n As Short
		For n = 32 To 126
			'UPGRADE_WARNING: Couldn't resolve default property of object p.TextWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TextSize(n, 1) = p.TextWidth(Chr(n)) / 2
			'UPGRADE_WARNING: Couldn't resolve default property of object p.TextHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		Dim X As Short
		Dim Y As Short
		Dim x1 As Short
		Dim y1 As Short
		Dim n As Short
		Dim topx As Short
		Dim topy As Short
		Static FullBmapLoaded As Boolean
		Dim strKey As String
		Dim DistColor As Short
		Dim DistPoint() As String
		Dim DistCount() As Short
		Dim v As productionDataType
		
		topx = OriginX + NbrHexWide * 2 + 2
		topy = OriginY + NbrHexTall
		'UPGRADE_WARNING: Lower bound of array ScreenInfo was changed from OriginX,OriginY to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim ScreenInfo(topx, topy)
		
		'112803 rjk: Moved Determination Unit code to EmpUnitView class
		GetUnitCount(ScreenInfo, topx, topy)
		
		'if supressbmap, the load static array and reload from that to speed
		'up screen redraws, since we only have to load the bmap from the data
		'base once.
		'load static array if not already loaded.
		If (Not FullBmapLoaded) Or (Not frmOptions.bSuppressBmapRefresh) Then
			'UPGRADE_WARNING: Lower bound of array FullBmap was changed from -1 * MapSizeX / 2,-1 * MapSizeY / 2 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim FullBmap(MapSizeX / 2, MapSizeY / 2)
			If Not (rsBmap.BOF And rsBmap.EOF) Then
				rsBmap.MoveFirst()
			End If
			While Not rsBmap.EOF
				FullBmap(rsBmap.Fields("x").Value, rsBmap.Fields("y").Value) = rsBmap.Fields("des").Value
				rsBmap.MoveNext()
			End While
			FullBmapLoaded = True
		End If
		'load bmap info
		If Not (rsBmap.BOF And rsBmap.EOF) Then
			rsBmap.MoveFirst()
		End If
		While Not rsBmap.EOF
			X = SIx(rsBmap.Fields("x").Value)
			Y = SIy(rsBmap.Fields("y").Value)
			If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
				ScreenInfo(X, Y).cDes = FullBmap(rsBmap.Fields("x").Value, rsBmap.Fields("y").Value)
				If ColorScheme = COLORSCHEME_GRADUATED Then
					ScreenInfo(X, Y).iValue = -2
				End If
			End If
			rsBmap.MoveNext()
		End While
		'load intellegence info
		'110103 rjk: Modified using the imported intelligence information to fill in missing bmap info.
		If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
			rsEnemySect.MoveFirst()
		End If
		While Not rsEnemySect.EOF
			X = SIx(rsEnemySect.Fields("x").Value)
			Y = SIy(rsEnemySect.Fields("y").Value)
			rsSectors.Seek("=", X, Y)
			If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy And rsSectors.NoMatch Then
				'110304 rjk: Switched to IsNull check
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rsEnemySect.Fields("des").Value) Then
					If ScreenInfo(X, Y).cDes = "" Then
						ScreenInfo(X, Y).cDes = "?"
					End If
				Else
					'112203 rjk: Override not known or unknown 112903 rjk: Added ShareBmap checks as well
					If ScreenInfo(X, Y).cDes = "" Or ScreenInfo(X, Y).cDes = "?" Or (ScreenInfo(X, Y).cDes >= "A" And ScreenInfo(X, Y).cDes <= "W") Or (ScreenInfo(X, Y).cDes >= "Y" And ScreenInfo(X, Y).cDes <= "Z") Or IsNumeric(ScreenInfo(X, Y).cDes) Or ScreenInfo(X, Y).cDes = vbNullChar Then
						ScreenInfo(X, Y).cDes = rsEnemySect.Fields("des").Value
					ElseIf ScreenInfo(X, Y).cDes <> rsEnemySect.Fields("des").Value And rsEnemySect.Fields("des").Value <> "?" Then 
						ScreenInfo(X, Y).cSecondDes = rsEnemySect.Fields("des").Value
					End If
				End If
				Select Case ColorScheme
					Case COLORSCHEME_NORMAL
						ScreenInfo(X, Y).bDisplayColor = rsEnemySect.Fields("owner").Value
					Case COLORSCHEME_GRADUATED '112203 rjk: Added the ability to show values from enemysectors in graduated mode
						If bIncludeEnemySectorsF4Display Then
							ScreenInfo(X, Y).bDisplayColor = rsEnemySect.Fields("owner").Value
							Select Case GradColorItem
								'x,y,des,owner,oldowner,eff,road,rail,defense,civ,mil,shell,gun,pet,food,iron,bar,timestamp"
								Case "eff", "road", "rail", "defense", "civ", "mil", "shell", "gun", "pet", "food", "iron", "bar"
									ScreenInfo(X, Y).iValue = rsEnemySect.Fields(GradColorItem).Value
								Case Else
									ScreenInfo(X, Y).iValue = -2
							End Select
						End If
					Case COLORSCHEME_CONDITIONAL '112503 rjk: Added
						If bIncludeEnemySectorsF4Display Then
							ScreenInfo(X, Y).bDisplayColor = rsEnemySect.Fields("owner").Value
							If ConditionMetEnemySector Then
								ScreenInfo(X, Y).iValue = ScreenInfo(X, Y).iValue + 1
							End If
						End If
				End Select
			End If
			rsEnemySect.MoveNext()
		End While
		'load sector info
		ReDim DistPoint(0)
		ReDim DistCount(0)
		If Not (rsSectors.BOF And rsSectors.EOF) Then
			rsSectors.MoveFirst()
		End If
		Dim sMobCost As Single
		While Not rsSectors.EOF
			X = SIx(rsSectors.Fields("x").Value)
			Y = SIy(rsSectors.Fields("y").Value)
			'for distribution based display, must build graduated array using all sectors,
			'   or colors will change based on map view
			If ColorScheme = COLORSCHEME_DISTRIBUTION Then
				x1 = rsSectors.Fields("dist_x").Value
				y1 = rsSectors.Fields("dist_y").Value
				DistColor = 0
				If Not (x1 = rsSectors.Fields("x").Value And y1 = rsSectors.Fields("y").Value) Then
					strKey = CStr(x1) & "," & CStr(y1)
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
				If bDeity And rsSectors.Fields("country").Value <> 0 And ColorScheme = COLORSCHEME_NORMAL Then
					ScreenInfo(X, Y).bOwner = False
				Else
					ScreenInfo(X, Y).bOwner = True
				End If
				If dspDesignation = enumDisplayDesignation.DD_NEW And rsSectors.Fields("sdes").Value <> " " Then '112203 Added the ability to show new designation
					ScreenInfo(X, Y).cDes = rsSectors.Fields("sdes").Value
				Else
					ScreenInfo(X, Y).cDes = rsSectors.Fields("des").Value
					If dspDesignation = enumDisplayDesignation.DD_BOTH And rsSectors.Fields("sdes").Value <> " " Then
						ScreenInfo(X, Y).cSecondDes = rsSectors.Fields("sdes").Value
					End If
				End If
				If bDeity And ColorScheme = COLORSCHEME_NORMAL Then
					ScreenInfo(X, Y).bDisplayColor = rsSectors.Fields("country").Value
				ElseIf bDeity And ColorScheme = COLORSCHEME_DISTRIBUTION And rsSectors.Fields("country").Value = 0 Then 
					ScreenInfo(X, Y).bDisplayColor = 0
				ElseIf rsSectors.Fields("*").Value = "*" And ColorScheme = COLORSCHEME_NORMAL Then  '100803 rjk: Added 'Not Yours' identify to Map display
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
								ScreenInfo(X, Y).iValue = CShort(PlagueRisk(rsSectors) * 10#)
							Case "mobility cost"
								sMobCost = SectorMobCost(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value, EmpCommon.enumMobType.MT_COMMODITY)
								If sMobCost > 9# Then
									ScreenInfo(X, Y).iValue = 999
								Else
									ScreenInfo(X, Y).iValue = CShort(100# * sMobCost)
								End If
							Case "food required" '110103 rjk: Added food required to the graduated colorscheme
								ScreenInfo(X, Y).iValue = FoodRequired(rsSectors)
							Case "food shortage" '110503 rjk: Added food shortage to the graduated colorscheme
								If FoodRequired(rsSectors) > rsSectors.Fields("food").Value Then
									ScreenInfo(X, Y).iValue = FoodRequired(rsSectors) - rsSectors.Fields("food").Value
								Else
									ScreenInfo(X, Y).iValue = 0
								End If
							Case "food excess" '112403 rjk: Added food shortage to the graduated colorscheme
								ScreenInfo(X, Y).iValue = rsSectors.Fields("food").Value - FoodRequired(rsSectors)
							Case "excess civilians" '050404 rjk: Added for Excess civilians
								'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								v = Production(rsSectors)
								If (v.prodValidFlag) Then
									ScreenInfo(X, Y).iValue = CShort(v.ExcessCiv)
								Else
									ScreenInfo(X, Y).iValue = 0
								End If
							Case "build cost"
								'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								v = Production(rsSectors)
								If (v.prodValidFlag) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object v.BuildAvailCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ScreenInfo(X, Y).iValue = CShort(v.BuildAvailCost)
								Else
									ScreenInfo(X, Y).iValue = 0
								End If
							Case "reacting planes" '100605 rjk: Added for reacting planes
								ScreenInfo(X, Y).iValue = UpdateScreenInfoWithReactingPlanes(X, Y)
							Case "production"
								'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								v = Production(rsSectors)
								If (v.prodValidFlag) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ScreenInfo(X, Y).iValue = CShort(v.ProdAmount)
								Else
									ScreenInfo(X, Y).iValue = 0
								End If
							Case Else
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								If IsDbNull(rsSectors.Fields(GradColorItem).Value) Then
									ScreenInfo(X, Y).iValue = -1
								Else
									ScreenInfo(X, Y).iValue = rsSectors.Fields(GradColorItem).Value
								End If
						End Select
						
						'102003 rjk: Modified to store delivery information as well
					Case COLORSCHEME_DELIVERIES
						Select Case rsSectors.Fields(strCommodity & "_del").Value
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
						ScreenInfo(X, Y).iValue = ScreenInfo(X, Y).iValue + (rsSectors.Fields(strCommodity & "_cut").Value * 8#)
						
					Case COLORSCHEME_DISTRIBUTION
						'112203 rjk: Added displaying the threshold
						ScreenInfo(X, Y).bDisplayColor = DistColor Mod (NUMBER_OF_PLAYER_COLORS + 1)
						ScreenInfo(X, Y).iValue = rsSectors.Fields(strCommodity & "_dist").Value
						
				End Select
			End If
			rsSectors.MoveNext()
		End While
		
		If ColorScheme = COLORSCHEME_GRADUATED And GradColorDspVal Then
			If (GradColorItem = "fert" Or GradColorItem = "ocontent") Then
				If Not (rsSea.BOF And rsSea.EOF) Then
					rsSea.MoveFirst()
				End If
				While Not rsSea.EOF
					X = SIx(rsSea.Fields("x").Value)
					Y = SIy(rsSea.Fields("y").Value)
					If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
						ScreenInfo(X, Y).iValue = rsSea.Fields(GradColorItem).Value
					End If
					rsSea.MoveNext()
				End While
			End If
		End If
		
		'Add event markers for distribution centers
		If ColorScheme = COLORSCHEME_DISTRIBUTION Then
			EventMarkers.Clear()
			For n = 1 To UBound(DistPoint)
				If ParseSectors(X, Y, DistPoint(n)) Then
					x1 = n
					While x1 > NUMBER_OF_PLAYER_COLORS
						x1 = x1 - NUMBER_OF_PLAYER_COLORS
					End While
					EventMarkers.Add(X, Y, x1, "DP for " & CStr(DistCount(n)) & " Hexs")
				End If
			Next n
		End If
	End Sub
	
	Function SIx(ByRef X As Short) As Short
		If X < OriginX Then
			SIx = X + MapSizeX
		Else
			SIx = X
		End If
	End Function
	
	Function SIy(ByRef Y As Short) As Short
		If Y < OriginY Then
			SIy = Y + MapSizeY
		Else
			SIy = Y
		End If
		
	End Function
	
	Public Sub UpdateScreenInfo(ByRef sx As Short, ByRef sy As Short)
		
		Dim ix As Short
		Dim iy As Short
		rsSectors.Seek("=", sx, sy)
		If Not rsSectors.NoMatch Then
			ix = SIx(sx)
			iy = SIy(sy)
			ScreenInfo(ix, iy).cDes = rsSectors.Fields("des").Value
		End If
	End Sub
	
	Public Function GetConditionalSectors() As String
		Dim strSector As String
		
		If ColorScheme = COLORSCHEME_CONDITIONAL Then
			rsSectors.MoveFirst()
			Do While rsSectors.EOF = False
				If ConditionMet Then
					If Len(strSector) = 0 Then
						strSector = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					Else
						strSector = strSector & "\" & SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
					End If
				End If
				rsSectors.MoveNext()
			Loop 
		Else
			strSector = ""
		End If
		GetConditionalSectors = strSector
	End Function
	
	'102303 rjk: Added to Count the number Multiple Sectors in the Sector field
	Public Function CountMultipleSectors(ByRef strSector As String) As Short
		Dim nCount As Short
		Dim pos As Short
		
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
	
	Public Sub GetUnitCount(ByRef ScreenInfo() As tScreenInfo, ByRef topx As Short, ByRef topy As Short)
		Dim X As Short
		Dim Y As Short
		Dim dExpiry As Date
		Dim dRecord As Date
		
		dExpiry = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -frmUnitView.unitView.iExpiry, Now)
		
		'load ships onto screen
		If Not (rsShip.BOF And rsShip.EOF) Then
			rsShip.MoveFirst()
		End If
		While Not rsShip.EOF
			'092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding
			'(ShowOnlyWarships And (rsShip.Fields("fir") > 0 Or rsShip.Fields("type") = "ls")) Then
			If Not frmOptions.bShowOnlyWarships Or (frmOptions.bShowOnlyWarships And (rsShip.Fields("fir").Value > 0 Or frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_S_LAND, rsShip.Fields("type").Value))) Then
				X = SIx(rsShip.Fields("x").Value)
				Y = SIy(rsShip.Fields("y").Value)
				If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
					ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_SHIPS) + 1
				End If
			End If
			rsShip.MoveNext()
		End While
		'load lands onto screen
		If Not (rsLand.BOF And rsLand.EOF) Then
			rsLand.MoveFirst()
		End If
		While Not rsLand.EOF
			X = SIx(rsLand.Fields("x").Value)
			Y = SIy(rsLand.Fields("y").Value)
			If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
				ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_LAND_UNITS) + 1
			End If
			rsLand.MoveNext()
		End While
		'load planes onto screen
		If Not (rsPlane.BOF And rsPlane.EOF) Then
			rsPlane.MoveFirst()
		End If
		While Not rsPlane.EOF
			X = SIx(rsPlane.Fields("x").Value)
			Y = SIy(rsPlane.Fields("y").Value)
			If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
				ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_PLANES) + 1
			End If
			rsPlane.MoveNext()
		End While
		'load nukes onto screen
		If Not (rsNuke.BOF And rsNuke.EOF) Then
			rsNuke.MoveFirst()
		End If
		While Not rsNuke.EOF
			X = SIx(rsNuke.Fields("x").Value)
			Y = SIy(rsNuke.Fields("y").Value)
			If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
				ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_NUKES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_OUR_NUKES) + 1
			End If
			rsNuke.MoveNext()
		End While
		'load enemy units onto screen
		If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
			rsEnemyUnit.MoveFirst()
		End If
		
		While Not rsEnemyUnit.EOF
			X = SIx(rsEnemyUnit.Fields("x").Value)
			Y = SIy(rsEnemyUnit.Fields("y").Value)
			If X >= OriginX And X <= topx And Y >= OriginY And Y <= topy Then
				dRecord = ConvertToLocalTimeFromWinACEUTC((rsEnemyUnit.Fields("timestamp").Value))
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If frmUnitView.unitView.iExpiry = 0 Or DateDiff(Microsoft.VisualBasic.DateInterval.Day, dExpiry, dRecord) > 0 Then
					Select Case rsEnemyUnit.Fields("class").Value
						Case "l"
							Select Case Nations.YourRelation(rsEnemyUnit.Fields("owner").Value)
								Case iREL_AT_WAR, iREL_HOSTILE, iREL_MOBILIZING, iREL_SITZKRIEG
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ENEMY_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ENEMY_LAND_UNITS) + 1
								Case iREL_ALLIED
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_LAND_UNITS) + 1
								Case iREL_FRIENDLY
									If frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_NEUTRAL Then
										ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_LAND_UNITS) + 1
									Else
										ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_LAND_UNITS) + 1
									End If
								Case Else 'iREL_NEUTRAL or unknown or offline 120103 rjk: Added offline
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_LAND_UNITS) + 1
							End Select
						Case "s"
							Select Case Nations.YourRelation(rsEnemyUnit.Fields("owner").Value)
								Case iREL_AT_WAR, iREL_HOSTILE, iREL_MOBILIZING, iREL_SITZKRIEG
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ENEMY_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ENEMY_SHIPS) + 1
								Case iREL_ALLIED
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_SHIPS) + 1
								Case iREL_FRIENDLY
									If frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_NEUTRAL Then
										ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_SHIPS) + 1
									Else
										ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_SHIPS) + 1
									End If
								Case Else 'iREL_NEUTRAL or unknown or offline 120103 rjk: Added offline
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_SHIPS) + 1
							End Select
						Case "p"
							Select Case Nations.YourRelation(rsEnemyUnit.Fields("owner").Value)
								Case iREL_AT_WAR, iREL_HOSTILE, iREL_MOBILIZING, iREL_SITZKRIEG
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ENEMY_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ENEMY_PLANES) + 1
								Case iREL_ALLIED
									ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_PLANES) + 1
								Case iREL_FRIENDLY
									If frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_NEUTRAL Then
										ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_PLANES) + 1
									Else
										ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_ALLIED_PLANES) + 1
									End If
							End Select
						Case Else 'iREL_NEUTRAL or unknown or offline 120103 rjk: Added offline
							ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_NEUTRAL_PLANES) + 1
					End Select
				Else
					Select Case rsEnemyUnit.Fields("class").Value
						Case "l"
							ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_EXPIRED_LAND_UNITS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_EXPIRED_LAND_UNITS) + 1
						Case "s"
							ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_EXPIRED_SHIPS) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_EXPIRED_SHIPS) + 1
						Case "p"
							ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_EXPIRED_PLANES) = ScreenInfo(X, Y).iUnitCounts(EmpUnitView.enumViewingUnits.VU_EXPIRED_PLANES) + 1
					End Select
				End If
			End If
			rsEnemyUnit.MoveNext()
		End While
	End Sub
	
	Public Function UpdateScreenInfoWithReactingPlanes(ByRef X As Short, ByRef Y As Short) As Short
		Dim hldIndex As String
		Dim iOpSecX As Short
		Dim iOpSecY As Short
		Dim iCount As Short
		
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
		rsMissions.MoveFirst()
		rsMissions.Seek("=", "p", "air defense")
		If rsMissions.NoMatch Then
			UpdateScreenInfoWithReactingPlanes = 0
			rsMissions.Index = hldIndex
			Exit Function
		End If
		Do Until rsMissions.EOF
			If rsMissions.Fields("type").Value <> "p" Or rsMissions.Fields("mission").Value <> "air defense" Then
				Exit Do
			End If
			ParseSectors(iOpSecX, iOpSecY, rsMissions.Fields("op sector").Value)
			If SectorDistance(X, Y, iOpSecX, iOpSecY) <= rsMissions.Fields("radius").Value Then
				iCount = iCount + 1
			End If
			rsMissions.MoveNext()
		Loop 
		rsMissions.Index = hldIndex
		UpdateScreenInfoWithReactingPlanes = iCount
	End Function
End Module