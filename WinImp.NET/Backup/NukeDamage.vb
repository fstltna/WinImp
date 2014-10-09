Option Strict Off
Option Explicit On
Module NukeDamage
	
	'Change Log:
	'200604 rjk: Created
	
	Public NukeDamageType As String
	Public NukeTargetSector As String
	Public bNukeAirBurst As Boolean
	
	Public Sub ShowNukeDamage(ByRef picMap As System.Windows.Forms.PictureBox)
		Dim MaxRadius As Short
		Dim StartX As Short
		Dim StartY As Short
		Dim iDistance As Short
		Dim iDirection As Short
		Dim iSpoke As Short
		Dim sx As Short
		Dim sy As Short
		Dim X As Short
		Dim Y As Short
		Dim PosX As Single
		Dim PosY As Single
		Dim xas, xa, ya, yas As Object
		Dim strID As String
		Dim bFission As Boolean
		Dim iPos As Short
		Dim hldIndex As String
		Dim iBlast As Short
		Dim iDamage As Short
		Dim HldColor As Integer
		Dim HldFontSize As Object
		Dim strDamage As String
		
		iPos = InStr(NukeDamageType, " - ")
		
		If iPos = 0 Then
			MsgBox("Invalid Nuke Type")
			Exit Sub
		End If
		
		strID = Left(NukeDamageType, iPos - 1)
		If InStr(NukeDamageType, "fission") > 0 Then
			bFission = True
		Else
			bFission = False
		End If
		
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		rsBuild.Seek("=", "n", strID)
		
		If rsBuild.NoMatch Then
			MsgBox("Nuke Type information not found")
			Exit Sub
		End If
		
		iBlast = rsBuild.Fields("stat1").Value
		iDamage = rsBuild.Fields("stat2").Value
		rsBuild.Index = hldIndex
		
		If bNukeAirBurst Then
			MaxRadius = iBlast
		Else
			MaxRadius = Fix(iBlast * 2 / 3)
		End If
		
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
		
		If Not ParseSectors(StartX, StartY, NukeTargetSector) Then
			Exit Sub
		End If
		
		HldColor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		picMap.ForeColor = System.Drawing.Color.White
		'UPGRADE_WARNING: Couldn't resolve default property of object HldFontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HldFontSize = picMap.Font.SizeInPoints
		picMap.Font = VB6.FontChangeSize(picMap.Font, picMap.Font.SizeInPoints * 0.7)
		
		Coord(StartX, StartY, PosX, PosY)
		strDamage = CStr(NukeSectorDamage(iBlast, iDamage, 0, bNukeAirBurst))
		'UPGRADE_ISSUE: PictureBox method picMap.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sx = PosX - (picMap.TextHeight(strDamage) / 2)
		'UPGRADE_ISSUE: PictureBox method picMap.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sy = PosY - (picMap.TextWidth(strDamage) / 2)
		If sx >= 0 And sy >= 0 Then
			'UPGRADE_ISSUE: PictureBox property picMap.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			picMap.CurrentX = sx
			'UPGRADE_ISSUE: PictureBox property picMap.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			picMap.CurrentY = sy
			'UPGRADE_ISSUE: PictureBox method picMap.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			picMap.Print(strDamage)
		End If
		
		'Now put out the nav markers
		For iDistance = 1 To MaxRadius
			For iDirection = 0 To 5
				For iSpoke = 0 To MaxRadius - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object xas(iDirection). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object xa(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					X = (iDistance * xa(iDirection)) + (iSpoke * xas(iDirection))
					'UPGRADE_WARNING: Couldn't resolve default property of object yas(iDirection). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ya(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Y = (iDistance * ya(iDirection)) + (iSpoke * yas(iDirection))
					Coord(StartX + X, StartY + Y, PosX, PosY)
					strDamage = CStr(NukeSectorDamage(iBlast, iDamage, SectorDistance(0, 0, X, Y), bNukeAirBurst))
					'UPGRADE_ISSUE: PictureBox method picMap.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					sx = PosX - (picMap.TextHeight(strDamage) / 2)
					'UPGRADE_ISSUE: PictureBox method picMap.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					sy = PosY - (picMap.TextWidth(strDamage) / 2)
					If sx >= 0 And sy >= 0 Then
						'UPGRADE_ISSUE: PictureBox property picMap.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						picMap.CurrentX = sx
						'UPGRADE_ISSUE: PictureBox property picMap.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						picMap.CurrentY = sy
						'UPGRADE_ISSUE: PictureBox method picMap.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						picMap.Print(strDamage)
					End If
				Next iSpoke
			Next iDirection
		Next iDistance
		
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(HldColor)
		'UPGRADE_WARNING: Couldn't resolve default property of object HldFontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.Font = VB6.FontChangeSize(picMap.Font, HldFontSize)
	End Sub
	
	'int
	'nukedamage(struct nchrstr *ncp, int range, int airburst)
	'{
	'    int dam;
	'    int rad;
	'
	'    rad = ncp->n_blast;
	'    if (airburst)
	'    rad = (int)(rad * 1.5);
	'    if (rad < range)
	'    return 0;
	'    if (airburst) {
	'    /* larger area, less center damage */
	'    dam = (int)((ncp->n_dam * 0.75) - (range * 20));
	'    } else {
	'    /* smaller area, more center damage */
	'    dam = (int)(ncp->n_dam / (range + 1.0));
	'    }
	'    if (dam < 5)
	'    dam = 0;
	'    return dam;
	'}
	
	Private Function NukeSectorDamage(ByRef iBlast As Short, ByRef iDamage As Short, ByRef iDistance As Short, ByRef bAirBurst As Boolean) As Short
		Dim iRad As Short
		
		iRad = iBlast
		If bAirBurst Then
			iRad = Fix(iRad * 1.5)
		End If
		If iRad < iDistance Then
			NukeSectorDamage = 0
			Exit Function
		End If
		If bAirBurst Then
			NukeSectorDamage = Fix((iDamage * 0.75) - (iDistance * 20))
		Else
			NukeSectorDamage = Fix(iDamage / (iDistance + 1#))
		End If
		If NukeSectorDamage < 5 Then
			NukeSectorDamage = 0
		End If
	End Function
End Module