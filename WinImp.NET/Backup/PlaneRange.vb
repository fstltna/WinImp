Option Strict Off
Option Explicit On
Module PlaneRange
	
	'Change Log:
	'010106 rjk: Created
	
	Public PlaneRangeType As String
	Public PlaneTargetSector As String
	Public PlaneRoundTrip As Boolean
	Public PlaneTech As Short
	
	Public Sub ShowPlaneRange(ByRef picMap As System.Windows.Forms.PictureBox)
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
		'Dim iPos As Integer
		Dim hldIndex As String
		Dim iRange As Short
		Dim HldColor As Integer
		Dim HldFontSize As Object
		Dim strDamage As String
		
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		rsBuild.Seek("=", "p", Trim(Right(PlaneRangeType, 5)))
		
		If rsBuild.NoMatch Then
			MsgBox("Plane Type information not found")
			Exit Sub
		End If
		MaxRadius = ComputePlaneRange(rsBuild.Fields("stat5").Value, rsBuild.Fields("tech").Value, PlaneTech)
		If PlaneRoundTrip Then
			MaxRadius = MaxRadius / 2
		End If
		rsBuild.Index = hldIndex
		
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
		
		If Not ParseSectors(StartX, StartY, PlaneTargetSector) Then
			Exit Sub
		End If
		
		HldColor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		picMap.ForeColor = System.Drawing.Color.White
		'UPGRADE_WARNING: Couldn't resolve default property of object HldFontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HldFontSize = picMap.Font.SizeInPoints
		picMap.Font = VB6.FontChangeSize(picMap.Font, picMap.Font.SizeInPoints * 0.7)
		
		Coord(StartX, StartY, PosX, PosY)
		strDamage = ""
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
				For iSpoke = 0 To iDistance - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object xas(iDirection). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object xa(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					X = (iDistance * xa(iDirection)) + (iSpoke * xas(iDirection))
					'UPGRADE_WARNING: Couldn't resolve default property of object yas(iDirection). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ya(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Y = (iDistance * ya(iDirection)) + (iSpoke * yas(iDirection))
					Coord(StartX + X, StartY + Y, PosX, PosY)
					If PlaneRoundTrip Then
						strDamage = CStr(iDistance * 2)
					Else
						strDamage = CStr(iDistance)
					End If
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
End Module