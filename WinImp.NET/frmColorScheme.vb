Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmColorScheme
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'100803 rjk: 'Not Your' colors, to be used in COLORSCHEME_NORMAL
	'102003 rjk: Added Delivery ColorScheme
	'110103 rjk: Added food required to the graduated scheme
	'110503 rjk: Added food shortage to the graduated scheme
	'112203 rjk: Added displaying the threshold for F4 Distribution
	'112203 rjk: Added the ability to show the new designation, current designation or both
	'            Added the ability to select graduated display from rsEnemySector table as well.
	'112403 rjk: Added the ability to show food shortages and food required (food excess)
	'112503 rjk: Added the ability to select conditional display from rsEnemySector table as well.
	'            Added a control for determining the default value when a field is missing or unknown
	'112903 rjk: Switched ViewUnit to unitView for controlling Unit View options
	'120303 rjk: Added Food required and Food Excess for F4 Conditional
	'121203 rjk: Switched to Items and Item objects.
	'121803 rjk: Added Plague Risk scaled by 10
	'122903 rjk: Change Show Value in Hex to Show Value to reflect functionality.
	'122903 rjk: Switched to database value from item object.
	'122903 rjk: Updated titles for Plague Risk.
	'123003 rjk: Fixed not saving special graduated values.
	'270104 rjk: Fill in the graduated display with a default item (prevents errors)
	'300304 rjk: Fixed to deal with strength fields being Null
	'050404 rjk: Added excess civilians to the graduated and conditional color schemes
	'180404 rjk: Set default and cancel buttons.  Change "Apply" to "OK".
	'270904 rjk: Added saving of the display mode (DD_CURRENT, DD_NEW and DD_BOTH).
	'160705 rjk: Added Reacting Planes.  Fixed Deity Problem with Mobility Cost.
	'090905 rjk: Made HideUnits a saved parameters across F4 calls and sessions.
	'030108 rjk: Increase size of the condition field list.
	'170508 rjk: Added FS (Food shortage) to the F4 Conditional.
	'220508 rjk: Added Build Cost to F4 graduated and conditional.
	
	'UPGRADE_WARNING: Lower bound of array HldC was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim HldC(NUMBER_OF_BASIC_COLORS) As Integer
	'UPGRADE_WARNING: Lower bound of array HldT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim HldT(NUMBER_OF_BASIC_COLORS) As Integer
	
	Dim oldColorScheme As Short
	Dim oldGradColorItem As String
	Dim oldGradColorHigh As Integer
	Dim oldGradColorLow As Integer
	Dim oldGradColorDspVal As Boolean
	Dim oldCondColorParms() As Object
	Dim oldCondColorParmNumber As Short
	Dim olddspDesignation As HexGrid.enumDisplayDesignation
	Dim olddspDesignationSave As HexGrid.enumDisplayDesignation
	'Dim oldViewUnits As Boolean
	
	Const NUMBER_OF_CONDITIONAL_ITEMS As Short = 20
	
	'112503 rjk: Added the ability to select conditional display from rsEnemySector table
	'UPGRADE_WARNING: Event chkMissing.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkMissing_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMissing.CheckStateChanged
		If chkMissing.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bConditionalMissing = True
		Else
			bConditionalMissing = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event chkEnemySectors.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkEnemySectors_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkEnemySectors.CheckStateChanged
		Dim Index As Short = chkEnemySectors.GetIndex(eventSender)
		If chkEnemySectors(Index).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bIncludeEnemySectorsF4Display = True
		Else
			bIncludeEnemySectorsF4Display = False
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Dim n As Short
		Dim m As Short
		
		'Restore Colors
		For n = 1 To NUMBER_OF_BASIC_COLORS
			BasicColors(n) = HldC(n)
			BasicText(n) = HldT(n)
		Next n
		
		'restore current values
		ColorScheme = oldColorScheme
		GradColorItem = oldGradColorItem
		GradColorHigh = oldGradColorHigh
		GradColorLow = oldGradColorLow
		GradColorDspVal = oldGradColorDspVal
		CondColorParmNumber = oldCondColorParmNumber
		HexGrid.dspDesignation = olddspDesignation
		frmOptions.dspDesignationSave = olddspDesignationSave
		
		'UPGRADE_WARNING: Lower bound of array CondColorParms was changed from 0,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim CondColorParms(CondColorParmNumber, 4)
		
		For n = 1 To CondColorParmNumber
			For m = 1 To 4
				'UPGRADE_WARNING: Couldn't resolve default property of object oldCondColorParms(n, m). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(n, m). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CondColorParms(n, m) = oldCondColorParms(n, m)
			Next m
		Next n
		
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		ApplyColors()
		frmOptions.SaveOptions()
		Me.Close()
	End Sub
	
	Private Sub ApplyColors()
		Dim n As Short
		Dim found As Boolean
		'112903 rjk: Removed ViewUnits and add unitView for controlling options.
		frmUnitView.unitView.Load()
		
		Dim eiItem As EmpItem
		If optScheme(COLORSCHEME_NORMAL).Checked Then
			'normal scheme
			ColorScheme = COLORSCHEME_NORMAL
		ElseIf optScheme(COLORSCHEME_CONDITIONAL).Checked Then 
			'conditional scheme
			If Not ParseParmString(cbCond.Text) Then
				Exit Sub
			End If
			ColorScheme = COLORSCHEME_CONDITIONAL
			'put conditional string in the combo box
			found = False
			n = 0
			While Not found And n < cbCond.Items.Count
				n = n + 1
				found = (VB6.GetItemString(cbCond, n - 1) = cbCond.Text)
			End While
			
			If Not found Then
				cbCond.Items.Insert(0, cbCond.Text)
			End If
		ElseIf optScheme(COLORSCHEME_GRADUATED).Checked Then 
			'graduated scheme
			If Len(txtHigh.Text) = 0 Or Len(txtLow.Text) = 0 Then
				MsgBox("Error:  Must have a high and low value.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Graduated Boundry Error")
				Exit Sub
			End If
			If Not (IsNumeric(txtHigh.Text) And IsNumeric(txtLow.Text)) Then
				MsgBox("Error:  High and low value must be numeric.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Graduated Boundry Error")
				Exit Sub
			End If
			If CInt(txtHigh.Text) < CInt(txtLow.Text) Then
				MsgBox("Error:  High value must not be less than low value.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Graduated Boundry Error")
				Exit Sub
			End If
			'passed error checks so set values
			ColorScheme = COLORSCHEME_GRADUATED
			
			
			eiItem = Items.FindByConditionalName((Combo1.Text))
			If eiItem Is Nothing Then
				GradColorItem = Combo1.Text
			Else
				GradColorItem = eiItem.DatabaseName
			End If
			GradColorHigh = CInt(txtHigh.Text)
			GradColorLow = CInt(txtLow.Text)
			GradColorDspVal = chkShowVal.CheckState
			
			If chkHide.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				frmUnitView.unitView.ClearAllUnits()
				SaveSetting(APPNAME, "Color Scheme", "Hide Units", CStr(True))
			Else
				SaveSetting(APPNAME, "Color Scheme", "Hide Units", CStr(False))
			End If
		ElseIf optScheme(COLORSCHEME_DISTRIBUTION).Checked Then  '112203 rjk: Added thresholds for Distribution
			'disribution scheme
			ColorScheme = COLORSCHEME_DISTRIBUTION
			frmUnitView.unitView.ClearAllUnits()
			strCommodity = VB.Left(cbDistCom.Text, 1)
		ElseIf optScheme(COLORSCHEME_DELIVERIES).Checked Then  '102003 rjk: Added ColorScheme for Deliveries
			'disribution scheme
			ColorScheme = COLORSCHEME_DELIVERIES
			strCommodity = VB.Left(cbCom.Text, 1)
			frmUnitView.unitView.ClearAllUnits()
		End If
		
		Select Case ColorScheme
			Case COLORSCHEME_NORMAL, COLORSCHEME_CONDITIONAL
				If optDesignation(1).Checked Then
					HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_CURRENT
					frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_CURRENT
				ElseIf optDesignation(2).Checked Then 
					HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_NEW
					frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_NEW
				Else
					HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_BOTH
					frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH
				End If
			Case COLORSCHEME_GRADUATED, COLORSCHEME_DISTRIBUTION, COLORSCHEME_DELIVERIES
				If optDesignation(1).Checked Then
					HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_CURRENT
					If frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH Or frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT Or frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_NEW Then
						frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT
					Else
						frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_CURRENT
					End If
				Else
					HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_NEW
					If frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH Or frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT Or frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_NEW Then
						frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_NEW
					Else
						frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_NEW
					End If
				End If
		End Select
		
		For n = 1 To NUMBER_OF_BASIC_COLORS
			BasicColors(n) = System.Drawing.ColorTranslator.ToOle(Picture1(n - 1).BackColor)
			BasicText(n) = System.Drawing.ColorTranslator.ToOle(Picture1(n - 1).ForeColor)
		Next n
		
		frmDrawMap.DrawHexes()
	End Sub
	
	Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click
		Dim Index As Short = cmdReset.GetIndex(eventSender)
		Select Case Index
			Case COLORSCHEME_NORMAL
				SetDefaultBasicColors()
			Case COLORSCHEME_CONDITIONAL
				SetDefaultComparisionColors()
			Case COLORSCHEME_GRADUATED
				SetDefaultGradColors()
			Case COLORSCHEME_DELIVERIES '102003 rjk: Added Deliveries ColorScheme
				SetDefaultDeliveryColors()
		End Select
		FillColors()
	End Sub
	
	Private Sub cmdTest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTest.Click
		ApplyColors()
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
		Dim v As productionDataType
		Dim low As Integer
		Dim high As Integer
		Dim n As Integer
		
		If rsSectors.BOF And rsSectors.EOF Then Exit Sub
		
		high = 0
		low = 99999
		rsSectors.MoveFirst()
		Dim sMobCost As Single
		Dim eiItem As EmpItem
		While Not rsSectors.EOF
			Select Case Combo1.Text
				Case "plague risk percentage chance"
					If PlagueRisk(rsSectors) > 1# Then
						n = CInt(Int(PlagueRisk(rsSectors))) '121803 rjk: Added Int to truncate
					Else
						n = 0
					End If
				Case "plague risk calculation X 10" '121803 rjk: added
					n = CInt(PlagueRisk(rsSectors) * 10#) '121803 rjk: Added Int to truncate
				Case "mobility cost"
					sMobCost = SectorMobCost(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value, EmpCommon.enumMobType.MT_COMMODITY)
					If sMobCost > 9# Then
						n = 999
					Else
						n = CInt(100# * sMobCost)
					End If
				Case "food required" '110103 rjk: Added food required to the graduated colorscheme
					n = FoodRequired(rsSectors)
				Case "food shortage" '110503 rjk: Added food shortage to the graduated colorscheme
					If FoodRequired(rsSectors) > rsSectors.Fields("food").Value Then
						n = FoodRequired(rsSectors) - rsSectors.Fields("food").Value
					Else
						n = 0
					End If
				Case "food excess" '112403 rjk: Added food excess to the graduated colorscheme
					n = rsSectors.Fields("food").Value - FoodRequired(rsSectors)
				Case "excess civilians" '050404 rjk: Added for Excess civilians
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If (v.prodValidFlag) Then
						n = CInt(v.ExcessCiv)
					Else
						n = 0
					End If
				Case "build cost"
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If (v.prodValidFlag) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object v.BuildAvailCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						n = CInt(v.BuildAvailCost)
					Else
						n = 0
					End If
				Case "production"
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If (v.prodValidFlag) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						n = CInt(v.ProdAmount)
					Else
						n = 0
					End If
				Case "reacting planes" '100605 rjk: Added for reacting planes
					n = UpdateScreenInfoWithReactingPlanes(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
				Case Else
					
					eiItem = Items.FindByConditionalName((Combo1.Text))
					If eiItem Is Nothing Then
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDbNull(rsSectors.Fields(Combo1.Text).Value) Then
							n = -1
						Else
							n = rsSectors.Fields(Combo1.Text).Value
						End If
					Else
						n = eiItem.DatabaseValue(rsSectors)
					End If
			End Select
			
			If n < low Then
				low = n
			End If
			If n > high Then
				high = n
			End If
			rsSectors.MoveNext()
		End While
		txtLow.Text = CStr(low)
		txtHigh.Text = CStr(high)
	End Sub
	
	Private Sub frmColorScheme_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim n As Short
		Dim m As Short
		Dim f As DAO.Field
		Dim strTemp As String
		
		Dim eiItem As EmpItem
		For	Each f In rsSectors.Fields
			eiItem = Items.FindByDatabaseName((f.Name))
			If Not (eiItem Is Nothing) Then
				Label1(0).Text = Label1(0).Text & eiItem.ConditionalName & ", "
				Combo1.Items.Add(eiItem.ConditionalName)
			Else
				Label1(0).Text = Label1(0).Text & f.Name & ", "
				If f.Type = DAO.DataTypeEnum.dbInteger Or f.Type = DAO.DataTypeEnum.dbLong Then
					If f.Name <> "x" And f.Name <> "y" Then
						Combo1.Items.Add(f.Name)
					End If
				End If
			End If
			If GradColorItem = f.Name Then
				'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				Combo1.SelectedIndex = Combo1.NewIndex
			End If
		Next f
		
		'120303 rjk: Added Food required and Food Excess for F4 Conditional
		'050404 rjk: Added EC for excess civilians
		'110605 rjk: Added RP for reacting planes
		Label1(0).Text = Label1(0).Text & "FR, FS, FE, EC, RP, PD, BC"
		Combo1.Items.Add("plague risk percentage chance")
		If GradColorItem = "plague risk percentage chance" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("plague risk calculation X 10") '121803 rjk: added
		If GradColorItem = "plague risk calculation X 10" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("mobility cost")
		If GradColorItem = "mobility cost" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("food required") '110103 rjk: Added food required to the graduated colorscheme
		If GradColorItem = "food required" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("food shortage") '110503 rjk: Added food shortage to the graduated colorscheme
		If GradColorItem = "food shortage" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("food excess") '112403 rjk: Added food excess to the graduated colorscheme
		If GradColorItem = "food excess" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("excess civilians") '050404 rjk: Added excess civilians to the graduated colorscheme
		If GradColorItem = "excess civilians" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("build cost")
		If GradColorItem = "build cost" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("reacting planes") '100605 rjk: Added reacting planes to the graduated colorscheme
		If GradColorItem = "reacting planes" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		Combo1.Items.Add("production")
		If GradColorItem = "production" Then
			'UPGRADE_ISSUE: ComboBox property Combo1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			Combo1.SelectedIndex = Combo1.NewIndex
		End If
		
		If Combo1.SelectedIndex = -1 Then
			Combo1.SelectedIndex = 0
			GradColorItem = Combo1.Text
		End If
		
		FrameNormal.SetBounds(frameCond.Left, frameCond.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		frameGrad.SetBounds(frameCond.Left, frameCond.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		'save current values
		oldColorScheme = ColorScheme
		oldGradColorItem = GradColorItem
		oldGradColorHigh = GradColorHigh
		oldGradColorLow = GradColorLow
		oldGradColorDspVal = GradColorDspVal
		oldCondColorParmNumber = CondColorParmNumber
		'UPGRADE_WARNING: Lower bound of array oldCondColorParms was changed from 0,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim oldCondColorParms(oldCondColorParmNumber, 4)
		olddspDesignation = HexGrid.dspDesignation
		olddspDesignationSave = frmOptions.dspDesignationSave
		
		'Retrieve Hide Unit option
		If CBool(GetSetting(APPNAME, "Color Scheme", "Hide Units", CStr(True))) Then
			chkHide.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkHide.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		'Save Colors
		For n = 1 To NUMBER_OF_BASIC_COLORS
			HldC(n) = BasicColors(n)
			HldT(n) = BasicText(n)
		Next n
		
		FillColors()
		
		'Center form on screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) \ 2)
		
		'set controls
		optScheme(ColorScheme).Checked = True
		Select Case HexGrid.dspDesignation
			Case HexGrid.enumDisplayDesignation.DD_BOTH
				optDesignation(3).Checked = True
			Case HexGrid.enumDisplayDesignation.DD_CURRENT
				optDesignation(1).Checked = True
			Case HexGrid.enumDisplayDesignation.DD_NEW
				optDesignation(2).Checked = True
		End Select
		txtHigh.Text = CStr(GradColorHigh)
		txtLow.Text = CStr(GradColorLow)
		
		If GradColorDspVal Then
			chkShowVal.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkShowVal.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		For n = 1 To CondColorParmNumber
			For m = 1 To 4
				'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				cbCond.Text = cbCond.Text & CStr(CondColorParms(n, m))
				'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(n, m). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oldCondColorParms(n, m). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oldCondColorParms(n, m) = CondColorParms(n, m)
			Next m
		Next n
		
		'load saved conditional items from the registry
		For n = 1 To NUMBER_OF_CONDITIONAL_ITEMS
			strTemp = GetSetting(APPNAME, "Colors", "Conditions" & CStr(n), vbNullString)
			If Len(strTemp) > 0 Then
				cbCond.Items.Add(strTemp)
			End If
		Next n
		
		chkEnemySectors(0).CheckState = System.Windows.Forms.CheckState.Unchecked
		'112503 rjk: Added the ability to select conditional display from rsEnemySector table
		bIncludeEnemySectorsF4Display = False
		bConditionalMissing = True
		
		'102003 rjk: Added for deliver routes
		LoadCommodityBox(cbCom)
		cbCom.SelectedIndex = 0
		LoadCommodityBox(cbDistCom) '112203 rjk: Added displaying the threshold for F4 Distribution
		cbDistCom.SelectedIndex = 0
	End Sub
	
	Private Sub frmColorScheme_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		'save conditonal items into the registry
		Dim n As Short
		For n = 1 To NUMBER_OF_CONDITIONAL_ITEMS
			If n <= cbCond.Items.Count Then
				SaveSetting(APPNAME, "Colors", "Conditions" & CStr(n), VB6.GetItemString(cbCond, n - 1))
			End If
		Next n
		
	End Sub
	
	'UPGRADE_WARNING: Event optScheme.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optScheme_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optScheme.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optScheme.GetIndex(eventSender)
			
			FrameNormal.Visible = False
			frameGrad.Visible = False
			frameCond.Visible = False
			frameDist.Visible = False
			frameDel.Visible = False
			If optScheme(COLORSCHEME_NORMAL).Checked Then
				FrameNormal.Visible = True
			ElseIf optScheme(COLORSCHEME_CONDITIONAL).Checked Then 
				frameCond.Visible = True
				'112503 rjk: Added the ability to select conditional display from rsEnemySector table
				If bIncludeEnemySectorsF4Display Then
					chkEnemySectors(1).CheckState = System.Windows.Forms.CheckState.Checked
				Else
					chkEnemySectors(1).CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
			ElseIf optScheme(COLORSCHEME_GRADUATED).Checked Then 
				frameGrad.Visible = True
				If bIncludeEnemySectorsF4Display Then
					chkEnemySectors(0).CheckState = System.Windows.Forms.CheckState.Checked
				Else
					chkEnemySectors(0).CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
			ElseIf optScheme(COLORSCHEME_DISTRIBUTION).Checked Then  '112203 rjk: Added frame for selection of distribution parameters
				frameDist.Visible = True
			ElseIf optScheme(COLORSCHEME_DELIVERIES).Checked Then  '102003 rjk: Added frame for deliver route parameters
				frameDel.Visible = True
			End If
			If optScheme(COLORSCHEME_NORMAL).Checked Or optScheme(COLORSCHEME_CONDITIONAL).Checked Then
				optDesignation(3).Enabled = True
				Select Case frmOptions.dspDesignationSave
					Case HexGrid.enumDisplayDesignation.DD_BOTH, HexGrid.enumDisplayDesignation.DD_BOTH_NEW, HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT
						optDesignation(3).Checked = True
					Case HexGrid.enumDisplayDesignation.DD_NEW
						optDesignation(2).Checked = True
					Case HexGrid.enumDisplayDesignation.DD_CURRENT
						optDesignation(1).Checked = True
				End Select
			Else
				optDesignation(3).Enabled = False
				Select Case frmOptions.dspDesignationSave
					Case HexGrid.enumDisplayDesignation.DD_NEW, HexGrid.enumDisplayDesignation.DD_BOTH_NEW, HexGrid.enumDisplayDesignation.DD_BOTH
						optDesignation(2).Checked = True
					Case HexGrid.enumDisplayDesignation.DD_CURRENT, HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT
						optDesignation(1).Checked = True
					Case HexGrid.enumDisplayDesignation.DD_BOTH
				End Select
			End If
		End If
	End Sub
	
	'Parse the conditional string
	Private Function ParseParmString(ByVal strLine As String) As Boolean
		Dim j As Short
		Dim n As Short
		Dim strSep As String
		Dim strCurToken As String
		Dim strTokens() As String
		ReDim strTokens(0)
		
		'strip spaces and leading question mark
		strLine = Trim(strLine)
		If VB.Left(strLine, 1) = "?" Then
			strLine = Mid(strLine, 2)
		End If
		
		If Len(Trim(strLine)) = 0 Then
			MsgBox("You must enter a valid conditional string for this scheme", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
			ParseParmString = False
			Exit Function
		End If
		
		'Break the string into an array of tokens
		strSep = "&|=#><"
		j = 0
		For n = 1 To Len(strLine)
			If InStr(strSep, Mid(strLine, n, 1)) > 0 Then
				If Len(Trim(strCurToken)) > 0 Then
					ReDim Preserve strTokens(j)
					strTokens(j) = Trim(strCurToken)
					j = j + 1
					strCurToken = vbNullString
				End If
				ReDim Preserve strTokens(j)
				strTokens(j) = Mid(strLine, n, 1)
				j = j + 1
			Else
				strCurToken = strCurToken & Mid(strLine, n, 1)
			End If
		Next n
		If Len(Trim(strCurToken)) > 0 Then
			ReDim Preserve strTokens(j)
			strTokens(j) = Trim(strCurToken)
			j = j + 1
			strCurToken = vbNullString
		End If
		
		'Now process the tokens
		Dim posToken As Short
		Dim strField As String
		
		CondColorParmNumber = 1
		'UPGRADE_WARNING: Lower bound of array CondColorParms was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim CondColorParms(MAX_COLOR_CONDITIONS, 4)
		
		For n = LBound(strTokens) To UBound(strTokens)
			posToken = (n Mod 4)
			Select Case posToken
				Case 0 'extract the field name
					strField = CheckName(strTokens(n))
					If Len(strField) = 0 Then
						MsgBox("Error: '" & strTokens(n) & "' is not a valid field name.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
						ParseParmString = False
						Exit Function
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CondColorParms(CondColorParmNumber, 1) = strField
					
				Case 1 'extract the operator
					If InStr("=#><", strTokens(n)) = 0 Then
						MsgBox("Error:  Expected '=', '#', '>', or '<' where '" & strTokens(n) & "' appeared.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
						ParseParmString = False
						Exit Function
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CondColorParms(CondColorParmNumber, 2) = strTokens(n)
					
				Case 2 'extract the comparision value
					'120303 rjk: Added Food required and Food Excess for F4 Conditional
					'050404 rjk: Added Excess civilians for F4 Conditional
					'110605 rjk: Added RP - Reacting Planes
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CondColorParms(CondColorParmNumber, 1) = "FR" Or CondColorParms(CondColorParmNumber, 1) = "FS" Or CondColorParms(CondColorParmNumber, 1) = "FE" Or CondColorParms(CondColorParmNumber, 1) = "EC" Or CondColorParms(CondColorParmNumber, 1) = "BC" Or CondColorParms(CondColorParmNumber, 1) = "RP" Or CondColorParms(CondColorParmNumber, 1) = "PD" Then
						If IsNumeric(CObj(strTokens(n))) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							CondColorParms(CondColorParmNumber, 3) = CInt(strTokens(n))
						Else
							MsgBox("Error:  Expected a number where '" & strTokens(n) & "' appeared.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
							ParseParmString = False
							Exit Function
						End If
					ElseIf rsSectors.Fields(CondColorParms(CondColorParmNumber, 1)).Type = DAO.DataTypeEnum.dbInteger Or rsSectors.Fields(CondColorParms(CondColorParmNumber, 1)).Type = DAO.DataTypeEnum.dbLong Then 
						If IsNumeric(CObj(strTokens(n))) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							CondColorParms(CondColorParmNumber, 3) = CInt(strTokens(n))
						Else
							MsgBox("Error:  Expected a number where '" & strTokens(n) & "' appeared.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
							ParseParmString = False
							Exit Function
						End If
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CondColorParms(CondColorParmNumber, 3) = strTokens(n)
					End If
					
				Case 3 'extract the and / or symbol
					If InStr("&|", strTokens(n)) = 0 Then
						MsgBox("Error:  Expected '&' (AND) or '|' (OR) where '" & strTokens(n) & "' appeared.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
						ParseParmString = False
						Exit Function
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object CondColorParms(CondColorParmNumber, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CondColorParms(CondColorParmNumber, 4) = strTokens(n)
					CondColorParmNumber = CondColorParmNumber + 1
			End Select
		Next 
		
		If posToken <> 2 Then
			Select Case posToken
				Case 0 'ended on the field name
					MsgBox("Error:  Expected '=', '#', '>', or '<' next.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
				Case 1 'ended on the operator
					MsgBox("Error:  Expected comparision value next", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
				Case 3 'ended on the and / or symbol
					MsgBox("Error:  Conditional string can't end with '&' (AND) or '|' (OR)", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Conditional String Error")
			End Select
			ParseParmString = False
			Exit Function
		End If
		ParseParmString = True
	End Function
	
	Private Function CheckName(ByRef strIn As String) As String
		'120303 rjk: Check for Food Required and Food Excess
		If InStr(strIn, "FR") = 1 Then
			CheckName = "FR"
			Exit Function
		ElseIf InStr(strIn, "FS") = 1 Then 
			CheckName = "FS"
			Exit Function
		ElseIf InStr(strIn, "FE") = 1 Then 
			CheckName = "FE"
			Exit Function
		ElseIf InStr(strIn, "EC") = 1 Then  '050404 rjk: Added Excess Civilians
			CheckName = "EC"
			Exit Function
		ElseIf InStr(strIn, "BC") = 1 Then 
			CheckName = "BC"
			Exit Function
		ElseIf InStr(strIn, "RP") = 1 Then  '110605 rjk: Added RP - Reacting Planes
			CheckName = "RP"
			Exit Function
		ElseIf InStr(strIn, "PD") = 1 Then 
			CheckName = "PD"
			Exit Function
		End If
		'Match the string fragment to a sector field name
		Dim eiItem As EmpItem
		
		eiItem = Items.FindByConditionalName(strIn)
		If Not (eiItem Is Nothing) Then
			CheckName = eiItem.DatabaseName
			Exit Function
		End If
		
		Dim f As DAO.Field
		
		For	Each f In rsSectors.Fields
			If InStr(f.Name, strIn) = 1 Then
				CheckName = f.Name
				Exit Function
			End If
		Next f
		CheckName = vbNullString
	End Function
	
	Private Sub Picture1_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture1.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = Picture1.GetIndex(eventSender)
		If Button = VB6.MouseButtonConstants.LeftButton Then
			cd1Color.Color = System.Drawing.ColorTranslator.FromOle(BasicColors(Index + 1))
			cd1Color.ShowDialog()
			BasicColors(Index + 1) = System.Drawing.ColorTranslator.ToOle(cd1Color.Color)
		Else
			cd1Color.Color = System.Drawing.ColorTranslator.FromOle(BasicText(Index + 1))
			cd1Color.ShowDialog()
			BasicText(Index + 1) = System.Drawing.ColorTranslator.ToOle(cd1Color.Color)
		End If
		FillColors()
	End Sub
	Private Sub FillColors()
		Dim n As Short
		Dim ch As String
		
		For n = 1 To NUMBER_OF_BASIC_COLORS
			Picture1(n - 1).BackColor = System.Drawing.ColorTranslator.FromOle(BasicColors(n))
		Next n
		
		For n = 1 To NUMBER_OF_BASIC_COLORS
			'UPGRADE_ISSUE: PictureBox property Picture1.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture1(n - 1).AutoRedraw = True
			Picture1(n - 1).ForeColor = System.Drawing.ColorTranslator.FromOle(BasicText(n))
			SetDrawingFont(Picture1(n - 1))
			Picture1(n - 1).Font = VB6.FontChangeSize(Picture1(n - 1).Font, 14)
			ch = "c"
			'UPGRADE_ISSUE: PictureBox method Picture1.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property Picture1.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture1(n - 1).CurrentX = (VB6.PixelsToTwipsX(Picture1(n - 1).ClientRectangle.Width) - Picture1(n - 1).TextWidth(ch)) / 2
			'UPGRADE_ISSUE: PictureBox method Picture1.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property Picture1.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture1(n - 1).CurrentY = (VB6.PixelsToTwipsY(Picture1(n - 1).ClientRectangle.Height) - Picture1(n - 1).TextHeight(ch)) / 2
			'UPGRADE_ISSUE: PictureBox method Picture1.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture1(n - 1).Print(ch)
		Next n
		
		FillGradBar()
	End Sub
	
	Private Sub FillGradBar()
		
		Const NUM_SEGMENTS As Short = 25
		Dim n As Integer
		Dim ncol As Integer
		'UPGRADE_ISSUE: PictureBox property Picture2.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Picture2.AutoRedraw = True
		For n = 1 To NUM_SEGMENTS
			ncol = BlendedColor(BasicColors(7), BasicColors(8), n / NUM_SEGMENTS)
			'UPGRADE_ISSUE: PictureBox method Picture2.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture2.Line (((n - 1) / NUM_SEGMENTS) * VB6.PixelsToTwipsX(Picture2.ClientRectangle.Width), 0) - ((n / NUM_SEGMENTS) * VB6.PixelsToTwipsX(Picture2.ClientRectangle.Width), VB6.PixelsToTwipsY(Picture2.ClientRectangle.Height)), ncol, BF
		Next n
	End Sub
End Class