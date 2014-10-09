Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmPromptAttack
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: Removed Dead Variables
	'092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding.
	'            General reformatting. Clear to ship combo to prevent invalid
	'            ships appearing in the list after the target was moved.
	'            Switched to SetUnitType and DisplayFirstSectorPanel
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'120703 rjk: Changed hldindex to hldIndex.
	'270104 rjk: St@r W@rs, must be on top of the sector to assault.
	'270304 rjk: Retro, assault does not required land or semi-land capability
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace ldump and sdump with GetLandUnits and GetShips.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'120108 rjk: Removed unused local variables iCountry and Des from txtOriginChange.
	'260408 rjk: Added assault in the same sector for Heavy Metal II
	'140708 rjk: Added GetLost to attack and assault.
	'160209 rjk: Allow assault of the sector when in it (new in server 4.3.19).
	'100410 rjk: Fixed Attack/Assault support to use y/n instead 1/0.
	
	Dim Support As Boolean
	Dim Assault As Boolean
	
	'UPGRADE_WARNING: Event cbShips.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cbShips.Change was upgraded to cbShips.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cbShips_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbShips.TextChanged
		LoadUnitBox()
	End Sub
	
	'UPGRADE_WARNING: Event cbShips.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbShips_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbShips.SelectedIndexChanged
		LoadUnitBox()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim strCmd2 As String
		Dim strSupport As String
		' Dim strTemp As String
		Dim X As Short
		' Dim sx As Integer
		' Dim sy As Integer
		
		strSupport = vbNullString
		
		For X = 0 To 3
			If Check1(X).CheckState = System.Windows.Forms.CheckState.Checked Then
				strSupport = strSupport & " y"
			Else
				strSupport = strSupport & " n"
			End If
		Next X
		
		If Label2.Text = "Attack" Then
			strCmd = "attack " & txtOrigin.Text
		Else
			strCmd = StringBetween((cbShips.Text), "#", "(")
			If Len(strCmd) = 0 Then
				strCmd = cbShips.Text
			End If
			strCmd = "assault " & txtOrigin.Text & " " & strCmd
		End If
		
		If strSupport <> " y y y y" Then
			strCmd = strCmd & strSupport
		End If
		
		'build header
		strCmd2 = "at1 "
		If Not Option1.Checked Then
			If Option2.Checked Then
				strCmd2 = strCmd2 & "/mil=29999;"
			ElseIf Option3.Checked And Len(Text1.Text) > 0 Then 
				strCmd2 = strCmd2 & "/mil=" & Trim(Text1.Text) & ";"
			End If
			strCmd2 = strCmd2 & "/sec=" & txtOrigin.Text & ";"
		End If
		
		If Not Option4.Checked Then
			If Option5.Checked Then
				strCmd2 = strCmd2 & "/occ=0;"
			ElseIf Option7.Checked Then 
				strCmd2 = strCmd2 & "/occ=29999;"
			ElseIf Option8.Checked And Len(Text2.Text) > 0 Then 
				strCmd2 = strCmd2 & "/occ=" & Trim(Text2.Text) & ";"
			End If
		End If
		
		'Attacking Units
		If Len(txtUnit.Text) > 0 Then
			strCmd2 = strCmd2 & "/unt=" & txtUnit.Text & ";"
		End If
		
		'Occupying Units
		If Len(txtUnit2.Text) > 0 Then
			strCmd2 = strCmd2 & "/ocu=" & txtUnit2.Text & ";"
		End If
		
		If Len(strCmd2) > 4 Then
			frmEmpCmd.SubmitEmpireCommand(strCmd2, False)
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			frmEmpCmd.SubmitEmpireCommand("at2", False)
		Else
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		GetLandUnits()
		If Label2.Text <> "Attack" Then
			GetShips()
		End If
		GetLost()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		txtUnit.Text = "Y"
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		txtUnit.Text = "N"
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		txtUnit2.Text = "N"
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		txtUnit2.Text = "Y"
	End Sub
	
	Private Sub frmPromptAttack_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		Support = True
	End Sub
	
	Private Sub frmPromptAttack_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	
	'UPGRADE_ISSUE: Frame event Frame1.Click was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Frame1_Click()
		Dim X As Short
		
		'Toggle support
		Support = Not Support
		
		If Support Then
			For X = 0 To 3
				Check1(X).CheckState = System.Windows.Forms.CheckState.Checked
			Next X
		Else
			For X = 0 To 3
				Check1(X).CheckState = System.Windows.Forms.CheckState.Unchecked
			Next X
		End If
		
	End Sub
	
	
	'UPGRADE_ISSUE: Label event Label2.Change was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Label2_Change()
		If Label2.Text = "Attack" Then
			Frame3.Enabled = True
			Frame3.Visible = True
			Frame6.Visible = False
			Assault = False
		Else
			Frame6.Visible = True
			Frame3.Visible = False
			Me.cbShips.Visible = True
			Me.Label1.Visible = True
			Assault = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event lstUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstUnits.SelectedIndexChanged
		Dim n As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		str_Renamed = vbNullString
		For n = 0 To lstUnits.Items.Count - 1
			If lstUnits.GetSelected(n) Then
				str_Renamed = str_Renamed & StringBetween(VB6.GetItemString(lstUnits, n), "#", "(") & "/"
			End If
		Next n
		If Len(str_Renamed) > 0 Then
			str_Renamed = VB.Left(str_Renamed, Len(str_Renamed) - 1)
		End If
		txtUnit.Text = str_Renamed
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			If Option3.Checked Then
				Text1.Focus()
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option8.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option8_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option8.CheckedChanged
		If eventSender.Checked Then
			If Option8.Checked Then
				Text2.Focus()
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Text1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		If Me.Visible Then
			If Len(Text1.Text) > 0 Then
				Option3.Checked = True
			Else
				Option1.Checked = True
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Text2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.TextChanged
		If Len(Text1.Text) > 0 Then
			Option8.Checked = True
		Else
			Option4.Checked = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim iEnemyMil As Short
		
		iEnemyMil = EnemyMil(txtOrigin.Text)
		
		If Assault Then
			LoadCombo(iEnemyMil)
			If cbShips.Items.Count > 0 Then
				cbShips.SelectedIndex = 0
			End If
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, True, True, True, txtOrigin.Text)
		End If
		
		If iEnemyMil > 0 Then
			Text1.Text = CStr(iEnemyMil * 1.5)
		End If
	End Sub
	
	Private Sub LoadCombo(ByRef DefMil As Short)
		
		Dim sx As Short
		Dim sy As Short
		Dim ass_str As Integer
		Dim n As Short
		Dim i As Short
		Dim strAdd As String
		Dim Done As Boolean
		Dim hldIndex As Object
		
		If Not ParseSectors(sx, sy, (txtOrigin.Text)) Then
			Exit Sub
		End If
		
		'save and set land unit index
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsLand.Index
		rsLand.Index = "location"
		
		cbShips.Items.Clear() '092403 rjk: Remove any ship not valid anymore
		
		If Not (rsShip.BOF And rsShip.EOF) Then
			rsShip.MoveFirst()
		End If
		While Not rsShip.EOF
			
			'choose the ships that could assault the target sector
			'270104 rjk: St@r W@rs, must be on top of the sector to assault.
			If (SectorDistance(sx, sy, rsShip.Fields("x").Value, rsShip.Fields("y").Value) = 1 And Not frmOptions.bStarWars) Or (sx = rsShip.Fields("x").Value And sy = rsShip.Fields("y").Value And frmOptions.bStarWars) Or (sx = rsShip.Fields("x").Value And sy = rsShip.Fields("y").Value And frmOptions.bHeavyMetal) Or (sx = rsShip.Fields("x").Value And sy = rsShip.Fields("y").Value And VersionCheck(4, 3, 19) <> WinAceRoutines.enumVersion.VER_LESS) Then
				'compute the number of ship-board military who could assault
				ass_str = rsShip.Fields("mil").Value - 1
				
				'092403 rjk: switched to UnitCharacteristicCheck to remove hard coding
				'        Select Case rsShip.Fields("type")
				'            Case "ls"
				'                ass_str = ass_str
				'            Case "tt", "frg"
				'                ass_str = Int(ass_str / 4#)
				'            Case Else
				'                ass_str = Int(ass_str / 10#)
				'        End Select
				If DefMil > 0 Then
					If Not frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_S_LAND, rsShip.Fields("type").Value) Then
						If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_S_SEMI_LAND, rsShip.Fields("type").Value) Then
							ass_str = Int(ass_str / 4#)
						Else
							ass_str = Int(ass_str / 10#)
						End If
					End If
				End If
				
				'check for lands on the ship
				If Not (rsLand.EOF And rsLand.BOF) Then
					rsLand.MoveFirst()
					rsLand.Seek("=", rsShip.Fields("x"), rsShip.Fields("y"))
					Done = rsLand.NoMatch
				Else
					Done = True
				End If
				
				While Not rsLand.EOF And Not Done
					'make sure they are still in the hex
					If rsShip.Fields("x").Value = rsLand.Fields("x").Value And rsShip.Fields("y").Value = rsLand.Fields("y").Value Then
						'make sure they are on the ship
						If rsLand.Fields("ship").Value = rsShip.Fields("id").Value Then
							'make sure they can assault
							'092403 rjk: switched to UnitCharacteristicCheck to remove hard coding
							'If InStr(" linf inf eng mar meng ", " " + rsLand.Fields("type") + " ") > 0 Then
							If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_L_ASSAULT, rsLand.Fields("type").Value) Then
								ass_str = ass_str + CShort(rsLand.Fields("mil").Value * (rsLand.Fields("eff").Value / 100#) * rsLand.Fields("att").Value * 0.5)
							End If
						End If
						rsLand.MoveNext()
					Else
						Done = True
					End If
				End While
				
				If ass_str > 0 Then
					
					'do a simple insertion sort into the list.
					n = 0
					i = cbShips.Items.Count
					While n <= cbShips.Items.Count - 1
						If VB6.GetItemData(cbShips, n) > ass_str Then
							n = n + 1
						Else
							i = n
							n = cbShips.Items.Count
						End If
					End While
					'build the display sting
					strAdd = rsShip.Fields("type").Value + "#" + CStr(rsShip.Fields("id").Value) + " (" + CStr(ass_str) + ")"
					If i = cbShips.Items.Count Then
						cbShips.Items.Add(strAdd)
					Else
						cbShips.Items.Insert(i, strAdd)
					End If
					'UPGRADE_ISSUE: ComboBox property cbShips.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					VB6.SetItemData(cbShips, cbShips.NewIndex, ass_str)
				End If
			End If
			rsShip.MoveNext()
		End While
		
		'restore land index
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsLand.Index = hldIndex
		
	End Sub
	
	Private Sub LoadUnitBox()
		Dim strShip As String
		Dim hldIndex As Object
		Dim hldIndex2 As Object
		Dim Done As Boolean
		Dim ass_str As Short
		Dim strEntry As String
		
		If Len(cbShips.Text) = 0 Then
			Exit Sub
		End If
		
		'save and set land unit index
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsLand.Index
		rsLand.Index = "location"
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex2 = rsShip.Index
		rsShip.Index = "PrimaryKey"
		
		On Error GoTo cbShips_Change_exit
		
		
		'get ship number
		strShip = StringBetween((cbShips.Text), "#", "(")
		If Len(strShip) = 0 Then
			strShip = cbShips.Text
		End If
		
		
		'Load unit box
		lstUnits.Items.Clear()
		rsShip.Seek("=", CShort(strShip))
		If Not rsShip.NoMatch Then
			'check for lands on the ship
			rsLand.MoveFirst()
			rsLand.Seek("=", rsShip.Fields("x"), rsShip.Fields("y"))
			Done = rsLand.NoMatch
			
			While Not rsLand.EOF And Not Done
				'make sure they are still in the hex
				If rsShip.Fields("x").Value = rsLand.Fields("x").Value And rsShip.Fields("y").Value = rsLand.Fields("y").Value Then
					'make sure they are on the ship
					If rsLand.Fields("ship").Value = rsShip.Fields("id").Value Then
						'make sure they can assault
						'092403 rjk: switched to UnitCharacteristicCheck to remove hard coding
						'If InStr(" linf inf eng mar meng ", " " + rsLand.Fields("type") + " ") > 0 Then
						If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_L_ASSAULT, rsLand.Fields("type").Value) Then
							ass_str = CShort(rsLand.Fields("mil").Value * (rsLand.Fields("eff").Value / 100#) * rsLand.Fields("att").Value * 0.5)
							strEntry = rsLand.Fields("type").Value + "#" + CStr(rsLand.Fields("id").Value) + " (" + CStr(ass_str) + ")"
							lstUnits.Items.Add(strEntry)
						End If
					End If
					rsLand.MoveNext()
				Else
					Done = True
				End If
			End While
			
		End If
		
cbShips_Change_exit: 
		
		'reset index
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsLand.Index = hldIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsShip.Index = hldIndex2
		
	End Sub
End Class