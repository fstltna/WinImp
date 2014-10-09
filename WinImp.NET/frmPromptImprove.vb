Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptImprove
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: Added Option Explicit and removed dead variables
	'092803 rjk: Added initial field selection. General reformatting.
	'            Made the options sector specific.
	'093003 rjk: Switched logic so if origin is not a single sector then all options are available.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'130506 rjk: Added South Pacific: Atlantis
	'210506 rjk: Fixed new SP: Atlantis new infrastructures percentages.
	'250506 rjk: Added Four SP: Atlantis infrastructures to the database,
	'            use runway -> min, radar -> gld, fort -> fert nad navigate -> oil.
	'280506 rjk: Fixed maximum percent for SP: Atlantis.
	'190606 rjk: Fixed maximum percentage for SP: Atlantis fields.
	
	Dim itype As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		strCmd = "improve " & itype & " " & txtOrigin.Text & " " & txtNum.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'frmEmpCmd.SubmitEmpireCommand "dump " + txtOrigin.Text, False
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptImprove.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptImprove_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		UpdateDesc()
		If optRoad.Enabled Then
			optRoad.Checked = True
		ElseIf optRail.Enabled Then 
			optRail.Checked = True
		ElseIf optDefense.Enabled Then 
			optDefense.Checked = True
		ElseIf optRadar.Enabled Then 
			optRadar.Checked = True
		ElseIf optFort.Enabled Then 
			optFort.Checked = True
		ElseIf optNavigation.Enabled Then 
			optNavigation.Checked = True
		ElseIf optRunway.Enabled Then 
			optRunway.Checked = True
		End If
		txtOrigin.Focus()
	End Sub
	
	Private Sub frmPromptImprove_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long   removed efj 8/2003
		If frmOptions.bSPAtlantis Then
			optRadar.Visible = True
			optRunway.Visible = True
			optFort.Visible = True
			optNavigation.Visible = True
		End If
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptImprove_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event optRoad.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRoad_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRoad.CheckedChanged
		If eventSender.Checked Then
			itype = optRoad.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRail.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRail_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRail.CheckedChanged
		If eventSender.Checked Then
			itype = optRail.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optDefense.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDefense_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDefense.CheckedChanged
		If eventSender.Checked Then
			itype = optDefense.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRadar.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRadar_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRadar.CheckedChanged
		If eventSender.Checked Then
			itype = optRadar.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optFort.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optFort_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFort.CheckedChanged
		If eventSender.Checked Then
			itype = optFort.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRunway.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRunway_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRunway.CheckedChanged
		If eventSender.Checked Then
			itype = optRunway.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optNavigation.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optNavigation_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNavigation.CheckedChanged
		If eventSender.Checked Then
			itype = optNavigation.Text
			UpdateDesc()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		UpdateDesc()
	End Sub
	
	Private Sub txtOrigin_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.DoubleClick
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptSectors)
		frmPromptSectors.strSectors = txtOrigin.Text
		frmPromptSectors.SetControls()
		frmPromptSectors.Text = "Select Sectors"
		frmPromptSectors.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptSectors.Width))
		frmPromptSectors.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptSectors.Height)) \ 2)
		frmPromptSectors.Show()
	End Sub
	
	Private Sub UpdateDesc()
		Dim sx As Short
		Dim sy As Short
		Dim MaxPercent As Short
		Dim hldIndex As String
		
		hldIndex = rsBuild.Index
		rsBuild.Index = "PrimaryKey"
		
		On Error Resume Next
		
		'093003 rjk: Switched logic so if origin is not a single sector then all options are available.
		'p1 = InStr(txtOrigin.Text, ",")
		'If p1 > 0 And Len(txtOrigin.Text) > p1 Then
		If Len(txtOrigin.Text) > 0 Then
			If ParseSectors(sx, sy, (txtOrigin.Text)) Then
				'sx = CInt(Left$(txtOrigin.Text, p1 - 1))
				'sy = CInt(Mid$(txtOrigin.Text, p1 + 1))
				'Get Sector Information
				rsSectors.Seek("=", sx, sy)
				If Not rsSectors.NoMatch Then
					lblDesc.Text = CStr(rsSectors.Fields("eff").Value) & "% "
					'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lblDesc.Text = lblDesc.Text + colDes.Item(rsSectors.Fields("des").Value)
					MaxPercent = 100
					If frmOptions.bSPAtlantis Then
						Select Case itype
							Case "runway"
								lblDesc.Text = lblDesc.Text & " w/ current " & itype & " of " & CStr(rsSectors.Fields("min").Value) & "% "
								MaxPercent = MaxPercent - rsSectors.Fields("min").Value
							Case "radar"
								lblDesc.Text = lblDesc.Text & " w/ current " & itype & " of " & CStr(rsSectors.Fields("gold").Value) & "% "
								MaxPercent = MaxPercent - rsSectors.Fields("gold").Value
							Case "fort"
								lblDesc.Text = lblDesc.Text & " w/ current " & itype & " of " & CStr(rsSectors.Fields("fert").Value) & "% "
								MaxPercent = MaxPercent - rsSectors.Fields("fert").Value
							Case "navigation"
								lblDesc.Text = lblDesc.Text & " w/ current " & itype & " of " & CStr(rsSectors.Fields("ocontent").Value) & "% "
								MaxPercent = MaxPercent - rsSectors.Fields("ocontent").Value
							Case Else
								lblDesc.Text = lblDesc.Text & " w/ current " & itype & " of " & CStr(rsSectors.Fields(itype).Value) & "% "
								MaxPercent = MaxPercent - rsSectors.Fields(itype).Value
						End Select
					Else
						lblDesc.Text = lblDesc.Text & " w/ current " & itype & " of " & CStr(rsSectors.Fields(itype).Value) & "% "
						MaxPercent = MaxPercent - rsSectors.Fields(itype).Value
					End If
					
					rsBuild.Seek("=", "i", VB.Left(itype, 5))
					If Not rsBuild.NoMatch Then
						'compute maximum
						If MaxPercent > (rsSectors.Fields("mob").Value / rsBuild.Fields("stat2").Value) Then
							MaxPercent = Int(rsSectors.Fields("mob").Value / rsBuild.Fields("stat2").Value)
						End If
						If MaxPercent > Int(rsSectors.Fields("lcm").Value / rsBuild.Fields("lcm").Value) Then
							MaxPercent = Int(rsSectors.Fields("lcm").Value / rsBuild.Fields("lcm").Value)
						End If
						If MaxPercent > Int(rsSectors.Fields("hcm").Value / rsBuild.Fields("hcm").Value) Then
							MaxPercent = Int(rsSectors.Fields("hcm").Value / rsBuild.Fields("hcm").Value)
						End If
					Else
						MaxPercent = 0
					End If
					txtNum.Text = CStr(MaxPercent)
					
					rsBuild.Seek("=", "i", "road")
					If CDbl(CStr(rsSectors.Fields("road").Value)) < 100 And Not rsBuild.NoMatch Then
						If Not optRoad.Enabled Then
							optRoad.Enabled = True
						End If
					Else
						optRoad.Enabled = False
					End If
					rsBuild.Seek("=", "i", "rail")
					If CDbl(CStr(rsSectors.Fields("rail").Value)) < 100 And Not rsBuild.NoMatch Then
						If Not optRail.Enabled Then
							optRail.Enabled = True
						End If
					Else
						optRail.Enabled = False
					End If
					rsBuild.Seek("=", "i", "defen")
					If CDbl(CStr(rsSectors.Fields("defense").Value)) < 100 And Not rsBuild.NoMatch Then
						If Not optDefense.Enabled Then
							optDefense.Enabled = True
						End If
					Else
						optDefense.Enabled = False
					End If
					If frmOptions.bSPAtlantis Then
						rsBuild.Seek("=", "i", "runwa")
						If CDbl(CStr(rsSectors.Fields("min").Value)) < 100 And Not rsBuild.NoMatch Then
							If Not optRunway.Enabled Then
								optRunway.Enabled = True
							End If
						Else
							optRunway.Enabled = False
						End If
						rsBuild.Seek("=", "i", "radar")
						If CDbl(CStr(rsSectors.Fields("gold").Value)) < 100 And Not rsBuild.NoMatch Then
							If Not optRadar.Enabled Then
								optRadar.Enabled = True
							End If
						Else
							optRadar.Enabled = False
						End If
						rsBuild.Seek("=", "i", "fort")
						If CDbl(CStr(rsSectors.Fields("fert").Value)) < 100 And Not rsBuild.NoMatch Then
							If Not optFort.Enabled Then
								optFort.Enabled = True
							End If
						Else
							optFort.Enabled = False
						End If
						rsBuild.Seek("=", "i", "navig")
						If CDbl(CStr(rsSectors.Fields("ocontent").Value)) < 100 And Not rsBuild.NoMatch Then
							If Not optNavigation.Enabled Then
								optNavigation.Enabled = True
							End If
						Else
							optNavigation.Enabled = False
						End If
					End If
				Else
					optRoad.Enabled = False
					optRail.Enabled = False
					optDefense.Enabled = False
					If frmOptions.bSPAtlantis Then
						optRunway.Enabled = False
						optRadar.Enabled = False
						optFort.Enabled = False
						optNavigation.Checked = False
					End If
				End If
			Else '093003 rjk: In case a multiple sector group is selected, switched to true.
				optRoad.Enabled = True
				optRail.Enabled = True
				optDefense.Enabled = True
				If frmOptions.bSPAtlantis Then
					optRunway.Enabled = True
					optRadar.Enabled = True
					optFort.Enabled = True
					optNavigation.Checked = True
				End If
			End If
		Else
			optRoad.Enabled = False
			optRail.Enabled = False
			optDefense.Enabled = False
			If frmOptions.bSPAtlantis Then
				optRunway.Enabled = False
				optRadar.Enabled = False
				optFort.Enabled = False
				optNavigation.Checked = False
			End If
		End If
		
		If optRoad.Enabled Or optRail.Enabled Or optDefense.Enabled Or frmOptions.bSPAtlantis Then
			cmdOK.Enabled = True
		Else
			cmdOK.Enabled = False
			txtOrigin.Focus()
		End If
		
		rsBuild.Index = hldIndex
	End Sub
End Class