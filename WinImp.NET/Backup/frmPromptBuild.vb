Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptBuild
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092803 rjk: Added timestamp ndump. General reformatting.
	'            Added initial field selection.
	'            Options now done based on the origin selected.
	'100103 rjk: Adjust positions of the options so the text boxes do not overlap.
	'100203 rjk: Fixed if a harbor is selected it will have the combo filled with ships
	'            (disabled option1 on the form).
	'100503 rjk: Removed timestamp ndump.
	'101403 rjk: Added ndump timestamp (try #2)
	'            Changed initial field selection to be the either direction or combo box.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'080304 rjk: Fixed crash when selecting a sector while building bridge span
	'080304 rjk: Changed to determine direction to build from sector selection on the map
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'280804 rjk: For 2K4, explore with m instead c when building bridge spans or towers.
	'180206 rjk: Replace ldump, sdump, ndump, lost with GetLandUnits, GetShips, GetNukes and GetLost.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'270407 rjk: Do not automatically select bridge tower when bridge span is available.
	
	Public strCmd As String
	Private bFirstLoad As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		strCmd = vbNullString
		Me.Close()
	End Sub
	
	Private Sub cmdDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDir.Click
		Dim Index As Short = cmdDir.GetIndex(eventSender)
		txtNum.Text = LCase(cmdDir(Index).Text)
		cmdOK.PerformClick()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim X As Short
		On Error Resume Next
		' Dim strCmd2 As String     removed 8/03 ej
		
		'Start with first list box results
		If Option1.Checked Then
			strCmd = "land"
			frmEmpCmd.HaveLands = True
		ElseIf Option2.Checked Then 
			strCmd = "plane"
			frmEmpCmd.HavePlanes = True
		ElseIf Option3.Checked Then 
			strCmd = "ship"
			frmEmpCmd.HaveShips = True
		ElseIf Option4.Checked Then 
			strCmd = "nuke"
			frmEmpCmd.HaveNukes = True
		ElseIf Option7.Checked Then 
			strCmd = "bridge"
		Else
			strCmd = "tower"
		End If
		
		strCmd = "build " & strCmd & " " & txtOrigin.Text
		If Option7.Checked Or Option8.Checked Then 'if  bridge then fill in
			If Len(Trim(txtNum.Text)) > 0 Then
				strCmd = strCmd & " " & txtNum.Text
			End If
			frmEmpCmd.SubmitEmpireCommand("bf1", False)
		Else
			strCmd = strCmd & " " & Trim(VB.Left(txtType.Text, 5)) & " "
			'Add number if input
			If Len(txtNum.Text) > 0 Then
				X = Val(txtNum.Text)
				If X < 1 Then
					MsgBox("Number String not valid", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Entry Error")
					txtNum.Focus()
					Exit Sub
				End If
				strCmd = strCmd & " " & CStr(Val(txtNum.Text))
			End If
			
			'Add tech level if input
			If Len(txtTechLevel.Text) > 0 Then
				If Trim(txtTechLevel.Text) <> "0" And Val(txtTechLevel.Text) < 1 Then
					MsgBox("Tech String not valid", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Entry Error")
					txtTechLevel.Focus()
					Exit Sub
				End If
				strCmd = strCmd & " " & CStr(Val(txtTechLevel.Text))
			End If
			
		End If
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'If we just built a bridge, try to explore it
		If Option7.Checked Or Option8.Checked Then
			If Len(Trim(txtNum.Text)) > 0 Then
				If frmOptions.b2K4 Then
					frmEmpCmd.SubmitEmpireCommand("explore m " & txtOrigin.Text & " 1 " & txtNum.Text & "h", True)
				Else
					frmEmpCmd.SubmitEmpireCommand("explore c " & txtOrigin.Text & " 1 " & txtNum.Text & "h", True)
				End If
			End If
			frmEmpCmd.SubmitEmpireCommand("bf2", False)
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If Option3.Checked Then
			GetShips()
		ElseIf Option1.Checked Then 
			GetLandUnits()
		ElseIf Option2.Checked Then 
			GetPlanes()
		ElseIf Option4.Checked Then 
			GetNukes()
		End If
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Private Sub cmdShow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShow.Click
		frmToolShow.Show()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptBuild.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptBuild_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'txtOrigin.SetFocus
		Dim sx As Short
		Dim sy As Short
		Dim strPath As String
		If bFirstLoad Then
			If Option8.Checked Or Option7.Checked Then
				txtNum.Focus()
			Else
				txtType.Focus()
			End If
		Else
			'080304 rjk: Fixed crash when selecting a sector while building bridge span
			'080304 rjk: Changed to determine direction to build from sector selection on the map
			If Option8.Checked Or Option7.Checked Then
				If txtOrigin.Text <> "" Then
					
					If ParseSectors(sx, sy, txtOrigin.Text) Then
						If frmDrawMap.SelX <> sx Or frmDrawMap.SelY <> sy Then
							rsBmap.Seek("=", frmDrawMap.SelX, frmDrawMap.SelY)
							If Not rsBmap.NoMatch Then
								If rsBmap.Fields("des").Value = "." Then
									strPath = EmpirePathDirection(frmDrawMap.SelX - sx, frmDrawMap.SelY - sy)
									If Len(strPath) = 1 And strPath <> "h" Then
										txtNum.Text = strPath
									Else
										txtNum.Text = ""
									End If
								Else
									txtNum.Text = ""
								End If
							Else
								txtNum.Text = ""
							End If
						End If
					End If
				End If
			End If
		End If
		bFirstLoad = False
	End Sub
	
	Private Sub frmPromptBuild_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		'Option3_Click 092803 rjk: Options now done based on the origin selected.
		'101403 rjk: Added bFirstLoad to control what the initial field is
		bFirstLoad = True
	End Sub
	
	Private Sub frmPromptBuild_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			'Option1 is Lands
			Label4.Width = Label3.Width
			Label4.Text = "type"
			txtType.Visible = True
			txtType.Left = txtOrigin.Left
			
			Label1.Visible = True
			Label5.Visible = True
			txtNum.Visible = True
			txtTechLevel.Visible = True
			txtNum.SetBounds(txtTechLevel.Left, Label5.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDir.Visible = False
			
			LoadTypebox("l")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			Label4.Width = Label3.Width
			Label4.Text = "type"
			txtType.Visible = True
			txtType.Left = txtOrigin.Left
			Label1.Visible = True
			Label5.Visible = True
			txtNum.Visible = True
			txtTechLevel.Visible = True
			txtNum.SetBounds(txtTechLevel.Left, Label5.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDir.Visible = False
			
			LoadTypebox("p")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			Label4.Width = Label3.Width
			Label4.Text = "type"
			txtType.Visible = True
			txtType.Left = txtOrigin.Left
			Label1.Visible = True
			Label5.Visible = True
			txtNum.Visible = True
			txtTechLevel.Visible = True
			txtNum.SetBounds(txtTechLevel.Left, Label5.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDir.Visible = False
			
			LoadTypebox("s")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option4.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option4_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option4.CheckedChanged
		If eventSender.Checked Then
			Label4.Width = Label3.Width
			Label4.Text = "type"
			txtType.Visible = True
			txtType.Left = txtOrigin.Left
			Label1.Visible = True
			Label5.Visible = True
			txtNum.Visible = True
			txtTechLevel.Visible = True
			txtNum.SetBounds(txtTechLevel.Left, Label5.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDir.Visible = False
			
			LoadTypebox("n")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option7.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option7_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option7.CheckedChanged
		If eventSender.Checked Then
			Label4.Width = Label1.Width
			Label4.Text = "Direction"
			txtType.Visible = False
			
			Label1.Visible = False
			Label5.Visible = False
			txtNum.Visible = True
			txtTechLevel.Visible = False
			
			txtNum.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtOrigin.Left) + VB6.PixelsToTwipsX(txtOrigin.Width) - VB6.PixelsToTwipsX(txtNum.Width)), Label4.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDir.Visible = True
			frmDir.SetBounds(Label4.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtNum.Top) + VB6.PixelsToTwipsY(txtNum.Height) * 1.5), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option8.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option8_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option8.CheckedChanged
		If eventSender.Checked Then
			Label4.Width = Label1.Width
			Label4.Text = "Direction"
			txtType.Visible = False
			
			Label1.Visible = False
			Label5.Visible = False
			txtNum.Visible = True
			txtTechLevel.Visible = False
			
			txtNum.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtOrigin.Left) + VB6.PixelsToTwipsX(txtOrigin.Width) - VB6.PixelsToTwipsX(txtNum.Width)), Label4.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDir.Visible = True
			frmDir.SetBounds(Label4.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtNum.Top) + VB6.PixelsToTwipsY(txtNum.Height) * 1.5), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		End If
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		If frmReport.Visible Then
			If frmReport.WindowState = System.Windows.Forms.FormWindowState.Normal Then
				
				'Size Report
				If VB6.PixelsToTwipsY(frmReport.Height) > VB6.PixelsToTwipsY(Me.Top) Then
					frmReport.Height = Me.Top
				End If
				
				'move report and shut off timer
				frmReport.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) - VB6.PixelsToTwipsY(frmReport.Height))
				frmReport.Left = Me.Left
				'UPGRADE_WARNING: Timer property Timer1.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
				Timer1.Interval = 0
			End If
		End If
	End Sub
	
	Private Sub LoadTypebox(ByRef UnitClass As String)
		txtType.Items.Clear()
		txtType.Text = vbNullString
		
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = UnitClass Then
					If TechLevel < 0 Or rsBuild.Fields("tech").Value <= TechLevel Then
						txtType.Items.Add(VB.Left(rsBuild.Fields("id").Value + "     ", 5) + rsBuild.Fields("desc").Value)
					End If
				End If
				rsBuild.MoveNext()
			End While
		End If
	End Sub
	
	Private Sub SelectOptions() ' 092803 rjk: Options now based on the sector selected.
		Dim secx As Short
		Dim secy As Short
		
		Option1.Enabled = False
		Option2.Enabled = False
		Option3.Enabled = False
		Option4.Enabled = False
		Option7.Enabled = False
		Option8.Enabled = False
		
		If ParseSectors(secx, secy, (txtOrigin.Text)) Then
			rsSectors.Seek("=", secx, secy)
			If Not rsSectors.NoMatch Then
				If rsSectors.Fields("coast").Value = 1 Then
					If Not Option7.Enabled Then
						Option7.Enabled = True
					End If
				Else
					If Option7.Enabled Then
						Option7.Enabled = False
					End If
				End If
				Select Case rsSectors.Fields("des").Value
					Case "h"
						Option3.Checked = True
						Option1.Enabled = False
						Option2.Enabled = False
						Option3.Enabled = True
						Option4.Enabled = False
						Option8.Enabled = False
					Case "*"
						Option2.Checked = True
						Option1.Enabled = False
						Option2.Enabled = True
						Option3.Enabled = False
						Option4.Enabled = False
						Option8.Enabled = False
					Case "!"
						Option1.Checked = True
						Option1.Enabled = True
						Option2.Enabled = False
						Option3.Enabled = False
						Option4.Enabled = False
						Option8.Enabled = False
					Case "n"
						Option4.Checked = True
						Option1.Enabled = False
						Option2.Enabled = False
						Option3.Enabled = False
						Option4.Enabled = True
						Option8.Enabled = False
					Case "="
						Option1.Enabled = False
						Option2.Enabled = False
						Option3.Enabled = False
						Option4.Enabled = False
						Option8.Enabled = True
						If Not Option7.Enabled Then
							Option8.Checked = True
						Else
							Option7.Checked = False
							Option8.Checked = False
						End If
					Case Else
						If Option7.Enabled Then
							Option7.Checked = True
						End If
				End Select
			End If
		ElseIf txtOrigin.Text <> "" Then 
			Option7.Checked = True
			Option1.Enabled = True
			Option2.Enabled = True
			Option3.Enabled = True
			Option4.Enabled = True
			Option8.Enabled = True
			Option7.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		SelectOptions()
	End Sub
End Class