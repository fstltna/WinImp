Option Strict Off
Option Explicit On
Friend Class frmPromptDes
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092803 rjk: general reformatting.
	'101403 rjk: Added check box for clearing old thresholds from old designation.
	'101703 rjk: Added Strength fields to Sector database
	'101903 rjk: Switched txtOrigin to txtMultOrigin to support multiple sector selection.
	'            Added bFirstLoad to only do initial field selection once.
	'102303 rjk: Added Sector count for Multiple Sector Selection
	'110703 rjk: Removed the threshold textbox when the label is blank or n/a
	'            Fixed to correctly save the threshold labels and values.
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'300505 rjk: Fixed defaults 'd' oil 10 and 'j,k' iron 600.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private strCodes As String
	Private strNewDes As String
	Private bAutoFill As Boolean
	Private bFirstLoad As Boolean '101903 rjk: Added bFirstLoad to prevent the focus from being drawn away from the txtMultOrigin field
	
	'UPGRADE_WARNING: Event cbType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbType.SelectedIndexChanged
		If cbType.SelectedIndex < 0 Then
			Exit Sub
		End If
		strNewDes = Mid(strCodes, VB6.GetItemData(cbType, cbType.SelectedIndex), 1)
		
		'Reset threshholds
		SetThreshholds()
	End Sub
	
	Private Sub cbType_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbType.DoubleClick
		cmdOK.PerformClick()
	End Sub
	
	'UPGRADE_WARNING: Event Check1.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
		'Reset threshholds
		SetThreshholds()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim X As Short
		Dim beginPos As Short
		Dim endPos As Short
		
		'Save the new defaults, if indicated
		If chkSave.Visible And chkSave.CheckState = System.Windows.Forms.CheckState.Checked Then
			For X = 0 To 3 '110703 rjk: Fixed to save the correct threshold labels
				SaveSetting(APPNAME, "Thresholds", strNewDes & CStr(X + 1) & " type", lblThresh(X).Text)
				SaveSetting(APPNAME, "Thresholds", strNewDes & CStr(X + 1) & " level", txtThresh(X).Text)
			Next X
		End If
		
		'101803 rjk: Logic redone and support function SetDesignation added for Multiple Sectors Selection
		If InStr(txtMultOrigin.Text, "\") > 0 Then
			beginPos = 1
			endPos = InStr(txtMultOrigin.Text, "\")
			While endPos > 0
				SetDesignation((Mid(txtMultOrigin.Text, beginPos, endPos - beginPos)))
				beginPos = endPos + 1
				endPos = InStr(beginPos, txtMultOrigin.Text, "\")
			End While
			SetDesignation((Mid(txtMultOrigin.Text, beginPos)))
		Else
			SetDesignation((txtMultOrigin.Text))
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Me.Close()
	End Sub
	
	'101903 rjk: Added Multiple Sector Selection
	Private Sub SetDesignation(ByRef strSector As String)
		Dim sx As Short
		Dim sy As Short
		Dim strCmd As String
		Dim strComm As String
		Dim X As Short
		
		'101403 rjk: Clear old thresholds from old designation when setting new thresholds
		If ParseSectors(sx, sy, strSector) And chkClearOldThresholds.CheckState Then
			'Get Sector Information
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				For X = 0 To 3 '110703 rjk: Switch to 0 to 3 to standardize with the rest of the code
					strComm = GetSetting(APPNAME, "Thresholds", rsSectors.Fields("des").Value + CStr(X + 1) + " type", vbNullString)
					If Len(strComm) > 0 And strComm <> "n/a" Then
						strCmd = "threshold " & strComm & " " & txtMultOrigin.Text & " " & " 0"
						frmEmpCmd.SubmitEmpireCommand(strCmd, False)
					End If
				Next X
			End If
		End If
		
		'Set Thresholds first (since des can change optional params)
		For X = 0 To 3
			If txtThresh(X).Visible And Len(txtThresh(X).Text) > 0 Then
				strCmd = "threshold " & lblThresh(X).Text & " " & strSector & " " & txtThresh(X).Text
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			End If
		Next X
		
		'now do the designation
		frmEmpCmd.SubmitEmpireCommand("designate " & strSector & " " & strNewDes, True)
	End Sub
	
	Private Sub cbType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cbType.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim ch As String
		Dim n As Short
		Dim i As Short
		
		ch = Chr(KeyAscii)
		n = InStr(strCodes, ch)
		If n > 0 Then
			strNewDes = ch
			For i = 0 To cbType.Items.Count - 1
				If VB6.GetItemData(cbType, i) = n Then
					cbType.SelectedIndex = i
				End If
			Next i
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptDes.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptDes_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If bFirstLoad Then '101903 rjk: Added bFirstLoad to prevent the focus from being drawn away from the txtMultOrigin field
			bFirstLoad = False
			cbType.Focus()
		End If
	End Sub
	
	Private Sub frmPromptDes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		bFirstLoad = True
		strCodes = LoadTypebox(cbType)
		cbType.SelectedIndex = 0
		chkSave.Visible = False
		
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptDes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub lblThresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblThresh.Click
		Dim Index As Short = lblThresh.GetIndex(eventSender)
		Dim newthresh As String
		
		'Allow a thresh title to be changed by clicking
		newthresh = InputBox("Set threshold type", "Change Threshold", lblThresh(Index).Text)
		If newthresh <> lblThresh(Index).Text Then
			lblThresh(Index).Text = newthresh
			If Len(newthresh) > 0 And newthresh <> "n/a" Then
				txtThresh(Index).Visible = True
			Else '110703 rjk: Removed the textbox when the label is blank or n/a
				txtThresh(Index).Visible = False
			End If
			If bAutoFill = False And Not chkSave.Visible Then
				chkSave.Visible = True
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtMultOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		On Error Resume Next
		
		If ParseSectors(sx, sy, (txtMultOrigin.Text)) Then
			'Get Sector Information
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				lblDesc.Text = "Sector " & txtMultOrigin.Text & " is currently a "
				lblDesc.Text = lblDesc.Text & CStr(rsSectors.Fields("eff").Value) & "% "
				'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lblDesc.Text = lblDesc.Text + colDes.Item(rsSectors.Fields("des").Value)
			End If
			'102303 rjk: Added Sector count for Multiple Sector Selection
		ElseIf InStr(txtMultOrigin.Text, "\") > 0 Then 
			lblDesc.Text = CStr(CountMultipleSectors((txtMultOrigin.Text))) & " Sectors"
		Else
			lblDesc.Text = ""
		End If
	End Sub
	
	Private Sub txtMultOrigin_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.DoubleClick
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptSectors)
		frmPromptSectors.strSectors = txtMultOrigin.Text
		frmPromptSectors.SetControls()
		frmPromptSectors.Text = "Select Sectors"
		frmPromptSectors.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptSectors.Width))
		frmPromptSectors.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptSectors.Height)) \ 2)
		frmPromptSectors.Show()
	End Sub
	
	Private Sub SetThreshholds()
		Dim thresh As String
		Dim n As Short
		
		bAutoFill = True
		'Clear all thresholds
		lblThresh(0).Text = "n/a"
		txtThresh(0).Visible = False
		txtThresh(0).Text = vbNullString
		lblThresh(1).Text = "n/a"
		txtThresh(1).Visible = False
		txtThresh(1).Text = vbNullString
		lblThresh(2).Text = "n/a"
		txtThresh(2).Visible = False
		txtThresh(2).Text = vbNullString
		lblThresh(3).Text = "n/a"
		txtThresh(3).Visible = False
		txtThresh(3).Text = vbNullString
		
		chkSave.Visible = False
		
		'If thresh use is off, ignore
		If Not Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			Exit Sub
		End If
		
		'get threshhold setting from the registry
		If Len(GetSetting(APPNAME, "Thresholds", strNewDes & CStr(1) & " type", vbNullString)) = 0 Then
			SetDefaultThresholds(strNewDes)
		Else
			For n = 0 To 3
				thresh = GetSetting(APPNAME, "Thresholds", strNewDes & CStr(n + 1) & " type", vbNullString)
				If Len(thresh) > 0 And thresh <> "n/a" Then
					lblThresh(n).Text = thresh
					txtThresh(n).Visible = True
					txtThresh(n).Text = GetSetting(APPNAME, "Thresholds", strNewDes & CStr(n + 1) & " level", vbNullString)
				Else '110703 rjk: Removed the textbox when the label is blank or n/a
					txtThresh(n).Visible = False
				End If
			Next n
		End If
		
		bAutoFill = False
	End Sub
	
	Private Sub SetDefaultThresholds(ByRef strNewDes As String)
		'now turn on the thresholds based on designations
		Select Case strNewDes
			Case "m"
				lblThresh(0).Text = "iron"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "1"
			Case "k"
				lblThresh(0).Text = "iron"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "600"
				lblThresh(1).Text = "hcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "1"
			Case "j"
				lblThresh(0).Text = "iron"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "600"
				lblThresh(1).Text = "lcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "1"
			Case "g"
				lblThresh(0).Text = "dust"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "1"
			Case "i"
				lblThresh(0).Text = "lcm"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "300"
				lblThresh(1).Text = "hcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "150"
				lblThresh(2).Text = "shell"
				txtThresh(2).Visible = True
				txtThresh(2).Text = "1"
			Case "d"
				lblThresh(0).Text = "oil"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "10"
				lblThresh(1).Text = "lcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "50"
				lblThresh(2).Text = "hcm"
				txtThresh(2).Visible = True
				txtThresh(2).Text = "100"
				lblThresh(3).Text = "gun"
				txtThresh(3).Visible = True
				txtThresh(3).Text = "1"
			Case "o"
				lblThresh(0).Text = "oil"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "1"
			Case "f"
				lblThresh(0).Text = "hcm"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "100"
				lblThresh(1).Text = "gun"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "7"
				lblThresh(2).Text = "shell"
				txtThresh(2).Visible = True
				txtThresh(2).Text = "50"
				lblThresh(3).Text = "military"
				txtThresh(3).Visible = True
				txtThresh(3).Text = "60"
				
			Case "*"
				lblThresh(0).Text = "petrol"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "1000"
				lblThresh(1).Text = "lcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "200"
				lblThresh(2).Text = "hcm"
				txtThresh(2).Visible = True
				txtThresh(2).Text = "100"
				lblThresh(3).Text = "shell"
				txtThresh(3).Visible = True
				txtThresh(3).Text = "200"
			Case "!"
				lblThresh(0).Text = "lcm"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "200"
				lblThresh(1).Text = "hcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "100"
				lblThresh(2).Text = "gun"
				txtThresh(2).Visible = True
				txtThresh(2).Text = vbNullString
				lblThresh(3).Text = "shell"
				txtThresh(3).Visible = True
				txtThresh(3).Text = vbNullString
			Case "e"
				lblThresh(0).Text = "civs"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "768"
				lblThresh(1).Text = "mil"
				txtThresh(1).Visible = True
				txtThresh(1).Text = "118"
			Case "t"
				lblThresh(0).Text = "oil"
				txtThresh(0).Visible = True
				txtThresh(0).Text = vbNullString
				lblThresh(1).Text = "lcm"
				txtThresh(1).Visible = True
				txtThresh(1).Text = vbNullString
				lblThresh(2).Text = "dust"
				txtThresh(2).Visible = True
				txtThresh(2).Text = vbNullString
			Case "l"
				lblThresh(0).Text = "lcm"
				txtThresh(0).Visible = True
				txtThresh(0).Text = vbNullString
			Case "n"
				lblThresh(0).Text = "rad"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "999"
			Case "p"
				lblThresh(0).Text = "lcm"
				txtThresh(0).Visible = True
				txtThresh(0).Text = vbNullString
			Case "%"
				lblThresh(0).Text = "petrol"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "1"
				lblThresh(1).Text = "oil"
				txtThresh(1).Visible = True
				txtThresh(1).Text = vbNullString
			Case "b"
				lblThresh(0).Text = "dust"
				txtThresh(0).Visible = True
				txtThresh(0).Text = vbNullString
			Case "a"
				lblThresh(0).Text = "food"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "60"
			Case "u"
				lblThresh(0).Text = "rads"
				txtThresh(0).Visible = True
				txtThresh(0).Text = "1"
		End Select
	End Sub
	
	'UPGRADE_WARNING: Event txtThresh.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtThresh_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThresh.TextChanged
		Dim Index As Short = txtThresh.GetIndex(eventSender)
		If bAutoFill = True Or chkSave.Visible Then
			Exit Sub
		End If
		
		chkSave.Visible = True
	End Sub
End Class