Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptThresh
	Inherits System.Windows.Forms.Form
	
	'Change Log
	'270903 rjk: General reformatting
	'171003 rjk: Added Strength fields to Sector database
	'191003 rjk: Added Multiple Sector Selection
	'231003 rjk: Added Sector count for Multiple Sectors
	'            Added sector description field
	'201103 rjk: Added option to control strength updates
	'251103 rjk: Added the ability to set the food levels based on food required
	'240104 rjk: Renamed the Supply check, switch to item object, added multiple sector selection,
	'            Removed the field transaction table, update the registry saving (add food required)
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private ftranslation(13) As String
	
	'UPGRADE_WARNING: Event cbCom.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbCom_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbCom.SelectedIndexChanged
		SetThreshLabel()
	End Sub
	
	Private Sub cbCom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cbCom.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim i As Short
		Dim ch As String
		On Error Resume Next
		
		ch = LCase(Chr(KeyAscii))
		If ch >= "a" And ch < "z" Then
			For i = 0 To cbCom.Items.Count - 1
				If VB.Left(VB6.GetItemString(cbCom, i), 1) = ch Then
					cbCom.SelectedIndex = i
				End If
			Next i
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'112503 rjk: Added the ability to set the food levels based on food required
	'UPGRADE_WARNING: Event chkFoodRequired.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkFoodRequired_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkFoodRequired.CheckStateChanged
		SetThreshLabel()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim iComm As EmpItem
		
		iComm = Items.FindByFormName((cbCom.Text))
		
		SetMultipleSectorThreshold(iComm, (txtMultOrigin.Text), Val(txtNum.Text), chkFoodRequired.CheckState <> System.Windows.Forms.CheckState.Unchecked, chkSupply.CheckState <> System.Windows.Forms.CheckState.Unchecked)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'frmEmpCmd.SubmitEmpireCommand "dump " + txtMultOrigin.Text, False
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Public Sub SetMultipleSectorThreshold(ByRef iComm As EmpItem, ByRef strMultOrigin As String, ByRef lThres As Integer, ByRef bFoodRequired As Boolean, ByRef bSupply As Boolean)
		Dim beginPos As Short
		Dim endPos As Short
		
		If InStr(strMultOrigin, "\") > 0 Then '101903 Multiple Sectors Selected
			beginPos = 1
			endPos = InStr(strMultOrigin, "\")
			While endPos > 0
				SetThreshold(Mid(strMultOrigin, beginPos, endPos - beginPos), iComm, lThres, bFoodRequired, bSupply)
				beginPos = endPos + 1
				endPos = InStr(beginPos, strMultOrigin, "\")
			End While
			SetThreshold(Mid(strMultOrigin, beginPos), iComm, lThres, bFoodRequired, bSupply)
		Else
			SetThreshold(strMultOrigin, iComm, lThres, bFoodRequired, bSupply)
		End If
	End Sub
	
	Private Sub SetThreshold(ByRef strStart As String, ByRef iComm As EmpItem, ByRef lThresh As Integer, ByRef bFoodRequired As Boolean, ByRef bSupply As Boolean)
		'112503 rjk: Added the ability to set the food levels based on food required
		If iComm.Letter = "f" And bFoodRequired Then
			AdjustedFoodLevel(strStart, lThresh)
		End If
		frmEmpCmd.SubmitEmpireCommand("threshold " & iComm.Letter & " " & strStart & " " & CStr(lThresh), True)
		
		If bSupply Then
			MeetThreshold(strStart, lThresh, iComm)
		End If
	End Sub
	
	Public Sub MeetThreshold(ByRef txtOrigin As String, ByRef lNewThresh As Integer, ByRef iComm As EmpItem)
		Dim DistX, sx, sy, DistY As Short
		Dim lShortage As Integer
		Dim strCmd As String
		
		ParseSectors(sx, sy, txtOrigin)
		rsSectors.Seek("=", sx, sy)
		lShortage = lNewThresh - iComm.DatabaseValue(rsSectors)
		
		DistX = rsSectors.Fields("dist_x").Value
		DistY = rsSectors.Fields("dist_y").Value
		
		rsSectors.Seek("=", DistX, DistY)
		
		If Not rsSectors.NoMatch Then
			If (lShortage > 0) Then
				lShortage = Min(lShortage, iComm.DatabaseValue(rsSectors))
				strCmd = "move " & iComm.Letter & "  " & SectorString(DistX, DistY) & " " & CStr(lShortage) & " " & txtOrigin
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			ElseIf (lShortage < 0) Then 
				strCmd = "move " & iComm.Letter & "  " & txtOrigin & " " & CStr(0 - lShortage) & " " & SectorString(DistX, DistY)
				frmEmpCmd.SubmitEmpireCommand(strCmd, False)
			End If
		End If
	End Sub
	
	Private Sub frmPromptThresh_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		LoadCommodityBox(cbCom)
		cbCom.SelectedIndex = 1 'set to civilians
		
		'Make Stay always on top
		SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		If CBool(GetSetting(APPNAME, "frmPromptThresh", "Food Required", CStr(False))) Then
			chkFoodRequired.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkFoodRequired.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If CBool(GetSetting(APPNAME, "frmPromptThresh", "Supply", CStr(False))) Then
			chkSupply.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSupply.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
	End Sub
	
	Private Sub frmPromptThresh_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		
		If chkFoodRequired.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptThresh", "Food Required", CStr(True))
		Else
			SaveSetting(APPNAME, "frmPromptThresh", "Food Required", CStr(False))
		End If
		If chkSupply.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptThresh", "Supply", CStr(True))
		Else
			SaveSetting(APPNAME, "frmPromptThresh", "Supply", CStr(False))
		End If
	End Sub
	
	Private Sub txtNum_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNum.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		cbCom_KeyPress(cbCom, New System.Windows.Forms.KeyPressEventArgs(Chr(KeyAscii)))
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtMultOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.TextChanged
		SetThreshLabel()
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
	
	Private Sub SetThreshLabel()
		Dim n As Short
		Dim sx As Short
		Dim sy As Short
		' Dim owner As Integer    8/2003 efj  removed
		On Error GoTo SetThreshLabel_Exit
		
		lblThresh.Text = vbNullString
		If ParseSectors(sx, sy, txtMultOrigin.Text) Then
			
			'Get Sector Information
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				n = rsSectors.Fields(VB.Left(VB6.GetItemString(cbCom, cbCom.SelectedIndex), 1) & "_dist").Value
				lblThresh.Text = "Current Threshold: " & CStr(n)
				'102303 rjk: Added sector label
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
		
		'112503 rjk: Added the ability to set the food levels based on food required
		If VB6.GetItemString(cbCom, cbCom.SelectedIndex) = "food" And InStr(txtMultOrigin.Text, "*") = 0 And InStr(txtMultOrigin.Text, "#") = 0 And InStr(txtMultOrigin.Text, ":") = 0 Then
			chkFoodRequired.Visible = True
			If chkFoodRequired.CheckState = System.Windows.Forms.CheckState.Checked Then
				lblFoodPercent.Visible = True
				Label2.Visible = False
			Else
				lblFoodPercent.Visible = False
				Label2.Visible = True
			End If
		Else
			chkFoodRequired.Visible = False
			lblFoodPercent.Visible = False
			Label2.Visible = True
		End If
SetThreshLabel_Exit: 
	End Sub
	
	'112503 rjk: Added the ability to set the food levels based on food required
	Public Sub AdjustedFoodLevel(ByRef strStart As String, ByRef lThresh As Integer)
		Dim sx As Short
		Dim sy As Short
		
		If ParseSectors(sx, sy, strStart) Then
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				lThresh = CInt(CStr(System.Math.Round(FoodRequired(rsSectors) * (1 + (lThresh / 100)), 0)))
			End If
		End If
	End Sub
End Class