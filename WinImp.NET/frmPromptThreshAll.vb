Option Strict Off
Option Explicit On
Friend Class frmPromptThreshAll
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: commented out dead Sub SetThresh
	'081103 efj: removal of dead variables
	'090703 rjk: SetThresh is used in the ThresholdAll function (right click or Ctrl 't')
	'091103 rjk: Change unc. workers to use word wrap.
	'091803 rjk: Moved label setting functionality to SetLabels. Call SetLabels when origin changes
	'092703 rjk: General reformatting.
	'093003 rjk: switched to x and y derived from txtOrigin instead of SelX and SelY
	'093003 rjk: Added to set initial field selection
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'240104 rjk: Renamed the Supply check, switched to item object, added multiple sector selection
	'            Removed Field Transaction table, Update registry saving, added food required option,
	'            Reorganized form.
	'280104 rjk: Added double-click option.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Dim arythresh(13) As Short
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp("threshold")
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim n As Short
		Dim strCmd As String
		On Error Resume Next
		Dim iComm As EmpItem
		Dim lThresh As Integer
		
		For n = 0 To 13
			iComm = Items.FindByFormLabel(lblThresh(n).Text)
			lThresh = Val(txtThresh(n).Text)
			If lThresh <> arythresh(n) Then
				frmPromptThresh.SetMultipleSectorThreshold(iComm, (txtMultOrigin.Text), lThresh, (chkFoodRequired.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkSupply.CheckState <> System.Windows.Forms.CheckState.Unchecked))
			End If
		Next n
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptThreshAll.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptThreshAll_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtMultOrigin.Focus() '093003 rjk: Added to set initial field selection
	End Sub
	
	Private Sub frmPromptThreshAll_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long   removed efj 8/2003
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		'SetLabels '092103 rjk: Moved code to SetLabels.
		
		'Put form in proper place
		Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
		
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
	
	Private Sub frmPromptThreshAll_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	
	'UPGRADE_WARNING: Event txtMultOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.TextChanged
		SetLabels()
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
	
	Private Sub txtThresh_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThresh.Enter
		Dim Index As Short = txtThresh.GetIndex(eventSender)
		If Len(txtThresh(Index).Text) > 0 Then
			txtThresh(Index).SelectionStart = 0
			txtThresh(Index).SelectionLength = Len(txtThresh(Index).Text)
		End If
	End Sub
	
	Private Sub SetLabels()
		Dim i As Short
		Dim X As Short
		Dim Y As Short
		
		lblThresh(0).Text = Items.FindByLetter("m").FormLabel
		lblThresh(1).Text = Items.FindByLetter("g").FormLabel
		lblThresh(2).Text = Items.FindByLetter("s").FormLabel
		lblThresh(3).Text = Items.FindByLetter("f").FormLabel
		lblThresh(4).Text = Items.FindByLetter("c").FormLabel
		lblThresh(5).Text = Items.FindByLetter("u").FormLabel
		lblThresh(6).Text = Items.FindByLetter("i").FormLabel
		lblThresh(7).Text = Items.FindByLetter("l").FormLabel
		lblThresh(8).Text = Items.FindByLetter("h").FormLabel
		lblThresh(9).Text = Items.FindByLetter("o").FormLabel
		lblThresh(10).Text = Items.FindByLetter("d").FormLabel
		lblThresh(11).Text = Items.FindByLetter("p").FormLabel
		lblThresh(12).Text = Items.FindByLetter("r").FormLabel
		lblThresh(13).Text = Items.FindByLetter("b").FormLabel
		
		If ParseSectors(X, Y, (txtMultOrigin.Text)) And InStr(txtMultOrigin.Text, "\") = 0 Then
			rsSectors.Seek("=", X, Y)
			
			For i = 0 To 13
				If Not rsSectors.NoMatch Then
					arythresh(i) = rsSectors.Fields(Items.FindByFormLabel(lblThresh(i).Text).Letter & "_dist").Value
					txtThresh(i).Text = CStr(arythresh(i))
				Else
					arythresh(i) = 0
					txtThresh(i).Text = ""
				End If
			Next i
		Else '093003 rjk: Otherwise blank values
			For i = 0 To 13
				arythresh(i) = 0
				txtThresh(i).Text = ""
			Next i
		End If
	End Sub
End Class