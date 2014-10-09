Option Strict Off
Option Explicit On
Friend Class frmPromptSimpleTerritory
	Inherits System.Windows.Forms.Form
	
	'Change Log
	'250904 rjk: Created to provide simple interface for territory.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private Sub cmdTool_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTool.Click
		frmDrawMap.PromptForm = frmPromptTerritory
		'Put form in proper place
		frmDrawMap.PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmDrawMap.PromptForm.Width)) \ 2)
		frmDrawMap.PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmDrawMap.PromptForm.Height))
		'Dialog box will take it from here
		frmDrawMap.PromptForm.Show()
		Me.Close()
	End Sub
	
	Private Sub frmPromptSimpleTerritory_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		cbField.SelectedIndex = 0
	End Sub
	
	Private Sub frmPromptSimpleTerritory_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		SetMultipleSectorTerritory((txtMultOrigin.Text), CStr(VB6.GetItemData(cbField, cbField.SelectedIndex)), txtTerrNumber.Text)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
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
	
	'UPGRADE_WARNING: Form event frmPromptSimpleTerritory.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptSimpleTerritory_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtMultOrigin.Focus()
	End Sub
	
	
	Private Sub SetMultipleSectorTerritory(ByRef strMultOrigin As String, ByRef strTerr As String, ByRef strTerrNumber As String)
		Dim beginPos As Short
		Dim endPos As Short
		
		If InStr(strMultOrigin, "\") > 0 Then
			beginPos = 1
			endPos = InStr(strMultOrigin, "\")
			While endPos > 0
				SetTerritory(Mid(strMultOrigin, beginPos, endPos - beginPos), strTerr, strTerrNumber)
				beginPos = endPos + 1
				endPos = InStr(beginPos, strMultOrigin, "\")
			End While
			SetTerritory(Mid(strMultOrigin, beginPos), strTerr, strTerrNumber)
		Else
			SetTerritory(strMultOrigin, strTerr, strTerrNumber)
		End If
	End Sub
	
	Private Sub SetTerritory(ByRef strStart As String, ByRef strTerr As String, ByRef strTerrNumber As String)
		frmEmpCmd.SubmitEmpireCommand("territory " & strStart & " " & strTerrNumber & " " & strTerr, True)
	End Sub
End Class