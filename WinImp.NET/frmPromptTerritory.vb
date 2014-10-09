Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptTerritory
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: commented out dead Sub ToggleHex
	'091103 rjk: added back ToggleHex, used when marking territories
	'092703 rjk: General Reformatting
	'100103 rjk: Clear list when switching territories so lists do not get combined.
	'101703 rjk: Added Strength fields to Sector database
	'102303 rjk: Deleted lblDescSect, not used
	'            Added Field combobox for selecting between terr, terr1, terr2 and terr3
	'            Added batch markers so all the commands are combined on the command display
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Dim strTerr As String
	
	Public Sub ToggleHex(ByRef sx As Short, ByRef sy As Short, ByRef desc As String)
		Dim searchindex As Integer
		Dim n As Short
		
		'seargh for item - if found delete it and exit
		searchindex = (2048 + CInt(sx)) * &HFFF + (sy + 2048)
		For n = 0 To lstSectors.Items.Count - 1
			If VB6.GetItemData(lstSectors, n) = searchindex Then
				lstSectors.Items.RemoveAt(n)
				Exit Sub
			End If
		Next n
		
		'Since entry was not found, add it
		lstSectors.Items.Add(New VB6.ListBoxItem(SectorString(sx, sy) & vbTab & desc, searchindex))
	End Sub
	
	'UPGRADE_WARNING: Event cbField.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbField_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbField.SelectedIndexChanged
		FillEditTerritoryCombo()
	End Sub
	
	'UPGRADE_WARNING: Event cbTerr.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbTerr_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbTerr.SelectedIndexChanged
		Dim strTerrField As String '102303 rjk: Added for access to the different territory fields
		
		lstSectors.Items.Clear() '100103 rjk: Clear list when switching territories so list do not get combined.
		
		strTerrField = GetTerritoryFieldDatabaseName '102303 rjk: Added for access to the different territory fields
		frmDrawMap.StartTerritory()
		rsSectors.MoveFirst()
		While Not rsSectors.EOF
			If rsSectors.Fields(strTerrField).Value = VB6.GetItemString(cbTerr, cbTerr.SelectedIndex) Then
				'draw hex in red
				frmDrawMap.ToggleTerritory(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
				
				'add to list
				'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(rsSectors.Fields(des).Value). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lstSectors.Items.Add(New VB6.ListBoxItem(SectorString((rsSectors.Fields("x").Value), (rsSectors.Fields("y").Value)) & vbTab & CStr(rsSectors.Fields("eff").Value) & "% " + colDes.Item(rsSectors.Fields("des").Value), (2048 + CInt(rsSectors.Fields("x").Value)) * &HFFF + rsSectors.Fields("y").Value + 2048))
			End If
			rsSectors.MoveNext()
		End While
	End Sub
	
	Private Sub cmdAddNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddNew.Click
		Dim n As Short
		
		'find first unused territory
		
		n = 1
		While InStr(strTerr, " " & CStr(n) & " ") > 0
			n = n + 1
		End While
		
		'now add that territory to the box
		strTerr = strTerr & " " & CStr(n) & " "
		cbTerr.Items.Add(CStr(n))
		'UPGRADE_ISSUE: ComboBox property cbTerr.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
		cbTerr.SelectedIndex = cbTerr.NewIndex
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim n As Short
		
		'102303 rjk: Added so all the commands are combined on the command display
		frmEmpCmd.SubmitEmpireCommand("bf1", False)
		
		'first wipe out the exsisting territory
		'102303 rjk: Modified for Multiple Territory field support
		'strCmd = "territory * ?terr=" + CStr(cbTerr.List(cbTerr.ListIndex)) + " 0"
		strCmd = "territory * ?terr" & CStr(VB6.GetItemData(cbField, cbField.SelectedIndex)) & "=" & CStr(VB6.GetItemString(cbTerr, cbTerr.SelectedIndex)) & " 0 " & CStr(VB6.GetItemData(cbField, cbField.SelectedIndex))
		
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		
		'then add the new one
		For n = 0 To lstSectors.Items.Count - 1
			strCmd = VB.Left(VB6.GetItemString(lstSectors, n), InStr(VB6.GetItemString(lstSectors, n), vbTab) - 1)
			strCmd = "territory " & strCmd & " " & CStr(VB6.GetItemString(cbTerr, cbTerr.SelectedIndex))
			'102303 rjk: Modified for Multiple Territory field support
			strCmd = strCmd & " " & CStr(VB6.GetItemData(cbField, cbField.SelectedIndex))
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		Next n
		'102303 rjk: Added so all the commands are combined on the command display
		frmEmpCmd.SubmitEmpireCommand("bf2", False)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Private Sub frmPromptTerritory_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long       removed 3/2003 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		cbField.SelectedIndex = 0
		FillEditTerritoryCombo() '102303 rjk: Code moved to separate function to accomodate multiple territory fields
	End Sub
	
	Private Sub frmPromptTerritory_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		frmDrawMap.MarkingTerritory = False
		frmDrawMap.DrawHexes()
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'102303 rjk: Added for access to the different territory fields
	Private Function GetTerritoryFieldDatabaseName() As String
		If cbField.SelectedIndex = 0 Then
			GetTerritoryFieldDatabaseName = "terr"
		Else
			GetTerritoryFieldDatabaseName = "terr" & CStr(VB6.GetItemData(cbField, cbField.SelectedIndex))
		End If
	End Function
	
	'102303 rjk: Moved the filling of edit territory combo box to separate so it can also be called
	'            cbField when difference territory field is selected.
	Private Sub FillEditTerritoryCombo()
		Dim strTerrField As String
		
		strTerrField = GetTerritoryFieldDatabaseName
		strTerr = vbNullString
		rsSectors.MoveFirst()
		cbTerr.Items.Clear()
		lstSectors.Items.Clear()
		While Not rsSectors.EOF
			If rsSectors.Fields(strTerrField).Value <> 0 Then
				If InStr(strTerr, " " & CStr(rsSectors.Fields(strTerrField).Value) & " ") = 0 Then
					strTerr = strTerr & " " & CStr(rsSectors.Fields(strTerrField).Value) & " "
					cbTerr.Items.Add(CStr(rsSectors.Fields(strTerrField).Value))
				End If
			End If
			rsSectors.MoveNext()
		End While
	End Sub
End Class