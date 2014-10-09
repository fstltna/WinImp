Option Strict Off
Option Explicit On
Friend Class frmUnitView
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'112703 rjk: Created
	'120303 rjk: Moved friendly unit to frmOption
	'            Changed Apply to Test and the changes are not saved.
	'140704 rjk: Added Expired Units.
	'290606 rjk: Added Our Nuke Units.
	
	Public unitView As New EmpUnitView
	
	'UPGRADE_WARNING: Event chkUnits.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkUnits_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUnits.CheckStateChanged
		Dim Index As Short = chkUnits.GetIndex(eventSender)
		UpdatePriorityList()
	End Sub
	
	Private Sub UpdatePriorityList()
		Dim chkIndex As EmpUnitView.enumViewingUnits
		Dim sortIndex As Short
		'UPGRADE_NOTE: location was upgraded to location_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim location_Renamed As Short
		Dim strUnitName As String
		
		For chkIndex = 0 To unitView.lastViewingUnit
			location_Renamed = -1
			For sortIndex = 0 To lstSort.Items.Count - 1
				If VB6.GetItemData(lstSort, sortIndex) = chkIndex Then
					location_Renamed = sortIndex
					Exit For
				End If
			Next sortIndex
			If chkUnits(chkIndex).CheckState <> System.Windows.Forms.CheckState.Checked Then
				If location_Renamed <> -1 Then
					lstSort.Items.RemoveAt((location_Renamed))
				End If
			Else
				If location_Renamed = -1 Then
					Select Case chkIndex
						Case 0, 1, 2, 3
							strUnitName = "Our "
						Case 4, 5, 6
							strUnitName = "Enemy "
						Case 7, 8, 9
							strUnitName = "Allied "
						Case 10, 11, 12
							strUnitName = "Neutral "
						Case 13, 14, 15
							strUnitName = "Expired "
					End Select
					lstSort.Items.Insert(lstSort.Items.Count, New VB6.ListBoxItem(strUnitName & chkUnits(chkIndex).Text, chkIndex))
					'UPGRADE_ISSUE: ListBox property lstSort.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					lstSort.SelectedIndex = lstSort.NewIndex
				End If
			End If
		Next chkIndex
		If lstSort.SelectedIndex >= lstSort.Items.Count Then
			lstSort.SelectedIndex = lstSort.Items.Count - 1
		End If
		If lstSort.SelectedIndex = lstSort.Items.Count - 1 Then
			cmdDown.Enabled = False
		Else
			cmdDown.Enabled = True
		End If
		If lstSort.SelectedIndex = 0 Then
			cmdUp.Enabled = False
		Else
			cmdUp.Enabled = True
		End If
	End Sub
	
	Private Sub cmdApply_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdApply.Click
		Apply()
		frmDrawMap.DrawHexes()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		unitView.Load()
		Me.Close()
	End Sub
	
	Private Sub cmdClearAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearAll.Click
		Dim Index As EmpUnitView.enumViewingUnits
		
		For Index = 0 To 11
			chkUnits(Index).CheckState = System.Windows.Forms.CheckState.Unchecked
		Next Index
	End Sub
	
	Private Sub cmdDown_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDown.Click
		Dim strUnitName As String
		Dim Data As Short
		Dim Index As EmpUnitView.enumViewingUnits
		
		strUnitName = VB6.GetItemString(lstSort, lstSort.SelectedIndex)
		Data = VB6.GetItemData(lstSort, lstSort.SelectedIndex)
		Index = lstSort.SelectedIndex
		lstSort.Items.RemoveAt(Index)
		lstSort.Items.Insert(Index + 1, New VB6.ListBoxItem(strUnitName, Data))
		lstSort.SelectedIndex = Index + 1
	End Sub
	
	Private Sub Apply()
		Dim Index As EmpUnitView.enumViewingUnits
		
		For Index = 0 To unitView.lastViewingUnit
			If chkUnits(Index).CheckState = System.Windows.Forms.CheckState.Unchecked Then
				unitView.bViewUnits(Index) = False
			Else
				unitView.bViewUnits(Index) = True
			End If
		Next Index
		If optFriendlyAlly.Checked Then
			frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_ALLIED
		Else
			frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_NEUTRAL
		End If
		For Index = 0 To lstSort.Items.Count - 1
			unitView.priorityList(Index) = VB6.GetItemData(lstSort, Index)
		Next Index
		unitView.priorityListCount = lstSort.Items.Count
		unitView.iExpiry = VB6.GetItemData(cbExpiry, cbExpiry.SelectedIndex)
		unitView.Save()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Apply()
		unitView.Save()
		frmOptions.SaveOptions()
		Me.Close()
	End Sub
	
	Private Sub cmdSetAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSetAll.Click
		Dim Index As EmpUnitView.enumViewingUnits
		
		For Index = 0 To unitView.lastViewingUnit
			chkUnits(Index).CheckState = System.Windows.Forms.CheckState.Checked
		Next Index
	End Sub
	
	Private Sub cmdUp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUp.Click
		Dim strUnitName As String
		Dim Data As Short
		Dim Index As Short
		
		strUnitName = VB6.GetItemString(lstSort, lstSort.SelectedIndex)
		Data = VB6.GetItemData(lstSort, lstSort.SelectedIndex)
		Index = lstSort.SelectedIndex
		lstSort.Items.RemoveAt(Index)
		lstSort.Items.Insert(Index - 1, New VB6.ListBoxItem(strUnitName, Data))
		lstSort.SelectedIndex = Index - 1
	End Sub
	
	Private Sub frmUnitView_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Index As EmpUnitView.enumViewingUnits
		
		UpdatePriorityList()
		lstSort.SelectedIndex = 0
		For Index = 0 To unitView.lastViewingUnit
			If unitView.bViewUnits(Index) Then
				chkUnits(Index).CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkUnits(Index).CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
		Next Index
		If frmOptions.friendlyUnit = frmOptions.enumFriendly.FRIEND_ALLIED Then
			optFriendlyAlly.Checked = True
		Else
			optFriendlyNeutral.Checked = True
		End If
		For Index = 0 To cbExpiry.Items.Count
			If VB6.GetItemData(cbExpiry, Index) = unitView.iExpiry Then
				cbExpiry.SelectedIndex = Index
				Exit For
			End If
		Next Index
	End Sub
	
	Private Sub frmUnitView_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		frmDrawMap.DrawHexes()
	End Sub
	
	'UPGRADE_WARNING: Event lstSort.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstSort_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstSort.SelectedIndexChanged
		UpdatePriorityList()
	End Sub
End Class