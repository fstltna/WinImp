Option Strict Off
Option Explicit On
Friend Class frmPromptOrder
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092703 rjk: general reformatting
	'110203 rjk: Changed the Cancel caption to Close
	'111603 rjk: Code Cleanup - remove the attempt to incorporate the order output into the form?
	'            Made strCmd a local variable. Made cmdOK_Click private.
	'112503 rjk: Added Food threshold for AutoFeed.
	'180104 rjk: Changed from orders to order for set(OK) command.
	'111104 rjk: Added the filling of the Order from the database.
	'131207 rjk: Added end cargo fields to Ship orders, rename cargo to start cargo
	'010108 rjk: Add GetOrders to update form after commands,
	'            make OrderLoadData public to allow data updates.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click 'Close
		Me.Close()
	End Sub
	
	Private Sub cmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClear.Click
		Dim strCmd As String
		'Command: order <SHIP/FLEET> c Clear Orders
		strCmd = "order " & txtUnit.Text & " c"
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		GetOrders(True)
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim strCmd2 As String
		Dim cargohold As Short
		Dim sx As Short
		Dim sy As Short
		
		strCmd = "order " & txtUnit.Text '180104 rjk: Changed from orders to order
		
		'Declares: order <SHIP/FLEET> d <dest1> [dest2|scuttle|-] Declare Orders
		frmEmpCmd.SubmitEmpireCommand("bf1", False) 'run as batch file
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			strCmd2 = strCmd & " d " & txtOrigin.Text
			If ParseSectors(sx, sy, (txtPath.Text)) Then
				strCmd2 = strCmd2 & " " & txtPath.Text
			Else
				If chkAutoScuttle.CheckState = System.Windows.Forms.CheckState.Checked Then
					strCmd2 = strCmd2 & " scuttle"
				End If
			End If
			frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
		End If
		
		'check autofeed
		If chkAutoFeed.CheckState = System.Windows.Forms.CheckState.Checked Then
			strCmd2 = strCmd & " l 1 start food " & txtAutoFeed.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			strCmd2 = strCmd & " l 1 end food " & txtAutoFeed.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
		End If
		
		'Cargo levels: Command: order <SHIP/FLEET> l <hold> <start/end> <COMM> <amount>
		For cargohold = 1 To 6
			If Len(Trim(txtStartCargo(cargohold).Text)) > 0 Then
				strCmd2 = strCmd & " l " & CStr(cargohold) & " start " & txtStartCargo(cargohold).Text & " " & txtStart(cargohold).Text
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			End If
			If Len(Trim(txtEndCargo(cargohold).Text)) > 0 Then
				strCmd2 = strCmd & " l " & CStr(cargohold) & " end " & txtEndCargo(cargohold).Text & " " & txtEnd(cargohold).Text
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			End If
		Next cargohold
		frmEmpCmd.SubmitEmpireCommand("bf2", False) 'run as batch file
		GetOrders(True)
	End Sub
	
	Private Sub cmdResume_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdResume.Click
		Dim strCmd As String
		'Command: order <SHIP/FLEET> r Resume Orders
		strCmd = "order " & txtUnit.Text & " r"
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		GetOrders(True)
	End Sub
	
	Private Sub cmdShow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShow.Click
		Dim strCmd As String
		Dim strUnit As String
		strUnit = txtUnit.Text
		
		If Len(txtUnit.Text) = 0 Then
			strUnit = "*"
		End If
		
		frmEmpCmd.SubmitEmpireCommand("br1", True)
		strCmd = "sorder " & strUnit
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		strCmd = "qorder " & strUnit
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		frmEmpCmd.SubmitEmpireCommand("br2", True)
	End Sub
	
	Private Sub cmdSuspend_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSuspend.Click
		Dim strCmd As String
		'Command: order <SHIP/FLEET> s Suspend Orders
		strCmd = "order " & txtUnit.Text & " s"
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		GetOrders(True)
	End Sub
	
	Private Sub frmPromptOrder_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		OrderClearForm()
	End Sub
	
	Private Sub frmPromptOrder_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		
		On Error Resume Next
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			txtPath.Focus()
		End If
	End Sub
	
	Public Sub OrderLoadData()
		Dim hldIndex As String
		Dim iCargoHold As Short
		Dim strErrorField As String
		
		If InStr(txtUnit.Text, "/") Or Len(txtUnit.Text) <= 0 Or Not IsNumeric(txtUnit.Text) Then
			OrderClearForm()
			Exit Sub
		End If
		
		hldIndex = rsShipOrders.Index
		
		On Error GoTo OrderLoadData_error
		
		rsShipOrders.Seek("=", txtUnit.Text)
		If Not rsShipOrders.NoMatch Then
			strErrorField = "start sector"
			txtOrigin.Text = rsShipOrders.Fields("start sector").Value
			strErrorField = "end sector"
			txtPath.Text = rsShipOrders.Fields("end sector").Value
			For iCargoHold = 1 To 6
				strErrorField = "start cargo " & CStr(iCargoHold)
				txtStartCargo(iCargoHold).Text = rsShipOrders.Fields("start cargo " & CStr(iCargoHold)).Value
				strErrorField = "end cargo " & CStr(iCargoHold)
				txtEndCargo(iCargoHold).Text = rsShipOrders.Fields("end cargo " & CStr(iCargoHold)).Value
				strErrorField = "start level " & CStr(iCargoHold)
				txtStart(iCargoHold).Text = rsShipOrders.Fields("start level " & CStr(iCargoHold)).Value
				strErrorField = "end level " & CStr(iCargoHold)
				txtEnd(iCargoHold).Text = rsShipOrders.Fields("end level " & CStr(iCargoHold)).Value
			Next iCargoHold
		Else
			OrderClearForm()
		End If
		
		rsShipOrders.Index = hldIndex
		Exit Sub
		
OrderLoadData_error: 
		EmpireError("OrderLoadData", "Database error:", strErrorField)
		rsShipOrders.Index = hldIndex
	End Sub
	
	Private Sub OrderClearForm()
		Dim iCargoHold As Short
		
		txtOrigin.Text = ""
		txtPath.Text = ""
		For iCargoHold = 1 To 6
			txtStartCargo(iCargoHold).Text = ""
			txtEndCargo(iCargoHold).Text = ""
			txtStart(iCargoHold).Text = ""
			txtEnd(iCargoHold).Text = ""
		Next iCargoHold
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		OrderLoadData()
	End Sub
End Class