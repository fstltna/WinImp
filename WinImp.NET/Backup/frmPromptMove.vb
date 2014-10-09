Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptMove
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: added Option Explicit and removed dead variables
	'081203 efj: added missing variable definitions
	'090703 rjk: ch needs to be a string not a Integer for cbCom_KeyPress
	'092803 rjk: General reformatting. Added a check to make sure a commodity was select
	'100203 rjk: Erase the sector description when not avaliable or valid
	'            Moved the default commodity for deliver to the form.
	'100903 rjk: Reduce the civilian value by one so the sector is not abandoned when double clicking.
	'            Added bulk move capability in both directions, one sector to many and many to one sector.
	'            The many is selected by the condition set in the display mode.
	'            You can not test in the cond move mode and it will not compute costs either.
	'            The commands are batched up.
	'101703 rjk: Added Strength fields to Sector database
	'101903 rjk: Removed Conditional Move code, replaced with Multiple Sector Selection code
	'102303 rjk: Modified Move to compute the desired quantity based on the following logic
	'            ### moves that many
	'            +### moves as many as needed to result in ### total in the To sector
	'            -### leaves ### behind in the From sector and moves the excess
	'            Added sector count to Sector Label when in Multiple Sector Selection Mode
	'103103 rjk: Changed the initial field to txtNum field.
	'111203 rjk: Fixed the moving of uw with -/+ logic.
	'111403 rjk: Set the Move key to be the default key.
	'111803 rjk: Moved DisplayPath to frmDrawMap
	'            Moved GetCommodityValueFromDB to EmpDataBase.bas
	'112003 rjk: Added option to control strength updates
	'121303 rjk: Switch from GetCommodityValueFromDB to Item.DatabaseValue and Item class
	'240104 rjk: Switched to use item object and single letter for commodities, fixed the selecting of the default commodity.
	'270104 rjk: Change the selection of the default commodity for deliver to new sector designation
	'180304 rjk: Fixed selecting the default to deal with sdes being a space
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'230704 rjk: Added Radius to Title for delivery for StarWars.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'301007 rjk: Rearrange to make more room for all long paths, fixed "from" label
	
	'This static var contains the last commodity moved so the box
	'can be reset to that commodity when it pops up again
	Private LastCommodityMoved As String
	Private bFirstLoad As Boolean
	
	Dim strAmount As String
	
	'UPGRADE_WARNING: Event cbCom.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbCom_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbCom.SelectedIndexChanged
		
		SetCurrentQuantityLabel()
		If txtNum.Visible Then
			txtNum.Focus()
		End If
		ComputePathCost()
	End Sub
	
	Private Sub SetCurrentQuantityLabel()
		Label4.Text = VB6.GetItemString(cbCom, cbCom.SelectedIndex)
		
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim sx As Short
		Dim sy As Short
		On Error Resume Next
		
		If ParseSectors(sx, sy, (txtMultOrigin.Text)) Then
			'Get Sector Information
			If LCase(Trim(Label2.Text)) = "move" Then
				rsSectors.Seek("=", sx, sy)
				If Not rsSectors.NoMatch Then
					'102303 rjk: Moved common code a function
					strAmount = CStr(Items.FindByFormName((Label4.Text)).DatabaseValue(rsSectors))
					'100903 rjk: Suppress the avail when for cond mode: many to one
					'            Changed to total avail for cond mode: one to many
					'101903 rjk: Changed to work for Multiple Sector Selection Instead
					'If chkCondMove = vbUnchecked Then
					If InStr(txtMultPath.Text, "\") = 0 Then
						Label4.Text = Label4.Text & " [" & strAmount & " avail]"
						'ElseIf optCondMoveTo.Value = True Then
					Else
						Label4.Text = Label4.Text & " [" & strAmount & " tot. avail]"
					End If
				End If
			Else
				'must be a delivery - put in cutoff
				Label4.Text = "cutoff for " & Label4.Text
			End If
		End If
	End Sub
	Private Sub cbCom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cbCom.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim ch As String
		Dim i As Short
		
		On Error Resume Next
		ch = LCase(Chr(KeyAscii))
		If ch >= "a" And ch < "z" Then
			i = FindByLetterCommodityBox(ch)
			If i <> -1 Then
				cbCom.SelectedIndex = i
			End If
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Function FindByLetterCommodityBox(ByRef ch As String) As Short
		Dim i As Short
		
		For i = 0 To cbCom.Items.Count - 1
			If VB.Left(VB6.GetItemString(cbCom, i), 1) = VB.Left(ch, 1) Then
				FindByLetterCommodityBox = i
				Exit Function
			End If
		Next i
		FindByLetterCommodityBox = -1
	End Function
	
	'UPGRADE_WARNING: Event chkDisplayPath.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDisplayPath_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDisplayPath.CheckStateChanged
		ComputePathCost()
	End Sub
	
	'UPGRADE_WARNING: Event chkShowAll.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkShowAll_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkShowAll.CheckStateChanged
		txtMultOrigin_TextChanged(txtMultOrigin, New System.EventArgs())
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strItem As String
		Dim beginPos As Short
		Dim endPos As Short
		
		'092103 rjk: Added check to ensure commodity has been selected
		If Len(cbCom.Text) = 0 Then
			MsgBox("No Commodity selected")
			Exit Sub
		End If
		
		strItem = Items.FindByFormName((cbCom.Text)).Letter
		
		'101903 rjk: Rework logic for Multiple Sector Selection, also supports function SetDeliveryRoute and MoveCommodity
		Select Case Label2.Text
			Case "Deliver"
				If InStr(txtMultOrigin.Text, "\") > 0 Then '101803 Multiple Sectors Selected
					beginPos = 1
					endPos = InStr(txtMultOrigin.Text, "\")
					While endPos > 0
						If Not SetDeliveryRoute(Mid(txtMultOrigin.Text, beginPos, endPos - beginPos), (txtMultPath.Text), strItem, (txtNum.Text)) Then
							Exit Sub
						End If
						beginPos = endPos + 1
						endPos = InStr(beginPos, txtMultOrigin.Text, "\")
					End While
					If Not SetDeliveryRoute(Mid(txtMultOrigin.Text, beginPos), (txtMultPath.Text), strItem, (txtNum.Text)) Then
						Exit Sub
					End If
				Else
					If Not SetDeliveryRoute((txtMultOrigin.Text), (txtMultPath.Text), strItem, (txtNum.Text)) Then
						Exit Sub
					End If
				End If
			Case "Move"
				'Save the last commodity moved for next time
				LastCommodityMoved = VB6.GetItemString(cbCom, cbCom.SelectedIndex)
				If InStr(txtMultOrigin.Text, "\") > 0 Then '101803 Multiple Sectors Selected
					beginPos = 1
					endPos = InStr(txtMultOrigin.Text, "\")
					While endPos > 0
						MoveCommodity(Mid(txtMultOrigin.Text, beginPos, endPos - beginPos), (txtMultPath.Text), strItem, (txtNum.Text), False)
						beginPos = endPos + 1
						endPos = InStr(beginPos, txtMultOrigin.Text, "\")
					End While
					MoveCommodity(Mid(txtMultOrigin.Text, beginPos), (txtMultPath.Text), strItem, (txtNum.Text), False)
				Else
					MoveCommodity((txtMultOrigin.Text), (txtMultPath.Text), strItem, (txtNum.Text), False)
				End If
		End Select
		
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	'101903 rjk: Added for Multiple Sector Selection
	Private Function SetDeliveryRoute(ByRef strStart As String, ByRef strPath As String, ByRef strItem As String, ByRef strThreshold As String) As Object
		Dim sx As Short
		Dim sy As Short
		Dim sx2 As Short
		Dim sy2 As Short
		
		If ParseSectors(sx, sy, strPath) Then
			If ParseSectors(sx2, sy2, strStart) Then
				If sx = sx2 And sy = sy2 Then
					strPath = "h" 'Same hex means cancel
				Else
					strPath = EmpirePathDirection(sx - sx2, sy - sy2)
					If Len(strPath) > 0 Then
						If VB.Right(strPath, 1) = "h" Then
							strPath = VB.Left(strPath, Len(strPath) - 1)
						End If
						If Len(strPath) <> 1 Then
							MsgBox("Must choose an adjacent sector for deliver")
							'UPGRADE_WARNING: Couldn't resolve default property of object SetDeliveryRoute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							SetDeliveryRoute = False
							Exit Function
						End If
					End If
				End If
			End If
		End If
		frmEmpCmd.SubmitEmpireCommand("deliver " & strItem & " " & strStart & " " & strThreshold & " " & strPath, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetDeliveryRoute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetDeliveryRoute = True
	End Function
	
	'101903 rjk: Added for Multiple Sector Selection
	Public Sub MoveCommodity(ByRef strStart As String, ByRef strPath As String, ByRef strItem As String, ByRef strQuantity As String, ByRef bTest As Boolean)
		Dim wp As Object
		Dim X As Short
		Dim beginPos As Short
		Dim endPos As Short
		Dim strAdjQuantity As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseWaypoints(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object wp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		wp = ParseWaypoints(strPath)
		Dim strDest As String
		If IsArray(wp) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strDest = wp(UBound(wp))
			strAdjQuantity = AdjustedQuantity(strStart, strDest, strItem, strQuantity)
			If CDbl(strAdjQuantity) <= 0 Then
				Exit Sub
			End If
			frmEmpCmd.SubmitEmpireCommand("bf1", True)
			For X = 1 To UBound(wp)
				If bTest Then
					'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmEmpCmd.SubmitEmpireCommand("test " & strItem & " " & strStart & " " & strAdjQuantity & " " + wp(X), True)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmEmpCmd.SubmitEmpireCommand("move " & strItem & " " & strStart & " " & strAdjQuantity & " " + wp(X), True)
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strStart = wp(X)
			Next X
			frmEmpCmd.SubmitEmpireCommand("bf2", True)
		ElseIf InStr(strPath, "\") > 0 Then 
			beginPos = 1
			endPos = InStr(strPath, "\")
			While endPos > 0
				MoveCommodity(strStart, Mid(strPath, beginPos, endPos - beginPos), strItem, strQuantity, bTest)
				beginPos = endPos + 1
				endPos = InStr(beginPos, strPath, "\")
			End While
			MoveCommodity(strStart, Mid(strPath, beginPos), strItem, strQuantity, bTest)
		Else
			strAdjQuantity = AdjustedQuantity(strStart, strPath, strItem, strQuantity)
			If Val(strAdjQuantity) <= 0 Then
				Exit Sub
			End If
			If bTest Then
				frmEmpCmd.SubmitEmpireCommand("test " & strItem & " " & strStart & " " & strAdjQuantity & " " & strPath, True)
			Else
				frmEmpCmd.SubmitEmpireCommand("move " & strItem & " " & strStart & " " & strAdjQuantity & " " & strPath, True)
			End If
		End If
	End Sub
	
	Private Sub cmdTest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTest.Click
		Dim strItem As String
		Dim beginPos As Short
		Dim endPos As Short
		
		strItem = cbCom.Text
		If VB.Left(strItem, 1) = "u" Then
			strItem = "un"
		End If
		
		'101903 rjk: Switched to Multiple Sector Selection
		If InStr(txtMultOrigin.Text, "\") > 0 Then
			beginPos = 1
			endPos = InStr(txtMultOrigin.Text, "\")
			While endPos > 0
				MoveCommodity(Mid(txtMultOrigin.Text, beginPos, endPos - beginPos), (txtMultPath.Text), strItem, (txtNum.Text), True)
				beginPos = endPos + 1
				endPos = InStr(beginPos, txtMultOrigin.Text, "\")
			End While
			MoveCommodity(Mid(txtMultOrigin.Text, beginPos), (txtMultPath.Text), strItem, (txtNum.Text), True)
		Else
			MoveCommodity((txtMultOrigin.Text), (txtMultPath.Text), strItem, (txtNum.Text), True)
		End If
	End Sub
	
	Private Sub cmdWayPoints_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdWayPoints.Click
		txtMultPath.Focus()
		txtMultPath_KeyPress(txtMultPath, New System.Windows.Forms.KeyPressEventArgs(Chr((32)))) 'simulate a space
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptMove.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptMove_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If bFirstLoad Then
			bFirstLoad = False
			Select Case LCase(Trim(Label2.Text))
				Case "deliver"
					cmdCancel.SetBounds(cmdTest.Left, cmdTest.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
					cmdTest.Visible = False
					chkDisplayPath.Visible = False
					chkShowAll.Visible = False '100203 rjk: Deliver already shows all.
					cmdOK.Text = "OK"
					txtNum.Text = "0"
					cmdWayPoints.Visible = False
					'txtMultPath.SetFocus 103103 rjk: Change to txtUnit, as txtMultPath is filled in from txtNum
					txtNum.Focus()
					If frmOptions.bStarWars Then
						Label1.Text = "Direction/ Radius"
						lblPathCost.Visible = False
					End If
					'101903 rjk: Removed the chkCond code, Added txtMultOrigin_Change to get initial quantity label correct
				Case "move"
					txtMultOrigin_TextChanged(txtMultOrigin, New System.EventArgs())
			End Select
		End If
	End Sub
	
	Private Sub frmPromptMove_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		bFirstLoad = True
		' Set parent for the toolbar to display on top of:
		' Dim sucess As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		'Put form in proper place
		Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
		
		If LastCommodityMoved = "" Then 'check for first time
			LastCommodityMoved = "civilians"
		End If
		
		LoadCommodityBox(cbCom)
		ComputePathCost()
	End Sub
	
	Private Sub frmPromptMove_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
	End Sub
	
	Private Sub txtMultOrigin_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.DoubleClick
		Dim strSector As String
		
		strSector = GetConditionalSectors
		If Len(strSector) > 0 Then
			txtMultOrigin.Text = strSector
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtNum.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtNum_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNum.TextChanged
		ComputePathCost()
	End Sub
	
	Private Sub txtNum_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNum.DoubleClick
		'Put the full amount in the text box
		If Len(strAmount) > 0 Then
			If cbCom.Text = "civilians" Then '100903 rjk: Reduce the civilian value by one so the sector is not abandoned.
				If Val(strAmount) = 1 Then
					txtNum.Text = "0"
				Else
					txtNum.Text = CStr(Val(strAmount) - 1)
				End If
			Else
				txtNum.Text = strAmount
			End If
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
	
	'UPGRADE_WARNING: Event txtMultPath.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultPath_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultPath.TextChanged
		' Dim p1 As Integer    8/2003 efj  removed
		Dim sx As Short
		Dim sy As Short
		On Error Resume Next
		
		If Len(txtMultPath.Text) = 0 Then
			lblDest.Text = vbNullString
			Exit Sub
		End If
		
		If ParseSectors(sx, sy, (txtMultPath.Text)) Then
			'Get Sector Information
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				lblDest.Text = CStr(rsSectors.Fields("eff").Value) & "% "
				'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lblDest.Text = lblDest.Text + colDes.Item(rsSectors.Fields("des").Value)
			Else '100203 rjk: Erase the sector description when not avaliable or valid
				lblDest.Text = ""
			End If
		Else '100203 rjk: Erase the sector description when not avaliable or valid
			lblDest.Text = ""
		End If
		
		'101903 rjk: Added Multiple Sector Selection Logic
		If InStr(txtMultPath.Text, "\") > 0 Then
			If InStr(txtMultOrigin.Text, "\") > 0 Then
				MsgBox("You can not have multiple sectors in start and end sectors")
				cmdOK.Enabled = False
				cmdTest.Enabled = False
			Else
				cmdOK.Enabled = True
				cmdTest.Enabled = True
			End If
			chkDisplayPath.Enabled = False
			cmdWayPoints.Enabled = False
			lblDest.Text = CStr(CountMultipleSectors((txtMultPath.Text))) & " Sectors"
		Else
			cmdWayPoints.Enabled = True
			If InStr(txtMultOrigin.Text, "\") = 0 Then
				chkDisplayPath.Enabled = True
			Else
				chkDisplayPath.Enabled = False
			End If
			cmdOK.Enabled = True
			cmdTest.Enabled = True
		End If
		SetCurrentQuantityLabel()
		ComputePathCost()
	End Sub
	
	Private Sub txtMultPath_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultPath.DoubleClick
		'Make sure starting sector is valid before prompting
		Dim sx As Short
		Dim sy As Short
		If Not ParseSectors(sx, sy, (txtMultOrigin.Text)) Then
			Exit Sub
		End If
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptPath)
		frmPromptPath.lblSector.Text = txtMultOrigin.Text
		frmPromptPath.lblEndSector.Text = txtMultOrigin.Text
		frmPromptPath.Text = "Select Movement Path"
		frmPromptPath.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptPath.Width))
		frmPromptPath.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptPath.Height)) \ 2)
		frmPromptPath.Show()
	End Sub
	
	'UPGRADE_WARNING: Event txtMultOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.TextChanged
		' Dim p1 As Integer    8/2003 efj  removed
		Dim sx As Short
		Dim sy As Short
		Dim n As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		On Error Resume Next
		
		Dim strCurrentComm As String
		If ParseSectors(sx, sy, (txtMultOrigin.Text)) Then
			'Get Sector Information
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				lblOrigin.Text = CStr(rsSectors.Fields("eff").Value) & "% "
				'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lblOrigin.Text = lblOrigin.Text + colDes.Item(rsSectors.Fields("des").Value)
				
				'remove all items that are not in sector
				LoadCommodityBox(cbCom)
				If LCase(Trim(Label2.Text)) = "move" Then
					If Not chkShowAll.CheckState = System.Windows.Forms.CheckState.Checked Then
						For n = cbCom.Items.Count - 1 To 0 Step -1
							'102303 rjk: Move common code to a function
							If Items.FindByFormName(VB6.GetItemString(cbCom, n)).DatabaseValue(rsSectors) = 0 Then
								cbCom.Items.RemoveAt(n)
							End If
						Next n
					End If
					For n = 0 To cbCom.Items.Count - 1
						If VB6.GetItemString(cbCom, n) = LastCommodityMoved Then
							If cbCom.Visible Then
								cbCom.SelectedIndex = n 'set to last commodity
							End If
							Exit For
						End If
					Next n
					If n = cbCom.Items.Count Then
						If cbCom.Visible Then
							cbCom.SelectedIndex = 0
						End If
					End If
				ElseIf LCase(Trim(Label2.Text)) = "deliver" Then 
					'270104 rjk: Change the selection of the default commodity for deliver to new sector designation
					'180304 rjk: Fixed to deal with sdes being a space
					If Trim(rsSectors.Fields("sdes").Value) <> "" And HexGrid.dspDesignation <> HexGrid.enumDisplayDesignation.DD_CURRENT Then
						SetDefaultCommodity((rsSectors.Fields("sdes").Value))
					Else
						SetDefaultCommodity((rsSectors.Fields("des").Value))
					End If
				End If
				chkShowAll.Enabled = True
				If InStr(txtMultPath.Text, "\") = 0 Then
					chkDisplayPath.Enabled = True
				Else
					chkDisplayPath.Enabled = False
				End If
			Else '100203 rjk: Erase the sector description when not avaliable or valid
				lblOrigin.Text = ""
				cbCom.Items.Clear()
				lblBestPath.Text = ""
				lblPathCost.Text = ""
				lblCost.Text = ""
				lblLeft.Text = ""
				chkShowAll.Enabled = False
				chkDisplayPath.Enabled = False
			End If
			cmdOK.Enabled = True
			cmdTest.Enabled = True
		Else '100203 rjk: Erase the sector description when not avaliable or valid
			'101903 rjk: Added Multiple Sector Selection Logic, removed condition logic
			lblOrigin.Text = ""
			If InStr(txtMultOrigin.Text, "\") > 0 Then
				
				strCurrentComm = cbCom.Text
				LoadCommodityBox(cbCom)
				For n = 0 To cbCom.Items.Count - 1
					If VB6.GetItemString(cbCom, n) = strCurrentComm Then
						cbCom.SelectedIndex = n 'set to last commodity
						Exit For
					End If
				Next n
				If n = cbCom.Items.Count Then
					cbCom.SelectedIndex = 0
				End If
				txtMultOrigin.Focus()
				If InStr(txtMultPath.Text, "\") > 0 Then
					MsgBox("You can not have multiple sectors in start and end sectors")
					cmdOK.Enabled = False
					cmdTest.Enabled = False
				Else
					cmdOK.Enabled = True
					cmdTest.Enabled = True
				End If
				lblOrigin.Text = CStr(CountMultipleSectors((txtMultOrigin.Text))) & " Sectors"
			Else
				cbCom.Items.Clear()
				cmdOK.Enabled = True
				cmdTest.Enabled = True
			End If
			lblBestPath.Text = ""
			lblPathCost.Text = ""
			lblCost.Text = ""
			lblLeft.Text = ""
			chkShowAll.Enabled = False
			chkDisplayPath.Enabled = False
		End If
		SetCurrentQuantityLabel()
		ComputePathCost()
	End Sub
	
	Private Sub txtMultPath_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMultPath.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 32 And (LCase(Trim(Label2.Text)) = "move") Then
			KeyAscii = 0
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmPromptWaypoints)
			frmPromptWaypoints.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptWaypoints.Width))
			frmPromptWaypoints.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptWaypoints.Height)) \ 2)
			frmPromptWaypoints.Show()
			frmPromptWaypoints.txtOrigin.Focus()
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub ComputePathCost()
		On Error Resume Next
		Dim pvar As Object
		Dim pcost As Single
		Dim sx As Short
		Dim sy As Short
		Dim pbdes As String
		
		lblLeft.Text = vbNullString
		lblCost.Text = vbNullString
		lblBestPath.Text = vbNullString
		lblPathCost.Text = vbNullString
		
		If LCase(Trim(Label2.Text)) <> "move" Then
			Exit Sub
		End If
		
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
		
		'Compute the path cost
		If ParseSectors(sx, sy, (txtMultPath.Text)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pvar = BestPath((txtMultOrigin.Text), (txtMultPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
			'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pcost = pvar(2)
			'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Len(CStr(pvar(1))) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lblBestPath.Text = "best path: " & CStr(pvar(1))
				lblPathCost.Text = "path cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
				If chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked Then
					'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmDrawMap.DisplayPath((txtMultOrigin.Text), CStr(pvar(1)))
				End If
			Else
				lblBestPath.Text = "best path: ??"
				lblPathCost.Text = "path cost: ??"
			End If
		Else
			pcost = PathCost((txtMultOrigin.Text), (txtMultPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
			If pcost > 0 Then
				lblBestPath.Text = "path: " & txtMultPath.Text
				lblPathCost.Text = "path cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
				If chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked Then
					frmDrawMap.DisplayPath((txtMultOrigin.Text), (txtMultPath.Text))
				End If
			End If
		End If
		
		'Compute the mobility cost
		'102303 rjk: Changed to Mobility cost to use AdjustedQuantity
		Dim nAdjQuantity As Short
		nAdjQuantity = Val(AdjustedQuantity((txtMultOrigin.Text), (txtMultPath.Text), (Items.FindByFormName(VB6.GetItemString(cbCom, cbCom.SelectedIndex)).Letter), (txtNum.Text)))
		'If pcost > 0 And Val(txtNum.Text) > 0 Then
		If pcost > 0 And nAdjQuantity > 0 Then
			If ParseSectors(sx, sy, (txtMultOrigin.Text)) Then
				'Get Sector Information
				rsSectors.Seek("=", sx, sy)
				If Not rsSectors.NoMatch Then
					If rsSectors.Fields("eff").Value < 60 Then
						pbdes = vbNullString
					Else
						pbdes = rsSectors.Fields("des").Value
					End If
					'pcost = pcost * MobilityWeight(Left$(cbCom.Text, 1), Val(txtNum.Text), pbdes)
					pcost = pcost * MobilityWeight(VB.Left(cbCom.Text, 1), nAdjQuantity, pbdes)
					lblCost.Text = "mob. cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
					If pcost > rsSectors.Fields("mob").Value Then
						lblLeft.Text = "Insuff. mobility"
					Else
						lblLeft.Text = "mob. left: " & CStr(Int(rsSectors.Fields("mob").Value - pcost))
					End If
				End If
			End If
		End If
	End Sub
	
	Private Sub SetDefaultCommodity(ByRef strDes As String)
		Dim i As Short
		
		i = -1
		'determine com box type based on sector
		'092103 rjk: rsSectors.Seek "=", MxPos, MyPos switched to SelX and SelY to work from Top Level Menu
		'100203 rjk: Moved com box type based on sector logic to the form.
		Select Case strDes
			Case "j" 'set to lcm
				i = FindByLetterCommodityBox("l")
			Case "k" 'set to hcm
				i = FindByLetterCommodityBox("h")
			Case "m" 'set to iron
				i = FindByLetterCommodityBox("i")
			Case "g" 'set to dust
				i = FindByLetterCommodityBox("d")
			Case "o" 'set to oil
				i = FindByLetterCommodityBox("o")
			Case "u" 'set to rads
				i = FindByLetterCommodityBox("r")
			Case "%" 'set to petrol
				i = FindByLetterCommodityBox("p")
			Case "a" 'set to food
				i = FindByLetterCommodityBox("f")
			Case "b" 'set to bars
				i = FindByLetterCommodityBox("b")
			Case "d" 'set to guns
				i = FindByLetterCommodityBox("g")
			Case "i" 'set to shells
				i = FindByLetterCommodityBox("s")
			Case "y" '240104 rjk: Added for St@r W@rs (uw85 robotics)
				i = FindByLetterCommodityBox("u")
		End Select
		If i <> -1 Then
			cbCom.SelectedIndex = i
		End If
	End Sub
	
	'102303 rjk: Function to compute the desired quantity based on the following logic
	'            ### moves that many
	'            +### moves as many as needed to result in ### total in the To sector
	'            -### leaves ### behind in the From sector and moves the excess
	Private Function AdjustedQuantity(ByRef strStart As String, ByRef strEnd As String, ByRef strItem As String, ByRef strQuantity As String) As String
		Dim sx As Short
		Dim sy As Short
		
		If Len(strQuantity) = 0 Then
			AdjustedQuantity = strQuantity
			Exit Function
		End If
		Select Case VB.Left(strQuantity, 1)
			Case "-"
				If Not ParseSectors(sx, sy, strStart) Then
					AdjustedQuantity = strQuantity
					Exit Function
				End If
				
				rsSectors.Seek("=", sx, sy)
				If rsSectors.NoMatch Then
					AdjustedQuantity = strQuantity
					Exit Function
				End If
				
				If (Items.FindByLetter(strItem).DatabaseValue(rsSectors) - System.Math.Abs(Val(strQuantity))) > 0 Then
					AdjustedQuantity = CStr(Items.FindByLetter(strItem).DatabaseValue(rsSectors) - System.Math.Abs(Val(strQuantity)))
				Else
					AdjustedQuantity = "0"
				End If
			Case "+"
				If Not ParseSectors(sx, sy, strEnd) Then
					AdjustedQuantity = strQuantity
					Exit Function
				End If
				
				rsSectors.Seek("=", sx, sy)
				If rsSectors.NoMatch Then
					AdjustedQuantity = strQuantity
					Exit Function
				End If
				
				If (Val(strQuantity) - Items.FindByLetter(strItem).DatabaseValue(rsSectors)) > 0 Then
					AdjustedQuantity = CStr(Val(strQuantity) - Items.FindByLetter(strItem).DatabaseValue(rsSectors))
				Else
					AdjustedQuantity = "0"
				End If
			Case Else
				AdjustedQuantity = strQuantity
		End Select
	End Function
End Class