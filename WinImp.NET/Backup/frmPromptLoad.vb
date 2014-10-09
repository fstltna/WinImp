Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptLoad
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'091803 rjk: Added Unit grid selection when activating this form
	'092103 rjk: Updated the information when the origin changes.
	'092803 rjk: Made OK button the default.
	'            Switched grid selection to use SetUnitType and DisplayFirstSectorPanel.
	'            General reformatting.
	'            Move the field selection logic from frmDrawMap to the form.
	'            Fixed to logic to account for ltend loading lunit from a ship
	'            Added logic for multiple units being loaded, ensure there are all at the same location
	'100903 rjk: Added bar to the list of database fields that is used to get the current quantities from the database
	'            Reduce the civilian value by one so the sector is not abandoned when double clicking.
	'101703 rjk: Added Strength fields to Sector database
	'111803 rjk: Added Multiple Sector Selection to the MoveAll command. Added -/+ logic to move quantities.
	'            Added Commodity Total display using this form.  Rearranged form to make it smaller.
	'            Added -/+ quantity logic to load/unload/lload/lunload.
	'112003 rjk: Added option to control strength updates
	'112103 rjk: Fixed so the quantities are cleared for CommodityTotal or MoveAll when no sector is present.
	'121303 rjk: Switch from GetCommodityValueFromDB to Item.DatabaseValue and Item class
	'240104 rjk: Fixed the commodities to use a single letter for command.
	'280204 rjk: Fixed Mob., Work. and Avail. for setsector, changed to lower case letters.
	'090404 rjk: Added all sectors option to Commodity Total
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'270904 rjk: Enabled land unit loading for 2K4.
	'180206 rjk: Replace pdump, ldump, and sdump with GetPlanes, GetLandUnits and GetShips.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'110307 rjk: Filter the load/unload/lload/lunload commands to relevant units only.
	'190508 rjk: Added sector total to CommodityTotal
	'310508 rjk: Allow loading of land units on trains for Heavy Metal II
	'020608 rjk: Allow unloading of land units on trains for Heavy Metal II
	
	Public strCmd As String
	Public strSector As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim n As Short
		Dim X As Short
		Dim strCmd2 As String
		Dim strCmd3 As String
		Dim strItem As String
		Dim beginPos As Short
		Dim endPos As Short
		On Error Resume Next
		
		'If this is a tend command, this holds the ship to be tended to
		If txtUnit3.Visible = True Then
			strCmd3 = txtUnit3.Text
		Else
			strCmd3 = vbNullString
		End If
		
		frmEmpCmd.SubmitEmpireCommand("bf1", False)
		Dim eiItem As EmpItem
		For n = 0 To 13
			If Len(txtThresh(n).Text) > 0 Then
				strItem = lblThresh(n).Text
				
				'Remove the amount portion of the caption
				X = InStr(strItem, " (")
				If X > 0 Then
					strItem = Trim(VB.Left(strItem, X))
				End If
				
				'If Left$(strItem, 1) = "u" Then
				'    strItem = "un"
				'End If
				eiItem = Items.FindByFormLabel(strItem)
				If Not (eiItem Is Nothing) Then
					strItem = eiItem.Letter
				End If
				If strCmd = "move" Then
					If InStr(txtMultOrigin.Text, "\") > 0 Then '111803 Multiple Sectors Selected
						beginPos = 1
						endPos = InStr(txtMultOrigin.Text, "\")
						While endPos > 0
							frmPromptMove.MoveCommodity(Mid(txtMultOrigin.Text, beginPos, endPos - beginPos), (txtMultDest.Text), strItem, txtThresh(n).Text, False)
							beginPos = endPos + 1
							endPos = InStr(beginPos, txtMultOrigin.Text, "\")
						End While
						frmPromptMove.MoveCommodity(Mid(txtMultOrigin.Text, beginPos), (txtMultDest.Text), (Items.FindByFormLabel(strItem).Letter), txtThresh(n).Text, False)
					Else
						frmPromptMove.MoveCommodity((txtMultOrigin.Text), (txtMultDest.Text), strItem, txtThresh(n).Text, False)
					End If
					'strCmd2 = Trim$(strCmd + " " + strItem + " " + txtMultOrigin.Text + " " + txtThresh(n).Text + " " + txtMultDest.Text)
				ElseIf strCmd = "give" Or strCmd = "setresource" Or strCmd = "setsector" Then 
					If InStr(txtMultOrigin.Text, "\") > 0 Then '111803 Multiple Sectors Selected
						beginPos = 1
						endPos = InStr(txtMultOrigin.Text, "\")
						While endPos > 0
							SetDeityLevel(strCmd, Mid(txtMultOrigin.Text, beginPos, endPos - beginPos), strItem, txtThresh(n).Text)
							beginPos = endPos + 1
							endPos = InStr(beginPos, txtMultOrigin.Text, "\")
						End While
						SetDeityLevel(strCmd, Mid(txtMultOrigin.Text, beginPos), strItem, txtThresh(n).Text)
					Else
						SetDeityLevel(strCmd, (txtMultOrigin.Text), strItem, txtThresh(n).Text)
					End If
					'strCmd2 = Trim$(strCmd + " " + strItem + " " + txtMultOrigin.Text + " " + txtThresh(n).Text)
				ElseIf strCmd = "load" Or strCmd = "unload" Or strCmd = "lload" Or strCmd = "lunload" Then 
					If InStr(txtThresh(n).Text, "-") > 0 Or InStr(txtThresh(n).Text, "+") > 0 Then '111803 -/+ quantity logic
						LoadCommodity(strCmd, (txtMultOrigin.Text), (txtUnit.Text), strItem, txtThresh(n).Text)
					Else
						strCmd2 = Trim(strCmd & " " & strItem & " " & txtUnit.Text & " " & txtThresh(n).Text)
						frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
					End If
					'strCmd2 = Trim$(strCmd + " " + strItem + " " + txtUnit.Text + " " + txtThresh(n).Text + " " + strCmd3)
				Else
					strCmd2 = Trim(strCmd & " " & strItem & " " & txtUnit.Text & " " & txtThresh(n).Text & " " & txtUnit3.Text)
					frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
				End If
			End If
		Next n
		
		If strCmd <> "move" Then
			If Option1.Checked And Len(txtUnit2.Text) > 0 Then
				strCmd2 = Trim(strCmd & " plane " & txtUnit.Text & " " & txtUnit2.Text & " " & strCmd3)
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			End If
			
			If Option2.Checked And Len(txtUnit2.Text) > 0 Then
				strCmd2 = Trim(strCmd & " land " & txtUnit.Text & " " & txtUnit2.Text & " " & strCmd3)
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			End If
		End If
		frmEmpCmd.SubmitEmpireCommand("bf2", False)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If strCmd <> "tend" Then
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
		End If
		If strCmd = "lload" Or strCmd = "lunload" Or strCmd = "ltend" Then
			GetLandUnits()
		ElseIf strCmd = "load" Or strCmd = "unload" Or strCmd = "tend" Or strCmd = "ltend" Then 
			GetShips()
			If (Option2.Checked And Len(txtUnit2.Text) > 0) Then
				GetLandUnits()
			End If
		End If
		
		'dump plane changes if any where loaded/unloaded
		If strCmd <> "move" Then
			If Option1.Checked And Len(txtUnit2.Text) > 0 Then
				GetPlanes()
			End If
		End If
		
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptLoad.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptLoad_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If (strCmd = "lload" And Not (frmOptions.b2K4 Or frmOptions.bHeavyMetal)) Or (strCmd = "lunload" And Not frmOptions.bHeavyMetal) Or strCmd = "ltend" Then
			Option2.Visible = False
		End If
		
		If strCmd = "move" Then
			Frame1.Visible = True
		Else
			Frame1.Visible = False
		End If
	End Sub
	
	Private Sub frmPromptLoad_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		SetLabels()
		
		Dim n As Short
		For n = 0 To 13
			txtThresh(n).Text = vbNullString
		Next n
		
		Select Case strCmd
			Case "move"
				txtMultOrigin.SetBounds(lblThresh(5).Left, txtUnit.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				txtMultDest.SetBounds(lblThresh(10).Left, txtUnit3.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Label3.SetBounds(txtThresh(5).Left, txtUnit3.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
				txtUnit.Visible = False
				txtUnit3.Visible = False
				txtUnit2.Visible = False
				Option1.Visible = False
				Option2.Visible = False
				txtMultOrigin.Visible = True
				txtMultDest.Visible = True
				Label1.Text = "from"
				Label3.Text = "to"
				Label2.Text = "Move"
				Label3.Visible = True
				'txtMultOrigin.Text = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
				txtMultOrigin.Text = strSector
				
				SetSectorLabels()
				Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
				Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
			Case "unload"
				Label2.Text = "Unload"
				Label1.Text = "from ship"
			Case "give", "setsector", "setresource"
				Select Case strCmd
					Case "give"
						Label2.Text = "Give"
					Case "setsector"
						Label2.Text = "Set Sector"
					Case "setresource"
						Label2.Text = "Set Resource"
				End Select
				txtMultOrigin.SetBounds(lblThresh(5).Left, txtUnit.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				txtMultDest.SetBounds(lblThresh(10).Left, txtUnit3.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Label3.SetBounds(txtThresh(5).Left, txtUnit3.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
				txtUnit.Visible = False
				txtUnit3.Visible = False
				txtUnit2.Visible = False
				Option1.Visible = False
				Option2.Visible = False
				txtMultOrigin.Visible = True
				txtMultDest.Visible = False
				Label1.Text = "to"
				Label3.Text = vbNullString
				Label3.Visible = True
				'txtMultOrigin.Text = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
				txtMultOrigin.Text = strSector
				SetSectorLabels()
			Case "lload"
				Label1.Text = "on land"
				SetSectorLabels()
			Case "ltend"
				Label2.Text = "Tend"
				Label1.Text = "from ship"
				Label3.Text = "to land"
				Label3.Visible = True
				txtUnit3.Visible = True
			Case "lunload"
				Label2.Text = "Unload"
				Label1.Text = "from land"
			Case "load"
				SetSectorLabels()
			Case "tend"
				Label2.Text = "Tend"
				Label1.Text = "from ship"
				Label3.Visible = True
				txtUnit3.Visible = True
			Case "commoditytotal"
				Label1.SetBounds(lblThresh(5).Left, txtUnit.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				txtMultOrigin.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Label1.Left) + VB6.PixelsToTwipsX(Label1.Width)), txtUnit.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				txtUnit.Visible = False
				txtUnit3.Visible = False
				txtUnit2.Visible = False
				Option1.Visible = False
				Option2.Visible = False
				txtMultOrigin.Visible = True
				txtMultDest.Visible = False
				Label1.Text = "for"
				Label2.Text = "Commodity Total"
				Label3.Visible = True
				Label3.Width = VB6.TwipsToPixelsX(1000)
				Label3.Text = "0 Sectors"
				txtMultOrigin.Text = strSector
				cmdOK.Visible = False
				cmdCancel.Text = "Close"
				For n = 0 To 13
					txtThresh(n).Enabled = False
					txtThresh(n).Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtThresh(n).Left) - VB6.PixelsToTwipsX(txtThresh(n).Width))
					txtThresh(n).Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtThresh(n).Width) + VB6.PixelsToTwipsX(txtThresh(n).Width))
				Next n
				
				SetSectorLabels()
				Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
				Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
		End Select
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		'Option2.Visible = True 092103 rjk: should be deal with above
		
	End Sub
	
	Private Sub frmPromptLoad_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Public Sub SetSectorLabels() '8/2003 efj  removed,  091103 rjk added back - it is used
		Dim sx As Short
		Dim sy As Short
		Dim t As Object
		Dim i As Short
		'100903 rjk: Added bar to the list so it is filled as well
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object t. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		t = New Object(){"mil", "gun", "shell", "food", "civ", "uw", "iron", "lcm", "hcm", "oil", "dust", "pet", "rad", "bar"}
		
		SetLabels()
		
		'Command: setsector
		'Give What (iron, gold, oil, uranium, fertility, owner, eff., mob., work, avail., oldown, mines)?
		
		'If the sector info is correct, fill with what available in sector (for Load, lload)
		Dim commodityTotal() As Integer
		Dim beginPos, endPos As Short
		Dim sectorCount As Short
		Dim lblTh As System.Windows.Forms.Label
		If strCmd = "commoditytotal" Then '111803 Commodity Total
			
			'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object t. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			t = New Object(){"mil", "gun", "shell", "food", "civ", "uw", "iron", "lcm", "hcm", "oil", "dust", "pet", "rad", "bar"}
			ReDim commodityTotal(UBound(t))
			If InStr(txtMultOrigin.Text, "\") > 0 Then
				beginPos = 1
				endPos = InStr(txtMultOrigin.Text, "\")
				While endPos > 0
					If ParseSectors(sx, sy, Mid(txtMultOrigin.Text, beginPos, endPos - beginPos)) Then
						rsSectors.Seek("=", sx, sy)
						If Not rsSectors.NoMatch Then
							sectorCount = sectorCount + 1
							For i = LBound(t) To UBound(t)
								commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
							Next 
						End If
					End If
					beginPos = endPos + 1
					endPos = InStr(beginPos, txtMultOrigin.Text, "\")
				End While
				If ParseSectors(sx, sy, Mid(txtMultOrigin.Text, beginPos)) Then
					rsSectors.Seek("=", sx, sy)
					If Not rsSectors.NoMatch Then
						sectorCount = sectorCount + 1
						For i = LBound(t) To UBound(t)
							commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
						Next 
					End If
				End If
			ElseIf ParseSectors(sx, sy, (txtMultOrigin.Text)) Then 
				rsSectors.Seek("=", sx, sy)
				If Not rsSectors.NoMatch Then
					sectorCount = sectorCount + 1
					For i = LBound(t) To UBound(t)
						commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
					Next 
				End If
			ElseIf txtMultOrigin.Text = "*" Then  '090404 rjk: Added an option for all sectors
				If Not (rsSectors.EOF And rsSectors.BOF) Then
					rsSectors.MoveFirst()
					While Not rsSectors.EOF
						If Not rsSectors.NoMatch Then
							sectorCount = sectorCount + 1
							For i = LBound(t) To UBound(t)
								commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
							Next 
						End If
						rsSectors.MoveNext()
					End While
				End If
			End If
			For i = LBound(t) To UBound(t)
				txtThresh(i).Text = CStr(commodityTotal(i))
			Next 
			Label3.Text = CStr(sectorCount) & " Sectors"
		ElseIf ParseSectors(sx, sy, strSector) Then 
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				If strCmd = "setresource" Or strCmd = "setsector" Then
					lblThresh(0).Text = lblThresh(0).Text & " (" & CStr(rsSectors.Fields("min").Value) & ")"
					lblThresh(1).Text = lblThresh(1).Text & " (" & CStr(rsSectors.Fields("gold").Value) & ")"
					lblThresh(2).Text = lblThresh(2).Text & " (" & CStr(rsSectors.Fields("fert").Value) & ")"
					lblThresh(3).Text = lblThresh(3).Text & " (" & CStr(rsSectors.Fields("oil").Value) & ")"
					lblThresh(6).Text = lblThresh(6).Text & " (" & CStr(rsSectors.Fields("uran").Value) & ")"
					
					If strCmd = "setsector" Then
						lblThresh(4).Text = lblThresh(4).Text & " (" & CStr(rsSectors.Fields("country").Value) & ")"
						lblThresh(5).Text = lblThresh(5).Text & " (" & CStr(rsSectors.Fields("eff").Value) & ")"
						lblThresh(9).Text = lblThresh(9).Text & " (" & CStr(rsSectors.Fields("mob").Value) & ")"
						lblThresh(8).Text = lblThresh(8).Text & " (" & CStr(rsSectors.Fields("work").Value) & ")"
						lblThresh(7).Text = lblThresh(7).Text & " (" & CStr(rsSectors.Fields("avail").Value) & ")"
					End If
				Else
					
					For	Each lblTh In lblThresh
						lblTh.Text = lblTh.Text & " (" & CStr(Items.FindByFormLabel((lblTh.Text)).DatabaseValue(rsSectors)) & ")"
					Next lblTh
				End If
			End If
		End If
	End Sub
	
	Public Sub SetShipLabels()
		'UPGRADE_WARNING: Arrays in structure rsTemp may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rsTemp As DAO.Recordset
		Dim hldIndex As String
		Dim isLand As Boolean
		Dim bLast As Boolean
		Dim bSecondFound As Boolean
		Dim strFirstSector As String
		Dim strUnit As String
		
		SetLabels()
		
		'check for ship or land, and choose the correct recordset
		If strCmd = "lload" Or strCmd = "lunload" Then '092803 move ltend to ship as the origin is a ship
			rsTemp = rsLand
			hldIndex = rsLand.Index
			rsTemp.Index = "PrimaryKey"
			isLand = True
		Else
			rsTemp = rsShip
			hldIndex = rsShip.Index
			rsTemp.Index = "PrimaryKey"
			isLand = False
		End If
		
		On Error GoTo Error_Exit
		
		'If there is a single unit in the box, set the labels based on its current load
		If Len(txtUnit.Text) > 0 And InStr(txtUnit.Text, "/") = 0 Then
			rsTemp.Seek("=", txtUnit.Text)
			If Not rsTemp.NoMatch Then
				If strCmd = "load" Or strCmd = "lload" Then
					strSector = CStr(rsTemp.Fields("x").Value) & " , " & CStr(rsTemp.Fields("y").Value)
					SetSectorLabels()
				Else
					lblThresh(0).Text = lblThresh(0).Text & " (" & CStr(rsTemp.Fields("mil").Value) & ")"
					lblThresh(1).Text = lblThresh(1).Text & " (" & CStr(rsTemp.Fields("gun").Value) & ")"
					lblThresh(2).Text = lblThresh(2).Text & " (" & CStr(rsTemp.Fields("shell").Value) & ")"
					lblThresh(3).Text = lblThresh(3).Text & " (" & CStr(rsTemp.Fields("food").Value) & ")"
					
					'Land units don't have civs
					If isLand Then
						lblThresh(4).Text = lblThresh(4).Text & " (N/A)"
						lblThresh(5).Text = lblThresh(5).Text & " (N/A)"
					Else
						lblThresh(4).Text = lblThresh(4).Text & " (" & CStr(rsTemp.Fields("civ").Value) & ")"
						lblThresh(5).Text = lblThresh(5).Text & " (" & CStr(rsTemp.Fields("uw").Value) & ")"
					End If
					lblThresh(6).Text = lblThresh(6).Text & " (" & CStr(rsTemp.Fields("iron").Value) & ")"
					lblThresh(7).Text = lblThresh(7).Text & " (" & CStr(rsTemp.Fields("lcm").Value) & ")"
					lblThresh(8).Text = lblThresh(8).Text & " (" & CStr(rsTemp.Fields("hcm").Value) & ")"
					lblThresh(9).Text = lblThresh(9).Text & " (" & CStr(rsTemp.Fields("oil").Value) & ")"
					lblThresh(10).Text = lblThresh(10).Text & " (" & CStr(rsTemp.Fields("dust").Value) & ")"
					lblThresh(11).Text = lblThresh(11).Text & " (" & CStr(rsTemp.Fields("petrol").Value) & ")"
					lblThresh(12).Text = lblThresh(12).Text & " (" & CStr(rsTemp.Fields("rad").Value) & ")"
					lblThresh(13).Text = lblThresh(13).Text & " (" & CStr(rsTemp.Fields("bar").Value) & ")"
				End If
			End If
			'092803 rjk: added logic for multiple units being loaded, ensure there are all at the same location
		ElseIf (strCmd = "load" Or strCmd = "lload") And Len(txtUnit.Text) > 0 Then 
			strUnit = txtUnit.Text
			bSecondFound = False
			Do 
				If InStr(strUnit, "/") <> 0 Then
					rsTemp.Seek("=", VB.Left(strUnit, InStr(strUnit, "/") - 1))
					strUnit = Mid(strUnit, InStr(strUnit, "/") + 1)
					bLast = False
				Else
					rsTemp.Seek("=", strUnit)
					bLast = True
				End If
				If Not rsTemp.NoMatch Then
					rsSectors.Seek("=", rsTemp.Fields("x").Value, rsTemp.Fields("y").Value)
					If Not rsSectors.NoMatch Then
						If (strCmd = "load" And rsSectors.Fields("des").Value = "h") Or strCmd = "lload" Then
							If strFirstSector = "" Then
								strFirstSector = CStr(rsTemp.Fields("x").Value) & " , " & CStr(rsTemp.Fields("y").Value)
							ElseIf strFirstSector <> (CStr(rsTemp.Fields("x").Value) & " , " & CStr(rsTemp.Fields("y").Value)) Then 
								bSecondFound = True
								Exit Do 'found a second source location for loading
							End If
						End If
					End If
				End If
			Loop While Not bLast
			If (Not bSecondFound) And strFirstSector <> "" Then 'only display load quantities if one loading point is being used
				strSector = strFirstSector
				SetSectorLabels()
			Else
				SetLabels()
			End If
		End If
		
Error_Exit: 
		
		rsTemp.Index = hldIndex
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
	End Sub
	
	Private Sub SetLabels()
		'Give What (iron, gold, oil, uranium, fertility, owner, eff., mob., work, avail., oldown, mines)?
		If strCmd = "setresource" Or strCmd = "setsector" Then
			lblThresh(0).Text = "iron"
			lblThresh(1).Text = "gold"
			lblThresh(2).Text = "fertility"
			lblThresh(3).Text = "oil"
			lblThresh(6).Text = "uranium"
			
			If strCmd = "setresource" Then
				lblThresh(4).Visible = False
				lblThresh(5).Visible = False
				lblThresh(7).Visible = False
				lblThresh(8).Visible = False
				lblThresh(9).Visible = False
				lblThresh(10).Visible = False
				lblThresh(13).Visible = False
				txtThresh(4).Visible = False
				txtThresh(5).Visible = False
				txtThresh(7).Visible = False
				txtThresh(8).Visible = False
				txtThresh(9).Visible = False
				txtThresh(10).Visible = False
				txtThresh(13).Visible = False
			Else
				lblThresh(4).Text = "owner"
				lblThresh(5).Text = "eff"
				lblThresh(9).Text = "mob"
				lblThresh(8).Text = "work"
				lblThresh(7).Text = "avail"
				lblThresh(10).Text = "oldown"
				lblThresh(13).Text = "mines"
			End If
			lblThresh(11).Visible = False
			lblThresh(12).Visible = False
			txtThresh(11).Visible = False
			txtThresh(12).Visible = False
		Else
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
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			' added by thomas lullier
			' to prevent extra click when loading planes
			
			'Call frmDrawMap.ChangeUnitDisplay("Plane", True) 091803 rjk done in GotFocus now
			txtUnit2.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			' added by thomas lullier
			' to prevent extra click when loading land units
			
			'Call frmDrawMap.ChangeUnitDisplay("Land", True) '091803 rjk done in GotFocus now
			txtUnit2.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtMultDest.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultDest_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultDest.TextChanged
		ComputePathCost()
	End Sub
	
	Private Sub txtMultDest_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultDest.DoubleClick
		Dim strSector As String
		If strCmd = "move" Then
			
			strSector = GetConditionalSectors
			If Len(strSector) > 0 Then
				txtMultDest.Text = strSector
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtMultOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMultOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.TextChanged
		'Only allow OK if a ship is selected
		cmdOK.Enabled = (Len(txtMultOrigin.Text) > 0)
		'112103 rjk: Removed so the quantity are cleared for CommodityTotal or MoveAll when no sector is present.
		'If Len(txtMultOrigin) > 0 Then ' 092103 rjk: Update Labels(quantity) when origin changes
		strSector = txtMultOrigin.Text
		SetSectorLabels()
		ComputePathCost()
		'End If
	End Sub
	
	Private Sub txtMultOrigin_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMultOrigin.DoubleClick
		Dim strSector As String
		If strCmd = "give" Or strCmd = "setresource" Or strCmd = "setsector" Then
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmPromptSectors)
			frmPromptSectors.strSectors = txtMultOrigin.Text
			frmPromptSectors.SetControls()
			frmPromptSectors.Text = "Select Sectors"
			frmPromptSectors.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptSectors.Width))
			frmPromptSectors.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptSectors.Height)) \ 2)
			frmPromptSectors.Show()
		ElseIf strCmd = "move" Or strCmd = "commoditytotal" Then 
			
			strSector = GetConditionalSectors
			If Len(strSector) > 0 Then
				txtMultOrigin.Text = strSector
			End If
		End If
	End Sub
	
	'added drk 5/31/03.  munged from the PromptMove form in order to display mobility
	'cost for the MoveAll command
	Private Sub ComputePathCost()
		On Error Resume Next
		Dim pvar As Object
		Dim pcost As Single
		Dim mcost As Single
		Dim sx As Short
		Dim sy As Short
		Dim pbdes As String
		Dim i As Short
		
		lblMobLeft.Text = vbNullString
		lblMcost.Text = vbNullString
		lblPathCost.Text = vbNullString
		
		If LCase(Trim(Label2.Text)) <> "move" Then
			Exit Sub
		End If
		
		'111803 rjk: Added Multiple Sector Selection Logic
		If InStr(txtMultDest.Text, "\") > 0 Then
			If InStr(txtMultOrigin.Text, "\") > 0 Then
				MsgBox("You can not have multiple sectors in start and end sectors")
				Exit Sub
			End If
		End If
		
		'If frmDrawMap.DisplayingPath Then
		'    frmDrawMap.DisplayingPath = False
		'    frmDrawMap.FinishPath
		'End If
		
		'Compute the path cost
		If ParseSectors(sx, sy, (txtMultDest.Text)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pvar = BestPath((txtMultOrigin.Text), (txtMultDest.Text), EmpCommon.enumMobType.MT_COMMODITY)
			'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pcost = pvar(2)
			'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Len(CStr(pvar(1))) > 0 Then
				'lblBestPath.Caption = "best path: " + CStr(pvar(1))
				lblPathCost.Text = "path cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
				'If chkDisplayPath.Value = vbChecked Then
				'    DisplayPath txtMultOrigin.Text, CStr(pvar(1))
				'End If
			Else
				'lblBestPath.Caption = "best path: ??"
				lblPathCost.Text = "path cost: ??"
			End If
		Else
			pcost = PathCost((txtMultOrigin.Text), (txtMultDest.Text), EmpCommon.enumMobType.MT_COMMODITY)
			If pcost > 0 Then
				'lblBestPath.Caption = "path: " + txtPath.Text
				lblPathCost.Text = "path cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
				'If chkDisplayPath.Value = vbChecked Then
				'    DisplayPath txtMultOrigin.Text, txtPath.Text
				'End If
			End If
		End If
		
		'Compute the mobility cost
		If ParseSectors(sx, sy, (txtMultOrigin.Text)) = False Or pcost <= 0 Then Exit Sub
		
		mcost = 0
		'Get Sector Information
		rsSectors.Seek("=", sx, sy)
		If rsSectors.NoMatch Then Exit Sub
		
		If rsSectors.Fields("eff").Value < 60 Then
			pbdes = vbNullString
		Else
			pbdes = rsSectors.Fields("des").Value
		End If
		
		For i = 0 To 13
			If Val(txtThresh(i).Text) > 0 Then
				mcost = mcost + pcost * MobilityWeight(VB.Left(lblThresh(i).Text, 1), Val(txtThresh(i).Text), pbdes)
			End If
		Next 
		
		lblMcost.Text = "mob. cost: " & CStr(CSng(CInt(CSng(mcost) * 1000) / 1000))
		If mcost > rsSectors.Fields("mob").Value Then
			lblMobLeft.Text = "Insuff. mobility"
		Else
			lblMobLeft.Text = "mob. left: " & CStr(Int(rsSectors.Fields("mob").Value - mcost))
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtThresh.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtThresh_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThresh.TextChanged
		Dim Index As Short = txtThresh.GetIndex(eventSender)
		ComputePathCost()
	End Sub
	
	Private Sub txtThresh_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThresh.DoubleClick
		Dim Index As Short = txtThresh.GetIndex(eventSender)
		Dim X As Short
		Dim Y As Short
		
		'put in the max in a sector
		X = InStr(lblThresh(Index).Text, "(")
		If X > 0 Then
			Y = InStr(lblThresh(Index).Text, ")")
			If Y > 0 Then
				'100903 rjk: Reduce the civilian value by one so the sector is not abandoned when double clicking.
				If InStr(lblThresh(Index).Text, "civilians") > 0 And (strCmd = "move" Or strCmd = "load" Or strCmd = "lload") Then
					If Val(Mid(lblThresh(Index).Text, X + 1, Y - X - 1)) > 1 Then
						txtThresh(Index).Text = CStr(Val(Mid(lblThresh(Index).Text, X + 1, Y - X - 1)) - 1)
					Else
						txtThresh(Index).Text = "0"
					End If
				Else
					txtThresh(Index).Text = Mid(lblThresh(Index).Text, X + 1, Y - X - 1)
				End If
			End If
		End If
		
		ComputePathCost()
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		'New unit - recalc whats on it for labels
		'If Not (strCmd = "load" Or strCmd = "lload") Then 092103 rjk: SetShipLabel now understands load and lload as well
		SetShipLabels()
		'End If
		
		'Only allow OK if a ship is selected
		cmdOK.Enabled = (Len(txtUnit.Text) > 0)
		
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		Select Case strCmd
			Case "load", "unload", "tend", "ltend"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "lload", "lunload"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "give", "setsector", "setresource", "move"
				frmDrawMap.DisplayFirstSectorPanel()
		End Select
	End Sub
	
	Private Sub txtUnit3_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit3.Enter
		Select Case strCmd
			Case "load", "unload", "tend"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "lload", "lunload", "ltend"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "give", "setsector", "setresource", "move"
				frmDrawMap.DisplayFirstSectorPanel()
		End Select
	End Sub
	
	Private Sub txtUnit2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit2.Enter
		Dim utType As frmDrawMap.enumUnitType
		
		If strCmd = "load" Then
			utType = frmDrawMap.enumUnitType.TYPE_SHIP_UNLOADED
		ElseIf strCmd = "unload" Then 
			utType = frmDrawMap.enumUnitType.TYPE_SHIP_LOADED
		ElseIf strCmd = "lload" Then 
			utType = frmDrawMap.enumUnitType.TYPE_LAND_UNLOADED
		ElseIf strCmd = "lunload" Then 
			utType = frmDrawMap.enumUnitType.TYPE_LAND_LOADED
		Else
			utType = frmDrawMap.enumUnitType.TYPE_ALL
		End If
		If Option1.Checked Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, utType, False, False, False)
		ElseIf Option2.Checked Then 
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, utType, False, False, False)
		End If
	End Sub
	
	Private Sub SetDeityLevel(ByRef strCmd As String, ByRef strOrigin As String, ByRef strItem As String, ByRef strQuantity As String)
		Dim strCmd2 As String
		
		strCmd2 = Trim(strCmd & " " & strItem & " " & strOrigin & " " & strQuantity)
		frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
	End Sub
	
	Private Sub LoadCommodity(ByRef strCmd As String, ByRef strOrigin As String, ByRef strUnits As String, ByRef strItem As String, ByRef strQuantity As String)
		Dim strCmd2 As String
		Dim strUnit As String
		'UPGRADE_WARNING: Arrays in structure rsTemp may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rsTemp As DAO.Recordset
		Dim AdjustedQuantity As String
		Dim bLast As Boolean
		Dim hldIndex As String
		
		'check for ship or land, and choose the correct recordset
		If strCmd = "lload" Or strCmd = "lunload" Then
			rsTemp = rsLand
			hldIndex = rsLand.Index
			rsTemp.Index = "PrimaryKey"
		Else
			rsTemp = rsShip
			hldIndex = rsShip.Index
			rsTemp.Index = "PrimaryKey"
		End If
		
		Do 
			If InStr(strUnits, "/") <> 0 Then
				strUnit = VB.Left(strUnits, InStr(strUnits, "/") - 1)
				strUnits = Mid(strUnits, InStr(strUnits, "/") + 1)
				bLast = False
			Else
				strUnit = strUnits
				bLast = True
			End If
			rsTemp.Seek("=", strUnit)
			If Not rsTemp.NoMatch Then
				If strCmd = "load" Or strCmd = "lload" Then
					Select Case VB.Left(strQuantity, 1)
						Case "-"
							rsSectors.Seek("=", rsTemp.Fields("x").Value, rsTemp.Fields("y").Value)
							If Not rsSectors.NoMatch Then
								If (Items.FindByLetter(strItem).DatabaseValue(rsSectors) - System.Math.Abs(Val(strQuantity))) > 0 Then
									AdjustedQuantity = CStr(Items.FindByLetter(strItem).DatabaseValue(rsSectors) - System.Math.Abs(Val(strQuantity)))
								Else
									AdjustedQuantity = "0"
								End If
							Else
								MsgBox("Ship or Land unit not in a loadable sector")
								Exit Sub
							End If
						Case "+"
							If (System.Math.Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsTemp)) > 0 Then
								AdjustedQuantity = CStr(System.Math.Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsTemp))
							Else
								AdjustedQuantity = "0"
							End If
					End Select
				ElseIf strCmd = "unload" Or strCmd = "lunload" Then 
					Select Case VB.Left(strQuantity, 1)
						Case "-"
							If (Items.FindByLetter(VB.Left(strItem, 1)).DatabaseValue(rsTemp) - System.Math.Abs(Val(strQuantity))) > 0 Then
								AdjustedQuantity = CStr(Items.FindByLetter(strItem).DatabaseValue(rsTemp) - System.Math.Abs(Val(strQuantity)))
							Else
								AdjustedQuantity = "0"
							End If
						Case "+"
							rsSectors.Seek("=", rsTemp.Fields("x").Value, rsTemp.Fields("y").Value)
							If Not rsSectors.NoMatch Then
								If (System.Math.Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsSectors)) > 0 Then
									AdjustedQuantity = CStr(System.Math.Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsSectors))
								Else
									AdjustedQuantity = "0"
								End If
							Else
								MsgBox("Ship or Land unit not in a loadable sector")
								Exit Sub
							End If
					End Select
				End If
				strCmd2 = Trim(strCmd & " " & strItem & " " & strUnit & " " & AdjustedQuantity)
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			End If
		Loop While Not bLast
		
	End Sub
End Class