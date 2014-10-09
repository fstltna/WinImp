Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptExportIntelligence
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'103003 rjk: Created.  Added parameters to ExportIntelligence.
	'            Moved ExportIntelligence to the form.
	'110403 rjk: Reorganized SendTelegram to send headers for each telegram and to deal with long lines better.
	'111203 rjk: Added some safety checks before sending the telegram.
	'111403 rjk: Do not send enemy intelligence about a sector we now own.
	'111803 rjk: Added Header and control for it.
	'            Added Data field selection and expand the sectors to dump everything possible.
	'111903 rjk: Added SetAll and ClearAll for data fields.
	'070304 rjk: Switch Enemy timestamp to UTC
	'110304 rjk: Replace null Enemy Sector designation with BMap if known otherwise "?"
	'090607 rjk: Remove function call chkComm_Click.
	'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.
	
	Dim filenumber As Short
	Dim FileName As String
	Dim strTelegram As String
	
	Enum enumSectorSearch
		ssFound
		ssNotFound
		ssError
	End Enum
	
	'UPGRADE_WARNING: Event cbTelegram.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbTelegram_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbTelegram.SelectedIndexChanged
		If cbTelegram.Text <> "" Then
			cmdOK.Enabled = True
		Else
			cmdOK.Enabled = False
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdClearAllData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearAllData.Click
		chkData(0).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(1).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(4).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(5).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(6).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(7).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(8).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(9).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(10).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(11).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(12).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(13).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(14).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(15).CheckState = System.Windows.Forms.CheckState.Unchecked
		chkData(16).CheckState = System.Windows.Forms.CheckState.Unchecked
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		ExportIntelligence()
		Me.Close()
	End Sub
	
	Private Sub cmdSetAllData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSetAllData.Click
		chkData(0).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(1).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(4).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(5).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(6).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(7).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(8).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(9).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(10).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(11).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(12).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(13).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(14).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(15).CheckState = System.Windows.Forms.CheckState.Checked
		chkData(16).CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub frmPromptExportIntelligence_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Index As Short
		
		Nations.FillListBox(cbNations)
		Nations.FillListBox(cbTelegram)
		
		For Index = 0 To cbNations.Items.Count - 1
			If VB6.GetItemData(cbNations, Index) = CountryNumber Then
				cbNations.Items.RemoveAt((Index))
				Exit For
			End If
		Next Index
		
		cbNations.Items.Insert(0, New VB6.ListBoxItem("All", -1))
		cbNations.SelectedIndex = 0
		
		cmdOK.Enabled = False 'Enabled either file is selected or destination for telegram is selected.
		
		
		'Put form in proper place
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
	End Sub
	
	Private Sub frmPromptExportIntelligence_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event optDestination.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDestination_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDestination.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optDestination.GetIndex(eventSender)
			If optDestination(0).Checked Then
				cmdOK.Enabled = True
				cbTelegram.Enabled = False
			ElseIf optDestination(1).Checked Then 
				cbTelegram.Enabled = True
				If cbTelegram.Text <> "" Then
					cmdOK.Enabled = True
				Else
					cmdOK.Enabled = False
				End If
			End If
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
	Public Sub ExportIntelligence()
		On Error GoTo Empire_Error
		
		
		Dim n As Short
		Dim nSect As Short
		Dim nUnit As Short
		Dim secx As Short
		Dim secy As Short
		Dim searchResult As enumSectorSearch
		
		filenumber = -1
		
		'111203 rjk: Do some safety checks before sending the telegram.
		If optDestination(1).Checked Then 'telegram
			If Nations.YourRelation(CShort(CStr(VB6.GetItemData(cbTelegram, cbTelegram.SelectedIndex)))) <> iREL_ALLIED And Nations.YourRelation(CShort(CStr(VB6.GetItemData(cbTelegram, cbTelegram.SelectedIndex)))) <> iREL_FRIENDLY Then
				If MsgBox("This export is not being sent to an ally or an friend, do you what to continue", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
					Exit Sub
				End If
			End If
			If VB6.GetItemData(cbNations, cbNations.SelectedIndex) = -1 And txtMultOrigin.Text = "*" Then
				If MsgBox("You have selected to export all your enemy intelligence about the items selected, is not correct?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
					Exit Sub
				End If
			End If
			If Nations.YourRelation(CShort(CStr(VB6.GetItemData(cbTelegram, cbTelegram.SelectedIndex)))) <> iREL_ALLIED And (chkItems(5).CheckState Or chkItems(6).CheckState Or chkItems(7).CheckState Or chkItems(8).CheckState Or chkItems(9).CheckState) Then
				If MsgBox("You are sending your own sector information to non-ally, do you what to continue", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
					Exit Sub
				End If
			End If
		End If
		
		If Not OpenDestination Then
			Exit Sub
		End If
		
		If chkHeader.CheckState = System.Windows.Forms.CheckState.Checked Then '111803 rjk: Added header option
			OutputIntelligence("+++ Sector Header x,y,des,owner,oldowner,eff,road,rail,defense,civ,mil,shell,gun,pet,food,iron,bar,timestamp", True)
		End If
		
		'now write out the sector file
		If chkItems(0).CheckState Then
			If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
				rsEnemySect.MoveFirst()
				While Not rsEnemySect.EOF
					If (VB6.GetItemData(cbNations, cbNations.SelectedIndex) = -1) Or (VB6.GetItemData(cbNations, cbNations.SelectedIndex) = rsEnemySect.Fields("owner").Value) Then
						If (InStr(txtDesignations.Text, rsEnemySect.Fields("des").Value) > 0) Or (Len(txtDesignations.Text) = 0) Then
							secx = rsEnemySect.Fields("x").Value
							secy = rsEnemySect.Fields("y").Value
							searchResult = SectorWithInSectorList(secx, secy)
							Select Case searchResult
								Case enumSectorSearch.ssFound
									rsSectors.Seek("=", secx, secy)
									If rsSectors.NoMatch Then '111403 rjk: Do not send enemy intelligence about a sector we now own.
										OffsetSectorLocation(secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text))
										'Write #filenumber, secX;
										OutputIntelligence(CStr(secx), False)
										'Write #filenumber, secY;
										OutputIntelligence(CStr(secy), False)
										nSect = nSect + 1
										For n = 2 To rsEnemySect.Fields.Count - 2
											'Write #filenumber, rsEnemySect.Fields(n);
											If n >= 5 And n <= 16 Then
												If chkData(n).CheckState = System.Windows.Forms.CheckState.Checked Then
													OutputIntelligence(rsEnemySect.Fields(n), False)
												Else
													OutputIntelligence("-1", False)
												End If
											ElseIf n = 2 Then 
												'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
												If IsDbNull(rsEnemySect.Fields(2).Value) Then
													rsBmap.Seek("=", secx, secy)
													If rsBmap.NoMatch Then
														OutputIntelligence("?", False)
													Else
														OutputIntelligence(rsBmap.Fields("des"), False)
													End If
												Else
													OutputIntelligence(rsEnemySect.Fields(n), False)
												End If
											Else
												OutputIntelligence(rsEnemySect.Fields(n), False)
											End If
										Next n
										'Write #filenumber, rsEnemySect.Fields(rsEnemySect.Fields.Count - 1)
										OutputIntelligence(rsEnemySect.Fields(rsEnemySect.Fields.Count - 1), True)
									End If
								Case enumSectorSearch.ssError
									CloseDestination(False)
									Exit Sub
							End Select
						End If
					End If
					rsEnemySect.MoveNext()
				End While
			End If
		End If
		If chkItems(5).CheckState Then
			If Not (rsSectors.BOF And rsSectors.EOF) Then
				rsSectors.MoveFirst()
				While Not rsSectors.EOF
					If (InStr(txtDesignations.Text, rsSectors.Fields("des").Value) > 0) Or (Len(txtDesignations.Text) = 0) Then
						secx = rsSectors.Fields("x").Value
						secy = rsSectors.Fields("y").Value
						searchResult = SectorWithInSectorList(secx, secy)
						Select Case searchResult
							Case enumSectorSearch.ssFound
								OffsetSectorLocation(secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text))
								nSect = nSect + 1
								OutputIntelligence(CStr(secx), False)
								OutputIntelligence(CStr(secy), False)
								OutputIntelligence(rsSectors.Fields("des"), False)
								OutputIntelligence(rsSectors.Fields("country"), False)
								OutputIntelligence("-1", False)
								If chkData(4).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("eff"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(6).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("road"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(7).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("rail"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(8).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("defense"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(9).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("civ"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(10).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("mil"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(11).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("shell"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(12).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("gun"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(13).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("pet"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(14).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("food"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(15).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("iron"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(16).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsSectors.Fields("bar"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								OutputIntelligence(GetWinACEUTC, True)
							Case enumSectorSearch.ssError
								CloseDestination(False)
								Exit Sub
						End Select
					End If
					rsSectors.MoveNext()
				End While
			End If
		End If
		OutputIntelligence("+++ " & CStr(nSect) & " Sectors exported", True)
		
		If chkHeader.CheckState = System.Windows.Forms.CheckState.Checked Then '111803 rjk: Added header option
			OutputIntelligence("+++ Unit Header id,type,x,y,owner,eff,mil,tech,timestamp,wake", True)
		End If
		
		'now write out the unit file
		If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
			rsEnemyUnit.MoveFirst()
			While Not rsEnemyUnit.EOF
				If (rsEnemyUnit.Fields("class").Value = "s" And chkItems(1).CheckState) Or (rsEnemyUnit.Fields("class").Value = "p" And chkItems(2).CheckState) Or (rsEnemyUnit.Fields("class").Value = "l" And chkItems(3).CheckState) Then
					If (VB6.GetItemData(cbNations, cbNations.SelectedIndex) = -1) Or (VB6.GetItemData(cbNations, cbNations.SelectedIndex) = rsEnemyUnit.Fields("owner").Value) Then
						secx = rsEnemyUnit.Fields("x").Value
						secy = rsEnemyUnit.Fields("y").Value
						searchResult = SectorWithInSectorList(secx, secy)
						Select Case searchResult
							Case enumSectorSearch.ssFound
								OffsetSectorLocation(secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text))
								nUnit = nUnit + 1
								OutputIntelligence(rsEnemyUnit.Fields("id"), False)
								OutputIntelligence(rsEnemyUnit.Fields("type"), False)
								OutputIntelligence(CStr(secx), False)
								OutputIntelligence(CStr(secy), False)
								OutputIntelligence(rsEnemyUnit.Fields("owner"), False)
								If chkData(5).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsEnemyUnit.Fields("eff"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(10).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsEnemyUnit.Fields("mil"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								If chkData(1).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsEnemyUnit.Fields("tech"), False)
								Else
									OutputIntelligence("-1", False)
								End If
								OutputIntelligence(rsEnemyUnit.Fields("timestamp"), False)
								OutputIntelligence(rsEnemyUnit.Fields("class"), False)
								If chkData(1).CheckState = System.Windows.Forms.CheckState.Checked Then
									OutputIntelligence(rsEnemyUnit.Fields("wake"), True)
								Else
									OutputIntelligence("", True)
								End If
							Case enumSectorSearch.ssError
								CloseDestination(False)
								Exit Sub
						End Select
					End If
				End If
				rsEnemyUnit.MoveNext()
			End While
		End If
		
		'now write out the our ship file
		If chkItems(6).CheckState Then
			If Not (rsShip.BOF And rsShip.EOF) Then
				rsShip.MoveFirst()
				While Not rsShip.EOF
					secx = rsShip.Fields("x").Value
					secy = rsShip.Fields("y").Value
					searchResult = SectorWithInSectorList(secx, secy)
					Select Case searchResult
						Case enumSectorSearch.ssFound
							OffsetSectorLocation(secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text))
							nUnit = nUnit + 1
							OutputIntelligence(rsShip.Fields("id"), False)
							OutputIntelligence(rsShip.Fields("type"), False)
							OutputIntelligence(CStr(secx), False)
							OutputIntelligence(CStr(secy), False)
							OutputIntelligence(rsShip.Fields("country"), False)
							If chkData(5).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsShip.Fields("eff"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							If chkData(10).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsShip.Fields("mil"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							If chkData(1).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsShip.Fields("tech"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							OutputIntelligence(GetWinACEUTC, False)
							OutputIntelligence("s", False) 'Class
							OutputIntelligence("", True) 'Wake
						Case enumSectorSearch.ssError
							CloseDestination(False)
							Exit Sub
					End Select
					rsShip.MoveNext()
				End While
			End If
		End If
		
		'now write out the our plane file
		If chkItems(7).CheckState Then
			If Not (rsPlane.BOF And rsPlane.EOF) Then
				rsPlane.MoveFirst()
				While Not rsPlane.EOF
					secx = rsPlane.Fields("x").Value
					secy = rsPlane.Fields("y").Value
					searchResult = SectorWithInSectorList(secx, secy)
					Select Case searchResult
						Case enumSectorSearch.ssFound
							OffsetSectorLocation(secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text))
							nUnit = nUnit + 1
							OutputIntelligence(rsPlane.Fields("id"), False)
							OutputIntelligence(rsPlane.Fields("type"), False)
							OutputIntelligence(CStr(secx), False)
							OutputIntelligence(CStr(secy), False)
							OutputIntelligence(rsPlane.Fields("country"), False)
							If chkData(5).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsPlane.Fields("eff"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							OutputIntelligence("-1", False)
							If chkData(1).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsPlane.Fields("tech"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							OutputIntelligence(GetWinACEUTC, False)
							OutputIntelligence("p", False) 'Class
							OutputIntelligence("", True) 'Wake
						Case enumSectorSearch.ssError
							CloseDestination(False)
							Exit Sub
					End Select
					rsPlane.MoveNext()
				End While
			End If
		End If
		
		'now write out the our land unit file
		If chkItems(8).CheckState Then
			If Not (rsLand.BOF And rsLand.EOF) Then
				rsLand.MoveFirst()
				While Not rsLand.EOF
					secx = rsLand.Fields("x").Value
					secy = rsLand.Fields("y").Value
					searchResult = SectorWithInSectorList(secx, secy)
					Select Case searchResult
						Case enumSectorSearch.ssFound
							OffsetSectorLocation(secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text))
							nUnit = nUnit + 1
							OutputIntelligence(rsLand.Fields("id"), False)
							OutputIntelligence(rsLand.Fields("type"), False)
							OutputIntelligence(CStr(secx), False)
							OutputIntelligence(CStr(secy), False)
							OutputIntelligence(rsLand.Fields("country"), False)
							If chkData(5).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsLand.Fields("eff"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							If chkData(10).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsLand.Fields("mil"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							If chkData(1).CheckState = System.Windows.Forms.CheckState.Checked Then
								OutputIntelligence(rsLand.Fields("tech"), False)
							Else
								OutputIntelligence("-1", False)
							End If
							OutputIntelligence(GetWinACEUTC, False)
							OutputIntelligence("l", False) 'Class
							OutputIntelligence("", True) 'Wake
						Case enumSectorSearch.ssError
							CloseDestination(False)
							Exit Sub
					End Select
					rsLand.MoveNext()
				End While
			End If
		End If
		
		OutputIntelligence("+++ " & CStr(nUnit) & " Units exported", True)
		CloseDestination(True)
		
		If optDestination(0).Checked Then 'file
			MsgBox(CStr(nSect) & " Sectors, " & CStr(nUnit) & " Units exported to file " & FileName)
		ElseIf optDestination(1).Checked Then  'telegram
			MsgBox(CStr(nSect) & " Sectors, " & CStr(nUnit) & " Units exported to " & cbTelegram.Text)
		End If
		
		Exit Sub
Empire_Error: 
		CloseDestination(False)
		EmpireError("ExportIntelligence", vbNullString, vbNullString)
	End Sub
	
	
	Private Function OpenDestination() As Boolean
		If optDestination(0).Checked Then 'file
			If VerifySubDirectory("Exports", True) Then
				FileName = My.Application.Info.DirectoryPath & "\Exports\" & GameName & " Intelligence.txt"
			Else
				FileName = My.Application.Info.DirectoryPath & "\" & GameName & " Intelligence.txt"
			End If
			filenumber = FreeFile ' Get unused file number.
			FileOpen(filenumber, FileName, OpenMode.Output)
			OpenDestination = True
		ElseIf optDestination(1).Checked Then  'telegram
			strTelegram = vbNullString
			OpenDestination = True
		End If
	End Function
	
	Private Sub OutputIntelligence(ByRef vField As Object, ByRef bEOL As Boolean)
		If optDestination(0).Checked Then 'file
			If filenumber <> -1 Then
			End If
			If bEOL Then
				WriteLine(filenumber, vField)
			Else
				Write(filenumber, vField)
			End If
		ElseIf optDestination(1).Checked Then  'telegram
			'UPGRADE_WARNING: Couldn't resolve default property of object vField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strTelegram = strTelegram & Chr(34) & CStr(vField) & Chr(34)
			If bEOL Then
				strTelegram = strTelegram & vbLf
			Else
				strTelegram = strTelegram & ","
			End If
		End If
	End Sub
	
	Private Function CloseDestination(ByRef bSend As Object) As Boolean
		If optDestination(0).Checked Then 'file
			FileClose(filenumber)
		ElseIf optDestination(1).Checked Then  'telegram
			'UPGRADE_WARNING: Couldn't resolve default property of object bSend. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If bSend Then
				SendTelegram((strTelegram))
			End If
			strTelegram = vbNullString
		End If
	End Function
	
	Private Function SectorWithInSectorList(ByRef secx As Short, ByRef secy As Short) As enumSectorSearch
		Dim sx As Short
		Dim sy As Short
		Dim beginPos As Short
		Dim endPos As Short
		
		SectorWithInSectorList = enumSectorSearch.ssError
		If Len(txtMultOrigin.Text) = 0 Then
			Exit Function
		ElseIf Len(txtMultOrigin.Text) = 1 And txtMultOrigin.Text = "*" Then 
			SectorWithInSectorList = enumSectorSearch.ssFound
		ElseIf InStr(txtMultOrigin.Text, ":") Then 
			SectorWithInSectorList = SectorWithInRange((txtMultOrigin.Text), secx, secy)
		ElseIf InStr(txtMultOrigin.Text, "#") > 0 Then 
			MsgBox("Realms not support in the sector box")
		ElseIf InStr(txtMultOrigin.Text, "?") > 0 Then 
			MsgBox("Conditionals not support in the sector box")
		ElseIf InStr(txtMultOrigin.Text, "\") Then 
			beginPos = 1
			endPos = InStr(txtMultOrigin.Text, "\")
			While endPos > 0
				If ParseSectors(sx, sy, Mid(txtMultOrigin.Text, beginPos, endPos - beginPos)) Then
					If sx = secx And sy = secy Then
						SectorWithInSectorList = enumSectorSearch.ssFound
						Exit Function
					End If
				Else
					MsgBox("Invalid sector format")
					Exit Function
				End If
				beginPos = endPos + 1
				endPos = InStr(beginPos, txtMultOrigin.Text, "\")
			End While
			If ParseSectors(sx, sy, Mid(txtMultOrigin.Text, beginPos)) Then
				If sx = secx And sy = secy Then
					SectorWithInSectorList = enumSectorSearch.ssFound
					Exit Function
				End If
			Else
				MsgBox("Invalid sector format")
				Exit Function
			End If
			SectorWithInSectorList = enumSectorSearch.ssNotFound
		ElseIf InStr(txtMultOrigin.Text, ",") Then 
			If ParseSectors(sx, sy, (txtMultOrigin.Text)) Then
				If sx = secx And sy = secy Then
					SectorWithInSectorList = enumSectorSearch.ssFound
				Else
					SectorWithInSectorList = enumSectorSearch.ssNotFound
				End If
			Else
				MsgBox("Invalid sector format")
			End If
		Else
			MsgBox("Invalid sector format")
		End If
	End Function
	
	Private Function SectorWithInRange(ByRef strSector As String, ByRef secx As Short, ByRef secy As Short) As enumSectorSearch
		Dim sx1 As Short
		Dim sx2 As Short
		Dim sy1 As Short
		Dim sy2 As Short
		
		SectorWithInRange = enumSectorSearch.ssNotFound
		
		If ParseSectorRange(sx1, sx2, sy1, sy2, strSector) Then
			If sx1 <= sx2 Then
				If sy1 <= sy2 Then
					If (secx >= sx1 And secx <= sx2) And (secy >= sy1 And secy <= sy2) Then
						SectorWithInRange = enumSectorSearch.ssFound
					End If
				Else
					If (secx >= sx1 And secx <= sx2) And (secy >= sy1 Or secy <= sy2) Then
						SectorWithInRange = enumSectorSearch.ssFound
					End If
				End If
			Else
				If sy1 <= sy2 Then
					If (secx >= sx1 Or secx <= sx2) And (secy >= sy1 And secy <= sy2) Then
						SectorWithInRange = enumSectorSearch.ssFound
					End If
				Else
					If (secx >= sx1 Or secx <= sx2) And (secy >= sy1 Or secy <= sy2) Then
						SectorWithInRange = enumSectorSearch.ssFound
					End If
				End If
			End If
		Else
			MsgBox("Invalid sector range format")
			SectorWithInRange = enumSectorSearch.ssError
		End If
	End Function
	
	Private Sub SendTelegram(ByRef strTelegram As String)
		Dim numTelegrams As Short
		Dim strTarget() As String
		Dim n As Short
		Dim strTemp As String
		Dim strSubject As String
		
		strSubject = "+++ Intelligence Report " & VB6.Format(Now, "ddd mmm dd hh:mm:ss yyyy") & vbLf
		'UPGRADE_WARNING: Lower bound of array strTarget was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim strTarget(1)
		numTelegrams = 1
		strTarget(numTelegrams) = strSubject
		Do 
			n = InStr(strTelegram, vbLf)
			If n = 0 Then
				strTemp = strTelegram
				strTelegram = ""
			Else
				strTemp = VB.Left(strTelegram, n - 1)
				strTelegram = Mid(strTelegram, n + 1)
			End If
			
			'just in case we get a real long line, split it up
			
			If Len(strTemp) > (1021 - Len(strSubject)) Then
				MsgBox("Error in generating telegram - line too long - telegram aborted")
				Exit Sub
			End If
			
			If (Len(strTarget(numTelegrams)) + Len(strTemp)) >= 1022 Then
				numTelegrams = numTelegrams + 1
				'UPGRADE_WARNING: Lower bound of array strTarget was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve strTarget(numTelegrams)
				strTarget(numTelegrams) = strSubject
			End If
			
			strTarget(numTelegrams) = strTarget(numTelegrams) & strTemp & vbLf
		Loop While Len(strTelegram) > 0
		
		'run thru multiple telegrams if necessay
		For n = 1 To numTelegrams
			frmEmpCmd.SubmitEmpireCommand("te1 " & strTarget(n), False)
			frmEmpCmd.SubmitEmpireCommand("telegram " & CStr(VB6.GetItemData(cbTelegram, cbTelegram.SelectedIndex)), False)
		Next n
	End Sub
End Class