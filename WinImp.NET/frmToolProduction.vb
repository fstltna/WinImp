Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmToolProduction
	Inherits System.Windows.Forms.Form
	
	'150304 rjk: Added from Jim Original code
	'240304 rjk: Limited the product value to 9999, change values to long to prevent overflow,
	'            Redesigned form, removed button and text boxes and replaces with label.
	'100404 rjk: Added 'e' enlist to Estimate Tool.
	
	Dim secx As Short
	Dim secy As Short
	Dim OriginChange As Boolean
	Dim pe As Single
	Dim ColItem As New Collection
	Dim newdes As String
	Dim v As Object
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	Private Sub frmToolProduction_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		Dim success As Integer
		success = SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		LoadItems()
	End Sub
	
	Private Sub frmToolProduction_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		
		'use origin change to avoid firing cascading events
		OriginChange = True
		
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			secx = sx
			secy = sy
			FillSectorInfo()
		Else
			HandleError("Not a valid Sector")
		End If
		
		OriginChange = False
		
	End Sub
	
	Private Sub FillSectorInfo()
		Dim n As Short
		Dim strTemp As String
		Dim v As productionDataType
		
		'get sector record
		rsSectors.Seek("=", secx, secy)
		If rsSectors.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		v = Production(rsSectors)
		'1 = new efficency
		'2 = production amount
		'3 = max production amount
		'4 = excess civ
		'5 = new designation
		'6 = units consumed
		'7 = item produced
		'8 = build cost in avail
		'9 = resource efficency
		'10 = level efficency
		'11 = production avail
		
		'If IsArray(v) Then
		If (v.prodValidFlag) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			newdes = CStr(v.newdes)
		Else
			HandleError("Cannot Calculate Production")
			Exit Sub
		End If
		
		'get sector type
		rsSectorType.Seek("=", newdes)
		If rsSectorType.NoMatch Then
			HandleError("Sector Type not found")
			Exit Sub
		End If
		
		For n = 1 To 3
			lblMaterial(n).Visible = False
			lblRequired(n).Visible = False
			lblResource(n).Text = ""
		Next n
		
		Line3.Visible = False
		Line4.Visible = False
		Line5.Visible = False
		Line6.Visible = False
		Line7.Visible = False
		Line8.Visible = False
		Line11.Visible = False
		Line12.Visible = False
		
		cmdPE.Visible = True
		lblProduct.Visible = True
		
		'set labels
		If rsSectorType.Fields("use1n").Value > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ColItem.item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strTemp = ColItem.Item(rsSectorType.Fields("use1s"))
			lblMaterial(1).Text = strTemp
			lblMaterial(1).Visible = True
			lblRequired(1).Visible = True
			lblResource(1).Text = CStr(rsSectors.Fields(strTemp).Value)
			Line3.Visible = True
			Line4.Visible = True
			Line11.Visible = True
			Line12.Visible = True
			
			If rsSectorType.Fields("use2n").Value > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ColItem.item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strTemp = ColItem.Item(rsSectorType.Fields("use2s"))
				lblMaterial(2).Text = strTemp
				lblMaterial(2).Visible = True
				lblRequired(2).Visible = True
				lblResource(2).Text = CStr(rsSectors.Fields(strTemp).Value)
				Line5.Visible = True
				Line6.Visible = True
				
				If rsSectorType.Fields("use3n").Value > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ColItem.item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strTemp = ColItem.Item(rsSectorType.Fields("use3s"))
					lblMaterial(3).Text = strTemp
					lblMaterial(3).Visible = True
					lblRequired(3).Visible = True
					lblResource(3).Text = CStr(rsSectors.Fields(strTemp).Value)
					Line7.Visible = True
					Line8.Visible = True
				End If
			End If
		End If
		
		lblType.Text = CStr(v.NewEff) & "% " + rsSectorType.Fields("desc").Value
		
		lblProduct.Text = Trim(rsSectorType.Fields("product").Value)
		If lblProduct.Text = "" Then lblProduct.Text = "avail"
		
		If newdes = "e" Then
			strTemp = "civ"
			lblMaterial(1).Text = strTemp
			lblMaterial(1).Visible = True
			lblRequired(1).Visible = True
			lblResource(1).Text = CStr(rsSectors.Fields(strTemp).Value)
			Line3.Visible = True
			Line4.Visible = True
			Line11.Visible = True
			Line12.Visible = True
			strTemp = "mil"
			lblMaterial(2).Text = strTemp
			lblMaterial(2).Visible = True
			lblRequired(2).Visible = True
			lblResource(2).Text = CStr(rsSectors.Fields(strTemp).Value)
			Line5.Visible = True
			Line6.Visible = True
			lblProduct.Text = "mil"
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object v.LevelEff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object v.ResourceEff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pe = (v.NewEff / 100) * v.ResourceEff * v.LevelEff
		cmdPE.Text = "PE: " & VB6.Format(CShort(v.NewEff * pe), "##0")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object v.ProductionAvail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lblResource(0).Text = CStr(v.ProductionAvail)
		'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		txtProduce.Text = CStr(v.ProdAmount)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object v.BuildAvailCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If v.BuildAvailCost > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object v.BuildAvailCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rtbReport.Text = CStr(v.BuildAvailCost) & " avail used to raise eff. to " & CStr(v.NewEff) & "%"
		Else
			rtbReport.Text = ""
		End If
	End Sub
	
	Private Sub HandleError(ByRef ErrMsg As String)
		Dim n As Short
		
		lblType.Text = ErrMsg
		Line3.Visible = False
		Line4.Visible = False
		Line5.Visible = False
		Line6.Visible = False
		Line7.Visible = False
		Line8.Visible = False
		Line11.Visible = False
		Line12.Visible = False
		cmdPE.Visible = False
		'Slider1.Visible = False
		lblProduct.Visible = False
		For n = 1 To 3
			lblMaterial(n).Visible = False
			lblRequired(n).Visible = False
		Next n
		
	End Sub
	
	Private Sub LoadItems()
		
		If ColItem.Count() > 0 Then
			Exit Sub
		End If
		
		With ColItem
			.Add("bar", "b")
			.Add("civ", "c")
			.Add("dust", "d")
			.Add("food", "f")
			.Add("gun", "g")
			.Add("hcm", "h")
			.Add("iron", "i")
			.Add("lcm", "l")
			.Add("mil", "m")
			.Add("oil", "o")
			.Add("pet", "p")
			.Add("rad", "r")
			.Add("shell", "s")
			.Add("uw", "u")
		End With
		
	End Sub
	
	'UPGRADE_WARNING: Event txtProduce.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtProduce_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProduce.TextChanged
		
		Dim units As Integer
		Dim avail As Integer
		Dim q1 As Integer
		Dim q2 As Integer
		Dim q3 As Integer
		Dim n As Integer
		Dim buildavail As Integer
		Dim NewEff As Short
		Dim new_pe As Single
		Dim save_Units As Short
		Dim Done As Boolean
		
		
		For n = 0 To 3
			lblRequired(n).Text = ""
		Next n
		
		rsSectors.Seek("=", secx, secy)
		rsSectorType.Seek("=", newdes)
		
		If rsSectors.NoMatch Or rsSectorType.NoMatch Then
			Exit Sub
		End If
		
		If Val(txtProduce.Text) <= 0 Then
			Exit Sub
		End If
		
		If Len(Trim(rsSectors.Fields("sdes").Value)) > 0 And rsSectors.Fields("eff").Value > 0 Then
			buildavail = 60 + Int((rsSectors.Fields("eff").Value + 3) / 4)
			NewEff = 60
		Else
			If rsSectors.Fields("eff").Value < 60 Then
				buildavail = 60 - rsSectors.Fields("eff").Value
				NewEff = 60
			Else
				buildavail = 0
				NewEff = rsSectors.Fields("eff").Value
			End If
		End If
		
		Done = False
		save_Units = -99
		
		Dim mil_requested As Integer
		Dim mil_required As Integer
		Dim recruit_rate As Single
		Do 
			new_pe = pe * (NewEff / 100#)
			
			If new_pe <= 0 Then
				Done = True
				units = 0
			Else
				If Val(txtProduce.Text) > 9999 Then
					txtProduce.Text = "9999"
				End If
				units = CShort((Val(txtProduce.Text) / new_pe) + 0.4999)
			End If
			
			If units = save_Units Then
				'units required didn't change - we are done
				Exit Do
			End If
			
			'set text
			If rsSectorType.Fields("use1n").Value > 0 Then
				q1 = rsSectorType.Fields("use1n").Value * units
				lblRequired(1).Text = CStr(q1)
				If rsSectorType.Fields("use2n").Value > 0 Then
					q2 = rsSectorType.Fields("use2n").Value * units
					lblRequired(2).Text = CStr(q2)
					If rsSectorType.Fields("use3n").Value > 0 Then
						q3 = rsSectorType.Fields("use3n").Value * units
						lblRequired(3).Text = CStr(q3)
					End If
				End If
			End If
			
			avail = q1 + q2 + q3
			
			If avail < buildavail Then
				avail = 2 * buildavail
				Done = True
			ElseIf NewEff = 100 Then 
				avail = avail + buildavail
				Done = True
			Else
				'figure out how much avail will be used increasing the sector
				'and run the calculation again with the new efficency to see
				'if it lowers the units that have to be produced
				n = (avail - buildavail) / 2
				If n > (100 - NewEff) Then
					n = 100 - NewEff
				End If
				buildavail = buildavail + n
				avail = avail + buildavail
				NewEff = NewEff + n
				save_Units = units
				Done = False
			End If
			
			If newdes = "e" Then
				If NewEff >= 60 Then
					
					recruit_rate = (rsVersion.Fields("ETU per update").Value / 20)
					
					mil_requested = Val(txtProduce.Text)
					
					mil_required = mil_requested / recruit_rate
					
					If mil_required > 10 Then
						mil_required = mil_required - 10
						lblRequired(2).Text = CStr(mil_required)
					Else
						lblRequired(2).Text = CStr(0)
					End If
					lblRequired(1).Text = CStr(mil_requested * 2)
				Else
					lblRequired(2).Text = CStr(0)
					lblRequired(1).Text = CStr(0)
				End If
			End If
		Loop Until Done
		
		If rsSectors.Fields("work").Value < 100 Then
			avail = CShort((avail / (CSng(rsSectors.Fields("work").Value) / 100#)) + 0.4999)
		End If
		
		lblRequired(0).Text = CStr(avail)
		If buildavail > 0 Then
			rtbReport.Text = CStr(buildavail) & " avail used to raise eff. to " & CStr(NewEff) & "%"
		Else
			rtbReport.Text = ""
		End If
		
		lblType.Text = CStr(NewEff) & "% " + rsSectorType.Fields("desc").Value
		cmdPE.Text = "PE: " & VB6.Format(CShort(NewEff * pe), "##0")
		
	End Sub
End Class