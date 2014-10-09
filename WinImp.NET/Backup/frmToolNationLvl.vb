Option Strict Off
Option Explicit On
Friend Class frmToolNationLvl
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'251003 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
	'140506 rjk: Removed education from tech production for South Pacific Atlantis.
	
	Dim NoReset As Boolean
	Dim EstimateMaxPop As Single
	Dim turn As Short
	
	Private Sub cmdNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNext.Click
		Dim civs As Integer
		
		turn = turn + 1
		NoReset = True
		
		'Set labels with current values
		lblEd.Text = "Turn " & CStr(turn)
		txtEd.Text = txtNewEd.Text
		lblNewEd.Text = "Turn " & CStr(turn + 1)
		
		lblHappy.Text = "Turn " & CStr(turn)
		txtHappy.Text = txtNewHappy.Text
		lblNewHappy.Text = "Turn " & CStr(turn + 1)
		
		lblTech.Text = "Turn " & CStr(turn)
		txtTech.Text = txtNewTech.Text
		lblNewTech.Text = "Turn " & CStr(turn + 1)
		
		lblRes.Text = "Turn " & CStr(turn)
		txtRes.Text = txtNewRes.Text
		lblNewRes.Text = "Turn " & CStr(turn + 1)
		
		civs = CInt(Val(txtPop.Text) * (1 + ((rsVersion.Fields("ETU per update").Value * rsVersion.Fields("Civ Birthrate").Value) / 1000)) + 0.499)
		If civs > EstimateMaxPop Then
			civs = EstimateMaxPop
		End If
		
		txtPop.Text = CStr(civs)
		txtPop2.Text = CStr(civs)
		txtNewEd.Text = VB6.Format(CalcEducation, "###0.00")
		txtNewHappy.Text = VB6.Format(CalcHappy, "###0.00")
		txtNewTech.Text = VB6.Format(CalcTechnology, "###0.00")
		txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		
		NoReset = False
		cmdNext.Text = "Turn " & CStr(turn + 1)
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'Set values if they are not in use
		If Education < 0 Then
			Education = Val(txtEd.Text)
		End If
		If TechLevel < 0 Then
			TechLevel = Val(txtTech.Text)
		End If
		If Happiness < 0 Then
			Happiness = Val(txtHappy.Text)
		End If
		If Research < 0 Then
			Research = Val(txtRes.Text)
		End If
		
		Me.Close()
	End Sub
	
	Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click
		LoadInitial()
	End Sub
	
	Private Sub frmToolNationLvl_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		LoadInitial()
	End Sub
	
	Private Sub LoadInitial()
		
		Dim G As Short
		Dim HS As Short
		Dim TU As Single
		Dim MU As Single
		Dim civs As Integer
		Dim v As productionDataType
		
		NoReset = True
		turn = 0
		
		'Set labels with current values
		lblEd.Text = "Current"
		txtEd.Text = CStr(Education)
		lblNewEd.Text = "Next Turn"
		
		lblHappy.Text = "Current"
		txtHappy.Text = CStr(Happiness)
		lblNewHappy.Text = "Next Turn"
		
		lblTech.Text = "Current"
		txtTech.Text = CStr(TechLevel)
		lblNewTech.Text = "Next Turn"
		
		lblRes.Text = "Current"
		txtRes.Text = CStr(Research)
		lblNewRes.Text = "Next Turn"
		
		EstimateMaxPop = 0
		'Scan production for Grads, tech units, etc.
		If Not (rsSectors.BOF And rsSectors.EOF) Then
			rsSectors.MoveFirst()
			While Not rsSectors.EOF
				
				'total the population
				civs = civs + rsSectors.Fields("civ").Value
				EstimateMaxPop = EstimateMaxPop + 1000
				
				If InStr("ltrp", rsSectors.Fields("des").Value) > 0 Or InStr("ltrp", rsSectors.Fields("sdes").Value) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v = Production(rsSectors)
					If v.prodValidFlag Then
						'UPGRADE_WARNING: Couldn't resolve default property of object v.newdes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Select Case CStr(v.newdes)
							Case "l"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								G = G + v.ProdAmount
							Case "t"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.unitsConsumed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								TU = TU + v.unitsConsumed ' ?? drk@unxsoft.com 10/15/02
							Case "r"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								MU = MU + v.ProdAmount
							Case "p"
								'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								HS = HS + v.ProdAmount
							Case Else
						End Select
					End If
				End If
				rsSectors.MoveNext()
			End While
		End If
		
		If civs > EstimateMaxPop Then
			EstimateMaxPop = civs
		End If
		
		txtGrads.Text = CStr(G)
		txtTUs.Text = CStr(TU)
		txtMeds.Text = CStr(MU)
		txtStrollers.Text = CStr(HS)
		txtPop.Text = CStr(civs)
		txtPop2.Text = CStr(civs)
		txtNewEd.Text = VB6.Format(CalcEducation, "###0.00")
		txtNewTech.Text = VB6.Format(CalcTechnology, "###0.00")
		txtNewHappy.Text = VB6.Format(CalcHappy, "###0.00")
		txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		cmdNext.Text = "Next Turn"
		Frame3.Text = "Happiness (" & VB6.Format(rsNation.Fields("HapNeeded").Value, "###0.00") & " needed)"
		NoReset = False
		
	End Sub
	
	Function CalcEducation() As Single
		
		Dim eduPE As Single
		Dim totalGrads As Single
		Dim eduDelta As Single
		
		
		Const EASY_EDU As Single = 5
		Const EDU_LOGBASE As Single = 4
		
		If Val(txtPop.Text) < 1 Then
			CalcEducation = 0
			Exit Function
		End If
		
		'drk@unxsoft.com 6/30/02.  This function was producing totally incorrect results.  I don't
		'know where the division by 12 comes in, and that second part doesn't look right either.
		'I just rewrote it from scratch.  this is accurate to 3 sig digits for 1 update in the
		'future, less for multiple updates away (but I think that is due to inaccuracies in calculating
		'the new population, not this routine)
		eduPE = rsVersion.Fields("Civs per graduate").Value / Val(txtPop.Text)
		totalGrads = Val(txtGrads.Text) * eduPE
		If (totalGrads = 0) Or (totalGrads - EASY_EDU + EDU_LOGBASE) <= 0 Then
			eduDelta = 0
		Else
			eduDelta = EASY_EDU + (totalGrads - EASY_EDU) / (System.Math.Log(totalGrads - EASY_EDU + EDU_LOGBASE) / System.Math.Log(EDU_LOGBASE))
		End If
		CalcEducation = (rsVersion.Fields("ETU per update").Value * eduDelta + rsVersion.Fields("Edu average").Value * Val(txtEd.Text)) / (rsVersion.Fields("ETU per update").Value + rsVersion.Fields("Edu average").Value)
		
		'Dim cg As Single
		'cg = val(txtPop.Text) / (rsVersion.Fields("ETU per update") * rsVersion.Fields("Civs per graduate") / 12)
		'cg = val(txtGrads.Text) / cg
		'
		'If cg > 5 Then
		'    cg = ((cg - 5) / (Log(cg - 1) / Log(4))) + 5
		'End If
		'
		'CalcEducation = (rsVersion.Fields("ETU per update") * cg + rsVersion.Fields("Edu average") * val(txtEd.Text)) _
		''    / (rsVersion.Fields("ETU per update") + rsVersion.Fields("Edu average"))
		'
		
	End Function
	Function CalcHappy() As Single
		
		Dim happyPE As Single
		Dim totalStrollers As Single
		Dim happyDelta As Single
		Dim ed As Single
		Dim hap_edu As Single
		
		Const EASY_HAPPY As Single = 5
		Const HAPPY_LOGBASE As Single = 6
		
		If Val(txtPop.Text) < 1 Then
			CalcHappy = 0
			Exit Function
		End If
		
		'drk@unxsoft.com 6/30/02.  this was also wrong.  I fixed it.  Note that the current info documentation
		'"Happiness" is incorrect since it says that Happiness is calculated the same way that Education is
		'but it is not.  There is the additional hap_edu factor
		
		ed = Val(txtEd.Text)
		hap_edu = 1.5 - ((ed + 10#) / (ed + 20#))
		
		happyPE = rsVersion.Fields("Civs per stroller").Value / Val(txtPop.Text)
		totalStrollers = Val(txtStrollers.Text) * happyPE * hap_edu
		If (totalStrollers = 0) Then
			happyDelta = 0
		Else
			happyDelta = EASY_HAPPY + (totalStrollers - EASY_HAPPY) / (System.Math.Log(totalStrollers - EASY_HAPPY + HAPPY_LOGBASE) / System.Math.Log(HAPPY_LOGBASE))
		End If
		CalcHappy = (rsVersion.Fields("ETU per update").Value * happyDelta + rsVersion.Fields("Happy average").Value * Val(txtHappy.Text)) / (rsVersion.Fields("ETU per update").Value + rsVersion.Fields("Happy average").Value)
		
		
		'Dim cg As Single
		'Dim ed As Single
		'Dim hap_edu As Single
		'ed = val(txtEd.Text)
		'hap_edu = 1.5 - ((ed + 10#) / (ed + 20#))
		'cg = val(txtPop.Text) / (rsVersion.Fields("ETU per update") * rsVersion.Fields("Civs per stroller") / 12)
		'cg = (val(txtStrollers.Text) * hap_edu) / cg
		'
		'If cg > 5 Then
		'    cg = ((cg - 5) / (Log(cg + 1) / Log(6))) + 5
		'End If
		'
		'CalcHappy = (rsVersion.Fields("ETU per update") * cg + rsVersion.Fields("Happy average") * val(txtHappy.Text)) _
		''    / (rsVersion.Fields("ETU per update") + rsVersion.Fields("Happy average"))
		
	End Function
	
	Function CalcTechnology() As Single
		
		Dim cg As Single
		Dim ed As Single
		Dim decay As Single
		
		Dim tech_decay_time As Single
		Dim tech_decay_rate As Single
		
		'First compute decay
		tech_decay_time = rsVersion.Fields("Tech Decay Time").Value
		tech_decay_rate = rsVersion.Fields("Tech Decay Rate").Value
		If tech_decay_time = 0 Or tech_decay_rate = 0 Then
			decay = 0
		Else
			decay = (Val(txtTech.Text) * rsVersion.Fields("ETU per update").Value * tech_decay_rate) / (100 * tech_decay_time)
		End If
		txtTechDecay.Text = VB6.Format(decay, "###0.00")
		
		ed = Val(txtEd.Text)
		If ed < 5 And Not frmOptions.bSPAtlantis Then
			txtTechBuild.Text = vbNullString
			CalcTechnology = Val(txtTech.Text) + Val(txtTechBleed.Text) - Val(txtTechDecay.Text)
			Exit Function
		End If
		
		If frmOptions.bSPAtlantis Then
			cg = Val(txtTUs.Text)
		Else
			cg = Val(txtTUs.Text) * ((ed - 5) / (ed + 5))
		End If
		
		If cg > rsVersion.Fields("Easy tech").Value Then
			cg = System.Math.Log(cg - rsVersion.Fields("Easy tech").Value + 1) / (System.Math.Log(rsVersion.Fields("Tech base").Value))
			If cg < 0 Then
				cg = 0
			End If
			cg = cg + rsVersion.Fields("Easy tech").Value
		End If
		
		txtTechBuild.Text = VB6.Format(cg, "####0.00")
		CalcTechnology = Val(txtTech.Text) + Val(txtTechBleed.Text) - Val(txtTechDecay.Text) + cg
		
	End Function
	
	Function CalcResearch() As Single
		
		Dim cg As Single
		Dim ed As Single
		Dim decay As Single
		Dim tech_decay_time As Single
		Dim tech_decay_rate As Single
		
		'First compute decay
		tech_decay_time = rsVersion.Fields("Tech Decay Time").Value
		tech_decay_rate = rsVersion.Fields("Tech Decay Rate").Value
		If tech_decay_time = 0 Or tech_decay_rate = 0 Then
			decay = 0
		Else
			decay = (Val(txtRes.Text) * rsVersion.Fields("ETU per update").Value * tech_decay_rate) / (100 * tech_decay_time)
		End If
		txtResDecay.Text = VB6.Format(decay, "###0.00")
		
		ed = Val(txtEd.Text)
		If ed < 5 Then
			txtResBuild.Text = vbNullString
			CalcResearch = Val(txtRes.Text) + Val(txtResBleed.Text) - Val(txtResDecay.Text)
			Exit Function
		End If
		
		cg = Val(txtMeds.Text) * ((ed - 5) / (ed + 5))
		
		If cg > 0.75 Then
			cg = System.Math.Log(cg + 0.25) / (System.Math.Log(2))
			If cg < 0 Then
				cg = 0
			End If
			cg = cg + 0.75
		End If
		
		txtResBuild.Text = VB6.Format(cg, "###0.00")
		CalcResearch = Val(txtRes.Text) + Val(txtResBleed.Text) - Val(txtResDecay.Text) + cg
		
	End Function
	'UPGRADE_WARNING: Event txtEd.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtEd_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEd.TextChanged
		If Not NoReset Then
			txtNewEd.Text = VB6.Format(CalcEducation, "###0.00")
			txtNewTech.Text = VB6.Format(CalcTechnology, "###0.00")
			txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtGrads.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtGrads_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGrads.TextChanged
		If Not NoReset Then
			txtNewEd.Text = VB6.Format(CalcEducation, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtHappy.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtHappy_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHappy.TextChanged
		If Not NoReset Then
			txtNewHappy.Text = VB6.Format(CalcHappy, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtMeds.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtMeds_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMeds.TextChanged
		If Not NoReset Then
			txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtPop.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPop_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPop.TextChanged
		If Not NoReset Then
			txtPop2.Text = txtPop.Text
			txtNewEd.Text = VB6.Format(CalcEducation, "###0.00")
			txtNewHappy.Text = VB6.Format(CalcHappy, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtRes.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRes.TextChanged
		If Not NoReset Then
			txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtResBleed.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtResBleed_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtResBleed.TextChanged
		If Not NoReset Then
			txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtResDecay.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtResDecay_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtResDecay.TextChanged
		If Not NoReset Then
			txtNewRes.Text = VB6.Format(CalcResearch, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtStrollers.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtStrollers_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStrollers.TextChanged
		If Not NoReset Then
			txtNewHappy.Text = VB6.Format(CalcHappy, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtTech.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTech_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTech.TextChanged
		If Not NoReset Then
			txtNewTech.Text = VB6.Format(CalcTechnology, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtTechBleed.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTechBleed_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTechBleed.TextChanged
		If Not NoReset Then
			txtNewTech.Text = VB6.Format(CalcTechnology, "###0.00")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtTUs.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTUs_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTUs.TextChanged
		If Not NoReset Then
			txtNewTech.Text = VB6.Format(CalcTechnology, "###0.00")
		End If
	End Sub
End Class