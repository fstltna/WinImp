VERSION 5.00
Begin VB.Form frmToolNationLvl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nation Levels"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2280
      TabIndex        =   44
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Turn"
      Height          =   375
      Left            =   960
      TabIndex        =   43
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Research"
      Height          =   1815
      Left            =   3000
      TabIndex        =   37
      Top             =   2280
      Width           =   3015
      Begin VB.TextBox txtResBuild 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtRes 
         Height          =   285
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMeds 
         Height          =   285
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtNewRes 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtResDecay 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtResBleed 
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Build"
         Height          =   255
         Left            =   1320
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblRes 
         Caption         =   "Current"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Med units"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblNewRes 
         Caption         =   "Next Turn"
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Decay"
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Bleed"
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Happiness"
      Height          =   1815
      Left            =   240
      TabIndex        =   32
      Top             =   2280
      Width           =   2415
      Begin VB.TextBox txtHappy 
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtStrollers 
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtNewHappy 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtPop2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblHappy 
         Caption         =   "Current"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Strollers"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblNewHappy 
         Caption         =   "Next Turn"
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Population"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Technology"
      Height          =   1815
      Left            =   3000
      TabIndex        =   25
      Top             =   240
      Width           =   3015
      Begin VB.TextBox txtTechBuild 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtTechBleed 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTechDecay 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtNewTech 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtTUs 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTech 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Build"
         Height          =   255
         Left            =   1320
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Bleed"
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Decay"
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblNewTech 
         Caption         =   "Next Turn"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tech units"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblTech 
         Caption         =   "Current"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Education"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.TextBox txtPop 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtNewEd 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtGrads 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtEd 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Population"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblNewEd 
         Caption         =   "Next Turn"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Grads"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblEd 
         Caption         =   "Current"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmToolNationLvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'251003 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
'140506 rjk: Removed education from tech production for South Pacific Atlantis.

Dim NoReset As Boolean
Dim EstimateMaxPop As Single
Dim turn As Integer

Private Sub cmdNext_Click()
Dim civs As Long

turn = turn + 1
NoReset = True

'Set labels with current values
lblEd.Caption = "Turn " + CStr(turn)
txtEd.Text = txtNewEd.Text
lblNewEd.Caption = "Turn " + CStr(turn + 1)

lblHappy.Caption = "Turn " + CStr(turn)
txtHappy.Text = txtNewHappy.Text
lblNewHappy.Caption = "Turn " + CStr(turn + 1)

lblTech.Caption = "Turn " + CStr(turn)
txtTech.Text = txtNewTech.Text
lblNewTech.Caption = "Turn " + CStr(turn + 1)

lblRes.Caption = "Turn " + CStr(turn)
txtRes.Text = txtNewRes.Text
lblNewRes.Caption = "Turn " + CStr(turn + 1)

civs = CLng(Val(txtPop.Text) * (1 + ((rsVersion.Fields("ETU per update") * rsVersion.Fields("Civ Birthrate")) / 1000)) + 0.499)
If civs > EstimateMaxPop Then
    civs = EstimateMaxPop
End If

txtPop.Text = CStr(civs)
txtPop2.Text = CStr(civs)
txtNewEd.Text = Format$(CalcEducation, "###0.00")
txtNewHappy.Text = Format$(CalcHappy, "###0.00")
txtNewTech.Text = Format$(CalcTechnology, "###0.00")
txtNewRes.Text = Format$(CalcResearch, "###0.00")

NoReset = False
cmdNext.Caption = "Turn " + CStr(turn + 1)
            
End Sub

Private Sub cmdOK_Click()
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

Unload Me
End Sub

Private Sub cmdReset_Click()
LoadInitial
End Sub

Private Sub Form_Load()
LoadInitial
End Sub

Private Sub LoadInitial()

Dim G As Integer
Dim HS As Integer
Dim TU As Single
Dim MU As Single
Dim civs As Long
Dim v As productionDataType

NoReset = True
turn = 0

'Set labels with current values
lblEd.Caption = "Current"
txtEd.Text = CStr(Education)
lblNewEd.Caption = "Next Turn"

lblHappy.Caption = "Current"
txtHappy.Text = CStr(Happiness)
lblNewHappy.Caption = "Next Turn"

lblTech.Caption = "Current"
txtTech.Text = CStr(TechLevel)
lblNewTech.Caption = "Next Turn"

lblRes.Caption = "Current"
txtRes.Text = CStr(Research)
lblNewRes.Caption = "Next Turn"

EstimateMaxPop = 0
'Scan production for Grads, tech units, etc.
If Not (rsSectors.BOF And rsSectors.EOF) Then
    rsSectors.MoveFirst
    While Not rsSectors.EOF
    
        'total the population
        civs = civs + rsSectors.Fields("civ")
        EstimateMaxPop = EstimateMaxPop + 1000
        
        If InStr("ltrp", rsSectors.Fields("des")) > 0 Or _
            InStr("ltrp", rsSectors.Fields("sdes")) > 0 Then
            v = Production(rsSectors)
            If v.prodValidFlag Then
                Select Case CStr(v.newdes)
                    Case "l"
                        G = G + v.ProdAmount
                    Case "t"
                        TU = TU + v.unitsConsumed ' ?? drk@unxsoft.com 10/15/02
                    Case "r"
                        MU = MU + v.ProdAmount
                    Case "p"
                        HS = HS + v.ProdAmount
                    Case Else
                End Select
            End If
        End If
        rsSectors.MoveNext
    Wend
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
txtNewEd.Text = Format$(CalcEducation, "###0.00")
txtNewTech.Text = Format$(CalcTechnology, "###0.00")
txtNewHappy.Text = Format$(CalcHappy, "###0.00")
txtNewRes.Text = Format$(CalcResearch, "###0.00")
cmdNext.Caption = "Next Turn"
Frame3.Caption = "Happiness (" + Format$(rsNation.Fields("HapNeeded"), "###0.00") + " needed)"
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
eduPE = rsVersion.Fields("Civs per graduate") / Val(txtPop.Text)
totalGrads = Val(txtGrads.Text) * eduPE
If (totalGrads = 0) Or (totalGrads - EASY_EDU + EDU_LOGBASE) <= 0 Then
    eduDelta = 0
Else
    eduDelta = EASY_EDU + (totalGrads - EASY_EDU) / (Log(totalGrads - EASY_EDU + EDU_LOGBASE) / Log(EDU_LOGBASE))
End If
CalcEducation = (rsVersion.Fields("ETU per update") * eduDelta + rsVersion.Fields("Edu average") * Val(txtEd.Text)) _
    / (rsVersion.Fields("ETU per update") + rsVersion.Fields("Edu average"))

'Dim cg As Single
'cg = val(txtPop.Text) / (rsVersion.Fields("ETU per update") * rsVersion.Fields("Civs per graduate") / 12)
'cg = val(txtGrads.Text) / cg
'
'If cg > 5 Then
'    cg = ((cg - 5) / (Log(cg - 1) / Log(4))) + 5
'End If
'
'CalcEducation = (rsVersion.Fields("ETU per update") * cg + rsVersion.Fields("Edu average") * val(txtEd.Text)) _
'    / (rsVersion.Fields("ETU per update") + rsVersion.Fields("Edu average"))
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

happyPE = rsVersion.Fields("Civs per stroller") / Val(txtPop.Text)
totalStrollers = Val(txtStrollers.Text) * happyPE * hap_edu
If (totalStrollers = 0) Then
    happyDelta = 0
Else
    happyDelta = EASY_HAPPY + (totalStrollers - EASY_HAPPY) / (Log(totalStrollers - EASY_HAPPY + HAPPY_LOGBASE) / Log(HAPPY_LOGBASE))
End If
CalcHappy = (rsVersion.Fields("ETU per update") * happyDelta + rsVersion.Fields("Happy average") * Val(txtHappy.Text)) _
    / (rsVersion.Fields("ETU per update") + rsVersion.Fields("Happy average"))


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
'    / (rsVersion.Fields("ETU per update") + rsVersion.Fields("Happy average"))

End Function

Function CalcTechnology() As Single

Dim cg As Single
Dim ed As Single
Dim decay As Single

Dim tech_decay_time As Single
Dim tech_decay_rate As Single

'First compute decay
tech_decay_time = rsVersion.Fields("Tech Decay Time")
tech_decay_rate = rsVersion.Fields("Tech Decay Rate")
If tech_decay_time = 0 Or tech_decay_rate = 0 Then
    decay = 0
Else
    decay = (Val(txtTech.Text) * rsVersion.Fields("ETU per update") * tech_decay_rate) / (100 * tech_decay_time)
End If
txtTechDecay.Text = Format$(decay, "###0.00")

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

If cg > rsVersion.Fields("Easy tech") Then
    cg = Log(cg - rsVersion.Fields("Easy tech") + 1) / (Log(rsVersion.Fields("Tech base")))
    If cg < 0 Then
        cg = 0
    End If
    cg = cg + rsVersion.Fields("Easy tech")
End If

txtTechBuild.Text = Format$(cg, "####0.00")
CalcTechnology = Val(txtTech.Text) + Val(txtTechBleed.Text) - Val(txtTechDecay.Text) + cg

End Function

Function CalcResearch() As Single

Dim cg As Single
Dim ed As Single
Dim decay As Single
Dim tech_decay_time As Single
Dim tech_decay_rate As Single

'First compute decay
tech_decay_time = rsVersion.Fields("Tech Decay Time")
tech_decay_rate = rsVersion.Fields("Tech Decay Rate")
If tech_decay_time = 0 Or tech_decay_rate = 0 Then
    decay = 0
Else
    decay = (Val(txtRes.Text) * rsVersion.Fields("ETU per update") * tech_decay_rate) / (100 * tech_decay_time)
End If
txtResDecay.Text = Format$(decay, "###0.00")

ed = Val(txtEd.Text)
If ed < 5 Then
    txtResBuild.Text = vbNullString
    CalcResearch = Val(txtRes.Text) + Val(txtResBleed.Text) - Val(txtResDecay.Text)
    Exit Function
End If

cg = Val(txtMeds.Text) * ((ed - 5) / (ed + 5))

If cg > 0.75 Then
    cg = Log(cg + 0.25) / (Log(2))
    If cg < 0 Then
        cg = 0
    End If
    cg = cg + 0.75
End If

txtResBuild.Text = Format$(cg, "###0.00")
CalcResearch = Val(txtRes.Text) + Val(txtResBleed.Text) - Val(txtResDecay.Text) + cg

End Function
Private Sub txtEd_Change()
If Not NoReset Then
    txtNewEd.Text = Format$(CalcEducation, "###0.00")
    txtNewTech.Text = Format$(CalcTechnology, "###0.00")
    txtNewRes.Text = Format$(CalcResearch, "###0.00")
End If
End Sub

Private Sub txtGrads_Change()
If Not NoReset Then
    txtNewEd.Text = Format$(CalcEducation, "###0.00")
End If
End Sub

Private Sub txtHappy_Change()
If Not NoReset Then
    txtNewHappy.Text = Format$(CalcHappy, "###0.00")
End If
End Sub

Private Sub txtMeds_Change()
If Not NoReset Then
    txtNewRes.Text = Format$(CalcResearch, "###0.00")
End If
End Sub

Private Sub txtPop_Change()
If Not NoReset Then
    txtPop2.Text = txtPop.Text
    txtNewEd.Text = Format$(CalcEducation, "###0.00")
    txtNewHappy.Text = Format$(CalcHappy, "###0.00")
End If
End Sub

Private Sub txtRes_Change()
If Not NoReset Then
    txtNewRes.Text = Format$(CalcResearch, "###0.00")
End If
End Sub

Private Sub txtResBleed_Change()
If Not NoReset Then
    txtNewRes.Text = Format$(CalcResearch, "###0.00")
End If
End Sub

Private Sub txtResDecay_Change()
If Not NoReset Then
    txtNewRes.Text = Format$(CalcResearch, "###0.00")
End If
End Sub

Private Sub txtStrollers_Change()
If Not NoReset Then
    txtNewHappy.Text = Format$(CalcHappy, "###0.00")
End If
End Sub

Private Sub txtTech_Change()
If Not NoReset Then
    txtNewTech.Text = Format$(CalcTechnology, "###0.00")
End If
End Sub

Private Sub txtTechBleed_Change()
If Not NoReset Then
    txtNewTech.Text = Format$(CalcTechnology, "###0.00")
End If
End Sub

Private Sub txtTUs_Change()
If Not NoReset Then
    txtNewTech.Text = Format$(CalcTechnology, "###0.00")
End If
End Sub
