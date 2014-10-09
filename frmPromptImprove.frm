VERSION 5.00
Begin VB.Form frmPromptImprove 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1650
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNavigation 
      Caption         =   "navigation"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optFort 
      Caption         =   "fort"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optRadar 
      Caption         =   "radar"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optRunway 
      Caption         =   "runway"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optDefense 
      Caption         =   "defense"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton optRail 
      Caption         =   "rail"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton optRoad 
      Caption         =   "road"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblDesc 
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "percent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Improve"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmPromptImprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081103 efj: Added Option Explicit and removed dead variables
'092803 rjk: Added initial field selection. General reformatting.
'            Made the options sector specific.
'093003 rjk: Switched logic so if origin is not a single sector then all options are available.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'130506 rjk: Added South Pacific: Atlantis
'210506 rjk: Fixed new SP: Atlantis new infrastructures percentages.
'250506 rjk: Added Four SP: Atlantis infrastructures to the database,
'            use runway -> min, radar -> gld, fort -> fert nad navigate -> oil.
'280506 rjk: Fixed maximum percent for SP: Atlantis.
'190606 rjk: Fixed maximum percentage for SP: Atlantis fields.

Dim itype As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
strCmd = "improve " + itype + " " + txtOrigin.Text + " " + txtNum.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
'frmEmpCmd.SubmitEmpireCommand "dump " + txtOrigin.Text, False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False
Unload Me
End Sub

Private Sub Form_Activate()
UpdateDesc
If optRoad.Enabled Then
    optRoad.Value = True
ElseIf optRail.Enabled Then
    optRail.Value = True
ElseIf optDefense.Enabled Then
    optDefense.Value = True
ElseIf optRadar.Enabled Then
    optRadar.Value = True
ElseIf optFort.Enabled Then
    optFort.Value = True
ElseIf optNavigation.Enabled Then
    optNavigation.Value = True
ElseIf optRunway.Enabled Then
    optRunway.Value = True
End If
txtOrigin.SetFocus
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim sucess As Long   removed efj 8/2003
If frmOptions.bSPAtlantis Then
    optRadar.Visible = True
    optRunway.Visible = True
    optFort.Visible = True
    optNavigation.Visible = True
End If
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub optRoad_Click()
itype = optRoad.Caption
UpdateDesc
End Sub

Private Sub optRail_Click()
itype = optRail.Caption
UpdateDesc
End Sub

Private Sub optDefense_Click()
itype = optDefense.Caption
UpdateDesc
End Sub

Private Sub optRadar_Click()
itype = optRadar.Caption
UpdateDesc
End Sub

Private Sub optFort_Click()
itype = optFort.Caption
UpdateDesc
End Sub

Private Sub optRunway_Click()
itype = optRunway.Caption
UpdateDesc
End Sub

Private Sub optNavigation_Click()
itype = optNavigation.Caption
UpdateDesc
End Sub

Private Sub txtOrigin_Change()
UpdateDesc
End Sub

Private Sub txtOrigin_DblClick()
Load frmPromptSectors
frmPromptSectors.strSectors = txtOrigin.Text
frmPromptSectors.SetControls
frmPromptSectors.Caption = "Select Sectors"
frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
frmPromptSectors.Show vbModeless
End Sub

Private Sub UpdateDesc()
Dim sx As Integer
Dim sy As Integer
Dim MaxPercent As Integer
Dim hldIndex As String

hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"

On Error Resume Next

'093003 rjk: Switched logic so if origin is not a single sector then all options are available.
'p1 = InStr(txtOrigin.Text, ",")
'If p1 > 0 And Len(txtOrigin.Text) > p1 Then
If Len(txtOrigin.Text) > 0 Then
    If ParseSectors(sx, sy, txtOrigin.Text) Then
        'sx = CInt(Left$(txtOrigin.Text, p1 - 1))
        'sy = CInt(Mid$(txtOrigin.Text, p1 + 1))
        'Get Sector Information
        rsSectors.Seek "=", sx, sy
        If Not rsSectors.NoMatch Then
            lblDesc = CStr(rsSectors.Fields("eff").Value) + "% "
            lblDesc = lblDesc + colDes.Item(rsSectors.Fields("des").Value)
            MaxPercent = 100
            If frmOptions.bSPAtlantis Then
                Select Case itype
                Case "runway"
                    lblDesc = lblDesc + " w/ current " + itype + " of " + CStr(rsSectors.Fields("min").Value) + "% "
                MaxPercent = MaxPercent - rsSectors.Fields("min").Value
                Case "radar"
                    lblDesc = lblDesc + " w/ current " + itype + " of " + CStr(rsSectors.Fields("gold").Value) + "% "
                MaxPercent = MaxPercent - rsSectors.Fields("gold").Value
                Case "fort"
                    lblDesc = lblDesc + " w/ current " + itype + " of " + CStr(rsSectors.Fields("fert").Value) + "% "
                MaxPercent = MaxPercent - rsSectors.Fields("fert").Value
                Case "navigation"
                    lblDesc = lblDesc + " w/ current " + itype + " of " + CStr(rsSectors.Fields("ocontent").Value) + "% "
                    MaxPercent = MaxPercent - rsSectors.Fields("ocontent").Value
                Case Else
                    lblDesc = lblDesc + " w/ current " + itype + " of " + CStr(rsSectors.Fields(itype).Value) + "% "
                    MaxPercent = MaxPercent - rsSectors.Fields(itype).Value
                End Select
            Else
                lblDesc = lblDesc + " w/ current " + itype + " of " + CStr(rsSectors.Fields(itype).Value) + "% "
                MaxPercent = MaxPercent - rsSectors.Fields(itype).Value
            End If
            
            rsBuild.Seek "=", "i", Left$(itype, 5)
            If Not rsBuild.NoMatch Then
                'compute maximum
                If MaxPercent > (rsSectors.Fields("mob").Value / rsBuild.Fields("stat2")) Then
                    MaxPercent = Int(rsSectors.Fields("mob").Value / rsBuild.Fields("stat2"))
                End If
                If MaxPercent > Int(rsSectors.Fields("lcm").Value / rsBuild.Fields("lcm")) Then
                    MaxPercent = Int(rsSectors.Fields("lcm").Value / rsBuild.Fields("lcm"))
                End If
                If MaxPercent > Int(rsSectors.Fields("hcm").Value / rsBuild.Fields("hcm")) Then
                    MaxPercent = Int(rsSectors.Fields("hcm").Value / rsBuild.Fields("hcm"))
                End If
            Else
                MaxPercent = 0
            End If
            txtNum.Text = CStr(MaxPercent)
            
            rsBuild.Seek "=", "i", "road"
            If CStr(rsSectors.Fields("road").Value) < 100 And Not rsBuild.NoMatch Then
                If Not optRoad.Enabled Then
                    optRoad.Enabled = True
                End If
            Else
                optRoad.Enabled = False
            End If
            rsBuild.Seek "=", "i", "rail"
            If CStr(rsSectors.Fields("rail").Value) < 100 And Not rsBuild.NoMatch Then
                If Not optRail.Enabled Then
                    optRail.Enabled = True
                End If
            Else
                optRail.Enabled = False
            End If
            rsBuild.Seek "=", "i", "defen"
            If CStr(rsSectors.Fields("defense").Value) < 100 And Not rsBuild.NoMatch Then
                If Not optDefense.Enabled Then
                    optDefense.Enabled = True
                End If
            Else
                optDefense.Enabled = False
            End If
            If frmOptions.bSPAtlantis Then
                rsBuild.Seek "=", "i", "runwa"
                If CStr(rsSectors.Fields("min").Value) < 100 And Not rsBuild.NoMatch Then
                    If Not optRunway.Enabled Then
                        optRunway.Enabled = True
                    End If
                Else
                    optRunway.Enabled = False
                End If
                rsBuild.Seek "=", "i", "radar"
                If CStr(rsSectors.Fields("gold").Value) < 100 And Not rsBuild.NoMatch Then
                    If Not optRadar.Enabled Then
                        optRadar.Enabled = True
                    End If
                Else
                    optRadar.Enabled = False
                End If
                rsBuild.Seek "=", "i", "fort"
                If CStr(rsSectors.Fields("fert").Value) < 100 And Not rsBuild.NoMatch Then
                    If Not optFort.Enabled Then
                        optFort.Enabled = True
                    End If
                Else
                    optFort.Enabled = False
                End If
                rsBuild.Seek "=", "i", "navig"
                If CStr(rsSectors.Fields("ocontent").Value) < 100 And Not rsBuild.NoMatch Then
                    If Not optNavigation.Enabled Then
                        optNavigation.Enabled = True
                    End If
                Else
                    optNavigation.Enabled = False
                End If
            End If
        Else
            optRoad.Enabled = False
            optRail.Enabled = False
            optDefense.Enabled = False
            If frmOptions.bSPAtlantis Then
                optRunway.Enabled = False
                optRadar.Enabled = False
                optFort.Enabled = False
                optNavigation = False
            End If
        End If
    Else '093003 rjk: In case a multiple sector group is selected, switched to true.
        optRoad.Enabled = True
        optRail.Enabled = True
        optDefense.Enabled = True
        If frmOptions.bSPAtlantis Then
            optRunway.Enabled = True
            optRadar.Enabled = True
            optFort.Enabled = True
            optNavigation = True
        End If
    End If
Else
    optRoad.Enabled = False
    optRail.Enabled = False
    optDefense.Enabled = False
    If frmOptions.bSPAtlantis Then
        optRunway.Enabled = False
        optRadar.Enabled = False
        optFort.Enabled = False
        optNavigation = False
    End If
End If

If optRoad.Enabled Or optRail.Enabled Or optDefense.Enabled Or frmOptions.bSPAtlantis Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
    txtOrigin.SetFocus
End If

rsBuild.Index = hldIndex
End Sub
