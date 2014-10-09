VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmToolProduction 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2280
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbReport 
      Height          =   975
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmToolProduction.frx":0000
   End
   Begin VB.CommandButton cmdPE 
      Caption         =   "PE: 97%"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtProduce 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Text            =   "999"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblMaterial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "oil"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblMaterial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lcm"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   21
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblMaterial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hcm"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   20
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMaterial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "avail"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblRequired 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   18
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblRequired 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblRequired 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Sector"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblRequired 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Required"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblResource 
      Alignment       =   1  'Right Justify
      Caption         =   "9999"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblResource 
      Alignment       =   1  'Right Justify
      Caption         =   "9999"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblResource 
      Alignment       =   1  'Right Justify
      Caption         =   "9999"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblResource 
      Alignment       =   1  'Right Justify
      Caption         =   "9999"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   2880
      X2              =   3000
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   2880
      X2              =   2760
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Line Line10 
      BorderWidth     =   4
      X1              =   3600
      X2              =   3480
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Line Line9 
      BorderWidth     =   4
      X1              =   3480
      X2              =   3600
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   2880
      X2              =   2880
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2160
      X2              =   2880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2880
      X2              =   2880
      Y1              =   1320
      Y2              =   1680
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   2160
      X2              =   2880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2880
      X2              =   2880
      Y1              =   960
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2160
      X2              =   2880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2160
      X2              =   3600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblProduct 
      Caption         =   "guns"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmToolProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'150304 rjk: Added from Jim Original code
'240304 rjk: Limited the product value to 9999, change values to long to prevent overflow,
'            Redesigned form, removed button and text boxes and replaces with label.
'100404 rjk: Added 'e' enlist to Estimate Tool.

Dim secx As Integer
Dim secy As Integer
Dim OriginChange As Boolean
Dim pe As Single
Dim ColItem As New Collection
Dim newdes As String
Dim v As Variant

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
Dim success As Long
success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
LoadItems
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer

'use origin change to avoid firing cascading events
OriginChange = True

If ParseSectors(sx, sy, txtOrigin.Text) Then
    secx = sx
    secy = sy
    FillSectorInfo
Else
    HandleError "Not a valid Sector"
End If

OriginChange = False

End Sub

Private Sub FillSectorInfo()
Dim n As Integer
Dim strTemp As String
Dim v As productionDataType

'get sector record
rsSectors.Seek "=", secx, secy
If rsSectors.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If

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
    newdes = CStr(v.newdes)
Else
    HandleError "Cannot Calculate Production"
    Exit Sub
End If

'get sector type
rsSectorType.Seek "=", newdes
If rsSectorType.NoMatch Then
    HandleError "Sector Type not found"
    Exit Sub
End If

For n = 1 To 3
    lblMaterial(n).Visible = False
    lblRequired(n).Visible = False
    lblResource(n).Caption = ""
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
If rsSectorType.Fields("use1n") > 0 Then
    strTemp = ColItem.item(rsSectorType.Fields("use1s"))
    lblMaterial(1).Caption = strTemp
    lblMaterial(1).Visible = True
    lblRequired(1).Visible = True
    lblResource(1).Caption = CStr(rsSectors.Fields(strTemp))
    Line3.Visible = True
    Line4.Visible = True
    Line11.Visible = True
    Line12.Visible = True
    
    If rsSectorType.Fields("use2n") > 0 Then
        strTemp = ColItem.item(rsSectorType.Fields("use2s"))
        lblMaterial(2).Caption = strTemp
        lblMaterial(2).Visible = True
        lblRequired(2).Visible = True
        lblResource(2).Caption = CStr(rsSectors.Fields(strTemp))
        Line5.Visible = True
        Line6.Visible = True

        If rsSectorType.Fields("use3n") > 0 Then
            strTemp = ColItem.item(rsSectorType.Fields("use3s"))
            lblMaterial(3).Caption = strTemp
            lblMaterial(3).Visible = True
            lblRequired(3).Visible = True
            lblResource(3).Caption = CStr(rsSectors.Fields(strTemp))
            Line7.Visible = True
            Line8.Visible = True
        End If
    End If
End If

lblType = CStr(v.NewEff) + "% " + rsSectorType.Fields("desc")

lblProduct.Caption = Trim(rsSectorType.Fields("product"))
If lblProduct.Caption = "" Then lblProduct.Caption = "avail"

If newdes = "e" Then
    strTemp = "civ"
    lblMaterial(1).Caption = strTemp
    lblMaterial(1).Visible = True
    lblRequired(1).Visible = True
    lblResource(1).Caption = CStr(rsSectors.Fields(strTemp))
    Line3.Visible = True
    Line4.Visible = True
    Line11.Visible = True
    Line12.Visible = True
    strTemp = "mil"
    lblMaterial(2).Caption = strTemp
    lblMaterial(2).Visible = True
    lblRequired(2).Visible = True
    lblResource(2).Caption = CStr(rsSectors.Fields(strTemp))
    Line5.Visible = True
    Line6.Visible = True
    lblProduct.Caption = "mil"
End If

pe = (v.NewEff / 100) * v.ResourceEff * v.LevelEff
cmdPE.Caption = "PE: " + Format(CInt(v.NewEff * pe), "##0")

lblResource(0).Caption = CStr(v.ProductionAvail)
txtProduce.Text = CStr(v.ProdAmount)

If v.BuildAvailCost > 0 Then
    rtbReport.Text = CStr(v.BuildAvailCost) + " avail used to raise eff. to " + CStr(v.NewEff) + "%"
Else
    rtbReport.Text = ""
End If
End Sub

Private Sub HandleError(ErrMsg As String)
Dim n As Integer

lblType.Caption = ErrMsg
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

If ColItem.Count > 0 Then
    Exit Sub
End If

With ColItem
    .Add "bar", "b"
    .Add "civ", "c"
    .Add "dust", "d"
    .Add "food", "f"
    .Add "gun", "g"
    .Add "hcm", "h"
    .Add "iron", "i"
    .Add "lcm", "l"
    .Add "mil", "m"
    .Add "oil", "o"
    .Add "pet", "p"
    .Add "rad", "r"
    .Add "shell", "s"
    .Add "uw", "u"
End With

End Sub

Private Sub txtProduce_Change()

Dim units As Long
Dim avail As Long
Dim q1 As Long
Dim q2 As Long
Dim q3 As Long
Dim n As Long
Dim buildavail As Long
Dim NewEff As Integer
Dim new_pe As Single
Dim save_Units As Integer
Dim Done As Boolean


For n = 0 To 3
    lblRequired(n).Caption = ""
Next n

rsSectors.Seek "=", secx, secy
rsSectorType.Seek "=", newdes

If rsSectors.NoMatch Or rsSectorType.NoMatch Then
    Exit Sub
End If

If Val(txtProduce.Text) <= 0 Then
    Exit Sub
End If

If Len(Trim(rsSectors.Fields("sdes"))) > 0 _
        And rsSectors.Fields("eff") > 0 Then
    buildavail = 60 + Int((rsSectors.Fields("eff") + 3) / 4)
    NewEff = 60
Else
    If rsSectors.Fields("eff") < 60 Then
        buildavail = 60 - rsSectors.Fields("eff")
        NewEff = 60
    Else
        buildavail = 0
        NewEff = rsSectors.Fields("eff")
    End If
End If

Done = False
save_Units = -99

Do
    new_pe = pe * (NewEff / 100#)

    If new_pe <= 0 Then
        Done = True
        units = 0
    Else
        If Val(txtProduce.Text) > 9999 Then
            txtProduce.Text = "9999"
        End If
        units = CInt((Val(txtProduce.Text) / new_pe) + 0.4999)
    End If
    
    If units = save_Units Then
        'units required didn't change - we are done
        Exit Do
    End If

    'set text
    If rsSectorType.Fields("use1n") > 0 Then
        q1 = rsSectorType.Fields("use1n") * units
        lblRequired(1).Caption = CStr(q1)
        If rsSectorType.Fields("use2n") > 0 Then
            q2 = rsSectorType.Fields("use2n") * units
           lblRequired(2).Caption = CStr(q2)
            If rsSectorType.Fields("use3n") > 0 Then
                q3 = rsSectorType.Fields("use3n") * units
               lblRequired(3).Caption = CStr(q3)
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
            Dim mil_requested As Long
            Dim mil_required As Long
            Dim recruit_rate As Single
            
            recruit_rate = (rsVersion.Fields("ETU per update") / 20)
            
            mil_requested = Val(txtProduce.Text)
            
            mil_required = mil_requested / recruit_rate
            
            If mil_required > 10 Then
                mil_required = mil_required - 10
                lblRequired(2).Caption = CStr(mil_required)
            Else
                lblRequired(2).Caption = CStr(0)
            End If
            lblRequired(1).Caption = CStr(mil_requested * 2)
        Else
            lblRequired(2).Caption = CStr(0)
            lblRequired(1).Caption = CStr(0)
        End If
    End If
Loop Until Done

If rsSectors.Fields("work") < 100 Then
    avail = CInt((avail / (CSng(rsSectors.Fields("work")) / 100#)) + 0.4999)
End If

lblRequired(0).Caption = CStr(avail)
If buildavail > 0 Then
    rtbReport.Text = CStr(buildavail) + " avail used to raise eff. to " + CStr(NewEff) + "%"
Else
    rtbReport.Text = ""
End If

lblType = CStr(NewEff) + "% " + rsSectorType.Fields("desc")
cmdPE.Caption = "PE: " + Format(CInt(NewEff * pe), "##0")

End Sub

