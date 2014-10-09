VERSION 5.00
Begin VB.Form frmPromptMove 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2265
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox cbCom 
      Columns         =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   4560
      TabIndex        =   11
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Move"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblDest 
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
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblOrigin 
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
      Left            =   2160
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Move"
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
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "to"
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
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "from"
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
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "frmPromptMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OriginalParenthWnd As Long
Dim strAmount As String

Private Sub cbCom_Click()
Label4.Caption = cbCom.List(cbCom.ListIndex)

Dim str As String
Dim sx As Integer
Dim sy As Integer
On Error Resume Next

If ParseSectors(sx, sy, txtOrigin.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        str = Left(Label4.Caption, 3)
        If str = "foo" Or str = "iro" Or str = "dus" Then
            str = Left(Label4.Caption, 4)
        ElseIf str = "unc" Then
            str = "uw"
        ElseIf str = "she" Then
            str = "shell"
        End If
        
        strAmount = CStr(rsSectors.Fields(str).Value)
        Label4.Caption = Label4.Caption + " [" + strAmount + " avail]"
    
    End If
End If


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String
Dim stritem As String

stritem = cbCom.Text
If Left(stritem, 1) = "u" Then
    stritem = "un"
End If

strCmd = "move " + stritem + " " + txtOrigin.Text + " " + txtNum.Text + " " + txtPath.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
frmEmpCmd.SubmitEmpireCommand "dump " + txtOrigin.Text, False
frmEmpCmd.SubmitEmpireCommand "dump " + txtPath.Text, False
frmEmpCmd.SubmitEmpireCommand "db2", False
Unload Me
End Sub

Private Sub cmdTest_Click()
Dim strCmd As String
Dim stritem As String

stritem = cbCom.Text
If Left(stritem, 1) = "u" Then
    stritem = "un"
End If

strCmd = "test " + stritem + " " + txtOrigin.Text + " " + txtNum.Text + " " + txtPath.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
End Sub

Private Sub Form_Load()

' Set parent for the toolbar to display on top of:
Dim sucess As Long
success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
LoadCommodityBox cbCom
cbCom.ListIndex = 1 'set to civilians
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtNum_DblClick()
'Put the full amount in the text box
If Len(strAmount) > 0 Then
    txtNum = strAmount
End If
End Sub

Private Sub txtPath_Change()
Dim p1 As Integer
Dim sx As Integer
Dim sy As Integer
On Error Resume Next

If Len(txtPath.Text) = 0 Then
    lblDest = 0
    Exit Sub
End If

If ParseSectors(sx, sy, txtPath.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        lblDest = CStr(rsSectors.Fields("eff").Value) + "% "
        lblDest = lblDest + colDes.Item(rsSectors.Fields("des").Value)
    End If
End If

End Sub

Private Sub txtPath_DblClick()
'Make sure starting sector is valid before prompting
Dim sx As Integer
Dim sy As Integer
If Not ParseSectors(sx, sy, txtOrigin.Text) Then
    Exit Sub
End If

Load frmPromptPath
frmPromptPath.lblSector.Caption = txtOrigin.Text
frmPromptPath.lblEndSector.Caption = txtOrigin.Text
frmPromptPath.Caption = "Select Movement Path"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless

End Sub

Private Sub txtOrigin_Change()

Dim p1 As Integer
Dim sx As Integer
Dim sy As Integer
On Error Resume Next

If ParseSectors(sx, sy, txtOrigin.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        lblOrigin = CStr(rsSectors.Fields("eff").Value) + "% "
        lblOrigin = lblOrigin + colDes.Item(rsSectors.Fields("des").Value)
    End If
End If

End Sub
