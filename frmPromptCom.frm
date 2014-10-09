VERSION 5.00
Begin VB.Form frmPromptCom 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1515
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtLotNumber 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   1335
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
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cbCom 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblOrigin 
      Alignment       =   1  'Right Justify
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
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Market"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  'Right Justify
      Caption         =   "Price"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblLotNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Lot Number"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblCommodity 
      Alignment       =   1  'Right Justify
      Caption         =   "Commodity"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmPromptCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081103 efj: Added Option Explicit
'101603 rjk: Added functionality for Buy, Sell, Reset and enhanced Market
'101703 rjk: Added Strength fields to Sector database
'101803 rjk: Added Route to using this form
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Public strSector As String
Private bFirstLoad As Boolean

Private Sub cbCom_Click()
Dim strComm As String

If Label2.Caption = "Buy" Then
    strComm = cbCom.Text
    If strComm = "unc. workers" Then
        strComm = "uw"
    End If
    If bFirstLoad Then
        frmEmpCmd.SubmitEmpireCommand "market all", False
    Else
        frmEmpCmd.SubmitEmpireCommand "market " + strComm, False
    End If
End If
End Sub

Private Sub cmdCancel_Click()
If Label2.Caption = "Buy" Or Label2.Caption = "Reset" Then
    If frmReport.Visible Then
        Unload frmReport
    End If
End If
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strComm As String
Dim strCmd As String
Dim sx As Integer
Dim sy As Integer

strComm = cbCom.Text
If strComm = "unc. workers" Then
    strComm = "uw"
End If

Select Case Label2.Caption
Case "Buy"
    If frmReport.Visible Then
        Unload frmReport
    End If
    strCmd = "buy " + strComm
    If txtLotNumber.Text <> "" Then
        strCmd = strCmd + " " + txtLotNumber.Text
        If txtPrice.Text <> "" Then
            strCmd = strCmd + " " + txtPrice.Text
            If ParseSectors(sx, sy, txtOrigin.Text) Then
                rsSectors.Seek "=", sx, sy
                If rsSectors.NoMatch Then
                    MsgBox "You do not own this sector"
                    Exit Sub
                End If
                If rsSectors.Fields("des").Value <> "h" And rsSectors.Fields("des").Value <> "w" Then
                    MsgBox "This is not a valid sector must be either harbor or warehouse"
                    Exit Sub
                End If
                strCmd = strCmd + " " + txtOrigin.Text
            End If
        End If
    End If
Case "Market"
    strCmd = "market " + strComm
    
Case "Reset"
    If frmReport.Visible Then
        Unload frmReport
    End If
    strCmd = "reset " + strComm
    If txtLotNumber.Text <> "" Then
        strCmd = strCmd + " " + txtLotNumber.Text
        If txtPrice.Text <> "" Then
            strCmd = strCmd + " " + txtPrice.Text
        End If
    End If
Case "Sell"
    strCmd = "sell " + strComm
    If txtOrigin.Text <> "" Then
        'Manual indicates multiple sectors are supported but does not seem to work
        'so currently will ensure we are starting with a valid harbor or warehouse
        If ParseSectors(sx, sy, txtOrigin.Text) Then
            rsSectors.Seek "=", sx, sy
            If rsSectors.NoMatch Then
                MsgBox "You do not own this sector"
                Exit Sub
            End If
            If rsSectors.Fields("des").Value <> "h" And rsSectors.Fields("des").Value <> "w" Then
                MsgBox "This is not a valid sector must be either harbor or warehouse"
                Exit Sub
            End If
            strCmd = strCmd + " " + txtOrigin.Text
            If txtLotNumber.Text <> "" Then
                strCmd = strCmd + " " + txtLotNumber.Text
                If txtPrice.Text <> "" Then
                    strCmd = strCmd + " " + txtPrice.Text
                End If
            End If
        End If
    End If
Case "Route"
    strCmd = "route " + strComm + " " + txtOrigin.Text
End Select
frmEmpCmd.SubmitEmpireCommand strCmd, True

'database update
If Label2.Caption <> "Market" Then
    frmEmpCmd.SubmitEmpireCommand "db1", False
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
    frmEmpCmd.SubmitEmpireCommand "db2", False
End If
Unload Me
End Sub

Private Sub Form_Activate()
If bFirstLoad Then
    txtOrigin.Text = strSector
    
    Select Case Label2.Caption
    Case "Buy"
    Case "Market"
        lblLotNumber.Visible = False
        txtLotNumber.Visible = False
        lblPrice.Visible = False
        txtPrice.Visible = False
        lblOrigin.Visible = False
        txtOrigin.Visible = False
        cbCom.AddItem "all", 0 'note: "All" does not work
        cbCom.AddItem "*", 1
    Case "Reset"
        frmEmpCmd.SubmitEmpireCommand "market all", False
        lblOrigin.Visible = False
        txtOrigin.Visible = False
        lblCommodity.Visible = False
        cbCom.Visible = False
    Case "Sell"
        lblLotNumber.Caption = "Quantity"
        lblOrigin.Caption = "from"
    Case "Route"
        LoadCommodityBox cbCom 'load back missing items.
        lblPrice.Visible = False
        txtPrice.Visible = False
        lblOrigin.Visible = False
        txtOrigin.Visible = False
    End Select
    If cbCom.Visible Then
        cbCom.ListIndex = 0
        cbCom.SetFocus
    Else
        txtLotNumber.SetFocus
    End If
    bFirstLoad = False
End If
End Sub

Private Sub Form_Load()
Dim Index As Integer

bFirstLoad = True

'Fill list box with commodities
LoadCommodityBox cbCom

Index = 0
While Index < cbCom.ListCount - 1
    If cbCom.List(Index) = "military" Or cbCom.List(Index) = "civilians" Then
        cbCom.RemoveItem Index
    End If
    Index = Index + 1
Wend

Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
top = frmDrawMap.top + frmDrawMap.Height - Height

'Make Stay always on top
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

'Manual indicates multiple sectors are supported but does not seem to work
'Private Sub txtOrigin_DblClick()
'If Label2.Caption = "Sell" Then
'    Load frmPromptSectors
'    frmPromptSectors.strSectors = txtOrigin.Text
'    frmPromptSectors.SetControls
'    frmPromptSectors.Caption = "Select Sectors"
'    frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
'    frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
'    frmPromptSectors.Show vbModeless
'End If
'End Sub
Private Sub txtOrigin_Change()
If Label2.Caption = "Route" Then
    Load frmPromptSectors
    frmPromptSectors.strSectors = txtOrigin.Text
    frmPromptSectors.SetControls
    frmPromptSectors.Caption = "Select Sectors"
    frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
    frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
    frmPromptSectors.Show vbModeless
End If
End Sub
