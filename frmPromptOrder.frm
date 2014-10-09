VERSION 5.00
Begin VB.Form frmPromptOrder 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Autofeed (fb, ft only)"
      Height          =   855
      Left            =   120
      TabIndex        =   52
      Top             =   1320
      Width           =   1815
      Begin VB.TextBox txtAutoFeed 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "30"
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox chkAutoFeed 
         Caption         =   "Autofeed"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblAutoFeed 
         Caption         =   "Threshold"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   2952
      TabIndex        =   31
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5688
      TabIndex        =   34
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      Height          =   375
      Left            =   4776
      TabIndex        =   33
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   375
      Left            =   3864
      TabIndex        =   32
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cargo Holds"
      Height          =   1935
      Left            =   2040
      TabIndex        =   36
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtEndCargo 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtEndCargo 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtEndCargo 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtEndCargo 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtEndCargo 
         Height          =   285
         Index           =   6
         Left            =   4680
         TabIndex        =   28
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtEndCargo 
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   24
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtStartCargo 
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   25
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   23
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtStartCargo 
         Height          =   285
         Index           =   6
         Left            =   4680
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Index           =   6
         Left            =   4680
         TabIndex        =   29
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Index           =   6
         Left            =   4680
         TabIndex        =   27
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtStartCargo 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   21
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   19
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtStartCargo 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   17
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtStartCargo 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtStartCargo 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   13
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "type"
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
         Left            =   1080
         TabIndex        =   54
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "level"
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
         Left            =   1080
         TabIndex        =   48
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "level"
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
         Left            =   1080
         TabIndex        =   47
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblHold 
         Alignment       =   2  'Center
         Caption         =   "6"
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
         Index           =   6
         Left            =   4800
         TabIndex        =   46
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblHold 
         Alignment       =   2  'Center
         Caption         =   "5"
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
         Index           =   5
         Left            =   4200
         TabIndex        =   45
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblHold 
         Alignment       =   2  'Center
         Caption         =   "4"
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
         Index           =   4
         Left            =   3600
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "end cargo"
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
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "type"
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
         Left            =   1080
         TabIndex        =   42
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblHold 
         Alignment       =   2  'Center
         Caption         =   "3"
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
         Index           =   3
         Left            =   3000
         TabIndex        =   41
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblHold 
         Alignment       =   2  'Center
         Caption         =   "2"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   40
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblHold 
         Alignment       =   2  'Center
         Caption         =   "1"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   39
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "start cargo"
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
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   30
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   35
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
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
      Left            =   7560
      TabIndex        =   37
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox chkAutoScuttle 
      Caption         =   "Auto-scuttle (ts only)"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "2nd dest"
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
      Left            =   240
      TabIndex        =   51
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "1st dest"
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
      Left            =   240
      TabIndex        =   50
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Order"
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
      TabIndex        =   49
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPromptOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092703 rjk: general reformatting
'110203 rjk: Changed the Cancel caption to Close
'111603 rjk: Code Cleanup - remove the attempt to incorporate the order output into the form?
'            Made strCmd a local variable. Made cmdOK_Click private.
'112503 rjk: Added Food threshold for AutoFeed.
'180104 rjk: Changed from orders to order for set(OK) command.
'111104 rjk: Added the filling of the Order from the database.
'131207 rjk: Added end cargo fields to Ship orders, rename cargo to start cargo
'010108 rjk: Add GetOrders to update form after commands,
'            make OrderLoadData public to allow data updates.

Private Sub cmdCancel_Click() 'Close
Unload Me
End Sub

Private Sub cmdClear_Click()
Dim strCmd As String
'Command: order <SHIP/FLEET> c Clear Orders
strCmd = "order " + txtUnit.Text + " c"
frmEmpCmd.SubmitEmpireCommand strCmd, True
GetOrders True
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String
Dim strCmd2 As String
Dim cargohold As Integer
Dim sx As Integer
Dim sy As Integer

strCmd = "order " + txtUnit.Text '180104 rjk: Changed from orders to order

'Declares: order <SHIP/FLEET> d <dest1> [dest2|scuttle|-] Declare Orders
frmEmpCmd.SubmitEmpireCommand "bf1", False  'run as batch file
If ParseSectors(sx, sy, txtOrigin.Text) Then
    strCmd2 = strCmd + " d " + txtOrigin.Text
        If ParseSectors(sx, sy, txtPath.Text) Then
            strCmd2 = strCmd2 + " " + txtPath.Text
        Else
            If chkAutoScuttle.Value = vbChecked Then
                strCmd2 = strCmd2 + " scuttle"
            End If
        End If
    frmEmpCmd.SubmitEmpireCommand strCmd2, True
End If

'check autofeed
If chkAutoFeed.Value = vbChecked Then
    strCmd2 = strCmd + " l 1 start food " + txtAutoFeed
    frmEmpCmd.SubmitEmpireCommand strCmd2, True
    strCmd2 = strCmd + " l 1 end food " + txtAutoFeed
    frmEmpCmd.SubmitEmpireCommand strCmd2, True
End If

'Cargo levels: Command: order <SHIP/FLEET> l <hold> <start/end> <COMM> <amount>
For cargohold = 1 To 6
    If Len(Trim$(txtStartCargo(cargohold).Text)) > 0 Then
        strCmd2 = strCmd + " l " + CStr(cargohold) + " start " + txtStartCargo(cargohold).Text + " " + txtStart(cargohold).Text
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    End If
    If Len(Trim$(txtEndCargo(cargohold).Text)) > 0 Then
        strCmd2 = strCmd + " l " + CStr(cargohold) + " end " + txtEndCargo(cargohold).Text + " " + txtEnd(cargohold).Text
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    End If
Next cargohold
frmEmpCmd.SubmitEmpireCommand "bf2", False  'run as batch file
GetOrders True
End Sub

Private Sub cmdResume_Click()
Dim strCmd As String
'Command: order <SHIP/FLEET> r Resume Orders
strCmd = "order " + txtUnit.Text + " r"
frmEmpCmd.SubmitEmpireCommand strCmd, True
GetOrders True
End Sub

Private Sub cmdShow_Click()
Dim strCmd As String
Dim strUnit As String
strUnit = txtUnit.Text

If Len(txtUnit.Text) = 0 Then
    strUnit = "*"
End If

frmEmpCmd.SubmitEmpireCommand "br1", True
strCmd = "sorder " + strUnit
frmEmpCmd.SubmitEmpireCommand strCmd, True
strCmd = "qorder " + strUnit
frmEmpCmd.SubmitEmpireCommand strCmd, True
frmEmpCmd.SubmitEmpireCommand "br2", True
End Sub

Private Sub cmdSuspend_Click()
Dim strCmd As String
'Command: order <SHIP/FLEET> s Suspend Orders
strCmd = "order " + txtUnit.Text + " s"
frmEmpCmd.SubmitEmpireCommand strCmd, True
GetOrders True
End Sub

Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
OrderClearForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer

On Error Resume Next
If ParseSectors(sx, sy, txtOrigin.Text) Then
    txtPath.SetFocus
End If
End Sub

Public Sub OrderLoadData()
Dim hldIndex As String
Dim iCargoHold As Integer
Dim strErrorField As String

If InStr(txtUnit.Text, "/") Or Len(txtUnit.Text) <= 0 Or Not IsNumeric(txtUnit.Text) Then
    OrderClearForm
    Exit Sub
End If

hldIndex = rsShipOrders.Index

On Error GoTo OrderLoadData_error

rsShipOrders.Seek "=", txtUnit.Text
If Not rsShipOrders.NoMatch Then
    strErrorField = "start sector"
    txtOrigin.Text = rsShipOrders.Fields("start sector")
    strErrorField = "end sector"
    txtPath.Text = rsShipOrders.Fields("end sector")
    For iCargoHold = 1 To 6
        strErrorField = "start cargo " & CStr(iCargoHold)
        txtStartCargo(iCargoHold).Text = rsShipOrders.Fields("start cargo " & CStr(iCargoHold))
        strErrorField = "end cargo " & CStr(iCargoHold)
        txtEndCargo(iCargoHold).Text = rsShipOrders.Fields("end cargo " & CStr(iCargoHold))
        strErrorField = "start level " & CStr(iCargoHold)
        txtStart(iCargoHold).Text = rsShipOrders.Fields("start level " & CStr(iCargoHold))
        strErrorField = "end level " & CStr(iCargoHold)
        txtEnd(iCargoHold).Text = rsShipOrders.Fields("end level " & CStr(iCargoHold))
    Next iCargoHold
Else
    OrderClearForm
End If

rsShipOrders.Index = hldIndex
Exit Sub

OrderLoadData_error:
EmpireError "OrderLoadData", "Database error:", strErrorField
rsShipOrders.Index = hldIndex
End Sub

Private Sub OrderClearForm()
Dim iCargoHold As Integer

txtOrigin.Text = ""
txtPath.Text = ""
For iCargoHold = 1 To 6
    txtStartCargo(iCargoHold).Text = ""
    txtEndCargo(iCargoHold).Text = ""
    txtStart(iCargoHold).Text = ""
    txtEnd(iCargoHold).Text = ""
Next iCargoHold
End Sub

Private Sub txtUnit_Change()
OrderLoadData
End Sub
