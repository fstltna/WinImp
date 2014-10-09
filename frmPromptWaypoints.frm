VERSION 5.00
Begin VB.Form frmPromptWaypoints 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Waypoints"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox lstPoints 
      Height          =   2205
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Waypoints"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmPromptWaypoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'08xx03 efj: removed dead variables
'092703 rjk: general reformatting

Public HoldSource As Form
' Public strSectors As String    8/2003 efj  removed

Private Receiver As Control
' Private Done As Integer    8/2003 efj  removed


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

Dim n As Integer
Dim x As Integer
Dim strSector As String

'Load all waypoints into a string
For x = 0 To lstPoints.ListCount - 1
    n = InStr(lstPoints.List(x), " - ")
    If n > 0 Then
        strSector = strSector + Left$(lstPoints.List(x), n) + ";"
    End If
Next x

'Put the string in the Source
HoldSource.Enabled = True
Receiver.Text = strSector
Receiver.SetFocus
Unload Me
End Sub

Private Sub Command1_Click()
lstPoints.Clear
End Sub

Private Sub Form_Activate()
On Error Resume Next
HoldSource.Enabled = False
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

'Set to recieve focus for clicks on DrawMap
Set HoldSource = frmDrawMap.PromptForm
Set frmDrawMap.PromptForm = Me
Set Receiver = HoldSource.ActiveControl
'Hide the txtOrigin box
'txtOrigin.Move lstPoints.Left, lstPoints.top
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = HoldSource
HoldSource.Enabled = True
Set HoldSource = Nothing
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer
Dim strText As String

On Error Resume Next
If ParseSectors(sx, sy, txtOrigin.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        strText = txtOrigin.Text + " - " + CStr(rsSectors.Fields("eff").Value) + "% "
        strText = strText + colDes.Item(rsSectors.Fields("des").Value)
    Else
        rsBmap.Seek "=", sx, sy
        If Not rsBmap.NoMatch Then
            strText = txtOrigin.Text + " - "
            strText = strText + colDes.Item(rsBmap.Fields("des").Value)
        Else
            strText = txtOrigin.Text + " - unknown"
        End If
    End If
    
    lstPoints.AddItem strText
End If
End Sub

