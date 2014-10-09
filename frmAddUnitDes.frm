VERSION 5.00
Begin VB.Form frmAddUnitDes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Unit Type"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Add"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Class"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
      Begin VB.OptionButton Option1 
         Caption         =   "Plane"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Land"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ship"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      Caption         =   $"frmAddUnitDes.frx":0000
      Height          =   1455
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Id"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmAddUnitDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UnitClass As String

Private Sub cmdCancel_Click()
UnitClass = vbNullString
Me.Hide
End Sub

Private Sub cmdOK_Click()

If Len(txtID.Text) = 0 Then
    MsgBox "Must enter a Unit id", vbExclamation + vbOKOnly
    Exit Sub
End If

If Len(txtDesc.Text) = 0 Then
    MsgBox "Must enter a Description", vbExclamation + vbOKOnly
    Exit Sub
End If

If Option1(0).Value Then
    UnitClass = "s"
ElseIf Option1(1).Value Then
    UnitClass = "l"
ElseIf Option1(2).Value Then
    UnitClass = "p"
End If

Me.Hide
End Sub
