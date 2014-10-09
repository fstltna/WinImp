VERSION 5.00
Begin VB.Form frmUnitView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unit Viewing Options"
   ClientHeight    =   3945
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmExpiry 
      Caption         =   "Expired Units"
      Height          =   2055
      Left            =   3240
      TabIndex        =   28
      Top             =   120
      Width           =   1935
      Begin VB.ComboBox cbExpiry 
         Height          =   315
         ItemData        =   "frmUnitView.frx":0000
         Left            =   120
         List            =   "frmUnitView.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Planes"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Land Units"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Ships"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label lblExpiry 
         Caption         =   "Expiry after days"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Test"
      Height          =   375
      Left            =   3240
      TabIndex        =   27
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame frmFriendly 
      Caption         =   "Friendly Relations"
      Height          =   615
      Left            =   3240
      TabIndex        =   24
      Top             =   2640
      Width           =   1935
      Begin VB.OptionButton optFriendlyNeutral 
         Caption         =   "Neutral"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optFriendlyAlly 
         Caption         =   "Allied"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdSetAll 
      Caption         =   "Set All"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Priority List"
      Height          =   3375
      Left            =   5400
      TabIndex        =   20
      Top             =   360
      Width           =   1935
      Begin VB.ListBox lstSort 
         Height          =   2985
         ItemData        =   "frmUnitView.frx":0084
         Left            =   120
         List            =   "frmUnitView.frx":0086
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdDown 
      Height          =   615
      Left            =   7440
      Picture         =   "frmUnitView.frx":0088
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Lower the Priority"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdUp 
      Height          =   615
      Left            =   7440
      Picture         =   "frmUnitView.frx":04CA
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Raise the Priority"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Frame frmNeutral 
      Caption         =   "Neutral Units"
      Height          =   1335
      Left            =   1680
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
      Begin VB.CheckBox chkUnits 
         Caption         =   "Ships"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Land Units"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Planes"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame frmAlly 
      Caption         =   "Allied Units"
      Height          =   1335
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox chkUnits 
         Caption         =   "Planes"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Land Units"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Ships"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame frmEnemy 
      Caption         =   "Enemy Units"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
      Begin VB.CheckBox chkUnits 
         Caption         =   "Ships"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Land Units"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Planes"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame frmOur 
      Caption         =   "Our Units"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox chkUnits 
         Caption         =   "Nukes"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Planes"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Land Units"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "Ships"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "frmUnitView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'112703 rjk: Created
'120303 rjk: Moved friendly unit to frmOption
'            Changed Apply to Test and the changes are not saved.
'140704 rjk: Added Expired Units.
'290606 rjk: Added Our Nuke Units.

Public unitView As New EmpUnitView

Private Sub chkUnits_Click(Index As Integer)
UpdatePriorityList
End Sub

Private Sub UpdatePriorityList()
Dim chkIndex As enumViewingUnits
Dim sortIndex As Integer
Dim location As Integer
Dim strUnitName As String

For chkIndex = 0 To unitView.lastViewingUnit
    location = -1
    For sortIndex = 0 To lstSort.ListCount - 1
        If lstSort.ItemData(sortIndex) = chkIndex Then
            location = sortIndex
            Exit For
        End If
    Next sortIndex
    If chkUnits(chkIndex) <> vbChecked Then
        If location <> -1 Then
            lstSort.RemoveItem (location)
        End If
    Else
        If location = -1 Then
            Select Case chkIndex
            Case 0, 1, 2, 3
                strUnitName = "Our "
            Case 4, 5, 6
                strUnitName = "Enemy "
            Case 7, 8, 9
                strUnitName = "Allied "
            Case 10, 11, 12
                strUnitName = "Neutral "
            Case 13, 14, 15
                strUnitName = "Expired "
            End Select
            lstSort.AddItem strUnitName + chkUnits(chkIndex).Caption, lstSort.ListCount
            lstSort.ItemData(lstSort.NewIndex) = chkIndex
            lstSort.ListIndex = lstSort.NewIndex
        End If
    End If
Next chkIndex
If lstSort.ListIndex >= lstSort.ListCount Then
    lstSort.ListIndex = lstSort.ListCount - 1
End If
If lstSort.ListIndex = lstSort.ListCount - 1 Then
    cmdDown.Enabled = False
Else
    cmdDown.Enabled = True
End If
If lstSort.ListIndex = 0 Then
    cmdUp.Enabled = False
Else
    cmdUp.Enabled = True
End If
End Sub

Private Sub cmdApply_Click()
Apply
frmDrawMap.DrawHexes
End Sub

Private Sub cmdCancel_Click()
unitView.Load
Unload Me
End Sub

Private Sub cmdClearAll_Click()
Dim Index As enumViewingUnits

For Index = 0 To 11
    chkUnits(Index) = vbUnchecked
Next Index
End Sub

Private Sub cmdDown_Click()
Dim strUnitName As String
Dim Data As Integer
Dim Index As enumViewingUnits

strUnitName = lstSort.List(lstSort.ListIndex)
Data = lstSort.ItemData(lstSort.ListIndex)
Index = lstSort.ListIndex
lstSort.RemoveItem Index
lstSort.AddItem strUnitName, Index + 1
lstSort.ItemData(lstSort.NewIndex) = Data
lstSort.ListIndex = Index + 1
End Sub

Private Sub Apply()
Dim Index As enumViewingUnits

For Index = 0 To unitView.lastViewingUnit
    If chkUnits(Index) = vbUnchecked Then
        unitView.bViewUnits(Index) = False
    Else
        unitView.bViewUnits(Index) = True
    End If
Next Index
If optFriendlyAlly Then
    frmOptions.friendlyUnit = FRIEND_ALLIED
Else
    frmOptions.friendlyUnit = FRIEND_NEUTRAL
End If
For Index = 0 To lstSort.ListCount - 1
    unitView.priorityList(Index) = lstSort.ItemData(Index)
Next Index
unitView.priorityListCount = lstSort.ListCount
unitView.iExpiry = cbExpiry.ItemData(cbExpiry.ListIndex)
unitView.Save
End Sub

Private Sub cmdOK_Click()
Apply
unitView.Save
frmOptions.SaveOptions
Unload Me
End Sub

Private Sub cmdSetAll_Click()
Dim Index As enumViewingUnits

For Index = 0 To unitView.lastViewingUnit
    chkUnits(Index) = vbChecked
Next Index
End Sub

Private Sub cmdUp_Click()
Dim strUnitName As String
Dim Data As Integer
Dim Index As Integer

strUnitName = lstSort.List(lstSort.ListIndex)
Data = lstSort.ItemData(lstSort.ListIndex)
Index = lstSort.ListIndex
lstSort.RemoveItem Index
lstSort.AddItem strUnitName, Index - 1
lstSort.ItemData(lstSort.NewIndex) = Data
lstSort.ListIndex = Index - 1
End Sub

Private Sub Form_Load()
Dim Index As enumViewingUnits

UpdatePriorityList
lstSort.ListIndex = 0
For Index = 0 To unitView.lastViewingUnit
    If unitView.bViewUnits(Index) Then
        chkUnits(Index) = vbChecked
    Else
        chkUnits(Index) = vbUnchecked
    End If
Next Index
If frmOptions.friendlyUnit = FRIEND_ALLIED Then
    optFriendlyAlly = True
Else
    optFriendlyNeutral = True
End If
For Index = 0 To cbExpiry.ListCount
    If cbExpiry.ItemData(Index) = unitView.iExpiry Then
        cbExpiry.ListIndex = Index
        Exit For
    End If
Next Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmDrawMap.DrawHexes
End Sub

Private Sub lstSort_Click()
UpdatePriorityList
End Sub
