VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColorScheme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Color Scheme"
   ClientHeight    =   7164
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5412
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7164
   ScaleWidth      =   5412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameCond 
      Caption         =   "Parameters"
      Height          =   3850
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5175
      Begin VB.CheckBox chkMissing 
         Caption         =   "Missing True"
         Height          =   255
         Left            =   2280
         TabIndex        =   70
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkEnemySectors 
         Caption         =   "Include Enemy Sectors"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cbCond 
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Colors"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   42
         Top             =   1560
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   4
         Left            =   3480
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   5
         Left            =   4320
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   32
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   4950
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Condition Met"
         Height          =   495
         Index           =   9
         Left            =   3360
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Condition Missed"
         Height          =   495
         Index           =   5
         Left            =   4200
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "des=k&&civ#0&&iron>0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "mob>0&&mil>50 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Examples: "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Available fields are:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Conditional String"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame frameGrad 
      Caption         =   "Parameters"
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   5175
      Begin VB.CheckBox chkEnemySectors 
         Caption         =   "Include Enemy Sectors"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   68
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Colors"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   41
         Top             =   2400
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   6
         Left            =   3120
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   7
         Left            =   3960
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   27
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   3000
         ScaleHeight     =   204
         ScaleWidth      =   1524
         TabIndex        =   26
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtHigh 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtLow 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkHide 
         Caption         =   "Hide Units"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkShowVal 
         Caption         =   "Show Value"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "High"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   31
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Low"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   30
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Graduation"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   29
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "High Value"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Low Value"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Graduate Color based on"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FrameNormal 
      Caption         =   "Parameters"
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   5175
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   0
         Left            =   300
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   50
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   8
         Left            =   4380
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Colors"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   1
         Left            =   1325
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   2
         Left            =   2350
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   3
         Left            =   3365
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Land Hex"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Not Yours"
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   49
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Uncharted"
         Height          =   255
         Index           =   3
         Left            =   3185
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Diety Hex"
         Height          =   255
         Index           =   2
         Left            =   2170
         TabIndex        =   47
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Sea Hex"
         Height          =   255
         Index           =   1
         Left            =   1135
         TabIndex        =   46
         Top             =   960
         Width           =   875
      End
   End
   Begin VB.Frame frameDel 
      Caption         =   "Parameters"
      Height          =   1935
      Left            =   120
      TabIndex        =   52
      Top             =   840
      Width           =   5175
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   10
         Left            =   4320
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   58
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   9
         Left            =   3480
         ScaleHeight     =   444
         ScaleWidth      =   444
         TabIndex        =   57
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Colors"
         Height          =   375
         Index           =   4
         Left            =   3600
         TabIndex        =   56
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cbCom 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Route Not Present"
         Height          =   495
         Index           =   12
         Left            =   4200
         TabIndex        =   60
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Route Present"
         Height          =   495
         Index           =   11
         Left            =   3360
         TabIndex        =   59
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCom 
         Caption         =   "Delivery Commodity"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frameDist 
      Caption         =   "Parameters"
      Height          =   1215
      Left            =   120
      TabIndex        =   61
      Top             =   840
      Width           =   5175
      Begin VB.ComboBox cbDistCom 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblDistCom 
         Caption         =   "Distribution Commodity"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame frameDesignation 
      Caption         =   "Designation Selection"
      Height          =   615
      Left            =   120
      TabIndex        =   64
      Top             =   4800
      Width           =   5175
      Begin VB.OptionButton optDesignation 
         Caption         =   "Display Both"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optDesignation 
         Caption         =   "Display New"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optDesignation 
         Caption         =   "Display Current"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   2040
      TabIndex        =   39
      Top             =   6600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4920
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   21
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Color Schemes"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.OptionButton optScheme 
         Caption         =   "Del."
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   55
         ToolTipText     =   "Delivery Routes"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optScheme 
         Caption         =   "Distribution"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optScheme 
         Caption         =   "Graduated"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optScheme 
         Caption         =   "Conditional"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optScheme 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "To change colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   36
      Top             =   5520
      Width           =   5175
      Begin VB.Label Label11 
         Caption         =   "Right click square to change text color"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Left click square to change background color"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmColorScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'100803 rjk: 'Not Your' colors, to be used in COLORSCHEME_NORMAL
'102003 rjk: Added Delivery ColorScheme
'110103 rjk: Added food required to the graduated scheme
'110503 rjk: Added food shortage to the graduated scheme
'112203 rjk: Added displaying the threshold for F4 Distribution
'112203 rjk: Added the ability to show the new designation, current designation or both
'            Added the ability to select graduated display from rsEnemySector table as well.
'112403 rjk: Added the ability to show food shortages and food required (food excess)
'112503 rjk: Added the ability to select conditional display from rsEnemySector table as well.
'            Added a control for determining the default value when a field is missing or unknown
'112903 rjk: Switched ViewUnit to unitView for controlling Unit View options
'120303 rjk: Added Food required and Food Excess for F4 Conditional
'121203 rjk: Switched to Items and Item objects.
'121803 rjk: Added Plague Risk scaled by 10
'122903 rjk: Change Show Value in Hex to Show Value to reflect functionality.
'122903 rjk: Switched to database value from item object.
'122903 rjk: Updated titles for Plague Risk.
'123003 rjk: Fixed not saving special graduated values.
'270104 rjk: Fill in the graduated display with a default item (prevents errors)
'300304 rjk: Fixed to deal with strength fields being Null
'050404 rjk: Added excess civilians to the graduated and conditional color schemes
'180404 rjk: Set default and cancel buttons.  Change "Apply" to "OK".
'270904 rjk: Added saving of the display mode (DD_CURRENT, DD_NEW and DD_BOTH).
'160705 rjk: Added Reacting Planes.  Fixed Deity Problem with Mobility Cost.
'090905 rjk: Made HideUnits a saved parameters across F4 calls and sessions.
'030108 rjk: Increase size of the condition field list.
'170508 rjk: Added FS (Food shortage) to the F4 Conditional.
'220508 rjk: Added Build Cost to F4 graduated and conditional.

Dim HldC(1 To NUMBER_OF_BASIC_COLORS) As Long
Dim HldT(1 To NUMBER_OF_BASIC_COLORS) As Long

Dim oldColorScheme As Integer
Dim oldGradColorItem As String
Dim oldGradColorHigh As Long
Dim oldGradColorLow As Long
Dim oldGradColorDspVal As Boolean
Dim oldCondColorParms() As Variant
Dim oldCondColorParmNumber As Integer
Dim olddspDesignation As enumDisplayDesignation
Dim olddspDesignationSave As enumDisplayDesignation
'Dim oldViewUnits As Boolean

Const NUMBER_OF_CONDITIONAL_ITEMS = 20

'112503 rjk: Added the ability to select conditional display from rsEnemySector table
Private Sub chkMissing_Click()
If chkMissing <> vbUnchecked Then
    bConditionalMissing = True
Else
    bConditionalMissing = False
End If
End Sub

Private Sub chkEnemySectors_Click(Index As Integer)
If chkEnemySectors(Index) <> vbUnchecked Then
    bIncludeEnemySectorsF4Display = True
Else
    bIncludeEnemySectorsF4Display = False
End If
End Sub

Private Sub cmdCancel_Click()
Dim n As Integer
Dim m As Integer

'Restore Colors
For n = 1 To NUMBER_OF_BASIC_COLORS
    BasicColors(n) = HldC(n)
    BasicText(n) = HldT(n)
Next n

'restore current values
ColorScheme = oldColorScheme
GradColorItem = oldGradColorItem
GradColorHigh = oldGradColorHigh
GradColorLow = oldGradColorLow
GradColorDspVal = oldGradColorDspVal
CondColorParmNumber = oldCondColorParmNumber
dspDesignation = olddspDesignation
frmOptions.dspDesignationSave = olddspDesignationSave

ReDim CondColorParms(CondColorParmNumber, 1 To 4)

For n = 1 To CondColorParmNumber
    For m = 1 To 4
        CondColorParms(n, m) = oldCondColorParms(n, m)
    Next m
Next n

Unload Me
End Sub

Private Sub cmdOK_Click()
ApplyColors
frmOptions.SaveOptions
Unload Me
End Sub

Private Sub ApplyColors()
Dim n As Integer
Dim found As Boolean
'112903 rjk: Removed ViewUnits and add unitView for controlling options.
frmUnitView.unitView.Load

If optScheme(COLORSCHEME_NORMAL).Value Then
    'normal scheme
    ColorScheme = COLORSCHEME_NORMAL
ElseIf optScheme(COLORSCHEME_CONDITIONAL).Value Then
    'conditional scheme
    If Not ParseParmString(cbCond.Text) Then
        Exit Sub
    End If
    ColorScheme = COLORSCHEME_CONDITIONAL
    'put conditional string in the combo box
    found = False
    n = 0
    While Not found And n < cbCond.ListCount
        n = n + 1
        found = (cbCond.List(n - 1) = cbCond.Text)
    Wend
    
    If Not found Then
        cbCond.AddItem cbCond.Text, 0
    End If
ElseIf optScheme(COLORSCHEME_GRADUATED).Value Then
    'graduated scheme
    If Len(txtHigh.Text) = 0 Or Len(txtLow.Text) = 0 Then
        MsgBox "Error:  Must have a high and low value.", _
            vbExclamation + vbOKOnly, "Graduated Boundry Error"
        Exit Sub
    End If
    If Not (IsNumeric(txtHigh.Text) And IsNumeric(txtLow.Text)) Then
        MsgBox "Error:  High and low value must be numeric.", _
            vbExclamation + vbOKOnly, "Graduated Boundry Error"
        Exit Sub
    End If
    If CLng(txtHigh.Text) < CLng(txtLow.Text) Then
        MsgBox "Error:  High value must not be less than low value.", _
            vbExclamation + vbOKOnly, "Graduated Boundry Error"
        Exit Sub
    End If
    'passed error checks so set values
    ColorScheme = COLORSCHEME_GRADUATED
    
    Dim eiItem As EmpItem
    
    Set eiItem = Items.FindByConditionalName(Combo1.Text)
    If eiItem Is Nothing Then
        GradColorItem = Combo1.Text
    Else
        GradColorItem = eiItem.DatabaseName
    End If
    GradColorHigh = CLng(txtHigh.Text)
    GradColorLow = CLng(txtLow.Text)
    GradColorDspVal = chkShowVal.Value
    
    If chkHide.Value <> vbUnchecked Then
        frmUnitView.unitView.ClearAllUnits
        SaveSetting APPNAME, "Color Scheme", "Hide Units", True
    Else
        SaveSetting APPNAME, "Color Scheme", "Hide Units", False
    End If
ElseIf optScheme(COLORSCHEME_DISTRIBUTION).Value Then '112203 rjk: Added thresholds for Distribution
    'disribution scheme
    ColorScheme = COLORSCHEME_DISTRIBUTION
    frmUnitView.unitView.ClearAllUnits
    strCommodity = Left$(cbDistCom.Text, 1)
ElseIf optScheme(COLORSCHEME_DELIVERIES).Value Then '102003 rjk: Added ColorScheme for Deliveries
    'disribution scheme
    ColorScheme = COLORSCHEME_DELIVERIES
    strCommodity = Left$(cbCom.Text, 1)
    frmUnitView.unitView.ClearAllUnits
End If

Select Case ColorScheme
Case COLORSCHEME_NORMAL, COLORSCHEME_CONDITIONAL
    If optDesignation(1) Then
        dspDesignation = DD_CURRENT
        frmOptions.dspDesignationSave = DD_CURRENT
    ElseIf optDesignation(2) Then
        dspDesignation = DD_NEW
        frmOptions.dspDesignationSave = DD_NEW
    Else
        dspDesignation = DD_BOTH
        frmOptions.dspDesignationSave = DD_BOTH
    End If
Case COLORSCHEME_GRADUATED, COLORSCHEME_DISTRIBUTION, COLORSCHEME_DELIVERIES
    If optDesignation(1) Then
        dspDesignation = DD_CURRENT
        If frmOptions.dspDesignationSave = DD_BOTH Or _
           frmOptions.dspDesignationSave = DD_BOTH_CURRENT Or _
           frmOptions.dspDesignationSave = DD_BOTH_NEW Then
            frmOptions.dspDesignationSave = DD_BOTH_CURRENT
        Else
            frmOptions.dspDesignationSave = DD_CURRENT
        End If
    Else
        dspDesignation = DD_NEW
        If frmOptions.dspDesignationSave = DD_BOTH Or _
           frmOptions.dspDesignationSave = DD_BOTH_CURRENT Or _
           frmOptions.dspDesignationSave = DD_BOTH_NEW Then
            frmOptions.dspDesignationSave = DD_BOTH_NEW
        Else
            frmOptions.dspDesignationSave = DD_NEW
        End If
    End If
End Select

For n = 1 To NUMBER_OF_BASIC_COLORS
    BasicColors(n) = Picture1(n - 1).BackColor
    BasicText(n) = Picture1(n - 1).ForeColor
Next n

frmDrawMap.DrawHexes
End Sub

Private Sub cmdReset_Click(Index As Integer)
Select Case Index
    Case COLORSCHEME_NORMAL
        SetDefaultBasicColors
    Case COLORSCHEME_CONDITIONAL
        SetDefaultComparisionColors
    Case COLORSCHEME_GRADUATED
        SetDefaultGradColors
    Case COLORSCHEME_DELIVERIES '102003 rjk: Added Deliveries ColorScheme
        SetDefaultDeliveryColors
End Select
FillColors
End Sub

Private Sub cmdTest_Click()
ApplyColors
End Sub

Private Sub Combo1_Click()
Dim v As productionDataType
Dim low As Long
Dim high As Long
Dim n As Long

If rsSectors.BOF And rsSectors.EOF Then Exit Sub

high = 0
low = 99999
rsSectors.MoveFirst
While Not rsSectors.EOF
    Select Case Combo1.Text
    Case "plague risk percentage chance"
        If PlagueRisk(rsSectors) > 1# Then
            n = CLng(Int(PlagueRisk(rsSectors))) '121803 rjk: Added Int to truncate
        Else
            n = 0
        End If
    Case "plague risk calculation X 10" '121803 rjk: added
        n = CLng(PlagueRisk(rsSectors) * 10#) '121803 rjk: Added Int to truncate
    Case "mobility cost"
        Dim sMobCost As Single
        sMobCost = SectorMobCost(rsSectors.Fields("x"), rsSectors.Fields("y"), MT_COMMODITY)
        If sMobCost > 9# Then
            n = 999
        Else
            n = CLng(100# * sMobCost)
        End If
    Case "food required" '110103 rjk: Added food required to the graduated colorscheme
        n = FoodRequired(rsSectors)
    Case "food shortage" '110503 rjk: Added food shortage to the graduated colorscheme
        If FoodRequired(rsSectors) > rsSectors.Fields("food") Then
            n = FoodRequired(rsSectors) - rsSectors.Fields("food")
        Else
            n = 0
        End If
    Case "food excess" '112403 rjk: Added food excess to the graduated colorscheme
        n = rsSectors.Fields("food") - FoodRequired(rsSectors)
    Case "excess civilians" '050404 rjk: Added for Excess civilians
        v = Production(rsSectors)
        If (v.prodValidFlag) Then
            n = CLng(v.ExcessCiv)
        Else
            n = 0
        End If
    Case "build cost"
        v = Production(rsSectors)
        If (v.prodValidFlag) Then
            n = CLng(v.BuildAvailCost)
        Else
            n = 0
        End If
    Case "production"
        v = Production(rsSectors)
        If (v.prodValidFlag) Then
            n = CLng(v.ProdAmount)
        Else
            n = 0
        End If
    Case "reacting planes" '100605 rjk: Added for reacting planes
        n = UpdateScreenInfoWithReactingPlanes(rsSectors.Fields("x"), rsSectors.Fields("y"))
    Case Else
        Dim eiItem As EmpItem
        
        Set eiItem = Items.FindByConditionalName(Combo1.Text)
        If eiItem Is Nothing Then
            If IsNull(rsSectors.Fields(Combo1.Text)) Then
                n = -1
            Else
                n = rsSectors.Fields(Combo1.Text)
            End If
        Else
            n = eiItem.DatabaseValue(rsSectors)
        End If
    End Select

    If n < low Then
        low = n
    End If
    If n > high Then
        high = n
    End If
    rsSectors.MoveNext
Wend
txtLow.Text = CStr(low)
txtHigh.Text = CStr(high)
End Sub

Private Sub Form_Load()
Dim n As Integer
Dim m As Integer
Dim f As Field
Dim strTemp As String
 
For Each f In rsSectors.Fields
    Dim eiItem As EmpItem
    Set eiItem = Items.FindByDatabaseName(f.Name)
    If Not (eiItem Is Nothing) Then
        Label1(0).Caption = Label1(0).Caption + eiItem.ConditionalName + ", "
        Combo1.AddItem eiItem.ConditionalName
    Else
        Label1(0).Caption = Label1(0).Caption + f.Name + ", "
        If f.Type = dbInteger Or f.Type = dbLong Then
            If f.Name <> "x" And f.Name <> "y" Then
                Combo1.AddItem f.Name
            End If
        End If
    End If
    If GradColorItem = f.Name Then
        Combo1.ListIndex = Combo1.NewIndex
    End If
Next

'120303 rjk: Added Food required and Food Excess for F4 Conditional
'050404 rjk: Added EC for excess civilians
'110605 rjk: Added RP for reacting planes
Label1(0).Caption = Label1(0).Caption + "FR, FS, FE, EC, RP, PD, BC"
Combo1.AddItem "plague risk percentage chance"
If GradColorItem = "plague risk percentage chance" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "plague risk calculation X 10" '121803 rjk: added
If GradColorItem = "plague risk calculation X 10" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "mobility cost"
If GradColorItem = "mobility cost" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "food required" '110103 rjk: Added food required to the graduated colorscheme
If GradColorItem = "food required" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "food shortage" '110503 rjk: Added food shortage to the graduated colorscheme
If GradColorItem = "food shortage" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "food excess" '112403 rjk: Added food excess to the graduated colorscheme
If GradColorItem = "food excess" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "excess civilians" '050404 rjk: Added excess civilians to the graduated colorscheme
If GradColorItem = "excess civilians" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "build cost"
If GradColorItem = "build cost" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "reacting planes" '100605 rjk: Added reacting planes to the graduated colorscheme
If GradColorItem = "reacting planes" Then
    Combo1.ListIndex = Combo1.NewIndex
End If
Combo1.AddItem "production"
If GradColorItem = "production" Then
    Combo1.ListIndex = Combo1.NewIndex
End If

If Combo1.ListIndex = -1 Then
    Combo1.ListIndex = 0
    GradColorItem = Combo1.Text
End If

FrameNormal.Move frameCond.Left, frameCond.top
frameGrad.Move frameCond.Left, frameCond.top

'save current values
oldColorScheme = ColorScheme
oldGradColorItem = GradColorItem
oldGradColorHigh = GradColorHigh
oldGradColorLow = GradColorLow
oldGradColorDspVal = GradColorDspVal
oldCondColorParmNumber = CondColorParmNumber
ReDim oldCondColorParms(oldCondColorParmNumber, 1 To 4)
olddspDesignation = dspDesignation
olddspDesignationSave = frmOptions.dspDesignationSave

'Retrieve Hide Unit option
If GetSetting(APPNAME, "Color Scheme", "Hide Units", True) Then
    chkHide.Value = vbChecked
Else
    chkHide.Value = vbUnchecked
End If

'Save Colors
For n = 1 To NUMBER_OF_BASIC_COLORS
    HldC(n) = BasicColors(n)
    HldT(n) = BasicText(n)
Next n

FillColors

'Center form on screen
Left = (Screen.Width - Width) \ 2
top = (Screen.Height - Height) \ 2

'set controls
optScheme(ColorScheme).Value = True
Select Case dspDesignation
Case DD_BOTH
    optDesignation(3) = True
Case DD_CURRENT
    optDesignation(1) = True
Case DD_NEW
    optDesignation(2) = True
End Select
txtHigh.Text = CStr(GradColorHigh)
txtLow.Text = CStr(GradColorLow)

If GradColorDspVal Then
    chkShowVal.Value = vbChecked
Else
    chkShowVal.Value = 0
End If
    
For n = 1 To CondColorParmNumber
    For m = 1 To 4
        cbCond.Text = cbCond.Text + CStr(CondColorParms(n, m))
        oldCondColorParms(n, m) = CondColorParms(n, m)
    Next m
Next n

'load saved conditional items from the registry
For n = 1 To NUMBER_OF_CONDITIONAL_ITEMS
    strTemp = GetSetting(APPNAME, "Colors", "Conditions" + CStr(n), vbNullString)
    If Len(strTemp) > 0 Then
        cbCond.AddItem strTemp
    End If
Next n
 
chkEnemySectors(0).Value = vbUnchecked
'112503 rjk: Added the ability to select conditional display from rsEnemySector table
bIncludeEnemySectorsF4Display = False
bConditionalMissing = True
 
'102003 rjk: Added for deliver routes
LoadCommodityBox cbCom
cbCom.ListIndex = 0
LoadCommodityBox cbDistCom '112203 rjk: Added displaying the threshold for F4 Distribution
cbDistCom.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

'save conditonal items into the registry
Dim n As Integer
For n = 1 To NUMBER_OF_CONDITIONAL_ITEMS
    If n <= cbCond.ListCount Then
        SaveSetting APPNAME, "Colors", "Conditions" + CStr(n), cbCond.List(n - 1)
    End If
Next n

End Sub

Private Sub optScheme_Click(Index As Integer)

FrameNormal.Visible = False
frameGrad.Visible = False
frameCond.Visible = False
frameDist.Visible = False
frameDel.Visible = False
If optScheme(COLORSCHEME_NORMAL).Value Then
    FrameNormal.Visible = True
ElseIf optScheme(COLORSCHEME_CONDITIONAL).Value Then
    frameCond.Visible = True
    '112503 rjk: Added the ability to select conditional display from rsEnemySector table
    If bIncludeEnemySectorsF4Display Then
        chkEnemySectors(1).Value = vbChecked
    Else
        chkEnemySectors(1).Value = vbUnchecked
    End If
ElseIf optScheme(COLORSCHEME_GRADUATED).Value Then
    frameGrad.Visible = True
    If bIncludeEnemySectorsF4Display Then
        chkEnemySectors(0).Value = vbChecked
    Else
        chkEnemySectors(0).Value = vbUnchecked
    End If
ElseIf optScheme(COLORSCHEME_DISTRIBUTION).Value Then '112203 rjk: Added frame for selection of distribution parameters
    frameDist.Visible = True
ElseIf optScheme(COLORSCHEME_DELIVERIES).Value Then '102003 rjk: Added frame for deliver route parameters
    frameDel.Visible = True
End If
If optScheme(COLORSCHEME_NORMAL).Value Or optScheme(COLORSCHEME_CONDITIONAL).Value Then
    optDesignation(3).Enabled = True
    Select Case frmOptions.dspDesignationSave
    Case DD_BOTH, DD_BOTH_NEW, DD_BOTH_CURRENT
        optDesignation(3).Value = True
    Case DD_NEW
        optDesignation(2).Value = True
    Case DD_CURRENT
        optDesignation(1).Value = True
    End Select
Else
    optDesignation(3).Enabled = False
    Select Case frmOptions.dspDesignationSave
    Case DD_NEW, DD_BOTH_NEW, DD_BOTH
        optDesignation(2).Value = True
    Case DD_CURRENT, DD_BOTH_CURRENT
        optDesignation(1).Value = True
    Case DD_BOTH
    End Select
End If
End Sub

'Parse the conditional string
Private Function ParseParmString(ByVal strLine As String) As Boolean
Dim j As Integer
Dim n As Integer
Dim strSep As String
Dim strCurToken As String
Dim strTokens() As String
ReDim strTokens(0 To 0)

'strip spaces and leading question mark
strLine = Trim$(strLine)
If Left$(strLine, 1) = "?" Then
    strLine = Mid$(strLine, 2)
End If

If Len(Trim$(strLine)) = 0 Then
    MsgBox "You must enter a valid conditional string for this scheme", _
        vbExclamation + vbOKOnly, "Conditional String Error"
    ParseParmString = False
    Exit Function
End If

'Break the string into an array of tokens
strSep = "&|=#><"
j = 0
For n = 1 To Len(strLine)
    If InStr(strSep, Mid$(strLine, n, 1)) > 0 Then
        If Len(Trim$(strCurToken)) > 0 Then
            ReDim Preserve strTokens(0 To j)
            strTokens(j) = Trim$(strCurToken)
            j = j + 1
            strCurToken = vbNullString
        End If
        ReDim Preserve strTokens(0 To j)
        strTokens(j) = Mid$(strLine, n, 1)
        j = j + 1
    Else
        strCurToken = strCurToken + Mid$(strLine, n, 1)
    End If
Next n
If Len(Trim$(strCurToken)) > 0 Then
    ReDim Preserve strTokens(0 To j)
    strTokens(j) = Trim$(strCurToken)
    j = j + 1
    strCurToken = vbNullString
End If
        
'Now process the tokens
Dim posToken As Integer
Dim strField As String

CondColorParmNumber = 1
ReDim CondColorParms(1 To MAX_COLOR_CONDITIONS, 1 To 4)

For n = LBound(strTokens) To UBound(strTokens)
    posToken = (n Mod 4)
    Select Case posToken
        Case 0  'extract the field name
            strField = CheckName(strTokens(n))
            If Len(strField) = 0 Then
                MsgBox "Error: '" + strTokens(n) + "' is not a valid field name.", _
                    vbExclamation + vbOKOnly, "Conditional String Error"
                ParseParmString = False
                Exit Function
            End If
            CondColorParms(CondColorParmNumber, 1) = strField
        
        Case 1  'extract the operator
            If InStr("=#><", strTokens(n)) = 0 Then
                MsgBox "Error:  Expected '=', '#', '>', or '<' where '" + strTokens(n) + "' appeared.", _
                    vbExclamation + vbOKOnly, "Conditional String Error"
                ParseParmString = False
                Exit Function
            End If
            CondColorParms(CondColorParmNumber, 2) = strTokens(n)
        
        Case 2  'extract the comparision value
            '120303 rjk: Added Food required and Food Excess for F4 Conditional
            '050404 rjk: Added Excess civilians for F4 Conditional
            '110605 rjk: Added RP - Reacting Planes
            If CondColorParms(CondColorParmNumber, 1) = "FR" Or _
                CondColorParms(CondColorParmNumber, 1) = "FS" Or _
                CondColorParms(CondColorParmNumber, 1) = "FE" Or _
                CondColorParms(CondColorParmNumber, 1) = "EC" Or _
                CondColorParms(CondColorParmNumber, 1) = "BC" Or _
                CondColorParms(CondColorParmNumber, 1) = "RP" Or _
                CondColorParms(CondColorParmNumber, 1) = "PD" Then
                If IsNumeric(CVar(strTokens(n))) Then
                    CondColorParms(CondColorParmNumber, 3) = CLng(strTokens(n))
                Else
                    MsgBox "Error:  Expected a number where '" + strTokens(n) + "' appeared.", _
                        vbExclamation + vbOKOnly, "Conditional String Error"
                    ParseParmString = False
                    Exit Function
                End If
            ElseIf rsSectors.Fields(CondColorParms(CondColorParmNumber, 1)).Type = dbInteger Or _
                rsSectors.Fields(CondColorParms(CondColorParmNumber, 1)).Type = dbLong Then
                If IsNumeric(CVar(strTokens(n))) Then
                    CondColorParms(CondColorParmNumber, 3) = CLng(strTokens(n))
                Else
                    MsgBox "Error:  Expected a number where '" + strTokens(n) + "' appeared.", _
                        vbExclamation + vbOKOnly, "Conditional String Error"
                    ParseParmString = False
                    Exit Function
                End If
            Else
                CondColorParms(CondColorParmNumber, 3) = strTokens(n)
            End If
        
        Case 3  'extract the and / or symbol
            If InStr("&|", strTokens(n)) = 0 Then
                MsgBox "Error:  Expected '&' (AND) or '|' (OR) where '" + strTokens(n) + "' appeared.", _
                    vbExclamation + vbOKOnly, "Conditional String Error"
                ParseParmString = False
                Exit Function
            End If
            CondColorParms(CondColorParmNumber, 4) = strTokens(n)
            CondColorParmNumber = CondColorParmNumber + 1
    End Select
Next

If posToken <> 2 Then
    Select Case posToken
        Case 0  'ended on the field name
            MsgBox "Error:  Expected '=', '#', '>', or '<' next.", _
                vbExclamation + vbOKOnly, "Conditional String Error"
        Case 1  'ended on the operator
            MsgBox "Error:  Expected comparision value next", _
                vbExclamation + vbOKOnly, "Conditional String Error"
        Case 3  'ended on the and / or symbol
            MsgBox "Error:  Conditional string can't end with '&' (AND) or '|' (OR)", _
                vbExclamation + vbOKOnly, "Conditional String Error"
    End Select
    ParseParmString = False
    Exit Function
End If
ParseParmString = True
End Function

Private Function CheckName(strIn As String) As String
'120303 rjk: Check for Food Required and Food Excess
If InStr(strIn, "FR") = 1 Then
    CheckName = "FR"
    Exit Function
ElseIf InStr(strIn, "FS") = 1 Then
    CheckName = "FS"
    Exit Function
ElseIf InStr(strIn, "FE") = 1 Then
    CheckName = "FE"
    Exit Function
ElseIf InStr(strIn, "EC") = 1 Then '050404 rjk: Added Excess Civilians
    CheckName = "EC"
    Exit Function
ElseIf InStr(strIn, "BC") = 1 Then
    CheckName = "BC"
    Exit Function
ElseIf InStr(strIn, "RP") = 1 Then '110605 rjk: Added RP - Reacting Planes
    CheckName = "RP"
    Exit Function
ElseIf InStr(strIn, "PD") = 1 Then
    CheckName = "PD"
    Exit Function
End If
'Match the string fragment to a sector field name
Dim eiItem As EmpItem

Set eiItem = Items.FindByConditionalName(strIn)
If Not (eiItem Is Nothing) Then
    CheckName = eiItem.DatabaseName
    Exit Function
End If

Dim f As Field

For Each f In rsSectors.Fields
    If InStr(f.Name, strIn) = 1 Then
        CheckName = f.Name
        Exit Function
    End If
Next
CheckName = vbNullString
End Function

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    cd1.Color = BasicColors(Index + 1)
    cd1.ShowColor
    BasicColors(Index + 1) = cd1.Color
Else
    cd1.Color = BasicText(Index + 1)
    cd1.ShowColor
    BasicText(Index + 1) = cd1.Color
End If
FillColors
End Sub
Private Sub FillColors()
Dim n As Integer
Dim ch As String

For n = 1 To NUMBER_OF_BASIC_COLORS
    Picture1(n - 1).BackColor = BasicColors(n)
Next n

For n = 1 To NUMBER_OF_BASIC_COLORS
    Picture1(n - 1).AutoRedraw = True
    Picture1(n - 1).ForeColor = BasicText(n)
    SetDrawingFont Picture1(n - 1)
    Picture1(n - 1).FontSize = 14
    ch = "c"
    Picture1(n - 1).CurrentX = (Picture1(n - 1).ScaleWidth - Picture1(n - 1).TextWidth(ch)) / 2
    Picture1(n - 1).CurrentY = (Picture1(n - 1).ScaleHeight - Picture1(n - 1).TextHeight(ch)) / 2
    Picture1(n - 1).Print ch
Next n

FillGradBar
End Sub

Private Sub FillGradBar()

Const NUM_SEGMENTS = 25
Dim n As Long
Dim ncol As Long
Picture2.AutoRedraw = True
For n = 1 To NUM_SEGMENTS
    ncol = BlendedColor(BasicColors(7), BasicColors(8), (n / NUM_SEGMENTS))
    Picture2.Line (((n - 1) / NUM_SEGMENTS) * Picture2.ScaleWidth, 0)-((n / NUM_SEGMENTS) * Picture2.ScaleWidth, Picture2.ScaleHeight), ncol, BF
Next n
End Sub


