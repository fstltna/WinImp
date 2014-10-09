VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelationsGrid 
   Caption         =   "Relations Grid"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetNews 
      Caption         =   "Analyse News"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Min             =   30
      Max             =   120
      SelStart        =   50
      TickStyle       =   3
      Value           =   50
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label3 
      Caption         =   "Col Width"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From Country"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Towards Country"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRelationsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'093003 rjk: General reformatting.
'112903 rjk: Added a Get news button and displaying of the telegram count in the grid.
'121203 rjk: Fixed relations grid titles in Deity mode
'180206 rjk: For 4.3.0 get country list, do not depend on relation.
'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun.
'            Switch nation relations to fill in xdump coun instead of relations.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

Public Sub go()
Dim X As Integer, Y As Integer
Dim strCmd As String

If Me.Visible = False Then Exit Sub

Me.MousePointer = vbHourglass

'okay, even after being reworked, the nationsList class is still a little
'kludgy for what we want to do.  live with it for now drk@unxsoft.com

MSFlexGrid1.Rows = Nations.ActiveCount + 1
MSFlexGrid1.Cols = Nations.ActiveCount + 1

Y = 1
'121203 rjk: Fixed relations grid titles in Deity mode
For X = 0 To Nations.Count
    If Nations.NationName(X) <> vbNullString Then
        MSFlexGrid1.TextMatrix(0, Y) = Nations.NationName(X)
        MSFlexGrid1.TextMatrix(Y, 0) = Nations.NationName(X)
        MSFlexGrid1.ColWidth(Y) = 800
        Y = Y + 1
    End If
Next

If Y > 1 Then
    'found some active nations
    
    '100703 rjk: Moved the nation list building logic to Relations report
    '            must do our relations first as this resets the nation file.
    '            Also must skip otherwise the rsNations file will be reset.
    '            Reset the grid and request a refill.
    Nations.Clear
    
    frmEmpCmd.SubmitEmpireCommand "db1" 'start database update
    If Not VersionCheck(4, 3, 0) = VER_LESS Then
        GetCountryList
    End If
    If VersionCheck(4, 3, 4) = VER_LESS Then
        frmEmpCmd.SubmitEmpireCommand "relation"
    
        rsNations.MoveFirst
        While (Not rsNations.EOF)
            If rsNations.Fields("Number") <> CountryNumber Then
                strCmd = "relation " & rsNations.Fields("Number")
                frmEmpCmd.SubmitEmpireCommand strCmd
            End If
            rsNations.MoveNext
        Wend
    Else
        GetCountryList
    End If
    frmEmpCmd.SubmitEmpireCommand "db2" 'finished with database update
    
End If
End Sub

'well, this is a kludge, but there were no good callback hooks
'so that I can know when the database update has completed, so I
'just had to add mine own.

'have frmEmpCmd call this every time it updates the nationsList object
Public Sub callback()
Dim X As Integer, Y As Integer, r As Integer

If Me.Visible = False Then Exit Sub
    
MSFlexGrid1.row = 1
MSFlexGrid1.col = 1
MSFlexGrid1.Redraw = False

For X = 0 To Nations.Count
    For Y = 0 To Nations.Count
        If (Nations.NationName(X) <> vbNullString And Nations.NationName(Y) <> vbNullString) Then
            r = Nations.Relation(Y, X)
            MSFlexGrid1.Text = Nations.RelationString(r)
            If (Nations.telegramCount(Y, X) > 0) And (Y <> X) Then
                MSFlexGrid1.Text = MSFlexGrid1.Text + " (" + CStr(Nations.telegramCount(Y, X)) + ")"
            End If
            Select Case r
                Case iREL_ALLIED: MSFlexGrid1.CellBackColor = RGB(0, 240, 0)
                Case iREL_FRIENDLY: MSFlexGrid1.CellBackColor = RGB(160, 255, 160)
                Case iREL_NEUTRAL: MSFlexGrid1.CellBackColor = RGB(0, 0, 0)
                Case iREL_HOSTILE: MSFlexGrid1.CellBackColor = RGB(255, 150, 140)
                Case iREL_AT_WAR: MSFlexGrid1.CellBackColor = RGB(240, 0, 0)
                Case iREL_MOBILIZING: MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                Case iREL_SITZKRIEG: MSFlexGrid1.CellBackColor = RGB(255, 50, 50)
            End Select
            
            If (MSFlexGrid1.row = MSFlexGrid1.Rows - 1) Then
                MSFlexGrid1.row = 1
                If MSFlexGrid1.col < MSFlexGrid1.Cols - 1 Then MSFlexGrid1.col = MSFlexGrid1.col + 1
            Else
                MSFlexGrid1.row = MSFlexGrid1.row + 1
            End If
            
        End If
    Next
Next

MSFlexGrid1.Redraw = True

Me.MousePointer = vbNormal
End Sub

Private Sub cmdGetNews_Click()
'Setup Designation dialog and execute
Set frmDrawMap.PromptForm = frmPromptNews
frmDrawMap.PromptUp = True

'Put form in proper place
frmPromptNews.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptNews.Width) \ 2
frmPromptNews.top = frmDrawMap.top + frmDrawMap.Height - frmPromptNews.Height

'prompt for parameters
frmPromptNews.Show vbModeless
End Sub


Private Sub Form_Resize()
MSFlexGrid1.Width = Max(Me.Width - 300, 0)
MSFlexGrid1.Height = Max(Me.Height - 300, 0)
End Sub

Private Sub Slider1_Change()
Dim X As Integer

For X = 0 To MSFlexGrid1.Cols - 1
    MSFlexGrid1.ColWidth(X) = Slider1.Value * 10 + 10
Next
End Sub
