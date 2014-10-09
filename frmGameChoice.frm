VERSION 5.00
Begin VB.Form frmGameChoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Empire Game"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Game"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Show Other Games"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Most Recently Played Games"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton cmdMRU 
         Caption         =   "cmdMRU"
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton cmdMRU 
         Caption         =   "cmdMRU"
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton cmdMRU 
         Caption         =   "cmdMRU"
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmdMRU 
         Caption         =   "cmdMRU"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdMRU 
         Caption         =   "cmdMRU"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmGameChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Change Log:
'081103 efj: removal of dead variables
'102703 rjk: Fixed unable to access database when file name (egp) is changed to UCase because copying files between machines.

Private Sub cmdMRU_Click(Index As Integer)
' Dim Reply As Integer   removed efj 8/2003

If Len(cmdMRU(Index).Caption) > 0 Then
    CloseEmpireDB
    Me.Hide
    GameName = cmdMRU(Index).Caption
    
    'Reset the database name, and create a new database
    dbsName = App.Path + "\Games\" + GameName + ".mdb"
    
    If Not DatabaseExists Then
        Exit Sub
    End If
    
    'Open the new database
    OpenEmpireDB
    
    'Connect to the server
    If IsDBOpen Then
        frmEmpCmd.ConnectToServer
    End If

End If

End Sub

Private Sub cmdNew_Click()
Dim ValidName As Boolean
Dim Reply As Integer

On Error GoTo FileNew_Error

FileNew_Reset:
ValidName = False
While Not ValidName
    GameName = InputBox("Enter the name of the Empire Game", "New Game Name", GameName)
    If Len(GameName) = 0 Then
        Exit Sub
    End If
    If Len(GameName) > 25 Then
        ValidName = False
        Reply = MsgBox("Invalid Name Entered", vbRetryCancel, "New Game Error")
        If Reply = vbCancel Then
            Exit Sub
        End If
    Else
        ValidName = True
    End If
Wend

'Check for game already existing
If Len(Dir(App.Path + "\Games\" + GameName + ".egp")) > 0 Then
    Reply = MsgBox("File named " + GameName + " already exists" + vbLf + vbLf + "Do you want to replace it?", vbYesNoCancel + vbQuestion, "New Game")
    If Reply = vbCancel Then
        Exit Sub
    End If
    If Reply = vbNo Then
        GoTo FileNew_Reset
    End If

    'Must be a yes - delete old file
    Kill App.Path + "\Games\" + GameName + ".egp"
End If

'First Close the current Database
CloseEmpireDB

'Reset the database name, and create a new database
dbsName = App.Path + "\Games\" + GameName + ".mdb"

' Make sure there isn't already a file with the name of
' the new database.
If Len(Dir(dbsName)) > 0 Then Kill dbsName

If Len(Dir(App.Path + "\Empdata.mdb")) = 0 Then
    Reply = MsgBox("Error: The EMPDATA.MDB database was not found - you may need to re-install this application" _
        , vbOKOnly + vbCritical, "New Game Error")
    Exit Sub
End If

'Copy the base database into a new database
FileCopy App.Path + "\Empdata.mdb", dbsName

'Open the new database
OpenEmpireDB

'Connect to the server
If IsDBOpen Then
    frmEmpCmd.ConnectToServer
End If

Me.Hide
Exit Sub

FileNew_Error:
End Sub

Private Sub cmdOpen_Click()
FileOpen
If IsDBOpen Then
    Me.Hide
End If
End Sub

Private Sub Form_Load()

' Dim n As Integer   removed efj 8/2003
Dim i As Integer

VerifySubDirectory "Games", True

'Center Form
Me.Left = (Screen.Width - Me.Width) / 2
Me.top = (Screen.Height - Me.Height) / 2

'Load MRU and Set command buttons
For i = 0 To 4
    cmdMRU(i).Caption = GetSetting(APPNAME, "MRUFiles", "MRU" + CStr(i + 1))
    If cmdMRU(i).Caption = vbNullString Then
        cmdMRU(i).Visible = False
    End If
Next i

End Sub

Public Sub FileOpen()

' Dim strIn As String   removed efj 8/2003
Dim n As Integer

On Error GoTo File_Open_Exit

ChDir App.Path

Load frmCommonDialog
' Set CancelError is True
frmCommonDialog.cdb.CancelError = True
' Set flags
frmCommonDialog.cdb.FLAGS = cdlOFNHideReadOnly
' Set filters
frmCommonDialog.cdb.Filter = "All Files (*.*)|*.*|Empire Games" & _
"(*.egp)|*.egp"
' Specify default filter
frmCommonDialog.cdb.FilterIndex = 2
' Display name of selected file
frmCommonDialog.cdb.FileName = vbNullString
If VerifySubDirectory("Games", True) Then
    frmCommonDialog.cdb.InitDir = App.Path + "\Games"
End If

' Display the open dialog box
frmCommonDialog.cdb.ShowOpen

GameName = frmCommonDialog.cdb.FileTitle
'102703 rjk: Added LCase - when copying files, the case can
'            changed to Upper Case.
n = InStr(LCase$(frmCommonDialog.cdb.FileTitle), ".egp")
If n > 0 Then
    GameName = Left$(frmCommonDialog.cdb.FileTitle, n - 1)
End If

'create the database if necessary
'Reset the database name, and create a new database
dbsName = App.Path + "\Games\" + GameName + ".mdb"
If Not DatabaseExists Then
    Exit Sub
End If

'Open the new database
OpenEmpireDB

'Connect to the server
If IsDBOpen Then
    frmEmpCmd.ConnectToServer
End If

File_Open_Exit:
Unload frmCommonDialog
End Sub


Private Function DatabaseExists() As Boolean
Dim Reply
'Make sure the database exists
If Len(Dir(dbsName)) = 0 Then
    Reply = MsgBox("Error: The database " + dbsName + " was not found - Do you want to re-create it ?" _
        , vbCritical + vbYesNo, "Open Database Error")
    If Reply = vbNo Then
        MsgBox "Database Error - Load Aborted", vbOKOnly + vbCritical, "Open Database Error"
        DatabaseExists = False
        Exit Function
    End If
    If Len(Dir(App.Path + "\Empdata.mdb")) = 0 Then
        Reply = MsgBox("Error: The EMPDATA.MDB database was not found - you may need to re-install this application" _
            , vbOKOnly + vbCritical, "Create Database Error")
        DatabaseExists = False
        Exit Function
    End If
    'Copy the base database into a new database
    FileCopy App.Path + "\Empdata.mdb", dbsName
    
End If
DatabaseExists = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then End 'if user closes us, we take the whole thing with us!
End Sub
