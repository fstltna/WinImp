Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmGameChoice
	Inherits System.Windows.Forms.Form
	
	' Change Log:
	'081103 efj: removal of dead variables
	'102703 rjk: Fixed unable to access database when file name (egp) is changed to UCase because copying files between machines.
	
	Private Sub cmdMRU_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMRU.Click
		Dim Index As Short = cmdMRU.GetIndex(eventSender)
		' Dim Reply As Integer   removed efj 8/2003
		
		If Len(cmdMRU(Index).Text) > 0 Then
			CloseEmpireDB()
			Me.Hide()
			GameName = cmdMRU(Index).Text
			
			'Reset the database name, and create a new database
			dbsName = My.Application.Info.DirectoryPath & "\Games\" & GameName & ".mdb"
			
			If Not DatabaseExists Then
				Exit Sub
			End If
			
			'Open the new database
			OpenEmpireDB()
			
			'Connect to the server
			If IsDBOpen Then
				frmEmpCmd.ConnectToServer()
			End If
			
		End If
		
	End Sub
	
	Private Sub cmdNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNew.Click
		Dim ValidName As Boolean
		Dim Reply As Short
		
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
				Reply = MsgBox("Invalid Name Entered", MsgBoxStyle.RetryCancel, "New Game Error")
				If Reply = MsgBoxResult.Cancel Then
					Exit Sub
				End If
			Else
				ValidName = True
			End If
		End While
		
		'Check for game already existing
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(My.Application.Info.DirectoryPath & "\Games\" & GameName & ".egp")) > 0 Then
			Reply = MsgBox("File named " & GameName & " already exists" & vbLf & vbLf & "Do you want to replace it?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, "New Game")
			If Reply = MsgBoxResult.Cancel Then
				Exit Sub
			End If
			If Reply = MsgBoxResult.No Then
				GoTo FileNew_Reset
			End If
			
			'Must be a yes - delete old file
			Kill(My.Application.Info.DirectoryPath & "\Games\" & GameName & ".egp")
		End If
		
		'First Close the current Database
		CloseEmpireDB()
		
		'Reset the database name, and create a new database
		dbsName = My.Application.Info.DirectoryPath & "\Games\" & GameName & ".mdb"
		
		' Make sure there isn't already a file with the name of
		' the new database.
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(dbsName)) > 0 Then Kill(dbsName)
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(My.Application.Info.DirectoryPath & "\Empdata.mdb")) = 0 Then
			Reply = MsgBox("Error: The EMPDATA.MDB database was not found - you may need to re-install this application", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "New Game Error")
			Exit Sub
		End If
		
		'Copy the base database into a new database
		FileCopy(My.Application.Info.DirectoryPath & "\Empdata.mdb", dbsName)
		
		'Open the new database
		OpenEmpireDB()
		
		'Connect to the server
		If IsDBOpen Then
			frmEmpCmd.ConnectToServer()
		End If
		
		Me.Hide()
		Exit Sub
		
FileNew_Error: 
	End Sub
	
	Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
		FileOpen_Renamed()
		If IsDBOpen Then
			Me.Hide()
		End If
	End Sub
	
	Private Sub frmGameChoice_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		' Dim n As Integer   removed efj 8/2003
		Dim i As Short
		
		VerifySubDirectory("Games", True)
		
		'Center Form
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
		
		'Load MRU and Set command buttons
		For i = 0 To 4
			cmdMRU(i).Text = GetSetting(APPNAME, "MRUFiles", "MRU" & CStr(i + 1))
			If cmdMRU(i).Text = vbNullString Then
				cmdMRU(i).Visible = False
			End If
		Next i
		
	End Sub
	
	'UPGRADE_NOTE: FileOpen was upgraded to FileOpen_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub FileOpen_Renamed()
		
		' Dim strIn As String   removed efj 8/2003
		Dim n As Short
		
		On Error GoTo File_Open_Exit
		
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = True
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.ShowReadOnly = False
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Empire Games" & "(*.egp)|*.egp"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = vbNullString
		If VerifySubDirectory("Games", True) Then
			frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath & "\Games"
		End If
		
		' Display the open dialog box
		frmCommonDialog.cdbOpen.ShowDialog()
		
		'UPGRADE_WARNING: CommonDialog property frmCommonDialog.cdb.FileTitle was upgraded to frmCommonDialog.cdb.FileName which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		GameName = frmCommonDialog.cdbOpen.FileName
		'102703 rjk: Added LCase - when copying files, the case can
		'            changed to Upper Case.
		'UPGRADE_WARNING: CommonDialog property frmCommonDialog.cdb.FileTitle was upgraded to frmCommonDialog.cdb.FileName which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		n = InStr(LCase(frmCommonDialog.cdbOpen.FileName), ".egp")
		If n > 0 Then
			'UPGRADE_WARNING: CommonDialog property frmCommonDialog.cdb.FileTitle was upgraded to frmCommonDialog.cdb.FileName which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			GameName = VB.Left(frmCommonDialog.cdbOpen.FileName, n - 1)
		End If
		
		'create the database if necessary
		'Reset the database name, and create a new database
		dbsName = My.Application.Info.DirectoryPath & "\Games\" & GameName & ".mdb"
		If Not DatabaseExists Then
			Exit Sub
		End If
		
		'Open the new database
		OpenEmpireDB()
		
		'Connect to the server
		If IsDBOpen Then
			frmEmpCmd.ConnectToServer()
		End If
		
File_Open_Exit: 
		frmCommonDialog.Close()
	End Sub
	
	
	Private Function DatabaseExists() As Boolean
		Dim Reply As Object
		'Make sure the database exists
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(dbsName)) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Reply. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Reply = MsgBox("Error: The database " & dbsName & " was not found - Do you want to re-create it ?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Open Database Error")
			If Reply = MsgBoxResult.No Then
				MsgBox("Database Error - Load Aborted", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Open Database Error")
				DatabaseExists = False
				Exit Function
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Len(Dir(My.Application.Info.DirectoryPath & "\Empdata.mdb")) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Reply. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Reply = MsgBox("Error: The EMPDATA.MDB database was not found - you may need to re-install this application", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Create Database Error")
				DatabaseExists = False
				Exit Function
			End If
			'Copy the base database into a new database
			FileCopy(My.Application.Info.DirectoryPath & "\Empdata.mdb", dbsName)
			
		End If
		DatabaseExists = True
	End Function
	
	Private Sub frmGameChoice_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If Me.Visible Then End 'if user closes us, we take the whole thing with us!
	End Sub
End Class