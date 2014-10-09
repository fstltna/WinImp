Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmOptions
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'120203 rjk: Created when switched to frmOptions and boolean options
	'120203 rjk: Switched to frmOptions and boolean options.
	'121203 rjk: Added item options
	'122903 rjk: Added Local Help option
	'260104 rjk: Added Theme Game option (St@rW@rs)
	'140204 rjk: Added ImageFileName to store the file name for Background Image
	'150204 rjk: Added to store the border color for the map
	'270304 rjk: Added Theme Game option for Retro
	'300304 rjk: Added RolloverScaling
	'080404 rjk: Added DefenseScaling Option
	'220504 rjk: Added Show Flash Chat
	'010704 rjk: Added FullDumpwithoutSea for deity option.
	'240804 rjk: Added 2K4 Theme Game Option
	'270804 rjk: Added Xdump option to allow use select of either xdump or dump
	'270904 rjk: Added saving the display mode (DD_CURRENT, DD_NEW and DD_BOTH).
	'151104 rjk: Added the ability to start with your Capital instead of 0,0
	'190105 rjk: Set UTF8 in the server when the option is changed or loaded.
	'100405 rjk: UTF8 encapsulated in server version check for 4.2.21 or newer.
	'240405 rjk: Added Custom Script Buttons.  Added Toolbar visibility controls.
	'100505 rjk: Switch Toolbar to listbox and custom buttons images.
	'250505 rjk: Added Telegram Warning, Soft Limit and Hard Limit.
	'260505 rjk: Switched UTF8 from an registry option to login option.
	'290505 rjk: Added ability show Commodity Ratio and an option to control it.
	'120705 rjk: Added option to the control display of Telegram Notification with INFORM OFF in the cmd window
	'            Added option to the control display of Update Commands in the cmd window
	'160705 rjk: Added option to getting of oil content from the sea.
	'210705 rjk: Added option to the control the suppress of the unit damage reports.
	'230705 rjk: Added EE9 Theme Game Option
	'120905 rjk: Changed the default for Suppress Get Oil Content to true.
	'171005 rjk: Fixed Scripts buttons crashing when there invalid image name.
	'281205 rjk: Added South Pacific Theme Game Option for building loaded ships.
	'190206 rjk: Remove xdump, use VersionCheck instead.
	'220406 rjk: Added option to control the logging of flash conversions.
	'050506 rjk: Added South Pacific Search for Atlantis Theme Game Option
	'110606 rjk: Enhanced script buttons to support jump points.
	'090706 rjk: Added Heavy Metal Theme Game Option
	'170906 rjk: Added options for controlling TTS
	'311206 rjk: Added short name options for unit grid.
	'090607 rjk: Removed the RolloverAvail scaling by sector efficiency.
	'301007 rjk: Added option the maximum path length when determining path costs.
	'090311 kab: Modified the date parses to deal more regional effects.
	
	Public bSound As Boolean
	Public bToolbar As Boolean
	Public bStatusBar As Boolean
	Public bClearResultNewCommand As Boolean
	Public bNoIron As Boolean
	Public bSuppressBmapRefresh As Boolean
	Public bShowOnlyWarships As Boolean
	Public bDefaultRetreat As Boolean
	Public bMaxProd As Boolean
	Public bSuppressStrength As Boolean
	Public playersTimeInterval As Short '0 is disabled
	Public bDisplayDumpHeaders As Boolean
	Public friendlyUnit As enumFriendly '120303 rjk: Moved from unitView
	Public bLocalHelp As Boolean '122903 rjk: Added Help
	Public bStarWars As Boolean
	Public bRetro As Boolean '270304 rjk: Added for Retro game
	Public b2K4 As Boolean '250804 rjk: Added for 2K4 game
	Public bEE9 As Boolean '230705 rjk: Added EE9 Theme Game Option
	Public bSP2005 As Boolean '281205 rjk: Added South Pacific Theme Game Option
	Public bSPAtlantis As Boolean '050506 rjk: Added South Pacific Search for Atlantis Theme Game Option
	Public bHeavyMetal As Boolean '090706 rjk: Added Heavy Metal Theme Game Option
	Public strImageFileName As String '140204 rjk: Added to store the file name for Background Image
	Public lBorderColor As Integer '150204 rjk: Added to store the border color for the map
	Public iDefenseScaling As Short '080404 rjk: Added DefenseScaling Option
	Public bFlashChat As Boolean '220504 rjk: Added Flash Chat
	Public bFlashPlayerColors As Boolean '060604 rjk: Added Player colors for Flash Window
	Public bFullDumpwithoutSea As Boolean '010704 rjk: Added for only non-sea sectors for Deities.
	Public dspDesignationSave As HexGrid.enumDisplayDesignation '250904 rjk: Save the operator selection.
	Public bCapital As Boolean '151104 rjk: Added the ability to start at your capital instead of 0,0
	Public bCommodityRatio As Boolean '290505 rjk: Added to control whether Commodity Ratio is shown
	Public bOilContentForSeaSectors As Boolean '110605 rjk: Added to control the Updates requests ocontent
	'120705 rjk: Added to the control display of Telegram Notification with INFORM OFF in the cmd window
	Public bSuppressTelegramNotification As Boolean
	'120705 rjk: Added to the control display of Update Commands in the cmd window
	Public bSuppressUpdateCommands As Boolean
	'210705 rjk: Added to the control the suppress of the unit damage reports
	Public bSuppressUnitDamageReports As Boolean
	Public bFLashLogging As Boolean '220406 rjk: Added to control the logging of flash conversions.
	'170906 rjk: Added options for controlling TTS
	Public bShortNameUnitGrid As Boolean '311206 rjk: Added to display short names in the unit grid.
	Public iMaximumPathLength As Short '301007 rjk: Added to control the maximum path length when
	'determining path costs.
	Public bTTSReportWindow As Boolean
	Public bTTSTelegrams As Boolean
	Public bTTSBulletins As Boolean
	Public bTTSFlashes As Boolean
	
	Enum enumFriendly
		FRIEND_ALLIED
		FRIEND_NEUTRAL
	End Enum
	
	'UPGRADE_WARNING: Event cbItem.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbItem_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbItem.SelectedIndexChanged
		FillItemFrame(Items.FindByLetter(VB.Left(cbItem.Text, 1)))
	End Sub
	
	'UPGRADE_WARNING: Event chkTelegram.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkTelegram_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTelegram.CheckStateChanged
		Dim Index As Short = chkTelegram.GetIndex(eventSender)
		If chkTelegram(Index).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			txtTelegram(Index).Enabled = True
			cmbTelegramSound(Index).Enabled = True
		Else
			txtTelegram(Index).Enabled = False
			cmbTelegramSound(Index).Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmbScriptButton.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbScriptButton_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbScriptButton.SelectedIndexChanged
		FillScriptButtonFrame(VB6.GetItemData(cmbScriptButton, cmbScriptButton.SelectedIndex))
	End Sub
	
	'UPGRADE_WARNING: Event cmbTelegram.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbTelegram_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbTelegram.SelectedIndexChanged
		LoadTelegramFrames()
	End Sub
	
	Private Sub cmdApply_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdApply.Click
		ApplyOptions()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImage.Click
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		
		frmCommonDialog.cdbOpen.Title = "Image File for Background"
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = False
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_NOTE: Variable frmCommonDialog.cdb was renamed frmCommonDialog.cdbOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="94ADAC4D-C65D-414F-A061-8FDC6B83C5EC"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckPathExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.CheckFileExists = True
		frmCommonDialog.cdbOpen.CheckPathExists = True
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Image Files (*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = strImageFileName
		
		frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath
		
		' Display the open dialog box
		frmCommonDialog.cdbOpen.ShowDialog()
		
		strImageFileName = frmCommonDialog.cdbOpen.FileName
		txtImageName.Text = strImageFileName
		bMapBitMapLoaded = False
		frmCommonDialog.Close()
	End Sub
	
	Private Sub cmdScriptBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdScriptBrowse.Click
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		
		frmCommonDialog.cdbOpen.Title = "Script File for Custom Script Button"
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = False
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_NOTE: Variable frmCommonDialog.cdb was renamed frmCommonDialog.cdbOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="94ADAC4D-C65D-414F-A061-8FDC6B83C5EC"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckPathExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.CheckFileExists = True
		frmCommonDialog.cdbOpen.CheckPathExists = True
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = txtScriptFileName.Text
		
		frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
		
		' Display the open dialog box
		frmCommonDialog.cdbOpen.ShowDialog()
		
		txtScriptFileName.Text = frmCommonDialog.cdbOpen.FileName
		
		frmCommonDialog.Close()
	End Sub
	
	Private Sub cmdScriptImageBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdScriptImageBrowse.Click
		ChDir(My.Application.Info.DirectoryPath)
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCommonDialog)
		
		frmCommonDialog.cdbOpen.Title = "Image File for Custom Script Button"
		' Set CancelError is True
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmCommonDialog.cdb.CancelError = False
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_NOTE: Variable frmCommonDialog.cdb was renamed frmCommonDialog.cdbOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="94ADAC4D-C65D-414F-A061-8FDC6B83C5EC"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property frmCommonDialog.cdb.Flags was upgraded to frmCommonDialog.cdbOpen.CheckPathExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		frmCommonDialog.cdbOpen.CheckFileExists = True
		frmCommonDialog.cdbOpen.CheckPathExists = True
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmCommonDialog.cdbOpen.Filter = "All Files (*.*)|*.*|Text Files (*.bmp)|*.bmp"
		' Specify default filter
		frmCommonDialog.cdbOpen.FilterIndex = 2
		' Display name of selected file
		frmCommonDialog.cdbOpen.FileName = txtScriptImageFileName.Text
		
		frmCommonDialog.cdbOpen.InitialDirectory = My.Application.Info.DirectoryPath
		
		' Display the open dialog box
		frmCommonDialog.cdbOpen.ShowDialog()
		
		txtScriptImageFileName.Text = frmCommonDialog.cdbOpen.FileName
		
		frmCommonDialog.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		ApplyOptions()
		SaveOptions()
		Me.Close()
	End Sub
	
	Private Sub frmOptions_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim i As Short
		'handle ctrl+tab to move to the next tab
		If Shift = VB6.ShiftConstants.CtrlMask And KeyCode = System.Windows.Forms.Keys.Tab Then
			i = tbsOptions.SelectedItem.Index
			If i = tbsOptions.Tabs.Count Then
				'last tab so we need to wrap to tab 1
				tbsOptions.SelectedItem = tbsOptions.Tabs(1)
			Else
				'increment the tab
				tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
			End If
		End If
	End Sub
	
	Private Sub frmOptions_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Index As Short
		
		'center the form
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		If frmEmpCmd.bAutoUpdate Then
			chkUpdate.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkUpdate.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If frmEmpCmd.bAutoRead Then
			chkRead.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkRead.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bSound Then
			chkSound.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSound.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bToolbar Then
			chkToolbar.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkToolbar.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bStatusBar Then
			chkStatusBar.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkStatusBar.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bClearResultNewCommand Then
			chkClear.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkClear.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bNoIron Then
			chkNoIron.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkNoIron.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bSuppressBmapRefresh Then
			chkSuppressBmap.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSuppressBmap.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bShowOnlyWarships Then
			chkWarship.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkWarship.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bDefaultRetreat Then
			chkRetreat.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkRetreat.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bMaxProd Then
			chkMaxProd.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkMaxProd.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bSuppressStrength Then
			chkSuppressStrength.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSuppressStrength.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		For Index = 0 To cmbPlayer.Items.Count - 1
			If VB6.GetItemData(cmbPlayer, Index) = playersTimeInterval Then
				cmbPlayer.SelectedIndex = Index
				Exit For
			End If
		Next Index
		If bDisplayDumpHeaders Then
			chkDumpHeaders.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkDumpHeaders.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If friendlyUnit = enumFriendly.FRIEND_NEUTRAL Then
			optFriendlyNeutral.Checked = True
		Else
			optFriendlyAlly.Checked = True
		End If
		If bLocalHelp Then
			chkLocalHelp.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkLocalHelp.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bStarWars Then
			optThemeGames(1).Checked = True
		ElseIf bRetro Then 
			optThemeGames(2).Checked = True
		ElseIf b2K4 Then 
			optThemeGames(3).Checked = True
		ElseIf bEE9 Then 
			optThemeGames(4).Checked = True
		ElseIf bSP2005 Then 
			optThemeGames(5).Checked = True
		ElseIf bSPAtlantis Then 
			optThemeGames(6).Checked = True
		ElseIf bHeavyMetal Then 
			optThemeGames(7).Checked = True
		Else
			optThemeGames(0).Checked = True
		End If
		txtImageName.Text = strImageFileName
		picBorderColor.BackColor = System.Drawing.ColorTranslator.FromOle(lBorderColor)
		txtDefenseScaling.Text = CStr(iDefenseScaling)
		If bFlashChat Then
			chkFlashChat.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkFlashChat.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bFLashLogging Then
			chkFlashLogging.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkFlashLogging.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bFlashPlayerColors Then
			chkFlashPlayerColors.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkFlashPlayerColors.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bFullDumpwithoutSea Then
			chkFullDumpwithoutSea.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkFullDumpwithoutSea.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bCapital Then
			chkCapital.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkCapital.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bCommodityRatio Then
			chkCommodityRatio.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkCommodityRatio.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bOilContentForSeaSectors Then
			chkSuppressOilContentAtSea.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkSuppressOilContentAtSea.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		If bSuppressTelegramNotification Then
			chkSuppressTelegramNotification.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSuppressTelegramNotification.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bSuppressUpdateCommands Then
			chkSuppressUpdateCommands.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSuppressUpdateCommands.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bSuppressUnitDamageReports Then
			chkSuppressUnitDamageReports.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkSuppressUnitDamageReports.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bTTSReportWindow Then
			chkTTSReportWindow.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkTTSReportWindow.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bTTSTelegrams Then
			chkTTSTelegrams.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkTTSTelegrams.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bTTSBulletins Then
			chkTTSBulletins.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkTTSBulletins.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bTTSFlashes Then
			chkTTSFlashes.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkTTSFlashes.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If bShortNameUnitGrid Then
			chkShortName.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkShortName.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		txtPathLength.Text = CStr(iMaximumPathLength)
		
		LoadTelegramPicture()
		LoadItemPicture()
		LoadToolbarPicture()
		LoadScriptButtonPicture()
	End Sub
	
	'UPGRADE_WARNING: Event optCustomScript.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optCustomScript_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCustomScript.CheckedChanged
		If eventSender.Checked Then
			UpdateCustomButtonDisplay(False)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optJumpPoint.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optJumpPoint_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optJumpPoint.CheckedChanged
		If eventSender.Checked Then
			UpdateCustomButtonDisplay(True)
		End If
	End Sub
	
	Private Sub picBorderColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picBorderColor.Click
		frmCommonDialog.cdbColor.Color = picBorderColor.BackColor
		frmCommonDialog.cdbColor.ShowDialog()
		picBorderColor.BackColor = frmCommonDialog.cdbColor.Color
	End Sub
	
	Private Sub tbsOptions_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tbsOptions.ClickEvent
		picView.Visible = False
		picUpdate.Visible = False
		picPlaying.Visible = False
		picTelegram.Visible = False
		picItem.Visible = False
		picThemeGame.Visible = False
		picScriptButton.Visible = False
		picToolbar.Visible = False
		picTTS.Visible = False
		Select Case tbsOptions.SelectedItem.key
			Case "view"
				picView.Visible = True
			Case "update"
				picUpdate.Visible = True
			Case "playing"
				picPlaying.Visible = True
			Case "telegram"
				picTelegram.Visible = True
			Case "item"
				picItem.Visible = True
			Case "theme"
				picThemeGame.Visible = True
			Case "custombuttons"
				picScriptButton.Visible = True
			Case "toolbar"
				picToolbar.Visible = True
			Case "tts"
				picTTS.Visible = True
		End Select
	End Sub
	
	Private Sub ApplyOptions()
		If Not CheckTelegramOptions Then
			Exit Sub
		End If
		'AutoUpdate
		If chkUpdate.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmEmpCmd.bAutoUpdate = True
		Else
			frmEmpCmd.bAutoUpdate = False
		End If
		frmDrawMap.UpdateAutoUpdate()
		'AutoRead
		If chkRead.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmEmpCmd.bAutoRead = True
		Else
			frmEmpCmd.bAutoRead = False
		End If
		frmDrawMap.UpdateAutoRead()
		'Use Sound
		If chkSound.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bSound = True
		Else
			bSound = False
		End If
		'Show Toolbar
		If chkToolbar.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bToolbar = True
		Else
			bToolbar = False
		End If
		'Show Status Bar
		If chkStatusBar.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bStatusBar = True
		Else
			bStatusBar = False
		End If
		frmDrawMap.UpdateBars()
		'Clear Result Box on New Command
		If chkClear.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If Not bClearResultNewCommand Then
				bClearResultNewCommand = True
				frmDrawMap.lstCmdResult.Items.Clear()
			End If
		Else
			bClearResultNewCommand = False
		End If
		'No Iron Game
		If chkNoIron.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If Not bNoIron Then
				bNoIron = True
				If rsNation.BOF Or rsNation.EOF Then
					rsNation.MoveFirst()
					rsNation.Edit()
					rsNation.Fields("NoIron").Value = True
					rsNation.Update()
				End If
			End If
		Else
			If bNoIron Then
				bNoIron = False
				If rsNation.BOF Or rsNation.EOF Then
					rsNation.MoveFirst()
					rsNation.Edit()
					rsNation.Fields("NoIron").Value = False
					rsNation.Update()
				End If
			End If
		End If
		'Suppress Bmap Refresh
		If chkSuppressBmap.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If Not bSuppressBmapRefresh Then
				bSuppressBmapRefresh = True
				If rsNation.BOF Or rsNation.EOF Then
					rsNation.MoveFirst()
					rsNation.Edit()
					rsNation.Fields("SuppressBmap").Value = True
					rsNation.Update()
				End If
			End If
		Else
			If bSuppressBmapRefresh Then
				bSuppressBmapRefresh = False
				If rsNation.BOF Or rsNation.EOF Then
					rsNation.MoveFirst()
					rsNation.Edit()
					rsNation.Fields("SuppressBmap").Value = False
					rsNation.Update()
				End If
			End If
		End If
		'Show Only Warships
		If chkWarship.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If Not bShowOnlyWarships Then
				bShowOnlyWarships = True
				frmDrawMap.DrawHexes()
			End If
		Else
			If bShowOnlyWarships Then
				bShowOnlyWarships = False
				frmDrawMap.DrawHexes()
			End If
		End If
		'Set Retreat on Nav/March
		If chkRetreat.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bDefaultRetreat = True
		Else
			bDefaultRetreat = False
		End If
		'Max. Prod instead of Pop.
		If chkMaxProd.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bMaxProd = True
		Else
			bMaxProd = False
		End If
		'SuppressStrength
		If chkSuppressStrength.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If Not bSuppressStrength Then
				bSuppressStrength = True
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				GetCurrentStrength()
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			End If
		Else
			bSuppressStrength = False
		End If
		playersTimeInterval = VB6.GetItemData(cmbPlayer, cmbPlayer.SelectedIndex)
		frmDrawMap.SetPlayerTimeInterval()
		If chkDumpHeaders.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bDisplayDumpHeaders = True
		Else
			bDisplayDumpHeaders = False
		End If
		If optFriendlyNeutral.Checked Then
			friendlyUnit = enumFriendly.FRIEND_NEUTRAL
		Else
			friendlyUnit = enumFriendly.FRIEND_ALLIED
		End If
		If chkLocalHelp.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bLocalHelp = True
		Else
			bLocalHelp = False
		End If
		
		bStarWars = False
		bRetro = False
		b2K4 = False
		bEE9 = False
		bSP2005 = False
		bSPAtlantis = False
		bHeavyMetal = False
		If optThemeGames(1).Checked Then
			bStarWars = True
		ElseIf optThemeGames(2).Checked Then 
			bRetro = True
		ElseIf optThemeGames(3).Checked Then 
			b2K4 = True
		ElseIf optThemeGames(4).Checked Then 
			bEE9 = True
		ElseIf optThemeGames(5).Checked Then 
			bSP2005 = True
		ElseIf optThemeGames(6).Checked Then 
			bSPAtlantis = True
		ElseIf optThemeGames(7).Checked Then 
			bHeavyMetal = True
		End If
		
		lBorderColor = System.Drawing.ColorTranslator.ToOle(picBorderColor.BackColor)
		bMapBitMapLoaded = False
		iDefenseScaling = Val(txtDefenseScaling.Text)
		If chkFlashChat.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bFlashChat = True
		Else
			bFlashChat = False
		End If
		If chkFlashPlayerColors.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bFlashPlayerColors = True
		Else
			bFlashPlayerColors = False
		End If
		If chkFlashLogging.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bFLashLogging = True
		Else
			bFLashLogging = False
		End If
		If chkFullDumpwithoutSea.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bFullDumpwithoutSea = True
		Else
			bFullDumpwithoutSea = False
		End If
		If chkCapital.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bCapital = True
		Else
			bCapital = False
		End If
		If chkCommodityRatio.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If Not bCommodityRatio Then
				bCommodityRatio = True
				frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
			End If
		Else
			If bCommodityRatio Then
				bCommodityRatio = False
				frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
			End If
		End If
		If chkSuppressOilContentAtSea.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bOilContentForSeaSectors = False
		Else
			bOilContentForSeaSectors = True
		End If
		If chkSuppressTelegramNotification.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bSuppressTelegramNotification = True
		Else
			bSuppressTelegramNotification = False
		End If
		If chkSuppressUpdateCommands.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bSuppressUpdateCommands = True
		Else
			bSuppressUpdateCommands = False
		End If
		If chkSuppressUnitDamageReports.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bSuppressUnitDamageReports = True
		Else
			bSuppressUnitDamageReports = False
		End If
		If chkTTSReportWindow.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bTTSReportWindow = True
		Else
			bTTSReportWindow = False
		End If
		If chkTTSTelegrams.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bTTSTelegrams = True
		Else
			bTTSTelegrams = False
		End If
		If chkTTSBulletins.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bTTSBulletins = True
		Else
			bTTSBulletins = False
		End If
		If chkTTSFlashes.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bTTSFlashes = True
		Else
			bTTSFlashes = False
		End If
		If chkShortName.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bShortNameUnitGrid = True
		Else
			bShortNameUnitGrid = False
		End If
		iMaximumPathLength = CShort(txtPathLength.Text)
		
		frmChat.ControlFlashForm()
		StoreItemNames(Items.FindByLetter(VB.Left(cbItem.Text, 1)))
		StoreScriptButtonFrame(VB6.GetItemData(cmbScriptButton, cmbScriptButton.SelectedIndex))
		StoreToolbarFrame()
		SaveTelegramFrames()
	End Sub
	
	Public Sub LoadOptions()
		'AutoUpdate and AutoRead are set during initial connection setup.
		frmDrawMap.UpdateAutoUpdate()
		frmDrawMap.UpdateAutoRead()
		
		bSound = CBool(GetSetting(APPNAME, "Options", "Use Sound", CStr(True)))
		bToolbar = CBool(GetSetting(APPNAME, "Options", "Show Toolbar", CStr(True)))
		bStatusBar = CBool(GetSetting(APPNAME, "Options", "Show StatusBar", CStr(True)))
		frmDrawMap.UpdateBars()
		
		bClearResultNewCommand = CBool(GetSetting(APPNAME, "Options", "Clear Result Box", CStr(True)))
		bDefaultRetreat = CBool(GetSetting(APPNAME, "Options", "Default Retreat", CStr(True)))
		bMaxProd = CBool(GetSetting(APPNAME, "Options", "Max Prod", CStr(True)))
		bSuppressStrength = CBool(GetSetting(APPNAME, "Options", "Suppress Strength", CStr(False)))
		playersTimeInterval = CShort(GetSetting(APPNAME, "Options", "Player Time Interval", CStr(0)))
		InitializePlayers()
		frmDrawMap.SetPlayerTimeInterval()
		
		bDisplayDumpHeaders = CBool(GetSetting(APPNAME, "Options", "Display Dump Headers", CStr(False)))
		friendlyUnit = CShort(GetSetting(APPNAME, "Options", "Friendly", CStr(enumFriendly.FRIEND_ALLIED)))
		
		If Not (rsNation.BOF And rsNation.EOF) Then
			rsNation.MoveFirst()
			bNoIron = rsNation.Fields("NoIron").Value
			bSuppressBmapRefresh = rsNation.Fields("SuppressBmap").Value
		Else
			bNoIron = False
			bSuppressBmapRefresh = False
		End If
		bLocalHelp = CBool(GetSetting(APPNAME, "Options", "Display Local Help", CStr(True)))
		bStarWars = CBool(GetSetting(APPNAME, "Options", "Star Wars", CStr(False)))
		bRetro = CBool(GetSetting(APPNAME, "Options", "Retro", CStr(False)))
		b2K4 = CBool(GetSetting(APPNAME, "Options", "2K4", CStr(False)))
		bEE9 = CBool(GetSetting(APPNAME, "Options", "EE9", CStr(False)))
		bSP2005 = CBool(GetSetting(APPNAME, "Options", "SP2005", CStr(False)))
		bSPAtlantis = CBool(GetSetting(APPNAME, "Options", "SPAtlantis", CStr(False)))
		bHeavyMetal = CBool(GetSetting(APPNAME, "Options", "HeavyMetal", CStr(False)))
		strImageFileName = GetSetting(APPNAME, "Options", "Map Image File Name", "")
		lBorderColor = CInt(GetSetting(APPNAME, "Options", "Map Border Color", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black).ToString))
		iDefenseScaling = CShort(GetSetting(APPNAME, "Options", "Defense Scaling", CStr(20)))
		bFlashChat = CBool(GetSetting(APPNAME, "Options", "Show Flash Chat", CStr(True)))
		bFlashPlayerColors = CBool(GetSetting(APPNAME, "Options", "Show Flash Player Colors", CStr(False)))
		bFLashLogging = CBool(GetSetting(APPNAME, "Options", "Flash Logging", CStr(False)))
		bFullDumpwithoutSea = CBool(GetSetting(APPNAME, "Options", "FullDump without Sea", CStr(False)))
		bCapital = CBool(GetSetting(APPNAME, "Options", "Capital", CStr(False)))
		bCommodityRatio = CBool(GetSetting(APPNAME, "Options", "CommodityRatio", CStr(True)))
		bOilContentForSeaSectors = CBool(GetSetting(APPNAME, "Options", "Oil Content Updates", CStr(True)))
		bSuppressTelegramNotification = CBool(GetSetting(APPNAME, "Options", "Suppress Telegram Notifications", CStr(False)))
		bSuppressUpdateCommands = CBool(GetSetting(APPNAME, "Options", "Suppress Update Commands", CStr(False)))
		dspDesignationSave = CShort(GetSetting(APPNAME, "Options", "Display Designation", CStr(HexGrid.enumDisplayDesignation.DD_BOTH)))
		bSuppressUnitDamageReports = CBool(GetSetting(APPNAME, "Options", "Suppress Unit Damage Reports", CStr(False)))
		bTTSReportWindow = CBool(GetSetting(APPNAME, "Options", "TTS Report Window", CStr(False)))
		bTTSTelegrams = CBool(GetSetting(APPNAME, "Options", "TTS Telegrams", CStr(False)))
		bTTSBulletins = CBool(GetSetting(APPNAME, "Options", "TTS Bulletins", CStr(False)))
		bTTSFlashes = CBool(GetSetting(APPNAME, "Options", "TTS Flashes", CStr(False)))
		bShortNameUnitGrid = CBool(GetSetting(APPNAME, "Options", "Short Name Unit Grid", CStr(False)))
		iMaximumPathLength = CShort(GetSetting(APPNAME, "Options", "Maximum Path Length", CStr(20)))
		
		Select Case dspDesignationSave
			Case HexGrid.enumDisplayDesignation.DD_CURRENT
				HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_CURRENT
			Case HexGrid.enumDisplayDesignation.DD_NEW
				HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_NEW
			Case HexGrid.enumDisplayDesignation.DD_BOTH, HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT, HexGrid.enumDisplayDesignation.DD_BOTH_NEW
				HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_BOTH
		End Select
		LoadScriptButtons()
		LoadToolbarVisibility()
		LoadTelegramOptions()
	End Sub
	
	Public Sub SaveOptions()
		'AutoUpdate and AutoRead are saved during initial connection setup.
		SaveSetting(APPNAME, "Options", "Use Sound", CStr(bSound))
		SaveSetting(APPNAME, "Options", "Show Toolbar", CStr(bToolbar))
		SaveSetting(APPNAME, "Options", "Show StatusBar", CStr(bStatusBar))
		SaveSetting(APPNAME, "Options", "Clear Result Box", CStr(bClearResultNewCommand))
		SaveSetting(APPNAME, "Options", "Default Retreat", CStr(bDefaultRetreat))
		SaveSetting(APPNAME, "Options", "Max Prod", CStr(bMaxProd))
		SaveSetting(APPNAME, "Options", "Suppress Strength", CStr(bSuppressStrength))
		SaveSetting(APPNAME, "Options", "Player Time Interval", CStr(playersTimeInterval))
		SaveSetting(APPNAME, "Options", "Display Dump Headers", CStr(bDisplayDumpHeaders))
		SaveSetting(APPNAME, "Options", "Friendly", CStr(friendlyUnit))
		
		If Not (rsNation.BOF And rsNation.EOF) Then
			rsNation.MoveFirst()
			rsNation.Edit()
			rsNation.Fields("NoIron").Value = bNoIron
			rsNation.Fields("SuppressBmap").Value = bSuppressBmapRefresh
			rsNation.Update()
		End If
		SaveSetting(APPNAME, "Options", "Display Local Help", CStr(bLocalHelp))
		SaveSetting(APPNAME, "Options", "Star Wars", CStr(bStarWars))
		SaveSetting(APPNAME, "Options", "Retro", CStr(bRetro))
		SaveSetting(APPNAME, "Options", "2K4", CStr(b2K4))
		SaveSetting(APPNAME, "Options", "EE9", CStr(bEE9))
		SaveSetting(APPNAME, "Options", "SP2005", CStr(bSP2005))
		SaveSetting(APPNAME, "Options", "SPAtlantis", CStr(bSPAtlantis))
		SaveSetting(APPNAME, "Options", "HeavyMetal", CStr(bHeavyMetal))
		SaveSetting(APPNAME, "Options", "Map Image File Name", strImageFileName)
		SaveSetting(APPNAME, "Options", "Map Border Color", CStr(lBorderColor))
		SaveSetting(APPNAME, "Options", "Defense Scaling", CStr(iDefenseScaling))
		SaveSetting(APPNAME, "Options", "Show Flash Chat", CStr(bFlashChat))
		SaveSetting(APPNAME, "Options", "Show Flash Player Colors", CStr(bFlashPlayerColors))
		SaveSetting(APPNAME, "Options", "Flash Logging", CStr(bFLashLogging))
		SaveSetting(APPNAME, "Options", "FullDump without Sea", CStr(bFullDumpwithoutSea))
		SaveSetting(APPNAME, "Options", "Capital", CStr(bCapital))
		SaveSetting(APPNAME, "Options", "CommodityRatio", CStr(bCommodityRatio))
		SaveSetting(APPNAME, "Options", "Oil Content Updates", CStr(bOilContentForSeaSectors))
		SaveSetting(APPNAME, "Options", "Suppress Telegram Notifications", CStr(bSuppressTelegramNotification))
		SaveSetting(APPNAME, "Options", "Suppress Update Commands", CStr(bSuppressUpdateCommands))
		SaveSetting(APPNAME, "Options", "Display Designation", CStr(dspDesignationSave))
		SaveSetting(APPNAME, "Options", "Suppress Unit Damage Reports", CStr(bSuppressUnitDamageReports))
		SaveSetting(APPNAME, "Options", "TTS Report Window", CStr(bTTSReportWindow))
		SaveSetting(APPNAME, "Options", "TTS Telegrams", CStr(bTTSTelegrams))
		SaveSetting(APPNAME, "Options", "TTS Bulletins", CStr(bTTSBulletins))
		SaveSetting(APPNAME, "Options", "TTS Flashes", CStr(bTTSFlashes))
		SaveSetting(APPNAME, "Options", "Short Name Unit Grid", CStr(bShortNameUnitGrid))
		SaveSetting(APPNAME, "Options", "Maximum Path Length", CStr(iMaximumPathLength))
		
		SaveScriptButtons()
		SaveToolbarVisibility()
		SaveTelegramOptions()
	End Sub
	
	Private Sub LoadItemPicture()
		Dim Index As Short
		
		cbItem.Items.Clear()
		For Index = 1 To Items.Count
			cbItem.Items.Add(Items(Index).Letter & " - " & Items(Index).ItemName)
		Next Index
		cbItem.SelectedIndex = 0
		FillItemFrame(Items.FindByLetter(VB.Left(cbItem.Text, 1)))
	End Sub
	
	Private Sub FillItemFrame(ByRef eiItem As EmpItem)
		If eiItem Is Nothing Then
			txtItemFormName.Text = ""
			txtItemSectorPanelLabel.Text = ""
			txtItemDistributionPanelLabel.Text = ""
			txtItemFormLabel.Text = ""
			txtItemDBName.Text = ""
			txtItemIntelligenceNames.Text = ""
			txtItemConditionalName.Text = ""
		Else
			txtItemFormName.Text = eiItem.FormName
			txtItemSectorPanelLabel.Text = eiItem.SectorPanelLabel
			txtItemDistributionPanelLabel.Text = eiItem.DistributionPanelLabel
			txtItemFormLabel.Text = eiItem.FormLabel
			txtItemDBName.Text = eiItem.DatabaseName
			txtItemIntelligenceNames.Text = eiItem.IntelligenceNames
			txtItemConditionalName.Text = eiItem.ConditionalName
		End If
	End Sub
	
	Private Sub StoreItemNames(ByRef eiItem As EmpItem)
		If Not (eiItem Is Nothing) Then
			eiItem.FormName = txtItemFormName.Text
			eiItem.SectorPanelLabel = txtItemSectorPanelLabel.Text
			eiItem.DistributionPanelLabel = txtItemDistributionPanelLabel.Text
			eiItem.FormLabel = txtItemFormLabel.Text
			eiItem.DatabaseName = txtItemDBName.Text
			eiItem.IntelligenceNames = txtItemIntelligenceNames.Text
			eiItem.ConditionalName = txtItemConditionalName.Text
		End If
		eiItem.SaveItem()
	End Sub
	
	'UPGRADE_WARNING: Event txtDefenseScaling.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDefenseScaling_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDefenseScaling.TextChanged
		Dim iDefScaling As Short
		
		iDefScaling = Val(txtDefenseScaling.Text)
		If iDefScaling < 0 Then
			iDefScaling = 0
		ElseIf iDefScaling > 500 Then 
			iDefScaling = 500
		End If
		txtDefenseScaling.Text = CStr(txtDefenseScaling.Text)
	End Sub
	
	'UPGRADE_WARNING: Event txtImageName.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtImageName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImageName.TextChanged
		strImageFileName = txtImageName.Text
		bMapBitMapLoaded = False
	End Sub
	
	Private Sub LoadScriptButtonPicture()
		Dim iIndex As Short
		
		cmbScriptButton.Items.Clear()
		For iIndex = 1 To NUMBER_SCRIPT_BUTTONS
			'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			cmbScriptButton.Items.Add(frmDrawMap.tbMain.Items.Item(28 + iIndex).ToolTipText)
			VB6.SetItemData(cmbScriptButton, cmbScriptButton.Items.Count - 1, iIndex)
		Next iIndex
		cmbScriptButton.SelectedIndex = 0
		FillScriptButtonFrame(1)
	End Sub
	
	Private Sub FillScriptButtonFrame(ByRef iIndex As Short)
		txtScriptFileName.Text = tScriptButtons(iIndex).strFileName
		If tScriptButtons(iIndex).bSendOutputReportWindow Then
			chkScriptReport.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkScriptReport.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		txtScriptImageFileName.Text = tScriptButtons(iIndex).strImageFileName
		cmbJumpPoint.SelectedIndex = tScriptButtons(iIndex).iJumpPoint
		
		If tScriptButtons(iIndex).bJumpPoint Then
			If tScriptButtons(iIndex).iJumpPoint >= 0 Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = True
			Else
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = False
			End If
			optJumpPoint.Checked = System.Windows.Forms.CheckState.Checked
			UpdateCustomButtonDisplay(True)
		Else
			If Len(tScriptButtons(iIndex).strFileName) > 0 Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = True
			Else
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = False
			End If
			UpdateCustomButtonDisplay(False)
			optCustomScript.Checked = System.Windows.Forms.CheckState.Checked
		End If
		
	End Sub
	
	Private Sub StoreScriptButtonFrame(ByRef iIndex As Short)
		tScriptButtons(iIndex).strFileName = txtScriptFileName.Text
		UpdateToolTipText(iIndex)
		If chkScriptReport.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			tScriptButtons(iIndex).bSendOutputReportWindow = True
		Else
			tScriptButtons(iIndex).bSendOutputReportWindow = False
		End If
		If tScriptButtons(iIndex).strImageFileName <> txtScriptImageFileName.Text Then
			If Len(txtScriptImageFileName.Text) = 0 Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).ImageIndex = 3
			Else
				On Error GoTo Error_Exit
				frmDrawMap.ImageTB2.Images.Add(System.Drawing.Image.FromFile(txtScriptImageFileName.Text))
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).ImageIndex = frmDrawMap.ImageTB2.Images.Count
			End If
			tScriptButtons(iIndex).strImageFileName = txtScriptImageFileName.Text
		End If
		tScriptButtons(iIndex).iJumpPoint = cmbJumpPoint.SelectedIndex
		If optJumpPoint.Checked = True Then
			If tScriptButtons(iIndex).iJumpPoint >= 0 Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = True
			Else
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = False
			End If
			tScriptButtons(iIndex).bJumpPoint = True
		Else
			If Len(tScriptButtons(iIndex).strFileName) > 0 Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = True
			Else
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = False
			End If
			tScriptButtons(iIndex).bJumpPoint = False
		End If
		UpdateToolTipText(iIndex)
		'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		VB6.SetItemString(cmbScriptButton, iIndex - 1, frmDrawMap.tbMain.Items.Item(28 + iIndex).ToolTipText)
		cmbScriptButton.SelectedIndex = iIndex - 1
		Exit Sub
		
Error_Exit: 
		EmpireError("Failed to load image for custom script button", CStr(iIndex), (txtScriptImageFileName.Text))
	End Sub
	
	Private Sub SaveScriptButtons()
		Dim iIndex As Short
		
		For iIndex = 1 To NUMBER_SCRIPT_BUTTONS
			SaveSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " File Name", tScriptButtons(iIndex).strFileName)
			SaveSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Output Report Window", CStr(tScriptButtons(iIndex).bSendOutputReportWindow))
			SaveSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Image File Name", tScriptButtons(iIndex).strImageFileName)
			SaveSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Jump Point", CStr(tScriptButtons(iIndex).bJumpPoint))
			SaveSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Jump Point Index", CStr(tScriptButtons(iIndex).iJumpPoint))
		Next iIndex
	End Sub
	
	Private Sub LoadScriptButtons()
		Dim iIndex As Short
		
		For iIndex = 1 To NUMBER_SCRIPT_BUTTONS
			tScriptButtons(iIndex).strFileName = GetSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " File Name", "")
			tScriptButtons(iIndex).bSendOutputReportWindow = CBool(GetSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Output Report Window", CStr(True)))
			tScriptButtons(iIndex).strImageFileName = GetSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Image File Name", "")
			tScriptButtons(iIndex).bJumpPoint = CBool(GetSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Jump Point", CStr(False)))
			tScriptButtons(iIndex).iJumpPoint = CShort(GetSetting(APPNAME, "Options", "Script Button " & CStr(iIndex) & " Jump Point Index", CStr(-1)))
			
			If Len(tScriptButtons(iIndex).strImageFileName) > 0 Then
				On Error GoTo skip_image
				frmDrawMap.ImageTB2.Images.Add(System.Drawing.Image.FromFile(tScriptButtons(iIndex).strImageFileName))
				'UPGRADE_WARNING: MSComctlLib.Buttons method tbMain.Buttons.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				frmDrawMap.tbMain.Items.Add(New System.Windows.Forms.ToolStripButton(28 + iIndex, "Script" & CStr(iIndex),  , System.Windows.Forms.ToolBarButtonStyle.PushButton, frmDrawMap.ImageTB2.Images.Count))
			Else
skip_image: 
				'UPGRADE_WARNING: MSComctlLib.Buttons method tbMain.Buttons.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				frmDrawMap.tbMain.Items.Add(New System.Windows.Forms.ToolStripButton(28 + iIndex, "Script" & CStr(iIndex),  , System.Windows.Forms.ToolBarButtonStyle.PushButton, 3))
			End If
			UpdateToolTipText(iIndex)
			If tScriptButtons(iIndex).bJumpPoint Then
				If tScriptButtons(iIndex).iJumpPoint >= 0 Then
					'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = True
				Else
					'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = False
				End If
			Else
				If Len(tScriptButtons(iIndex).strFileName) > 0 Then
					'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = True
				Else
					'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					frmDrawMap.tbMain.Items.Item(28 + iIndex).Visible = False
				End If
			End If
		Next iIndex
	End Sub
	
	Private Sub UpdateToolTipText(ByRef iIndex As Short)
		'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		frmDrawMap.tbMain.Items.Item(28 + iIndex).ToolTipText = GetCustomButtonDescription(iIndex)
	End Sub
	
	Private Function GetCustomButtonDescription(ByRef iIndex As Object) As String
		Dim iPos As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object iIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If tScriptButtons(iIndex).bJumpPoint Then
			'UPGRADE_WARNING: Couldn't resolve default property of object iIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If tScriptButtons(iIndex).iJumpPoint >= 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object iIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetCustomButtonDescription = "Jump Point " & CStr(tScriptButtons(iIndex).iJumpPoint + 1)
			Else
				GetCustomButtonDescription = "Jump Point"
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object iIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iPos = InStrRev(tScriptButtons(iIndex).strFileName, "\")
			If iPos > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object iIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetCustomButtonDescription = "Custom Script " & " - " & Mid(tScriptButtons(iIndex).strFileName, iPos + 1)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object iIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetCustomButtonDescription = "Custom Script " & " - " & tScriptButtons(iIndex).strFileName
			End If
		End If
	End Function
	
	Private Sub UpdateCustomButtonDisplay(ByRef bJumpPoint As Boolean)
		If bJumpPoint Then
			lblScriptFileName.Text = "Jump Point"
			txtScriptFileName.Visible = False
			chkScriptReport.Visible = False
			cmdScriptBrowse.Visible = False
			cmbJumpPoint.Visible = True
			fraScriptButton.Text = "Jump Point Button"
		Else
			lblScriptFileName.Text = "Script File Name"
			txtScriptFileName.Visible = True
			chkScriptReport.Visible = True
			cmdScriptBrowse.Visible = True
			cmbJumpPoint.Visible = False
			fraScriptButton.Text = "Custom Script Button"
		End If
	End Sub
	
	Private Sub LoadToolbarPicture()
		Dim iIndex As Short
		
		lstToolbar.Items.Clear()
		For iIndex = 1 To 28
			'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_ISSUE: MSComctlLib.IButton property tbMain.Buttons.Item.Style was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If frmDrawMap.tbMain.Items.Item(iIndex).Style = System.Windows.Forms.ToolBarButtonStyle.PushButton Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				lstToolbar.Items.Add(New VB6.ListBoxItem(frmDrawMap.tbMain.Items.Item(iIndex).ToolTipText, iIndex))
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				If frmDrawMap.tbMain.Items.Item(iIndex).Visible Then
					'UPGRADE_ISSUE: ListBox property lstToolbar.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					lstToolbar.SetItemChecked(lstToolbar.NewIndex, True)
				Else
					'UPGRADE_ISSUE: ListBox property lstToolbar.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					lstToolbar.SetItemChecked(lstToolbar.NewIndex, False)
				End If
			End If
		Next iIndex
	End Sub
	
	Private Sub StoreToolbarFrame()
		Dim iIndex As Short
		
		For iIndex = 0 To lstToolbar.Items.Count - 1
			If lstToolbar.GetItemChecked(iIndex) Then
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(VB6.GetItemData(lstToolbar, iIndex)).Visible = True
			Else
				'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				frmDrawMap.tbMain.Items.Item(VB6.GetItemData(lstToolbar, iIndex)).Visible = False
			End If
		Next iIndex
	End Sub
	
	Private Sub LoadToolbarVisibility()
		Dim iIndex As Short
		For iIndex = 1 To 28
			'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			frmDrawMap.tbMain.Items.Item(iIndex).Visible = CBool(GetSetting(APPNAME, "Options", "Toolbar " & CStr(iIndex) & " Visibility", CStr(True)))
		Next iIndex
	End Sub
	
	Private Sub SaveToolbarVisibility()
		Dim iIndex As Short
		
		For iIndex = 1 To 28
			'UPGRADE_WARNING: Lower bound of collection frmDrawMap.tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			SaveSetting(APPNAME, "Options", "Toolbar " & CStr(iIndex) & " Visibility", CStr(frmDrawMap.tbMain.Items.Item(iIndex).Visible))
		Next iIndex
	End Sub
	
	Private Sub LoadTelegramOptions()
		Dim iIndex As Short
		Dim iIndexFrame As Short
		
		For iIndex = WinAceRoutines.enumTelegramOptionType.totTelegram To WinAceRoutines.enumTelegramOptionType.totMOTD
			For iIndexFrame = WinAceRoutines.enumTelegramLevel.tlWarning To WinAceRoutines.enumTelegramLevel.tlHardLimit
				tTelegramOptions(iIndex, iIndexFrame).bEnabled = CBool(GetSetting(APPNAME, "Options", "Telegram " & CStr(iIndex) & " " & CStr(iIndexFrame) & " Enabled", CStr(False)))
				tTelegramOptions(iIndex, iIndexFrame).iColumn = CShort(GetSetting(APPNAME, "Options", "Telegram " & CStr(iIndex) & " " & CStr(iIndexFrame) & " Column", CStr(0)))
				tTelegramOptions(iIndex, iIndexFrame).eTelegramSound = CShort(GetSetting(APPNAME, "Options", "Telegram " & CStr(iIndex) & " " & CStr(iIndexFrame) & " Sound", CStr(WinAceRoutines.enumTelegramSound.tsNoSound)))
			Next iIndexFrame
		Next iIndex
	End Sub
	
	Private Sub LoadTelegramPicture()
		Dim iIndex As Short
		Dim iIndexSound As Short
		
		cmbTelegram.Items.Clear()
		For iIndex = WinAceRoutines.enumTelegramOptionType.totTelegram To WinAceRoutines.enumTelegramOptionType.totMOTD
			Select Case iIndex
				Case WinAceRoutines.enumTelegramOptionType.totTelegram
					cmbTelegram.Items.Add(New VB6.ListBoxItem("Telegram", WinAceRoutines.enumTelegramOptionType.totTelegram))
				Case WinAceRoutines.enumTelegramOptionType.totAnnouncement
					cmbTelegram.Items.Add(New VB6.ListBoxItem("Announcement", WinAceRoutines.enumTelegramOptionType.totAnnouncement))
				Case WinAceRoutines.enumTelegramOptionType.totFlash
					cmbTelegram.Items.Add(New VB6.ListBoxItem("Flash", WinAceRoutines.enumTelegramOptionType.totFlash))
				Case WinAceRoutines.enumTelegramOptionType.totMOTD
					If bDeity Then
						cmbTelegram.Items.Add(New VB6.ListBoxItem("Message of the Day (motd)", WinAceRoutines.enumTelegramOptionType.totMOTD))
					End If
			End Select
		Next iIndex
		For iIndex = WinAceRoutines.enumTelegramLevel.tlWarning To WinAceRoutines.enumTelegramLevel.tlHardLimit
			cmbTelegramSound(iIndex).Items.Clear()
			For iIndexSound = WinAceRoutines.enumTelegramSound.tsNoSound To WinAceRoutines.enumTelegramSound.ts3Beeps
				Select Case iIndexSound
					Case WinAceRoutines.enumTelegramSound.tsNoSound
						cmbTelegramSound(iIndex).Items.Add("No Sound")
					Case WinAceRoutines.enumTelegramSound.ts1Beep
						cmbTelegramSound(iIndex).Items.Add("1 Beep")
					Case WinAceRoutines.enumTelegramSound.ts2Beeps
						cmbTelegramSound(iIndex).Items.Add("2 Beeps")
					Case WinAceRoutines.enumTelegramSound.ts3Beeps
						cmbTelegramSound(iIndex).Items.Add("3 Beeps")
				End Select
				'UPGRADE_ISSUE: ComboBox property cmbTelegramSound.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				VB6.SetItemData(cmbTelegramSound(iIndex), cmbTelegramSound(iIndex).NewIndex, iIndexSound)
			Next iIndexSound
		Next iIndex
		cmbTelegram.SelectedIndex = 0
		LoadTelegramFrames()
	End Sub
	
	Private Sub LoadTelegramFrames()
		Dim iIndex As Short
		
		For iIndex = WinAceRoutines.enumTelegramLevel.tlWarning To WinAceRoutines.enumTelegramLevel.tlHardLimit
			If tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).bEnabled Then
				chkTelegram(iIndex).CheckState = System.Windows.Forms.CheckState.Checked
				txtTelegram(iIndex).Enabled = True
				cmbTelegramSound(iIndex).Enabled = True
			Else
				chkTelegram(iIndex).CheckState = System.Windows.Forms.CheckState.Unchecked
				txtTelegram(iIndex).Enabled = False
				cmbTelegramSound(iIndex).Enabled = False
			End If
			txtTelegram(iIndex).Text = CStr(tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).iColumn)
			cmbTelegramSound(iIndex).SelectedIndex = tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).eTelegramSound
		Next iIndex
	End Sub
	
	Private Sub SaveTelegramFrames()
		Dim iIndex As Short
		
		For iIndex = WinAceRoutines.enumTelegramLevel.tlWarning To WinAceRoutines.enumTelegramLevel.tlHardLimit
			If chkTelegram(iIndex).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).bEnabled = True
			Else
				tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).bEnabled = False
			End If
			tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).iColumn = CShort(txtTelegram(iIndex).Text)
			tTelegramOptions(VB6.GetItemData(cmbTelegram, cmbTelegram.SelectedIndex), iIndex).eTelegramSound = cmbTelegramSound(iIndex).SelectedIndex
		Next iIndex
	End Sub
	
	Private Sub SaveTelegramOptions()
		Dim iIndex As Short
		Dim iIndexFrame As Short
		
		For iIndex = WinAceRoutines.enumTelegramOptionType.totTelegram To WinAceRoutines.enumTelegramOptionType.totMOTD
			For iIndexFrame = WinAceRoutines.enumTelegramLevel.tlWarning To WinAceRoutines.enumTelegramLevel.tlHardLimit
				SaveSetting(APPNAME, "Options", "Telegram " & CStr(iIndex) & " " & CStr(iIndexFrame) & " Enabled", CStr(tTelegramOptions(iIndex, iIndexFrame).bEnabled))
				SaveSetting(APPNAME, "Options", "Telegram " & CStr(iIndex) & " " & CStr(iIndexFrame) & " Column", CStr(tTelegramOptions(iIndex, iIndexFrame).iColumn))
				SaveSetting(APPNAME, "Options", "Telegram " & CStr(iIndex) & " " & CStr(iIndexFrame) & " Sound", CStr(tTelegramOptions(iIndex, iIndexFrame).eTelegramSound))
			Next iIndexFrame
		Next iIndex
	End Sub
	
	Private Function CheckTelegramOptions() As Boolean
		Dim iPrevious As Short
		Dim iIndex As Short
		
		On Error GoTo Telegram_Error
		iPrevious = 0
		For iIndex = WinAceRoutines.enumTelegramLevel.tlWarning To WinAceRoutines.enumTelegramLevel.tlHardLimit
			If chkTelegram(iIndex).CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				If iPrevious >= CShort(txtTelegram(iIndex).Text) Then
					MsgBox("Values for Enabled Telegram Columns must be of increasing value, 0 < Warning < Soft Limit < Hard Limit", MsgBoxStyle.OKOnly)
					CheckTelegramOptions = False
					Exit Function
				End If
				iPrevious = CShort(txtTelegram(iIndex).Text)
			End If
		Next iIndex
		CheckTelegramOptions = True
		Exit Function
		
Telegram_Error: 
		MsgBox("Enabled Columns must be Numeric", MsgBoxStyle.OKOnly)
		CheckTelegramOptions = False
	End Function
End Class