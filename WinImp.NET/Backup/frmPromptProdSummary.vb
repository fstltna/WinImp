Option Strict Off
Option Explicit On
Friend Class frmPromptProdSummary
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'092703 rjk: General reformatting
	'102503 rjk: Replace Done with bFirstLoad, set txtOrigin as initial field
	'            Removed DontMarkReportSectors, Enabled the user selected of reports
	
	'Dim DontMarkReportSectors As Boolean 102503 rjk: Removed, always enabled
	Dim bFirstLoad As Boolean
	
	Private Sub cmdAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAll.Click
		ProdSummaryReport(0, 1, (Me.ckfood).CheckState, (Me.ckOvrPop).CheckState, (Me.chkIdle).CheckState, (Me.chkBuild).CheckState)
		Me.Close()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'Dim strCmd As String    8/2003 efj  removed
		Dim x1 As Short
		' Dim X2 As Integer    8/2003 efj  removed
		Dim y1 As Short
		' Dim Y2 As Integer    8/2003 efj  removed
		
		If Not ParseSectors(x1, y1, (txtOrigin.Text)) Then
			MsgBox("Must fill in 'from' sector",  , "Error")
			txtOrigin.Focus()
			Exit Sub
		End If
		
		ProdSummaryReport(x1, y1, (Me.ckfood).CheckState, (Me.ckOvrPop).CheckState, (Me.chkIdle).CheckState, (Me.chkBuild).CheckState)
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptProdSummary.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptProdSummary_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If bFirstLoad Then
			txtOrigin.Focus()
			bFirstLoad = False
		End If
	End Sub
	
	Private Sub frmPromptProdSummary_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		'102503 rjk: removed, frame always present
		'Set Check Boxes
		'If DontMarkReportSectors Then
		'    Me.ckfood = 0
		'    Me.ckOvrPop = 0
		'    Me.chkIdle = 0
		'    Me.chkBuild = 0
		'Else
		'    Me.ckfood = vbChecked
		'    Me.ckOvrPop = vbChecked
		'    Me.chkIdle = vbChecked
		'    Me.chkBuild = vbChecked
		'End If
		
		bFirstLoad = True
	End Sub
	
	Private Sub frmPromptProdSummary_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'102503 rjk: removed, frame always present
	'Private Sub frameShow_Click()
	'Toggle support
	'DontMarkReportSectors = Not DontMarkReportSectors
	
	'If DontMarkReportSectors Then
	'    Me.ckfood = 0
	'    Me.ckOvrPop = 0
	'    Me.chkIdle = 0
	'    Me.chkBuild = 0
	'Else
	'    Me.ckfood = vbChecked
	'    Me.ckOvrPop = vbChecked
	'    Me.chkIdle = vbChecked
	'    Me.chkBuild = vbChecked
	'End If
	'End Sub
End Class