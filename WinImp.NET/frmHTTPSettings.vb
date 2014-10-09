Option Strict Off
Option Explicit On
Friend Class frmHTTPSettings
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'080605 rjk: Added XML RPC Parameters - http proxy through firewalls
	
	Public strXMLRPCServerName As String
	Public strXMLRPCServerParameters As String
	Public bProxyEnabled As Boolean
	Public strProxyServerName As String
	Public strProxyUser As String
	Public strProxyPassword As String
	
	'UPGRADE_WARNING: Event chkProxySettings.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkProxySettings_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProxySettings.CheckStateChanged
		If chkProxySettings.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			txtProxyServerName.Enabled = True
			txtProxyUser.Enabled = True
			txtProxyPassword.Enabled = True
		Else
			txtProxyServerName.Enabled = False
			txtProxyUser.Enabled = False
			txtProxyPassword.Enabled = False
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		strXMLRPCServerName = txtXMLRPCServerName.Text
		strXMLRPCServerParameters = txtXMLRPCServerParameters.Text
		If chkProxySettings.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bProxyEnabled = True
		Else
			bProxyEnabled = False
		End If
		strProxyServerName = txtProxyServerName.Text
		strProxyUser = txtProxyUser.Text
		strProxyPassword = txtProxyPassword.Text
		SaveXMLRPCParameters()
		Me.Close()
	End Sub
	
	Private Sub frmHTTPSettings_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		txtXMLRPCServerName.Text = strXMLRPCServerName
		txtXMLRPCServerParameters.Text = strXMLRPCServerParameters
		If bProxyEnabled Then
			chkProxySettings.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkProxySettings.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		txtProxyServerName.Text = strProxyServerName
		txtProxyUser.Text = strProxyUser
		txtProxyPassword.Text = strProxyPassword
		chkProxySettings_CheckStateChanged(chkProxySettings, New System.EventArgs())
	End Sub
	
	Public Sub LoadXMLRPCParameters()
		bXMLRPC = CBool(GetSetting(APPNAME, "XML RPC Options", "Use XMLRPC ", CStr(False)))
		strXMLRPCServerName = GetSetting(APPNAME, "XML RPC Options", "Server Name", "http://emproxy.tryba.nl")
		strXMLRPCServerParameters = GetSetting(APPNAME, "XML RPC Options", "Server Parameters", "empire.servlets/servlet/XmlRpc")
		bProxyEnabled = CBool(GetSetting(APPNAME, "XML RPC Options", "Proxy Enabled", CStr(False)))
		strProxyServerName = GetSetting(APPNAME, "XML RPC Options", "Proxy Server Name", "")
		strProxyUser = GetSetting(APPNAME, "XML RPC Options", "Proxy User", "")
		strProxyPassword = GetSetting(APPNAME, "XML RPC Options", "Proxy Password", "")
	End Sub
	
	Public Sub SaveXMLRPCParameters()
		SaveSetting(APPNAME, "XML RPC Options", "Use XMLRPC ", CStr(bXMLRPC))
		SaveSetting(APPNAME, "XML RPC Options", "Server Name", strXMLRPCServerName)
		SaveSetting(APPNAME, "XML RPC Options", "Server Parameters", strXMLRPCServerParameters)
		SaveSetting(APPNAME, "XML RPC Options", "Proxy Enabled", CStr(bProxyEnabled))
		SaveSetting(APPNAME, "XML RPC Options", "Proxy Server Name", strProxyServerName)
		SaveSetting(APPNAME, "XML RPC Options", "Proxy User", strProxyUser)
		SaveSetting(APPNAME, "XML RPC Options", "Proxy Password", strProxyPassword)
	End Sub
End Class