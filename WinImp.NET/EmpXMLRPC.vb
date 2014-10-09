Option Strict Off
Option Explicit On
Friend Class EmpXMLRPC
	
	'Change Log:
	'040605 rjk: Created.
	'070805 rjk: Added an error trap for Send and Receive.
	'            Added a fake prompt response for Send.
	'210506 rjk: Switch XMLRPC to asychronous communications.
	'221007 rjk: Handle empty response from send.
	'            Fixed error handler.
	
	Private strSession As String
	Private bStatus As Boolean
	Private strHoldResponse As String
	Enum enumXMLRPCstatus
		XML_WAITING_CLOSE
		XML_CLOSED
		XML_WAITING_OPEN
		XML_OPENED
		XML_WAITING_SEND
		XML_WAITING_RECEIVE
		XML_WAITING_HASDATA
	End Enum
	Private eStatus As enumXMLRPCstatus
	'UPGRADE_ISSUE: MSXML2.ServerXMLHTTP40 object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Public hProxyServer As MSXML2.ServerXMLHTTP40
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		strSession = ""
		strHoldResponse = ""
		bStatus = False
		eStatus = enumXMLRPCstatus.XML_CLOSED
		'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		hProxyServer = Nothing
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub OpenConnection(ByRef strHost As String, ByRef iPort As Short)
		If CreateAndSendHTTPMessage("raw.createSession", strHost, CStr(iPort)) Then
			eStatus = enumXMLRPCstatus.XML_WAITING_OPEN
		End If
		'If Not hProxyServer Is Nothing Then
		'    If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
		'        strSession = hProxyServer.responseXML.selectSingleNode("//params").Text
		'        If Len(strSession) > 0 Then
		'            frmEmpCmd.tmrXMLRPC.Enabled = True
		'            bStatus = True
		'        Else
		'            frmEmpCmd.tmrXMLRPC.Enabled = False
		'            bStatus = False
		'        End If
		'    Else
		'        frmEmpCmd.tmrXMLRPC.Enabled = False
		'        bStatus = False
		'    End If
		'Else
		'    frmEmpCmd.tmrXMLRPC.Enabled = False
		'    bStatus = False
		'    strSession = ""
		'End If
	End Sub
	
	Public Sub CloseConnection()
		If CreateAndSendHTTPMessage("raw.destroySession", strSession) Then
			eStatus = enumXMLRPCstatus.XML_WAITING_CLOSE
		End If
		'bStatus = False
		'frmEmpCmd.tmrXMLRPC.Enabled = False
	End Sub
	
	Public Function Status() As Boolean
		Status = bStatus
	End Function
	
	Public Sub Send(ByRef strMessage As String)
		If CreateAndSendHTTPMessage("raw.sendCommand", strSession, strMessage, CBool(True)) Then
			eStatus = enumXMLRPCstatus.XML_WAITING_SEND
		End If
		'If Not hProxyServer Is Nothing Then
		'    If Not hProxyServer.responseXML.selectSingleNode("//data") Is Nothing Then
		'        Set nlMessages = hProxyServer.responseXML.selectSingleNode("//data").childNodes
		'        For Each nMessage In nlMessages
		'          strHoldResponse = strHoldResponse + nMessage.Text + vbLf
		'        Next
		'    Else
		'        EmpireError "XmlRpc Send", "Invalid Response from the XMLRPC server", hProxyServer.responseText
		'        strHoldResponse = strHoldResponse + "6 0 640" + vbLf
		'    End If
		'End If
	End Sub
	
	Public Function Receive() As String
		'If Len(strHoldResponse) > 0 Then
		'    Receive = strHoldResponse
		'    strHoldResponse = ""
		'    Exit Function
		'End If
		
		If CreateAndSendHTTPMessage("raw.Read", strSession) Then
			eStatus = enumXMLRPCstatus.XML_WAITING_RECEIVE
		End If
		'If Not hProxyServer Is Nothing Then
		'    If Not hProxyServer.responseXML.selectSingleNode("//data") Is Nothing Then
		'        Set nlMessages = hProxyServer.responseXML.selectSingleNode("//data").childNodes
		'        For Each nMessage In nlMessages
		'          Receive = Receive + nMessage.Text + vbLf
		'        Next
		'    Else
		'        EmpireError "XmlRpc Receive", "Invalid Response from the XMLRPC server", hProxyServer.responseText
		'        Receive = ""
		'    End If
		'Else
		'    Receive = ""
		'End If
	End Function
	
	Public Function HasData() As Boolean
		'If Len(strHoldResponse) > 0 Then
		'    HasData = True
		'    Exit Function
		'End If
		
		If eStatus = enumXMLRPCstatus.XML_OPENED Then
			If CreateAndSendHTTPMessage("raw.hasData", strSession) Then
				eStatus = enumXMLRPCstatus.XML_WAITING_HASDATA
			End If
		End If
		'If Not hProxyServer Is Nothing Then
		'    If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
		'        HasData = CBool(hProxyServer.responseXML.selectSingleNode("//params").Text)
		'    End If
		'Else
		'End If
		HasData = False
	End Function
	
	Private Function CreateAndSendHTTPMessage(ByRef strFunction As String, Optional ByRef vOption1 As Object = Nothing, Optional ByRef vOption2 As Object = Nothing, Optional ByRef vOption3 As Object = Nothing) As Boolean
		Dim strBody As String
		CreateAndSendHTTPMessage = False
		
		On Error GoTo Create_Error
		hProxyServer = New MSXML2.ServerXMLHTTP40
		'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.open. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hProxyServer.open("POST ", frmHTTPSettings.strXMLRPCServerName & "/" & frmHTTPSettings.strXMLRPCServerParameters, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.setRequestHeader. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hProxyServer.setRequestHeader("User-Agent", "WinACE XMLRPC")
		'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.setRequestHeader. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hProxyServer.setRequestHeader("Content-Type", "text/xml")
		If frmHTTPSettings.bProxyEnabled Then
			'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.setProxy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hProxyServer.setProxy(MSXML2._SXH_PROXY_SETTING.SXH_PROXY_SET_PROXY, frmHTTPSettings.strProxyServerName)
			If Len(frmHTTPSettings.strProxyUser) > 0 Or Len(frmHTTPSettings.strProxyPassword) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.setProxyCredentials. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hProxyServer.setProxyCredentials(frmHTTPSettings.strProxyUser, frmHTTPSettings.strProxyPassword)
			End If
		End If
		On Error GoTo Send_Error
		strBody = "<?xml version=""1.0""?>" & vbCrLf & "<methodCall>" & vbCrLf & "  <methodName>" & strFunction & "</methodName>" & vbCrLf & "  <params>" & vbCrLf
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not IsNothing(vOption1) Then
			strBody = strBody & ParameterString(vOption1)
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(vOption2) Then
				strBody = strBody & ParameterString(vOption2)
				'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
				If Not IsNothing(vOption3) Then
					strBody = strBody & ParameterString(vOption3)
				End If
			End If
		End If
		strBody = strBody & "  </params>" & vbCrLf & "</methodCall>" & vbCrLf
		Dim ehXMLRPC As EmpXMLRPCEvent
		ehXMLRPC = New EmpXMLRPCEvent
		'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.OnReadyStateChange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hProxyServer.OnReadyStateChange = ehXMLRPC.OnReadyStateChange
		'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.setTimeouts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hProxyServer.setTimeouts(2000000, 60000, 300000, 3600000)
		'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.Send. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hProxyServer.Send(strBody)
		CreateAndSendHTTPMessage = True
		Exit Function
		
Send_Error: 
		'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		hProxyServer = Nothing
		eStatus = enumXMLRPCstatus.XML_CLOSED
		bStatus = False
		frmEmpCmd.tmrXMLRPC.Enabled = False
		Exit Function
		
Create_Error: 
		'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		hProxyServer = Nothing
		eStatus = enumXMLRPCstatus.XML_CLOSED
		bStatus = False
		frmEmpCmd.tmrXMLRPC.Enabled = False
		MsgBox("You can not use HTTP - XMLRPC capability on Windows 95/98/ME", MsgBoxStyle.OKOnly)
	End Function
	
	Public Sub Complete()
		Dim strBuffer As String
		Dim nlMessages As MSXML2.IXMLDOMNodeList
		Dim nMessage As MSXML2.IXMLDOMNode
		
		If hProxyServer Is Nothing Then
			Exit Sub
		End If
		Select Case eStatus
			Case enumXMLRPCstatus.XML_WAITING_OPEN
				'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSession = hProxyServer.responseXML.selectSingleNode("//params").Text
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
					If Len(strSession) > 0 Then
						frmEmpCmd.tmrXMLRPC.Enabled = True
						eStatus = enumXMLRPCstatus.XML_OPENED
						bStatus = True
					Else
						frmEmpCmd.tmrXMLRPC.Enabled = False
						eStatus = enumXMLRPCstatus.XML_CLOSED
						bStatus = False
					End If
				Else
					frmEmpCmd.tmrXMLRPC.Enabled = False
					eStatus = enumXMLRPCstatus.XML_CLOSED
					bStatus = False
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
				End If
			Case enumXMLRPCstatus.XML_WAITING_CLOSE
				frmEmpCmd.tmrXMLRPC.Enabled = False
				eStatus = enumXMLRPCstatus.XML_CLOSED
				bStatus = False
				'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				hProxyServer = Nothing
			Case enumXMLRPCstatus.XML_WAITING_SEND
				'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not hProxyServer.responseXML.selectSingleNode("//data") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nlMessages = hProxyServer.responseXML.selectSingleNode("//data").childNodes
					For	Each nMessage In nlMessages
						If (Len(nMessage.Text) > 0) Then
							strBuffer = strBuffer & nMessage.Text & vbLf
						End If
					Next nMessage
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
					eStatus = enumXMLRPCstatus.XML_OPENED
					frmEmpCmd.ProcessServerResponse(strBuffer)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmpireError("XmlRpc Send", "Invalid Response from the XMLRPC server", hProxyServer.responseText)
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
					eStatus = enumXMLRPCstatus.XML_OPENED
					frmEmpCmd.ProcessServerResponse("6 0 640" & vbLf)
				End If
			Case enumXMLRPCstatus.XML_WAITING_HASDATA
				'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CBool(hProxyServer.responseXML.selectSingleNode("//params").Text) Then
						eStatus = enumXMLRPCstatus.XML_OPENED
						'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						hProxyServer = Nothing
						Receive()
					Else
						'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						hProxyServer = Nothing
						eStatus = enumXMLRPCstatus.XML_OPENED
					End If
				Else
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
					eStatus = enumXMLRPCstatus.XML_OPENED
				End If
			Case enumXMLRPCstatus.XML_WAITING_RECEIVE
				'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not hProxyServer.responseXML.selectSingleNode("//data") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nlMessages = hProxyServer.responseXML.selectSingleNode("//data").childNodes
					For	Each nMessage In nlMessages
						strBuffer = strBuffer & nMessage.Text & vbLf
					Next nMessage
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
					eStatus = enumXMLRPCstatus.XML_OPENED
					frmEmpCmd.ProcessServerResponse(strBuffer)
				Else
					eStatus = enumXMLRPCstatus.XML_OPENED
					'UPGRADE_WARNING: Couldn't resolve default property of object hProxyServer.responseText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmpireError("XmlRpc Receive", "Invalid Response from the XMLRPC server", hProxyServer.responseText)
					'UPGRADE_NOTE: Object hProxyServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					hProxyServer = Nothing
				End If
		End Select
	End Sub
	Private Function ParameterString(ByRef vParam As Object) As String
		Dim strParam As String
		Dim strString As String
		strParam = "    <param>" & vbCrLf & "      <value>" & vbCrLf
		
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(vParam)
			Case VariantType.String
				'UPGRADE_WARNING: Couldn't resolve default property of object vParam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strString = Replace(CStr(vParam), "&", "&amp;", 1, -1, 1)
				strString = Replace(strString, "<", "&lt;", 1, -1, 1)
				strString = Replace(strString, ">", "&gt;", 1, -1, 1)
				strString = Replace(strString, "'", "&apos;", 1, -1, 1)
				strString = Replace(strString, """", "&quot;", 1, -1, 1)
				strParam = strParam & "        <string>" & strString & "</string>" & vbCrLf
			Case VariantType.Boolean
				'UPGRADE_WARNING: Couldn't resolve default property of object vParam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CBool(vParam) = True Then
					strParam = strParam & "        <boolean>1</boolean>" & vbCrLf
				Else
					strParam = strParam & "        <boolean>0</boolean>" & vbCrLf
				End If
			Case Else
				ParameterString = ""
				Exit Function
		End Select
		
		ParameterString = strParam & "      </value>" & vbCrLf & "    </param>" & vbCrLf
	End Function
End Class