Option Strict Off
Option Explicit On
Friend Class EmpXMLRPCEvent
	
	Sub OnReadyStateChange()
		'UPGRADE_WARNING: Couldn't resolve default property of object xmlrpcServer.hProxyServer.ReadyState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If xmlrpcServer.hProxyServer.ReadyState = 4 Then
			xmlrpcServer.Complete()
			'    Select Case eStatus
			'    Case XML_WAITING_OPEN
			'        If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
			'            strSession = hProxyServer.responseXML.selectSingleNode("//params").Text
			'            If Len(strSession) > 0 Then
			'                frmEmpCmd.tmrXMLRPC.Enabled = True
			'                bStatus = True
			'            Else
			'               frmEmpCmd.tmrXMLRPC.Enabled = False
			''               bStatus = False
			'           End If
			'       Else
			'          frmEmpCmd.tmrXMLRPC.Enabled = False
			''          bStatus = False
			'      End If
			'     Else
			'         frmEmpCmd.tmrXMLRPC.Enabled = False
			'         bStatus = False
			'        strSession = ""
			'    End If
			'   End Select
		End If
		
	End Sub
End Class