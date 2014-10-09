VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpXMLRPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
Public hProxyServer As MSXML2.ServerXMLHTTP40

Private Sub Class_Initialize()
strSession = ""
strHoldResponse = ""
bStatus = False
eStatus = XML_CLOSED
Set hProxyServer = Nothing
End Sub

Public Sub OpenConnection(strHost As String, iPort As Integer)
If CreateAndSendHTTPMessage("raw.createSession", strHost, CStr(iPort)) Then
    eStatus = XML_WAITING_OPEN
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
    eStatus = XML_WAITING_CLOSE
End If
'bStatus = False
'frmEmpCmd.tmrXMLRPC.Enabled = False
End Sub

Public Function Status() As Boolean
Status = bStatus
End Function

Public Sub Send(strMessage As String)
If CreateAndSendHTTPMessage("raw.sendCommand", strSession, strMessage, CBool(True)) Then
    eStatus = XML_WAITING_SEND
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
    eStatus = XML_WAITING_RECEIVE
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

If eStatus = XML_OPENED Then
    If CreateAndSendHTTPMessage("raw.hasData", strSession) Then
        eStatus = XML_WAITING_HASDATA
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

Private Function CreateAndSendHTTPMessage(strFunction As String, Optional vOption1 As Variant, _
                                   Optional vOption2 As Variant, Optional vOption3 As Variant) As Boolean
Dim strBody As String
CreateAndSendHTTPMessage = False

On Error GoTo Create_Error:
Set hProxyServer = New MSXML2.ServerXMLHTTP40
hProxyServer.open "POST ", _
                  frmHTTPSettings.strXMLRPCServerName + "/" + _
                  frmHTTPSettings.strXMLRPCServerParameters, _
                  True
hProxyServer.setRequestHeader "User-Agent", "WinACE XMLRPC"
hProxyServer.setRequestHeader "Content-Type", "text/xml"
If frmHTTPSettings.bProxyEnabled Then
    hProxyServer.setProxy SXH_PROXY_SET_PROXY, frmHTTPSettings.strProxyServerName
    If Len(frmHTTPSettings.strProxyUser) > 0 Or _
       Len(frmHTTPSettings.strProxyPassword) > 0 Then
        hProxyServer.setProxyCredentials frmHTTPSettings.strProxyUser, frmHTTPSettings.strProxyPassword
    End If
End If
On Error GoTo Send_Error
strBody = "<?xml version=""1.0""?>" + vbCrLf + _
    "<methodCall>" + vbCrLf + _
    "  <methodName>" + strFunction + "</methodName>" + vbCrLf + _
    "  <params>" + vbCrLf
If Not IsMissing(vOption1) Then
    strBody = strBody + ParameterString(vOption1)
    If Not IsMissing(vOption2) Then
        strBody = strBody + ParameterString(vOption2)
        If Not IsMissing(vOption3) Then
            strBody = strBody + ParameterString(vOption3)
        End If
    End If
End If
strBody = strBody + "  </params>" + vbCrLf + "</methodCall>" + vbCrLf
Dim ehXMLRPC As EmpXMLRPCEvent
Set ehXMLRPC = New EmpXMLRPCEvent
hProxyServer.OnReadyStateChange = ehXMLRPC
hProxyServer.setTimeouts 2000000, 60000, 300000, 3600000
hProxyServer.Send strBody
CreateAndSendHTTPMessage = True
Exit Function

Send_Error:
Set hProxyServer = Nothing
eStatus = XML_CLOSED
bStatus = False
frmEmpCmd.tmrXMLRPC = False
Exit Function

Create_Error:
Set hProxyServer = Nothing
eStatus = XML_CLOSED
bStatus = False
frmEmpCmd.tmrXMLRPC = False
MsgBox "You can not use HTTP - XMLRPC capability on Windows 95/98/ME", vbOKOnly
End Function

Public Sub Complete()
Attribute Complete.VB_UserMemId = 0
Dim strBuffer As String
Dim nlMessages As IXMLDOMNodeList
Dim nMessage As IXMLDOMNode

If hProxyServer Is Nothing Then
    Exit Sub
End If
Select Case eStatus
Case XML_WAITING_OPEN
    If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
        strSession = hProxyServer.responseXML.selectSingleNode("//params").Text
        Set hProxyServer = Nothing
        If Len(strSession) > 0 Then
            frmEmpCmd.tmrXMLRPC.Enabled = True
            eStatus = XML_OPENED
            bStatus = True
        Else
            frmEmpCmd.tmrXMLRPC.Enabled = False
            eStatus = XML_CLOSED
            bStatus = False
        End If
    Else
        frmEmpCmd.tmrXMLRPC.Enabled = False
        eStatus = XML_CLOSED
        bStatus = False
        Set hProxyServer = Nothing
    End If
Case XML_WAITING_CLOSE
    frmEmpCmd.tmrXMLRPC.Enabled = False
    eStatus = XML_CLOSED
    bStatus = False
    Set hProxyServer = Nothing
Case XML_WAITING_SEND
    If Not hProxyServer.responseXML.selectSingleNode("//data") Is Nothing Then
        Set nlMessages = hProxyServer.responseXML.selectSingleNode("//data").childNodes
        For Each nMessage In nlMessages
            If (Len(nMessage.Text) > 0) Then
                strBuffer = strBuffer + nMessage.Text + vbLf
            End If
        Next
        Set hProxyServer = Nothing
        eStatus = XML_OPENED
        frmEmpCmd.ProcessServerResponse strBuffer
    Else
        EmpireError "XmlRpc Send", "Invalid Response from the XMLRPC server", hProxyServer.responseText
        Set hProxyServer = Nothing
        eStatus = XML_OPENED
        frmEmpCmd.ProcessServerResponse "6 0 640" + vbLf
    End If
Case XML_WAITING_HASDATA
    If Not hProxyServer.responseXML.selectSingleNode("//params") Is Nothing Then
        If CBool(hProxyServer.responseXML.selectSingleNode("//params").Text) Then
            eStatus = XML_OPENED
            Set hProxyServer = Nothing
            Receive
        Else
            Set hProxyServer = Nothing
            eStatus = XML_OPENED
        End If
    Else
        Set hProxyServer = Nothing
        eStatus = XML_OPENED
    End If
Case XML_WAITING_RECEIVE
    If Not hProxyServer.responseXML.selectSingleNode("//data") Is Nothing Then
        Set nlMessages = hProxyServer.responseXML.selectSingleNode("//data").childNodes
        For Each nMessage In nlMessages
          strBuffer = strBuffer + nMessage.Text + vbLf
        Next
        Set hProxyServer = Nothing
        eStatus = XML_OPENED
        frmEmpCmd.ProcessServerResponse strBuffer
    Else
        eStatus = XML_OPENED
        EmpireError "XmlRpc Receive", "Invalid Response from the XMLRPC server", hProxyServer.responseText
        Set hProxyServer = Nothing
    End If
End Select
End Sub
Private Function ParameterString(vParam As Variant) As String
Dim strParam As String
Dim strString As String
strParam = "    <param>" + vbCrLf + "      <value>" + vbCrLf

Select Case VarType(vParam)
Case vbString
    strString = Replace(CStr(vParam), "&", "&amp;", 1, -1, 1)
    strString = Replace(strString, "<", "&lt;", 1, -1, 1)
    strString = Replace(strString, ">", "&gt;", 1, -1, 1)
    strString = Replace(strString, "'", "&apos;", 1, -1, 1)
    strString = Replace(strString, """", "&quot;", 1, -1, 1)
    strParam = strParam + "        <string>" + strString + "</string>" + vbCrLf
Case vbBoolean
  If CBool(vParam) = True Then
    strParam = strParam + "        <boolean>1</boolean>" + vbCrLf
  Else
    strParam = strParam + "        <boolean>0</boolean>" + vbCrLf
  End If
Case Else
    ParameterString = ""
    Exit Function
End Select

ParameterString = strParam + "      </value>" + vbCrLf + "    </param>" + vbCrLf
End Function

