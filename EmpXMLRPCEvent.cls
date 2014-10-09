VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpXMLRPCEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
If xmlrpcServer.hProxyServer.ReadyState = 4 Then
    xmlrpcServer.Complete
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

