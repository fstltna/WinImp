VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpSectorMarker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim sm As New Collection
Const SECTOR_MARKER_DURATION = 5 'default duration in minutes


Public Sub Add(sx As Integer, sy As Integer, Nation As Integer, Message As String)

Dim col As Collection
Dim ncol As New Collection

'first look for it
For Each col In sm
    If col.item("x") = sx And col.item("y") = sy Then
        col.Remove "msg"
        col.Remove "ts"
        col.Remove "nation"
        col.Add Message, "msg"
        col.Add Now, "ts"
        col.Add Nation, "nation"
        Exit Sub
    End If
Next col

ncol.Add sx, "x"
ncol.Add sy, "y"
ncol.Add Message, "msg"
ncol.Add Nation, "nation"
ncol.Add Now, "ts"

sm.Add ncol

End Sub

Public Function Find(sx As Integer, sy As Integer) As String

Dim col As Collection

'first look for it
For Each col In sm
    If col.item("x") = sx And col.item("y") = sy Then
        Find = col.item("msg")
        Exit Function
    End If
Next col

Find = vbNullString

End Function

Public Sub Clear()

While Me.Count > 0
    Me.Remove 1
Wend

End Sub

Public Function Count() As Integer
Count = sm.Count
End Function

Public Function X(Index As Integer) As Integer
If Index < 1 Or Index > sm.Count Then
    X = 0
    Exit Function
End If

X = sm.item(Index).item("x")
End Function
Public Function Y(Index As Integer) As Integer
If Index < 1 Or Index > sm.Count Then
    Y = 0
    Exit Function
End If
Y = sm.item(Index).item("y")
End Function
'Public Function Message(Index As Integer) As String    8/2003 efj  removed
'If Index < 1 Or Index > sm.Count Then
'    Message = vbnullstring
'    Exit Function
'End If
'Message = sm.Item(Index).Item("msg")
'End Function
Public Function Nation(Index As Integer) As Integer
If Index < 1 Or Index > sm.Count Then
    Nation = 0
    Exit Function
End If
Nation = sm.item(Index).item("nation")
End Function

Public Sub Remove(Index As Integer)
Dim col As Collection
Set col = sm.item(Index)
col.Remove "x"
col.Remove "y"
col.Remove "msg"
col.Remove "ts"
sm.Remove Index
End Sub
Public Sub RemoveExpired()

Dim col As Collection
Dim n As Integer
For n = sm.Count To 1 Step -1
    Set col = sm.item(n)
    If Now > DateAdd("n", CDbl(SECTOR_MARKER_DURATION), col.item("ts")) Then
        Me.Remove n
    End If
Next n

End Sub
