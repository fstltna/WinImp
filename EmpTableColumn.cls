VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpTableColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'050206 rjk: Created

Private mstrName As String
Private miType As Integer
Private miFlags As Integer
Private miLength As Integer
Private miTable As Integer

Public Property Let Name(ByVal vData As String)
mstrName = vData
End Property

Public Property Get Name() As String
Name = mstrName
End Property

Public Property Let ColType(ByVal vData As Integer)
miType = vData
End Property

Public Property Get ColType() As Integer
ColType = miType
End Property

Public Property Let Flags(ByVal vData As Integer)
miFlags = vData
End Property

Public Property Get Flags() As Integer
Flags = miType
End Property

Public Property Let Length(ByVal vData As Integer)
miLength = vData
End Property

Public Property Get Length() As Integer
Length = miLength
End Property

Public Property Let Table(ByVal vData As Integer)
miTable = vData
End Property

Public Property Get Table() As Integer
Table = miType
End Property

