VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpNationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Change Log:
'05xx02 drk: This class reworked by Daniel Kiracofe <drk@unxsoft.com> to support storing the full 2D grid of country relations
'10xx02 drk: SLOW_WAR options
'081103 efj: Commented out dead function "TheirRelation"
'113003 rjk: Added telegram Count to object.
'130304 rjk: Modified NationNumber to partial matches on the name as long as it is unique.
'            Used in Air Combat History.
'130604 rjk: Added GetCountryNumbers for flash window.
'180206 rjk: Remove top and replace UBound().  Fix returning an empty country list.
'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun by changing the
'            relation values to match the server code.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

Private NationList() As String
Private NationRelations() As Integer ' 2d array.  first index specifies country relations are FROM, second specifies country relations are TOWARDS
Private natTelegramCount() As Integer ' 2d array  first index specifies country relations are FROM, second specifies country relations are TOWARDS

Public Sub AddNation(NationNumber As Integer, NationName As String)

If NationNumber < 0 Or NationNumber > 255 Then
    Exit Sub
End If

If UBound(NationList) < NationNumber Then
    ReDim Preserve NationList(NationNumber)
    'vb doesn't seem to like to redim both dimensions of 2d arrays.  take the easy way out and make it one dimension static.  a little bigger than necessary, but not too much bigger
    ReDim Preserve NationRelations(0 To 255, 0 To NationNumber)
    If UBound(natTelegramCount, 2) < NationNumber Then
        ReDim Preserve natTelegramCount(0 To 255, 0 To NationNumber)
    End If
End If

If Len(NationName) > 0 Then
    NationList(NationNumber) = NationName
End If

End Sub

Public Function NationName(NationNumber As Integer) As String

If UBound(NationList) < NationNumber Or LBound(NationList) > NationNumber Then
    NationName = vbNullString
Else
    NationName = NationList(NationNumber)
End If

End Function

Public Function NationNumber(NationName As String) As Integer

'Simple liner search - could improve, but should be small sample
Dim top As Integer
Dim X As Integer
Dim iNameLength As Integer

iNameLength = Len(NationName)

NationNumber = -1

top = UBound(NationList)
For X = LBound(NationList) To top
    If Left$(NationList(X), iNameLength) = NationName Then
        If Len(NationList(X)) = iNameLength Then 'Exact match
            NationNumber = X
            Exit Function
        Else
            If NationNumber = -1 Then 'Partial Match
                NationNumber = X
            Else
                NationNumber = -1 'More then one Partial Match
                Exit Function
            End If
        End If
    End If
Next X

'Not Found or Partial Match
End Function

Public Sub FillListBox(lb As Object)
Dim X As Integer

lb.Clear
For X = LBound(NationList) To UBound(NationList)
    If Len(NationList(X)) > 0 Then
        lb.AddItem NationList(X)
        lb.ItemData(lb.NewIndex) = X
    End If
Next X
End Sub

Public Function GetCountryNumbers() As Integer()
Dim X As Integer
Dim iCountryNumbers() As Integer
Dim iCount As Integer

For X = LBound(NationList) To UBound(NationList)
    If Len(NationList(X)) > 0 Then
        iCount = iCount + 1
        ReDim Preserve iCountryNumbers(iCount)
        iCountryNumbers(iCount) = X
    End If
Next X
GetCountryNumbers = iCountryNumbers
End Function

Private Sub Class_Initialize()
Me.Clear
ReDim natTelegramCount(0 To 255, 0 To 0)
    
'Default Values, will be replaced with xdump nation-relationships
iREL_AT_WAR = 0
iREL_SITZKRIEG = 1  'SLOW_WAR only
iREL_MOBILIZING = 2 'SLOW_WAR only
iREL_HOSTILE = 3
iREL_NEUTRAL = 4
iREL_FRIENDLY = 5
iREL_ALLIED = 6

End Sub

Public Sub Clear()
ReDim NationList(0 To 0)
ReDim NationRelations(0 To 255, 0 To 0)
End Sub

Public Sub ClearTelegramCount()
ReDim natTelegramCount(0 To 255, 0 To UBound(NationList))
End Sub

Public Function YourRelation(NationNumber As Integer) As Integer

If UBound(NationList) < NationNumber Or LBound(NationList) > NationNumber Then
    YourRelation = 0
Else
    YourRelation = NationRelations(CountryNumber, NationNumber)
End If

End Function

'    removed efj 8/2003
'Public Function TheirRelation(NationNumber As Integer) As Integer
'
'If UBound(NationList) < NationNumber Or LBound(NationList) > NationNumber Then
'    TheirRelation = 0
'Else
'    TheirRelation = NationRelations(NationNumber, CountryNumber)
'End If
'
'End Function
Public Property Get Relation(NationFrom As Integer, NationTowards As Integer) As Integer

If UBound(NationList) < NationFrom Or LBound(NationList) > NationFrom Or UBound(NationList) < NationTowards Or LBound(NationList) > NationTowards Then
    Relation = 0
Else
    Relation = NationRelations(NationFrom, NationTowards)
End If


End Property

Public Property Let Relation(NationFrom As Integer, NationTowards As Integer, r As Integer)

If UBound(NationList) < NationFrom Or LBound(NationList) > NationFrom Or UBound(NationList) < NationTowards Or LBound(NationList) > NationTowards Then
    Exit Property
Else
    NationRelations(NationFrom, NationTowards) = r
End If

End Property


Public Function RelationIndex(Relationship As String) As Integer

Select Case Trim$(Relationship)
    Case "Allied"
        RelationIndex = iREL_ALLIED
    Case "Friendly"
        RelationIndex = iREL_FRIENDLY
    Case "Neutral"
        RelationIndex = iREL_NEUTRAL
    Case "Hostile"
        RelationIndex = iREL_HOSTILE
    Case "At War"
        RelationIndex = iREL_AT_WAR
    Case "Mobilizing"
        RelationIndex = iREL_MOBILIZING
    Case "Sitzkrieg"
        RelationIndex = iREL_SITZKRIEG
    Case Else
        RelationIndex = 0
End Select
End Function

Public Function RelationString(r As Integer) As String
    Select Case r
    Case iREL_ALLIED
        RelationString = "Allied"
    Case iREL_FRIENDLY
        RelationString = "Friendly"
    Case iREL_NEUTRAL
        RelationString = "Neutral"
    Case iREL_HOSTILE
        RelationString = "Hostile"
    Case iREL_AT_WAR
        RelationString = "At War"
    Case iREL_MOBILIZING
        RelationString = "Mobilizing"
    Case iREL_SITZKRIEG
        RelationString = "Sitzkrieg"
    End Select
End Function

Public Function Count() As Integer
Count = UBound(NationList)
End Function

Public Function ActiveCount() As Integer
    Dim X As Integer
    ActiveCount = 0
    For X = 0 To UBound(NationList)
        If NationList(X) <> vbNullString Then
            ActiveCount = ActiveCount + 1
        End If
    Next
        
End Function

Public Property Get telegramCount(from As Integer, dest As Integer) As Integer
If dest <= UBound(natTelegramCount, 2) Then
    telegramCount = natTelegramCount(from, dest)
Else
    telegramCount = 0
End If
End Property

Public Property Let telegramCount(from As Integer, dest As Integer, Value As Integer)
If dest <= UBound(natTelegramCount, 2) Then
    natTelegramCount(from, dest) = Value
End If
End Property

