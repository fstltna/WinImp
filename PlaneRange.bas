Attribute VB_Name = "PlaneRange"
Option Explicit

'Change Log:
'010106 rjk: Created

Public PlaneRangeType As String
Public PlaneTargetSector As String
Public PlaneRoundTrip As Boolean
Public PlaneTech As Integer

Public Sub ShowPlaneRange(picMap As PictureBox)
Dim MaxRadius As Integer
Dim StartX As Integer
Dim StartY As Integer
Dim iDistance As Integer
Dim iDirection As Integer
Dim iSpoke As Integer
Dim sx As Integer
Dim sy As Integer
Dim X As Integer
Dim Y As Integer
Dim PosX As Single
Dim PosY As Single
Dim xa, ya, xas, yas As Variant
'Dim iPos As Integer
Dim hldIndex As String
Dim iRange As Integer
Dim HldColor As Long
Dim HldFontSize As Variant
Dim strDamage As String

hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"
rsBuild.Seek "=", "p", Trim$(Right$(PlaneRangeType, 5))

If rsBuild.NoMatch Then
    MsgBox "Plane Type information not found"
    Exit Sub
End If
MaxRadius = ComputePlaneRange(rsBuild.Fields("stat5"), rsBuild.Fields("tech"), PlaneTech)
If PlaneRoundTrip Then
    MaxRadius = MaxRadius / 2
End If
rsBuild.Index = hldIndex

xa = Array(-2, -1, 1, 2, 1, -1)
ya = Array(0, -1, -1, 0, 1, 1)
xas = Array(1, 2, 1, -1, -2, -1)
yas = Array(-1, 0, 1, 1, 0, -1)

If Not ParseSectors(StartX, StartY, PlaneTargetSector) Then
    Exit Sub
End If

HldColor = picMap.ForeColor
picMap.ForeColor = vbWhite
HldFontSize = picMap.Font.Size
picMap.Font.Size = picMap.Font.Size * 0.7

Coord StartX, StartY, PosX, PosY
strDamage = ""
sx = PosX - (picMap.TextHeight(strDamage) / 2)
sy = PosY - (picMap.TextWidth(strDamage) / 2)
If sx >= 0 And sy >= 0 Then
    picMap.CurrentX = sx
    picMap.CurrentY = sy
    picMap.Print strDamage
End If

'Now put out the nav markers
For iDistance = 1 To MaxRadius
    For iDirection = 0 To 5
        For iSpoke = 0 To iDistance - 1
            X = (iDistance * xa(iDirection)) + (iSpoke * xas(iDirection))
            Y = (iDistance * ya(iDirection)) + (iSpoke * yas(iDirection))
            Coord StartX + X, StartY + Y, PosX, PosY
            If PlaneRoundTrip Then
                strDamage = CStr(iDistance * 2)
            Else
                strDamage = CStr(iDistance)
            End If
            sx = PosX - (picMap.TextHeight(strDamage) / 2)
            sy = PosY - (picMap.TextWidth(strDamage) / 2)
            If sx >= 0 And sy >= 0 Then
                picMap.CurrentX = sx
                picMap.CurrentY = sy
                picMap.Print strDamage
            End If
        Next iSpoke
    Next iDirection
Next iDistance

picMap.ForeColor = HldColor
picMap.Font.Size = HldFontSize
End Sub
