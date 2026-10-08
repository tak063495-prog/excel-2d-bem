VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CElementRule"
Option Explicit
Public Count As Long
Public XYZ As Variant
Public Normals As Variant
Public Weight As Variant
Public N As Variant
Public Sub Initialize(ByVal geom As CElementData, ByVal rule As CIntegrationRule)
    Dim i As Long, j As Long, l As Long, point(1 To 3) As Double, normal(1 To 3) As Double
    Dim ds(1 To 3) As Double, dt(1 To 3) As Double, jac As Double, field(1 To 6) As Double
    Dim s() As Double, t() As Double, w() As Double
    Count = rule.Count
    s = rule.A: t = rule.B: w = rule.W
    ReDim XYZ(1 To Count, 1 To geom.D) As Double: ReDim Normals(1 To Count, 1 To geom.D) As Double
    ReDim Weight(1 To Count) As Double: ReDim N(1 To Count, 1 To geom.K) As Double
    For i = 1 To Count
        MapGeometry geom, s(i), t(i), point, normal, jac, ds, dt
        FieldShape geom.D, geom.ElementOrder, s(i), t(i), field
        Weight(i) = w(i) * jac
        For j = 1 To geom.D: XYZ(i, j) = point(j): Normals(i, j) = normal(j): Next j
        For l = 1 To geom.K: N(i, l) = field(l): Next l
    Next i
End Sub
