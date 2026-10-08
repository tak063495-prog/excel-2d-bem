Attribute VB_Name = "BEM_Operator"
Option Explicit
Public Sub ElementInfluence(ByVal sr As CSolveRegion, ByVal target As Long, ByVal element As Long, ByRef gb() As Double, ByRef hb() As Double)
    Dim geom As CElementData, rule As CIntegrationRule, prepared As CElementRule, own As Boolean
    Dim x(1 To 3) As Double, y(1 To 3) As Double, normal(1 To 3) As Double, s As Double, t As Double, distance As Double
    Dim uk(1 To 3, 1 To 3) As Double, tk(1 To 3, 1 To 3) As Double, du(1 To 3, 1 To 3, 1 To 3) As Double, dt(1 To 3, 1 To 3, 1 To 3) As Double
    Dim positions() As Double, directions() As Double, weights() As Double, shapes() As Double, xyz() As Double
    Dim q As Long, a As Long, b As Long, l As Long, d As Long, k As Long, ownLocal As Long, wg As Double, wh As Double
    d = sr.Ref.Dimension: k = sr.Ref.FieldCount: xyz = sr.X
    ReDim gb(1 To k, 1 To d, 1 To d): ReDim hb(1 To k, 1 To d, 1 To d)
    For a = 1 To d: x(a) = xyz(target, a): Next a
    Set geom = sr.Geometries(element + 1): own = (target - 1) \ k = element
    If own Then
        ownLocal = (target - 1) Mod k + 1: Set rule = SelfRule(d, sr.Ref.ElementOrder, ownLocal)
    Else
        ClosestPoint geom, x, s, t, distance
        If distance / geom.Length > 0.2 Then Set prepared = geom.Regular Else Set rule = NearRule(geom, s, t, distance, BoundaryDE)
    End If
    If prepared Is Nothing Then Set prepared = New CElementRule: prepared.Initialize geom, rule
    positions = prepared.XYZ: directions = prepared.Normals: weights = prepared.Weight: shapes = prepared.N
    For q = 1 To prepared.Count
        For a = 1 To d: y(a) = positions(q, a): normal(a) = directions(q, a): Next a
        Kelvin sr, x, y, normal, uk, tk, du, dt
        For l = 1 To k
            wg = weights(q) * shapes(q, l): wh = wg
            If own And l = ownLocal Then wh = wh - weights(q)
            For a = 1 To d: For b = 1 To d
                gb(l, a, b) = gb(l, a, b) + wg * uk(a, b): hb(l, a, b) = hb(l, a, b) + wh * tk(a, b)
            Next b: Next a
        Next l
    Next q
End Sub
Public Sub ApplyRegion(ByVal sr As CSolveRegion, ByRef traction() As Double, ByRef displacement() As Double, ByRef result() As Double)
    Dim g() As Double, h() As Double, i As Long, j As Long, value As Double
    If sr.Backend = "fmm" Then
        sr.Fmm.Apply traction, displacement, result
    Else
        g = sr.G: h = sr.H: ReDim result(1 To sr.ND)
        For i = 1 To sr.ND
            value = 0
            For j = 1 To sr.ND: value = value + g(i, j) * traction(j) + h(i, j) * displacement(j): Next j
            result(i) = value
        Next i
    End If
End Sub
Public Sub SetDenseDiagonal(ByVal sr As CSolveRegion)
    Dim g() As Double, h() As Double, dg() As Double, dh() As Double, i As Long, a As Long, b As Long, d As Long
    g = sr.G: h = sr.H: d = sr.Ref.Dimension
    ReDim dg(1 To sr.NF, 1 To d, 1 To d): ReDim dh(1 To sr.NF, 1 To d, 1 To d)
    For i = 1 To sr.NF: For a = 1 To d: For b = 1 To d
        dg(i, a, b) = g((i - 1) * d + a, (i - 1) * d + b): dh(i, a, b) = h((i - 1) * d + a, (i - 1) * d + b)
    Next b: Next a: Next i
    sr.DiagG = dg: sr.DiagH = dh
End Sub
