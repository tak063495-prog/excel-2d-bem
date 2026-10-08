Attribute VB_Name = "BEM_Closest"
Option Explicit
Private Function Clamp(ByVal value As Double, ByVal lower As Double, ByVal upper As Double) As Double
    Clamp = MaxDouble(lower, MinDouble(upper, value))
End Function
Private Function CubicValue(ByVal x As Double, ByVal c3 As Double, ByVal c2 As Double, ByVal c1 As Double, ByVal c0 As Double) As Double
    CubicValue = ((c3 * x + c2) * x + c1) * x + c0
End Function
Private Sub LineCandidate(ByRef av() As Double, ByRef bv() As Double, ByRef zv() As Double, ByVal s As Double, ByRef best As Double, ByRef bestS As Double)
    Dim i As Long, value As Double, distance2 As Double
    For i = 1 To 3: value = av(i) * s * s + bv(i) * s + zv(i): distance2 = distance2 + value * value: Next i
    If distance2 < best Then best = distance2: bestS = s
End Sub
Public Sub ClosestLine(ByRef startP() As Double, ByRef midP() As Double, ByRef endP() As Double, ByRef x() As Double, ByVal quadratic As Boolean, ByRef s As Double, ByRef distance As Double)
    Dim av(1 To 3) As Double, bv(1 To 3) As Double, zv(1 To 3) As Double
    Dim c3 As Double, c2 As Double, c1 As Double, c0 As Double, i As Long, j As Long, best As Double
    Dim cuts As Collection, disc As Double, r1 As Double, r2 As Double, left As Double, right As Double, fleft As Double, fright As Double, middle As Double, fmid As Double
    Set cuts = New Collection: cuts.Add -1#: best = 1E+250
    For i = 1 To 3
        bv(i) = (endP(i) - startP(i)) / 2
        If quadratic Then av(i) = (startP(i) + endP(i)) / 2 - midP(i): zv(i) = midP(i) - x(i) Else zv(i) = (startP(i) + endP(i)) / 2 - x(i)
        c3 = c3 + 2 * av(i) * av(i): c2 = c2 + 3 * av(i) * bv(i)
        c1 = c1 + bv(i) * bv(i) + 2 * av(i) * zv(i): c0 = c0 + bv(i) * zv(i)
    Next i
    LineCandidate av, bv, zv, -1, best, s: LineCandidate av, bv, zv, 1, best, s
    If c3 <= 1E-28 * MaxDouble(Abs(c1), 1E-250) Then
        If c1 > 0 Then LineCandidate av, bv, zv, Clamp(-c0 / c1, -1, 1), best, s
    Else
        disc = 4 * c2 * c2 - 12 * c3 * c1
        If disc > 0 Then
            r1 = (-2 * c2 - Sqr(disc)) / (6 * c3): r2 = (-2 * c2 + Sqr(disc)) / (6 * c3)
            If r1 > -1 And r1 < 1 Then cuts.Add r1
            If r2 > -1 And r2 < 1 Then cuts.Add r2
        End If
        cuts.Add 1#
        For i = 1 To cuts.Count - 1
            left = cuts(i): right = cuts(i + 1): fleft = CubicValue(left, c3, c2, c1, c0): fright = CubicValue(right, c3, c2, c1, c0)
            LineCandidate av, bv, zv, left, best, s: LineCandidate av, bv, zv, right, best, s
            If fleft * fright < 0 Then
                For j = 1 To 60
                    middle = (left + right) / 2: fmid = CubicValue(middle, c3, c2, c1, c0)
                    If fleft * fmid <= 0 Then right = middle Else left = middle: fleft = fmid
                Next j
                LineCandidate av, bv, zv, (left + right) / 2, best, s
            End If
        Next i
    End If
    distance = Sqr(best)
End Sub
Private Sub ProjectTriangle(ByRef a As Double, ByRef b As Double)
    a = MaxDouble(a, 0): b = MaxDouble(b, 0)
    If a + b > 1 Then a = Clamp((a - b + 1) / 2, 0, 1): b = 1 - a
End Sub
Private Function DistanceAt(ByVal geom As CElementData, ByRef x() As Double, ByVal a As Double, ByVal b As Double) As Double
    Dim p(1 To 3) As Double, normal(1 To 3) As Double, ds(1 To 3) As Double, dt(1 To 3) As Double, jac As Double, i As Long
    MapGeometry geom, a, b, p, normal, jac, ds, dt
    For i = 1 To geom.D: DistanceAt = DistanceAt + (p(i) - x(i)) ^ 2: Next i
End Function
Public Sub ClosestPoint(ByVal geom As CElementData, ByRef x() As Double, ByRef bestA As Double, ByRef bestB As Double, ByRef distance As Double)
    Dim startP(1 To 3) As Double, midP(1 To 3) As Double, endP(1 To 3) As Double
    Dim ds(1 To 3) As Double, dt(1 To 3) As Double, p(1 To 3) As Double, normal(1 To 3) As Double
    Dim i As Long, edge As Long, side As Long, j As Long, seed As Long, iteration As Long, stepIndex As Long
    Dim g00 As Double, g01 As Double, g11 As Double, z0 As Double, z1 As Double, det As Double, a As Double, b As Double, jac As Double
    Dim s As Double, dist As Double, best As Double, trial As Double, stepA As Double, stepB As Double, newA As Double, newB As Double, factor As Double
    For i = 1 To geom.D: startP(i) = geom.Coords(1, i): endP(i) = geom.Coords(geom.NG, i): midP(i) = geom.Coords(2, i): Next i
    If geom.D = 2 Then
        ClosestLine startP, midP, endP, x, geom.ElementOrder = 2, bestA, distance: bestB = 0: Exit Sub
    End If
    For i = 1 To 3
        ds(i) = geom.Coords(2, i) - geom.Coords(1, i): dt(i) = geom.Coords(3, i) - geom.Coords(1, i)
        g00 = g00 + ds(i) * ds(i): g01 = g01 + ds(i) * dt(i): g11 = g11 + dt(i) * dt(i)
        z0 = z0 + ds(i) * (x(i) - geom.Coords(1, i)): z1 = z1 + dt(i) * (x(i) - geom.Coords(1, i))
    Next i
    det = g00 * g11 - g01 * g01: a = (g11 * z0 - g01 * z1) / det: b = (g00 * z1 - g01 * z0) / det
    best = 1E+250
    If a >= 0 And b >= 0 And a + b <= 1 Then best = DistanceAt(geom, x, a, b): bestA = a: bestB = b
    For side = 1 To 3
        edge = side Mod 3 + 1
        For i = 1 To 3
            startP(i) = geom.Coords(side, i): endP(i) = geom.Coords(edge, i)
            If geom.ElementOrder = 2 Then midP(i) = geom.Coords(side + 3, i)
        Next i
        ClosestLine startP, midP, endP, x, geom.ElementOrder = 2, s, dist
        If dist * dist < best Then
            best = dist * dist
            Select Case side
            Case 1: bestA = (s + 1) / 2: bestB = 0
            Case 2: bestA = (1 - s) / 2: bestB = (s + 1) / 2
            Case 3: bestA = 0: bestB = (1 - s) / 2
            End Select
        End If
    Next side
    If Not geom.Planar Then
        For seed = 0 To 1
            If seed = 0 Then a = (g11 * z0 - g01 * z1) / det: b = (g00 * z1 - g01 * z0) / det Else a = 1 / 3: b = 1 / 3
            ProjectTriangle a, b
            For iteration = 1 To 60
                MapGeometry geom, a, b, p, normal, jac, ds, dt
                g00 = 0: g01 = 0: g11 = 0: z0 = 0: z1 = 0: dist = 0
                For i = 1 To 3
                    g00 = g00 + ds(i) ^ 2: g01 = g01 + ds(i) * dt(i): g11 = g11 + dt(i) ^ 2
                    z0 = z0 + ds(i) * (p(i) - x(i)): z1 = z1 + dt(i) * (p(i) - x(i)): dist = dist + (p(i) - x(i)) ^ 2
                Next i
                det = g00 * g11 - g01 * g01
                stepA = (g11 * z0 - g01 * z1) / det: stepB = (g00 * z1 - g01 * z0) / det
                If Abs(stepA) + Abs(stepB) < 1E-12 Then Exit For
                factor = 1
                For stepIndex = 1 To 20
                    newA = a - factor * stepA: newB = b - factor * stepB: ProjectTriangle newA, newB
                    trial = DistanceAt(geom, x, newA, newB)
                    If trial < dist Then Exit For
                    factor = factor / 2
                Next stepIndex
                If trial >= dist Then Exit For
                a = newA: b = newB
            Next iteration
            trial = DistanceAt(geom, x, a, b)
            If trial < best Then best = trial: bestA = a: bestB = b
        Next seed
    End If
    distance = Sqr(best)
End Sub
