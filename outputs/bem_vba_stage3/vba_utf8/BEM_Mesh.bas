Attribute VB_Name = "BEM_Mesh"
Option Explicit
Public Const PI As Double = 3.1415926535897931

Public Sub CompensatedAdd(ByRef total As Double, ByRef compensation As Double, ByVal value As Double)
    Dim adjusted As Double, updated As Double
    adjusted = value - compensation: updated = total + adjusted
    compensation = (updated - total) - adjusted: total = updated
End Sub
Public Function Add3(ByVal a As Variant, ByVal b As Variant) As Variant
    Add3 = Array(a(0) + b(0), a(1) + b(1), a(2) + b(2))
End Function
Public Function Sub3(ByVal a As Variant, ByVal b As Variant) As Variant
    Sub3 = Array(a(0) - b(0), a(1) - b(1), a(2) - b(2))
End Function
Public Function Mul3(ByVal a As Variant, ByVal factor As Double) As Variant
    Mul3 = Array(a(0) * factor, a(1) * factor, a(2) * factor)
End Function
Public Function Dot3(ByVal a As Variant, ByVal b As Variant) As Double
    Dot3 = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function
Public Function Norm3(ByVal a As Variant) As Double
    Norm3 = Sqr(Dot3(a, a))
End Function
Public Function Cross3(ByVal a As Variant, ByVal b As Variant) As Variant
    Cross3 = Array(a(1) * b(2) - a(2) * b(1), a(2) * b(0) - a(0) * b(2), a(0) * b(1) - a(1) * b(0))
End Function
Private Function ArcPoint(ByVal r As CBemRegion, ByVal theta As Double) As Variant
    ArcPoint = Array(r.Origin(0) + r.Radius * Cos(theta), r.Origin(1) + r.Radius * Sin(theta), 0#)
End Function
Public Function SpherePoint(ByVal r As CBemRegion, ByVal theta As Double, ByVal phi As Double) As Variant
    SpherePoint = Add3(r.Origin, Array(r.Radius * Sin(phi) * Cos(theta), r.Radius * Sin(phi) * Sin(theta), r.Radius * Cos(phi)))
End Function
Private Function ProjectMid(ByVal r As CBemRegion, ByVal a As Variant, ByVal b As Variant) As Variant
    Dim p As Variant, v As Variant
    p = Mul3(Add3(a, b), 0.5)
    If r.ShapeName = "SPHERE" Then
        v = Sub3(p, r.Origin)
        If Norm3(v) < r.ModelScale * 0.00000000001 Then Fail "球の分割が粗すぎます。"
        p = Add3(r.Origin, Mul3(v, r.Radius / Norm3(v)))
    End If
    ProjectMid = p
End Function
Private Sub Segment(ByVal r As CBemRegion, ByVal a As Variant, ByVal b As Variant, ByVal boundary As String, Optional ByVal curvedMid As Variant)
    Dim ia As Long, ib As Long, im As Long, p As Variant
    ia = r.NodeId(a): ib = r.NodeId(b)
    If r.ElementOrder = 2 Then
        If IsMissing(curvedMid) Then p = Mul3(Add3(a, b), 0.5) Else p = curvedMid
        im = r.NodeId(p)
        r.PushElement Array(ia, im, ib), boundary
    Else
        r.PushElement Array(ia, ib), boundary
    End If
End Sub
Private Sub Triangle(ByVal r As CBemRegion, ByVal a As Variant, ByVal b As Variant, ByVal c As Variant, ByVal outward As Variant, ByVal boundary As String)
    Dim ids As Variant, tmp As Variant, n As Variant
    n = Cross3(Sub3(b, a), Sub3(c, a))
    If Dot3(n, outward) < 0 Then tmp = b: b = c: c = tmp
    If r.ElementOrder = 2 Then
        ids = Array(r.NodeId(a), r.NodeId(b), r.NodeId(c), r.NodeId(ProjectMid(r, a, b)), r.NodeId(ProjectMid(r, b, c)), r.NodeId(ProjectMid(r, c, a)))
    Else
        ids = Array(r.NodeId(a), r.NodeId(b), r.NodeId(c))
    End If
    r.PushElement ids, boundary
End Sub

Public Sub GenerateGeometry(ByVal r As CBemRegion)
    Dim nx As Long, ny As Long, nz As Long, i As Long, j As Long, k As Long
    Dim a As Variant, b As Variant, c As Variant, d As Variant, u As Variant, v As Variant
    Dim corners As Variant, labels As Variant, axis As Long, side As Long, axes As Variant
    Dim nu As Long, nv As Long, t As Double, ph As Double, n As Variant, boundary As String, swap As Variant, clockwise As Boolean
    nx = r.Divisions(0): ny = r.Divisions(1): nz = r.Divisions(2)
    Select Case r.ShapeName
    Case "RECT"
        corners = Array(Array(0#, 0#, 0#), Array(r.Sizes(0), 0#, 0#), Array(r.Sizes(0), r.Sizes(1), 0#), Array(0#, r.Sizes(1), 0#))
        labels = Array("BOTTOM", "RIGHT", "TOP", "LEFT")
        For side = 0 To 3
            nu = nx: If side Mod 2 = 1 Then nu = ny
            a = Add3(r.Origin, corners(side)): b = Add3(r.Origin, corners((side + 1) Mod 4))
            For i = 0 To nu - 1
                u = Add3(a, Mul3(Sub3(b, a), i / nu)): v = Add3(a, Mul3(Sub3(b, a), (i + 1) / nu))
                Segment r, u, v, CStr(labels(side))
            Next i
        Next side
    Case "CIRCLE"
        For i = 0 To nx - 1
            a = ArcPoint(r, 2 * PI * i / nx): b = ArcPoint(r, 2 * PI * ((i + 1) Mod nx) / nx)
            Segment r, a, b, "OUTER", ArcPoint(r, 2 * PI * (i + 0.5) / nx)
        Next i
    Case "POLYGON"
        corners = ReadPolygon(r)
        clockwise = PolygonArea(corners) < 0
        For side = 0 To UBound(corners)
            a = corners(side): b = corners((side + 1) Mod (UBound(corners) + 1))
            If clockwise Then swap = a: a = b: b = swap
            nu = Ceiling(Norm3(Sub3(b, a)) / MeshSize())
            If nx > 0 Then nu = nx
            For i = 0 To nu - 1
                Segment r, Add3(a, Mul3(Sub3(b, a), i / nu)), Add3(a, Mul3(Sub3(b, a), (i + 1) / nu)), "EDGE" & CStr(side + 1)
            Next i
        Next side
    Case "BOX"
        labels = Array("XMIN", "XMAX", "YMIN", "YMAX", "ZMIN", "ZMAX")
        For axis = 0 To 2
            Select Case axis
            Case 0: axes = Array(1, 2): nu = ny: nv = nz
            Case 1: axes = Array(0, 2): nu = nx: nv = nz
            Case 2: axes = Array(0, 1): nu = nx: nv = ny
            End Select
            For side = 0 To 1
                n = Array(0#, 0#, 0#): n(axis) = 2 * side - 1
                boundary = labels(2 * axis + side)
                For i = 0 To nu - 1
                    For j = 0 To nv - 1
                        a = FacePoint(r, axis, side, axes, i / nu, j / nv)
                        b = FacePoint(r, axis, side, axes, (i + 1) / nu, j / nv)
                        c = FacePoint(r, axis, side, axes, (i + 1) / nu, (j + 1) / nv)
                        d = FacePoint(r, axis, side, axes, i / nu, (j + 1) / nv)
                        Triangle r, a, b, c, n, boundary: Triangle r, a, c, d, n, boundary
                    Next j
                Next i
            Next side
        Next axis
    Case "SPHERE"
        For i = 0 To nx - 1
            t = 2 * PI * i / nx
            For j = 0 To ny - 1
                ph = PI * j / ny
                a = SpherePoint(r, t, ph): b = SpherePoint(r, 2 * PI * ((i + 1) Mod nx) / nx, ph)
                c = SpherePoint(r, 2 * PI * ((i + 1) Mod nx) / nx, PI * (j + 1) / ny)
                d = SpherePoint(r, t, PI * (j + 1) / ny)
                If j > 0 Then Triangle r, a, b, c, Sub3(Mul3(Add3(Add3(a, b), c), 1 / 3), r.Origin), "OUTER"
                If j < ny - 1 Then Triangle r, a, c, d, Sub3(Mul3(Add3(Add3(a, c), d), 1 / 3), r.Origin), "OUTER"
            Next j
        Next i
    End Select
    ValidateTopology r
End Sub

Private Function FacePoint(ByVal r As CBemRegion, ByVal axis As Long, ByVal side As Long, ByVal axes As Variant, ByVal u As Double, ByVal v As Double) As Variant
    Dim p As Variant: p = Array(0#, 0#, 0#)
    p(axis) = r.Sizes(axis) * side: p(axes(0)) = r.Sizes(axes(0)) * u: p(axes(1)) = r.Sizes(axes(1)) * v
    FacePoint = Add3(r.Origin, p)
End Function

Public Sub ValidateTopology(ByVal r As CBemRegion)
    Dim edges As Object, links As Object, seen As Object, starts As Object, adj As Object
    Dim e As Long, j As Long, a As Long, b As Long, mid As Long, p As Long, sign As Long
    Dim ids As Variant, q As Variant, uses As Variant, key As Variant, c As Variant, n As Variant, jac As Double
    Dim area As Double, areaScale As Double, compensation As Double, term As Double, volume As Double, queue As Collection, row As Collection, other As Variant
    Set edges = CreateObject("Scripting.Dictionary"): Set adj = CreateObject("Scripting.Dictionary")
    Set starts = CreateObject("Scripting.Dictionary")
    For e = 0 To r.Elements.Count - 1
        ids = r.Elements(e + 1)
        n = r.Differential(e, IIf(r.Dimension = 2, 0, 1 / 3), 1 / 3, jac)
        If r.Dimension = 2 Then
            a = ids(0): b = ids(UBound(ids))
            If starts.Exists(CStr(a)) Then Fail r.RegionName & ": 輪郭が分岐しています。"
            starts.Add CStr(a), b
            q = Sub3(r.Vertices(a + 1), r.Origin): c = Sub3(r.Vertices(b + 1), r.Origin)
            term = q(0) * c(1) - q(1) * c(0)
            CompensatedAdd area, compensation, term: areaScale = areaScale + Abs(term)
        Else
            Set row = New Collection: adj.Add CStr(e), row
            q = r.Vertices(ids(0) + 1): c = r.Vertices(ids(1) + 1)
            volume = volume + Dot3(Sub3(q, r.Origin), Cross3(Sub3(c, r.Origin), Sub3(r.Vertices(ids(2) + 1), r.Origin))) / 6
            For j = 0 To 2
                a = ids(j): b = ids((j + 1) Mod 3): mid = -1
                If r.ElementOrder = 2 Then mid = ids(3 + j)
                sign = 1: If a > b Then sign = -1
                key = CStr(MinLong(a, b)) & ":" & CStr(MaxLong(a, b))
                If edges.Exists(key) Then
                    uses = edges(key)
                    If uses(3) <> 1 Or uses(1) + sign <> 0 Or uses(2) <> mid Then Fail r.RegionName & ": 表面の共有辺が不整合です。"
                    adj(CStr(e)).Add uses(0): adj(CStr(uses(0))).Add e
                    edges(key) = Array(uses(0), uses(1), uses(2), 2)
                Else
                    edges.Add key, Array(e, sign, mid, 1)
                End If
            Next j
            'Check normals/Jacobians at corners and edge midpoints as well as centroid.
            For Each q In Array(Array(0#, 0#), Array(1#, 0#), Array(0#, 1#), Array(0.5, 0#), Array(0.5, 0.5), Array(0#, 0.5))
                c = r.Differential(e, CDbl(q(0)), CDbl(q(1)), jac)
                If Dot3(c, n) <= 0 Then Fail r.RegionName & ": 二次三角形が折れています。"
            Next q
        End If
    Next e
    Set seen = CreateObject("Scripting.Dictionary")
    If r.Dimension = 2 Then
        ids = r.Elements(1): p = ids(0)
        For e = 1 To r.Elements.Count
            If seen.Exists(CStr(p)) Or Not starts.Exists(CStr(p)) Then Fail r.RegionName & ": 単一の閉輪郭ではありません。"
            seen.Add CStr(p), True: p = starts(CStr(p))
        Next e
        If p <> ids(0) Or area <= 64 * 2.220446049250313E-16 * areaScale Then Fail r.RegionName & ": 反時計回りの非退化な閉輪郭が必要です。"
    Else
        For Each key In edges.Keys
            uses = edges(key): If uses(3) <> 2 Then Fail r.RegionName & ": 表面が閉じていません。"
        Next key
        Set queue = New Collection: queue.Add 0: seen.Add "0", True
        Do While queue.Count > 0
            p = queue(1): queue.Remove 1
            For Each other In adj(CStr(p))
                If Not seen.Exists(CStr(other)) Then seen.Add CStr(other), True: queue.Add other
            Next other
        Loop
        If seen.Count <> r.Elements.Count Or volume <= 0 Then Fail r.RegionName & ": 単一の外向き閉表面が必要です。"
    End If
End Sub

Public Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinLong = a Else MinLong = b
End Function
Public Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function
Public Function ElementMeasure(ByVal r As CBemRegion, ByVal e As Long) As Double
    Dim x As Variant, w As Variant, i As Long, jac As Double, n As Variant
    Dim a As Variant, b As Variant, c As Variant, ids As Variant
    If r.Dimension = 2 Then
        x = Array(-0.906179845938664, -0.538469310105683, 0#, 0.538469310105683, 0.906179845938664)
        w = Array(0.236926885056189, 0.478628670499366, 0.568888888888889, 0.478628670499366, 0.236926885056189)
        For i = 0 To 4
            n = r.Differential(e, x(i), 0, jac): ElementMeasure = ElementMeasure + w(i) * jac
        Next i
    Else
        '7-point degree-5 Dunavant triangle rule; weights sum to 1/2.
        x = Array(Array(1 / 3, 1 / 3), Array(0.470142064105115, 0.470142064105115), Array(0.05971587178977, 0.470142064105115), Array(0.470142064105115, 0.05971587178977), Array(0.101286507323456, 0.101286507323456), Array(0.797426985353087, 0.101286507323456), Array(0.101286507323456, 0.797426985353087))
        w = Array(0.1125, 0.066197076394253, 0.066197076394253, 0.066197076394253, 0.062969590272414, 0.062969590272414, 0.062969590272414)
        For i = 0 To 6
            n = r.Differential(e, x(i)(0), x(i)(1), jac): ElementMeasure = ElementMeasure + w(i) * jac
        Next i
    End If
End Function
Public Function MaxEdge(ByVal r As CBemRegion, ByVal e As Long) As Double
    Dim ids As Variant, i As Long, j As Long, length As Double
    ids = r.Elements(e + 1)
    If r.Dimension = 2 Then
        MaxEdge = Norm3(Sub3(r.Vertices(ids(0) + 1), r.Vertices(ids(UBound(ids)) + 1)))
    Else
        For i = 0 To 2
            j = (i + 1) Mod 3: length = Norm3(Sub3(r.Vertices(ids(i) + 1), r.Vertices(ids(j) + 1)))
            If length > MaxEdge Then MaxEdge = length
        Next i
    End If
End Function

Private Function Cross2(ByVal a As Variant, ByVal b As Variant) As Double
    Cross2 = a(0) * b(1) - a(1) * b(0)
End Function
Public Function InsideRoundMesh(ByVal r As CBemRegion, ByVal point As Variant) As Boolean
    'Generated circle/sphere meshes are star-shaped about Origin. Find the ray's
    'triangle/line cone, then intersect the actual quadratic geometry by Newton.
    Dim delta As Variant, direction As Variant, distance As Double, ids As Variant, e As Long
    Dim a As Variant, b As Variant, c As Variant, u As Variant, v As Variant, w As Variant, n As Variant
    Dim da As Variant, db As Variant, residual As Variant, minusDir As Variant
    Dim t As Double, s As Double, aa As Double, bb As Double, det As Double, denom As Double
    Dim d00 As Double, d01 As Double, d11 As Double, d20 As Double, d21 As Double, stepA As Double, stepB As Double, stepT As Double
    Dim iteration As Long, tol As Double, coordTol As Double, candidate As Boolean
    delta = Sub3(point, r.Origin): distance = Norm3(delta): tol = r.ModelScale * 0.000000001
    If distance <= tol Then InsideRoundMesh = True: Exit Function
    If distance >= r.Radius Then Exit Function
    direction = Mul3(delta, 1 / distance): minusDir = Mul3(direction, -1): coordTol = 0.000000001
    For e = 0 To r.Elements.Count - 1
        ids = r.Elements(e + 1): a = Sub3(r.Vertices(ids(0) + 1), r.Origin): candidate = False
        If r.Dimension = 2 Then
            b = Sub3(r.Vertices(ids(UBound(ids)) + 1), r.Origin): u = Sub3(b, a)
            denom = Cross2(direction, u)
            If Abs(denom) > r.ModelScale * 0.000000000001 Then
                t = Cross2(a, u) / denom: aa = Cross2(a, direction) / denom
                candidate = t > 0 And aa >= -coordTol And aa <= 1 + coordTol
                s = 2 * aa - 1
            End If
        Else
            b = Sub3(r.Vertices(ids(1) + 1), r.Origin): c = Sub3(r.Vertices(ids(2) + 1), r.Origin)
            u = Sub3(b, a): v = Sub3(c, a): n = Cross3(u, v): denom = Dot3(n, direction)
            If denom > r.ModelScale ^ 2 * 0.00000000000001 Then
                t = Dot3(n, a) / denom: w = Sub3(Mul3(direction, t), a)
                d00 = Dot3(u, u): d01 = Dot3(u, v): d11 = Dot3(v, v): d20 = Dot3(w, u): d21 = Dot3(w, v)
                det = d00 * d11 - d01 * d01
                aa = (d11 * d20 - d01 * d21) / det: bb = (d00 * d21 - d01 * d20) / det
                candidate = t > 0 And aa >= -coordTol And bb >= -coordTol And aa + bb <= 1 + coordTol
            End If
        End If
        If candidate Then
            If r.ElementOrder = 2 Then
                For iteration = 1 To 12
                    If r.Dimension = 2 Then
                        residual = Sub3(r.LocalPosition(e, s), Mul3(direction, t))
                        da = Mul3(Sub3(r.LocalPosition(e, s + 0.001), r.LocalPosition(e, s - 0.001)), 500)
                        det = Cross2(da, minusDir)
                        stepA = Cross2(residual, minusDir) / det: stepT = Cross2(da, residual) / det
                        s = s - stepA: t = t - stepT
                    Else
                        residual = Sub3(r.LocalPosition(e, aa, bb), Mul3(direction, t))
                        da = Mul3(Sub3(r.LocalPosition(e, aa + 0.001, bb), r.LocalPosition(e, aa - 0.001, bb)), 500)
                        db = Mul3(Sub3(r.LocalPosition(e, aa, bb + 0.001), r.LocalPosition(e, aa, bb - 0.001)), 500)
                        det = Dot3(da, Cross3(db, minusDir))
                        stepA = Dot3(residual, Cross3(db, minusDir)) / det
                        stepB = Dot3(da, Cross3(residual, minusDir)) / det
                        stepT = Dot3(da, Cross3(db, residual)) / det
                        aa = aa - stepA: bb = bb - stepB: t = t - stepT
                    End If
                    If Norm3(residual) < r.ModelScale * 0.00000000001 Then Exit For
                Next iteration
                If iteration > 12 Then Fail r.RegionName & ": 評価点の内部判定が収束しません。境界から離すかメッシュを細かくしてください。"
            End If
            InsideRoundMesh = distance < t - tol
            Exit Function
        End If
    Next e
End Function
