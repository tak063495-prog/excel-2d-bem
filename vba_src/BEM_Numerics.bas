Attribute VB_Name = "BEM_Numerics"
Option Explicit
Public BoundaryOrder As Long, BoundaryDE As Long, SelfOrder As Long
Public MaxUnknowns As Long, MaxGaussLevel As Long, MaxDELevel As Long, MaxPointBudget As Long
Public DisplacementTolerance As Double, StressTolerance As Double, LinearTolerance As Double
Public QuadratureMode As String, InteriorNearRatio As Double, InteriorMinDistanceRatio As Double
Private GaussCache As Object, DECache As Object, RegularCache As Object, SelfCache As Object
Private TriInverse(1 To 6, 1 To 6) As Double
Private Initialized As Boolean

Public Sub InitNumerics()
    Dim rows As Variant, i As Long, j As Long
    If Initialized Then Exit Sub
    Set GaussCache = CreateObject("Scripting.Dictionary"): Set DECache = CreateObject("Scripting.Dictionary")
    Set RegularCache = CreateObject("Scripting.Dictionary"): Set SelfCache = CreateObject("Scripting.Dictionary")
    rows = Array( _
        Array(2.33464825772518, 0.243918474687705, 0.243918474687705, -0.975673898750821, 0.128862590401052, -0.975673898750821), _
        Array(-6.82445759368836, -2.64299802761341, 0#, 9.46745562130177, -1.10453648915187, 1.10453648915187), _
        Array(-6.82445759368836, 0#, -2.64299802761341, 1.10453648915187, -1.10453648915187, 9.46745562130177), _
        Array(4.733727810650888, 4.733727810650887, 0#, -9.46745562130177, 0#, 0#), _
        Array(9.467455621301775, 0#, 0#, -9.467455621301774, 9.467455621301774, -9.467455621301775), _
        Array(4.733727810650889, 0#, 4.733727810650887, 0#, 0#, -9.467455621301775))
    For i = 1 To 6: For j = 1 To 6: TriInverse(i, j) = rows(i - 1)(j - 1): Next j: Next i
    Initialized = True
End Sub
Public Function SolveSetting(ByVal row As Long, ByVal fallback As Double) As Double
    If TextAt(WS("操作"), row, 2) = "" Then SolveSetting = fallback Else SolveSetting = NumberAt(WS("操作"), row, 2)
End Function
Public Function IntegerSetting(ByVal row As Long, ByVal fallback As Long, ByVal lower As Long, ByVal upper As Long) As Long
    Dim value As Double: value = SolveSetting(row, fallback)
    If value <> Fix(value) Or value < lower Or value > upper Then Fail "操作!B" & row & ": " & lower & "～" & upper & "の整数を指定してください。"
    IntegerSetting = CLng(value)
End Function
Public Sub ReadSolveOptions()
    InitNumerics
    MaxUnknowns = IntegerSetting(34, 1200, 1, 4000)
    BoundaryOrder = IntegerSetting(35, 0, 0, 128)
    If BoundaryOrder = 0 Then BoundaryOrder = IIf(CurrentDimension() = 2, 48, 12)
    If BoundaryOrder < 4 Then Fail "境界Gauss次数は0（自動）または4～128です。"
    SelfOrder = IntegerSetting(36, 24, 8, 128)
    BoundaryDE = IntegerSetting(37, -1, -1, 4)
    If BoundaryDE = -1 Then BoundaryDE = IIf(CurrentDimension() = 2, 3, 1)
    DisplacementTolerance = SolveSetting(38, 1E-8): StressTolerance = SolveSetting(39, 0.001)
    If DisplacementTolerance <= 0 Or StressTolerance <= 0 Then Fail "内部場の積分許容差は正の数が必要です。"
    MaxDELevel = IntegerSetting(40, 6, 2, 8): MaxGaussLevel = IntegerSetting(41, 6, 2, 8)
    MaxPointBudget = IntegerSetting(42, 200000, 1, 2000000)
    LinearTolerance = SolveSetting(43, 1E-9)
    If LinearTolerance <= 0 Or LinearTolerance >= 1 Then Fail "線形残差許容差は0より大きく1より小さい数です。"
    QuadratureMode = LCase$(TextAt(WS("操作"), 59, 2)): If QuadratureMode = "" Then QuadratureMode = "auto"
    If QuadratureMode <> "auto" And QuadratureMode <> "gauss" And QuadratureMode <> "de" Then Fail "内点積分方式はauto、gauss、deです。"
    InteriorNearRatio = SolveSetting(60, 0.2): InteriorMinDistanceRatio = SolveSetting(61, 1E-10)
    If InteriorNearRatio <= 0 Or InteriorMinDistanceRatio <= 0 Then Fail "内点近接判定・最小距離比は正の数です。"
End Sub

Public Sub KelvinTensor(ByVal sr As CSolveRegion, ByRef x() As Double, ByRef y() As Double, ByRef uk() As Double, ByRef tensor() As Double)
    Dim d As Long, i As Long, j As Long, k As Long, h(1 To 3) As Double, rho As Double, nu As Double, cu As Double, ct As Double, identity As Double
    d = sr.Ref.Dimension: nu = sr.NuEff
    For i = 1 To d: h(i) = y(i) - x(i): rho = rho + h(i) * h(i): Next i
    rho = Sqr(rho): If rho <= 0 Then Fail "FMMの展開点が一致しました。遠方判定を確認してください。"
    For i = 1 To d: h(i) = h(i) / rho: Next i
    cu = 1 / (IIf(d = 2, 8, 16) * PI * sr.Mu * (1 - nu)): ct = -1 / (IIf(d = 2, 4, 8) * PI * (1 - nu) * rho ^ (d - 1))
    For i = 1 To d: For j = 1 To d
        identity = IIf(i = j, 1#, 0#)
        If d = 2 Then uk(i, j) = cu * (-(3 - 4 * nu) * identity * Log(rho / sr.MeshScale) + h(i) * h(j)) Else uk(i, j) = cu * ((3 - 4 * nu) * identity + h(i) * h(j)) / rho
        For k = 1 To d
            tensor(i, j, k) = ct * ((1 - 2 * nu) * (identity * h(k) - h(i) * IIf(j = k, 1#, 0#) + h(j) * IIf(i = k, 1#, 0#)) + d * h(i) * h(j) * h(k))
        Next k
    Next j: Next i
End Sub
Public Sub FieldParameters(ByVal d As Long, ByVal order As Long, ByVal fieldIndex As Long, ByRef a As Double, ByRef b As Double)
    Dim nodes As Variant
    a = 0: b = 0
    If d = 2 Then
        If order = 2 Then a = (fieldIndex - 2) * 2 / 3
    Else
        a = 1 / 3: b = 1 / 3
        If order = 2 Then
            nodes = Array(Array(0#, 0#), Array(1#, 0#), Array(0#, 1#), Array(0.5, 0#), Array(0.5, 0.5), Array(0#, 0.5))
            a = 1 / 3 + 0.65 * (nodes(fieldIndex - 1)(0) - 1 / 3)
            b = 1 / 3 + 0.65 * (nodes(fieldIndex - 1)(1) - 1 / 3)
        End If
    End If
End Sub
Public Sub FieldShape(ByVal d As Long, ByVal order As Long, ByVal a As Double, ByVal b As Double, ByRef n() As Double)
    Dim mon(1 To 6) As Double, i As Long, j As Long
    If order = 0 Then n(1) = 1: Exit Sub
    If d = 2 Then
        n(1) = 1.125 * a * a - 0.75 * a: n(2) = 1 - 2.25 * a * a: n(3) = 1.125 * a * a + 0.75 * a
    Else
        mon(1) = 1: mon(2) = a: mon(3) = b: mon(4) = a * a: mon(5) = a * b: mon(6) = b * b
        For j = 1 To 6
            n(j) = 0
            For i = 1 To 6: n(j) = n(j) + mon(i) * TriInverse(i, j): Next i
        Next j
    End If
End Sub

Public Sub MapGeometry(ByVal g As CElementData, ByVal a As Double, ByVal b As Double, ByRef point() As Double, ByRef normal() As Double, ByRef jac As Double, ByRef ds() As Double, ByRef dt() As Double)
    Dim n(1 To 6) As Double, da(1 To 6) As Double, db(1 To 6) As Double, c As Double, i As Long, j As Long
    Dim coords() As Double: coords = g.Coords
    If g.D = 2 Then
        If g.ElementOrder = 0 Then
            n(1) = (1 - a) / 2: n(2) = (1 + a) / 2: da(1) = -0.5: da(2) = 0.5
        Else
            n(1) = a * (a - 1) / 2: n(2) = 1 - a * a: n(3) = a * (a + 1) / 2
            da(1) = a - 0.5: da(2) = -2 * a: da(3) = a + 0.5
        End If
    Else
        c = 1 - a - b
        If g.ElementOrder = 0 Then
            n(1) = c: n(2) = a: n(3) = b: da(1) = -1: da(2) = 1: db(1) = -1: db(3) = 1
        Else
            n(1) = c * (2 * c - 1): n(2) = a * (2 * a - 1): n(3) = b * (2 * b - 1)
            n(4) = 4 * c * a: n(5) = 4 * a * b: n(6) = 4 * b * c
            da(1) = 1 - 4 * c: da(2) = 4 * a - 1: da(4) = 4 * (c - a): da(5) = 4 * b: da(6) = -4 * b
            db(1) = 1 - 4 * c: db(3) = 4 * b - 1: db(4) = -4 * a: db(5) = 4 * a: db(6) = 4 * (c - b)
        End If
    End If
    For j = 1 To 3
        point(j) = coords(1, j): ds(j) = 0: dt(j) = 0
        For i = 1 To g.NG
            point(j) = point(j) + n(i) * (coords(i, j) - coords(1, j))
            ds(j) = ds(j) + da(i) * (coords(i, j) - coords(1, j)): dt(j) = dt(j) + db(i) * (coords(i, j) - coords(1, j))
        Next i
    Next j
    If g.D = 2 Then
        normal(1) = ds(2): normal(2) = -ds(1): normal(3) = 0
    Else
        normal(1) = ds(2) * dt(3) - ds(3) * dt(2)
        normal(2) = ds(3) * dt(1) - ds(1) * dt(3)
        normal(3) = ds(1) * dt(2) - ds(2) * dt(1)
    End If
    jac = Sqr(normal(1) ^ 2 + normal(2) ^ 2 + normal(3) ^ 2)
    If jac <= g.ModelScale ^ (g.D - 1) * 1E-14 Then Fail "積分点で幾何Jacobianが退化しています。"
    For j = 1 To 3: normal(j) = normal(j) / jac: Next j
End Sub

Public Function Gauss01Rule(ByVal order As Long) As CIntegrationRule
    Dim r As CIntegrationRule, i As Long, j As Long, iteration As Long, z As Double, prior As Double, p0 As Double, p1 As Double, p2 As Double, deriv As Double, weight As Double
    InitNumerics
    If GaussCache.Exists(CStr(order)) Then Set Gauss01Rule = GaussCache(CStr(order)): Exit Function
    Set r = New CIntegrationRule: r.Allocate order
    For i = 1 To (order + 1) \ 2
        z = Cos(PI * (i - 0.25) / (order + 0.5))
        For iteration = 1 To 50
            p0 = 1: p1 = z
            For j = 2 To order: p2 = ((2 * j - 1) * z * p1 - (j - 1) * p0) / j: p0 = p1: p1 = p2: Next j
            deriv = order * (z * p1 - p0) / (z * z - 1): prior = z: z = z - p1 / deriv
            If Abs(z - prior) < 2E-16 Then Exit For
        Next iteration
        weight = 1 / ((1 - z * z) * deriv * deriv)
        r.StorePoint i, (1 - z) / 2, 0, weight
        r.StorePoint order + 1 - i, (1 + z) / 2, 0, weight
    Next i
    r.Count = order: GaussCache.Add CStr(order), r: Set Gauss01Rule = r
End Function
Private Function SinhSafe(ByVal value As Double) As Double
    If Abs(value) < 0.0001 Then SinhSafe = value * (1 + value * value / 6) Else SinhSafe = (Exp(value) - Exp(-value)) / 2
End Function
Private Function CoshSafe(ByVal value As Double) As Double
    CoshSafe = (Exp(value) + Exp(-value)) / 2
End Function
Public Function AsinhSafe(ByVal value As Double) As Double
    If Abs(value) < 0.0001 Then
        AsinhSafe = value * (1 - value * value / 6)
    ElseIf value >= 0 Then
        AsinhSafe = Log(value + Sqr(value * value + 1))
    Else
        AsinhSafe = -Log(-value + Sqr(value * value + 1))
    End If
End Function
Private Function Expm1Safe(ByVal value As Double) As Double
    If Abs(value) < 0.0001 Then Expm1Safe = value * (1 + value * (0.5 + value * (1 / 6 + value / 24))) Else Expm1Safe = Exp(value) - 1
End Function
Private Function Log1pSafe(ByVal value As Double) As Double
    If Abs(value) < 0.0001 Then Log1pSafe = value * (1 - value * (0.5 - value * (1 / 3 - value / 4))) Else Log1pSafe = Log(1 + value)
End Function
Public Function DE01Rule(ByVal level As Long) As CIntegrationRule
    Dim r As CIntegrationRule, stepSize As Double, value As Double, z As Double, a As Double, small As Double, b As Double, w As Double, i As Long, limit As Long
    InitNumerics
    If DECache.Exists(CStr(level)) Then Set DE01Rule = DECache(CStr(level)): Exit Function
    stepSize = 0.5 / 2 ^ level: limit = CLng(4 / stepSize)
    Set r = New CIntegrationRule: r.Allocate 2 * limit + 1
    For i = -limit To limit
        value = i * stepSize: z = PI * SinhSafe(value): small = Exp(-Abs(z))
        If z >= 0 Then a = 1 / (1 + small): b = small / (1 + small) Else a = small / (1 + small): b = 1 / (1 + small)
        w = stepSize * PI * CoshSafe(value) * a * b
        If a > 0 And a < 1 And w > 0 Then r.Add a, 0, w
    Next i
    DECache.Add CStr(level), r: Set DE01Rule = r
End Function
Public Function RegularRule(ByVal d As Long, ByVal order As Long) As CIntegrationRule
    Dim r As CIntegrationRule, g As CIntegrationRule, i As Long, j As Long, key As String
    Call InitNumerics
    key = d & ":" & order
    If RegularCache.Exists(key) Then Set RegularRule = RegularCache(key): Exit Function
    Set g = Gauss01Rule(order): Set r = New CIntegrationRule: r.Allocate order ^ (d - 1)
    For i = 1 To order
        If d = 2 Then
            r.Add 2 * g.A(i) - 1, 0, 2 * g.W(i)
        Else
            For j = 1 To order: r.Add g.A(i), (1 - g.A(i)) * g.A(j), g.W(i) * g.W(j) * (1 - g.A(i)): Next j
        End If
    Next i
    RegularCache.Add key, r: Set RegularRule = r
End Function
Public Function SelfRule(ByVal d As Long, ByVal order As Long, ByVal fieldIndex As Long) As CIntegrationRule
    Dim r As CIntegrationRule, radial As CIntegrationRule, angular As CIntegrationRule, a As Double, b As Double, extent As Double
    Dim i As Long, j As Long, side As Long, sign As Long, vertices As Variant, p As Variant, q As Variant, det As Double, key As String, x As Double, y As Double
    Call InitNumerics
    key = d & ":" & order & ":" & fieldIndex & ":" & SelfOrder & ":" & BoundaryDE
    If SelfCache.Exists(key) Then Set SelfRule = SelfCache(key): Exit Function
    FieldParameters d, order, fieldIndex, a, b
    Set r = New CIntegrationRule
    If d = 2 Then
        Set radial = DE01Rule(BoundaryDE): r.Allocate 2 * radial.Count
        For side = 0 To 1
            sign = 2 * side - 1: extent = IIf(side = 0, a + 1, 1 - a)
            For i = 1 To radial.Count
                x = a + sign * extent * radial.A(i)
                If Abs(x - a) > 1E-12 Then r.Add x, 0, extent * radial.W(i)
            Next i
        Next side
    Else
        Set radial = Gauss01Rule(SelfOrder): Set angular = Gauss01Rule(MaxLong(48, SelfOrder))
        r.Allocate 3 * radial.Count * angular.Count: vertices = Array(Array(0#, 0#), Array(1#, 0#), Array(0#, 1#))
        For side = 0 To 2
            p = vertices(side): q = vertices((side + 1) Mod 3): det = Abs((p(0) - a) * (q(1) - b) - (p(1) - b) * (q(0) - a))
            For i = 1 To radial.Count
                For j = 1 To angular.Count
                    x = (1 - angular.A(j)) * (p(0) - a) + angular.A(j) * (q(0) - a)
                    y = (1 - angular.A(j)) * (p(1) - b) + angular.A(j) * (q(1) - b)
                    r.Add a + radial.A(i) * x, b + radial.A(i) * y, radial.A(i) * radial.W(i) * angular.W(j) * det
                Next j
            Next i
        Next side
    End If
    SelfCache.Add key, r: Set SelfRule = r
End Function

Public Function NearRule(ByVal geom As CElementData, ByVal a As Double, ByVal b As Double, ByVal distance As Double, ByVal level As Long) As CIntegrationRule
    Dim r As CIntegrationRule, radial As CIntegrationRule, angular As CIntegrationRule
    Dim p(1 To 3) As Double, normal(1 To 3) As Double, ds(1 To 3) As Double, dt(1 To 3) As Double, jac As Double
    Dim g00 As Double, g01 As Double, g11 As Double, speed As Double, delta As Double
    Dim side As Long, i As Long, j As Long, sign As Long, extent As Double, upper As Double, rho As Double
    Dim vertices As Variant, va As Variant, vb As Variant, ex As Double, ey As Double, length As Double, nx As Double, ny As Double, height As Double
    Dim tx As Double, ty As Double, wa As Double, wb As Double, angle As Double, cp As Double, sp As Double, limit As Double, Amap As Double, av As Double, angularWeight As Double, radialJac As Double
    MapGeometry geom, a, b, p, normal, jac, ds, dt
    For i = 1 To geom.D: g00 = g00 + ds(i) ^ 2: g01 = g01 + ds(i) * dt(i): g11 = g11 + dt(i) ^ 2: Next i
    speed = Sqr((g00 + g11 + Sqr((g00 - g11) ^ 2 + 4 * g01 * g01)) / 2)
    delta = MaxDouble(distance / speed, 1E-15): Set radial = DE01Rule(level): Set r = New CIntegrationRule
    If geom.D = 2 Then
        r.Allocate 2 * radial.Count
        For side = 0 To 1
            sign = 2 * side - 1: extent = IIf(side = 0, a + 1, 1 - a)
            If extent > 0 Then
                upper = AsinhSafe(extent / delta)
                For i = 1 To radial.Count: r.Add a + sign * delta * SinhSafe(upper * radial.A(i)), 0, delta * CoshSafe(upper * radial.A(i)) * upper * radial.W(i): Next i
            End If
        Next side
    Else
        Set angular = Gauss01Rule(12 * 2 ^ MinLong(level, 4))
        r.Allocate 3 * radial.Count * angular.Count: vertices = Array(Array(0#, 0#), Array(1#, 0#), Array(0#, 1#))
        For side = 0 To 2
            va = vertices(side): vb = vertices((side + 1) Mod 3)
            ex = vb(0) - va(0): ey = vb(1) - va(1): length = Sqr(ex * ex + ey * ey)
            nx = ey / length: ny = -ex / length: height = (va(0) - a) * nx + (va(1) - b) * ny
            If height > 1E-14 Then
                tx = -ny: ty = nx
                wa = AsinhSafe(((va(0) - a) * tx + (va(1) - b) * ty) / height)
                wb = AsinhSafe(((vb(0) - a) * tx + (vb(1) - b) * ty) / height)
                If wb <= wa Then Fail "三角形の極座標分割が不整合です。"
                For j = 1 To angular.Count
                    angle = wa + (wb - wa) * angular.A(j): cp = 1 / CoshSafe(angle): sp = SinhSafe(angle) * cp
                    limit = height * CoshSafe(angle): angularWeight = angular.W(j) * (wb - wa) * cp
                    Amap = Log1pSafe((limit / delta) ^ 2)
                    For i = 1 To radial.Count
                        av = Amap * radial.A(i): rho = delta * Sqr(Expm1Safe(av))
                        radialJac = 0.5 * delta * delta * Amap * Exp(av) * radial.W(i)
                        r.Add a + rho * (cp * nx + sp * tx), b + rho * (cp * ny + sp * ty), angularWeight * radialJac
                    Next i
                Next j
            End If
        Next side
    End If
    If r.Count = 0 Then Fail "近接積分点を生成できませんでした。"
    Set NearRule = r
End Function

Public Sub Kelvin(ByVal sr As CSolveRegion, ByRef x() As Double, ByRef y() As Double, ByRef normal() As Double, ByRef uk() As Double, ByRef tk() As Double, ByRef du() As Double, ByRef dt() As Double, Optional ByVal withGradient As Boolean = False)
    Dim d As Long, i As Long, j As Long, l As Long, h(1 To 3) As Double, rho As Double, hn As Double, alpha As Double, identity As Double
    Dim cu As Double, ct As Double, hh As Double, f As Double, di As Double, dj As Double, dhn As Double, dhI As Double, dhJ As Double, df As Double
    d = sr.Ref.Dimension
    For i = 1 To d: h(i) = y(i) - x(i): rho = rho + h(i) * h(i): Next i
    rho = Sqr(rho)
    If rho <= 0 Then Fail "基本解の評価点が一致しています。特異積分が必要です。"
    For i = 1 To d: h(i) = h(i) / rho: hn = hn + h(i) * normal(i): Next i
    alpha = 1 - 2 * sr.NuEff
    cu = 1 / (IIf(d = 2, 8, 16) * PI * sr.Mu * (1 - sr.NuEff))
    ct = 1 / (IIf(d = 2, 4, 8) * PI * (1 - sr.NuEff))
    For i = 1 To d
        For j = 1 To d
            identity = IIf(i = j, 1#, 0#): hh = h(i) * h(j)
            If d = 2 Then uk(i, j) = cu * (-(3 - 4 * sr.NuEff) * identity * Log(rho / sr.MeshScale) + hh) Else uk(i, j) = cu * ((3 - 4 * sr.NuEff) * identity + hh) / rho
            f = alpha * (identity * hn - h(i) * normal(j) + h(j) * normal(i)) + d * hh * hn
            tk(i, j) = -ct * f / rho ^ (d - 1)
            If withGradient Then
                For l = 1 To d
                    di = IIf(i = l, 1#, 0#): dj = IIf(j = l, 1#, 0#)
                    du(i, j, l) = cu * ((3 - 4 * sr.NuEff) * identity * h(l) - di * h(j) - h(i) * dj + d * hh * h(l)) / rho ^ (d - 1)
                    dhI = di - h(i) * h(l): dhJ = dj - h(j) * h(l): dhn = normal(l) - hn * h(l)
                    df = alpha * (identity * dhn - dhI * normal(j) + dhJ * normal(i)) + d * ((dhI * h(j) + h(i) * dhJ) * hn + hh * dhn)
                    dt(i, j, l) = ct * (df - (d - 1) * f * h(l)) / rho ^ d
                Next l
            End If
        Next j
    Next i
End Sub
