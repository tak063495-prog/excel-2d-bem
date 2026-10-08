Option Explicit
Public Function QA_Matrix(ByVal name As String, ByVal which As String) As Variant
    Dim sr As CSolveRegion: Set sr = SolveData(name)
    If which = "G" Then QA_Matrix = sr.G Else QA_Matrix = sr.H
End Function

Public Function QA_InitAffine() As Variant
    Dim name As Variant, sr As CSolveRegion, r As CBemRegion, e As Long, l As Long, j As Long, geom As CElementData
    Dim a As Double, b As Double, jac As Double, p As Variant, n As Variant
    Call ReadSolveOptions
    Set SolveData = CreateObject("Scripting.Dictionary")
    For Each name In Models.Keys
        Set r = Models(name): Set sr = New CSolveRegion: Set sr.Ref = r
        sr.NF = r.Elements.Count * r.FieldCount: sr.ND = sr.NF * r.Dimension
        sr.Mu = r.Young / (2 * (1 + r.Poisson)): sr.NuEff = r.Poisson
        If r.Dimension = 2 And TextAt(WS("操作"), 6, 2) = "plane_stress" Then sr.NuEff = r.Poisson / (1 + r.Poisson)
        sr.IsPlaneStrain = r.Dimension = 2 And TextAt(WS("操作"), 6, 2) = "plane_strain"
        sr.Lambda = 2 * sr.Mu * sr.NuEff / (1 - 2 * sr.NuEff): sr.MeshScale = r.ModelScale
        sr.AllocateFields: Set sr.Geometries = New Collection
        For e = 0 To r.Elements.Count - 1
            Set geom = New CElementData: geom.Initialize r, e: sr.Geometries.Add geom
            For l = 1 To r.FieldCount
                p = r.CollocationPoint(e, l - 1): FieldParameters r.Dimension, r.ElementOrder, l, a, b
                n = r.Differential(e, a, b, jac)
                For j = 1 To r.Dimension: sr.SetField e * r.FieldCount + l, j, p(j - 1), n(j - 1): Next j
            Next l
        Next e
        SolveData.Add CStr(name), sr
    Next name
    QA_InitAffine = QA_AffineFields()
End Function
Public Function QA_Closest() As Variant
    Dim r As CBemRegion, geom As CElementData, e As Long, d As Long, j As Long, p As Variant
    Dim x(1 To 3) As Double, s As Double, t As Double, distance As Double, result() As Variant
    Set r = Models.Items()(0): d = r.Dimension: p = r.Points(1)
    For j = 1 To d: x(j) = p(j - 1): Next j
    ReDim result(1 To r.Elements.Count, 1 To 3)
    For e = 0 To r.Elements.Count - 1
        Set geom = New CElementData: geom.Initialize r, e
        ClosestPoint geom, x, s, t, distance
        result(e + 1, 1) = s: result(e + 1, 2) = t: result(e + 1, 3) = distance
    Next e
    QA_Closest = result
End Function
Public Function QA_Kernel() As String
    Dim sr As CSolveRegion, r As CBemRegion, d As Long, i As Long, j As Long, l As Long, parts As Collection
    Dim x(1 To 3) As Double, y(1 To 3) As Double, n(1 To 3) As Double
    Dim u(1 To 3, 1 To 3) As Double, t(1 To 3, 1 To 3) As Double, du(1 To 3, 1 To 3, 1 To 3) As Double, dt(1 To 3, 1 To 3, 1 To 3) As Double
    Set sr = New CSolveRegion: Set r = Models.Items()(0): Set sr.Ref = r: d = r.Dimension
    sr.Mu = r.Young / (2 * (1 + r.Poisson)): sr.NuEff = r.Poisson: sr.MeshScale = 0.8
    If d = 2 And TextAt(WS("操作"), 6, 2) = "plane_stress" Then sr.NuEff = r.Poisson / (1 + r.Poisson)
    x(1) = 0.12: x(2) = -0.17: x(3) = 0.09: y(1) = 0.78: y(2) = 0.28: y(3) = 0.61
    If d = 2 Then n(1) = 0.6: n(2) = 0.8 Else n(1) = 1 / 3: n(2) = 2 / 3: n(3) = 2 / 3
    Kelvin sr, x, y, n, u, t, du, dt, True
    Set parts = New Collection
    For i = 1 To d: For j = 1 To d: parts.Add JNum(u(i, j)): Next j: Next i
    For i = 1 To d: For j = 1 To d: parts.Add JNum(t(i, j)): Next j: Next i
    For i = 1 To d: For j = 1 To d: For l = 1 To d: parts.Add JNum(du(i, j, l)): Next l: Next j: Next i
    For i = 1 To d: For j = 1 To d: For l = 1 To d: parts.Add JNum(dt(i, j, l)): Next l: Next j: Next i
    QA_Kernel = JoinCollection(parts, ",")
End Function
Public Function QA_Gauss(ByVal order As Long) As Variant
    Dim r As CIntegrationRule, a() As Variant, i As Long
    Set r = Gauss01Rule(order): ReDim a(1 To r.Count, 1 To 2)
    For i = 1 To r.Count: a(i, 1) = r.A(i): a(i, 2) = r.W(i): Next i
    QA_Gauss = a
End Function
Public Function QA_AffineFields() As Variant
    Dim name As Variant, sr As CSolveRegion, d As Long, i As Long, j As Long, index As Long, u() As Double, t() As Double
    Dim ex As Double, ey As Double, ez As Double, strain(1 To 3) As Double, stress(1 To 3) As Double, trace As Double, rows As Collection, good As Boolean
    For Each name In SolveData.Keys
        Set sr = SolveData(name): d = sr.Ref.Dimension
        ex = 100 / sr.Ref.Young: ey = -sr.Ref.Poisson * ex: ez = ey
        If sr.IsPlaneStrain Then ex = ex * (1 - sr.Ref.Poisson ^ 2): ey = -sr.Ref.Poisson * (1 + sr.Ref.Poisson) * 100 / sr.Ref.Young
        strain(1) = ex: strain(2) = ey: strain(3) = ez
        trace = ex + ey: If d = 3 Then trace = trace + ez
        For j = 1 To d: stress(j) = 2 * sr.Mu * strain(j) + sr.Lambda * trace: Next j
        u = sr.U: t = sr.T
        For i = 1 To sr.NF
            For j = 1 To d
                index = (i - 1) * d + j
                u(index) = strain(j) * sr.X(i, j): t(index) = stress(j) * sr.Normal(i, j)
            Next j
        Next i
        sr.U = u: sr.T = t
    Next name
    Set rows = New Collection: good = True: EvaluateAllFields rows, good
    QA_AffineFields = RowsMatrix(rows, 17)
End Function
