Attribute VB_Name = "BEM_Solver"
Option Explicit
Public SolveData As Object
Public SolveDOF As Long
Public SolveResidual As Double
Public SolveConverged As Boolean
Public SolveAllChecks As Boolean
Public SolveSeconds As Double
Public InterfaceJump As Double
Public InterfaceImbalance As Double
Public LastSolvedFingerprint As String
Private SystemLU() As Double, Permutation() As Long, RowScale() As Double, RHS() As Double, Solution() As Double

Private Sub InitializeRegion(ByVal sr As CSolveRegion, ByVal r As CBemRegion)
    Dim e As Long, i As Long, j As Long, a As Double, b As Double, geom As CElementData, rule As CIntegrationRule
    Dim p As Variant, n As Variant, jac As Double, lower As Variant, upper As Variant, point As Variant
    Dim localPoint(1 To 3) As Double, normal(1 To 3) As Double, ds(1 To 3) As Double, dt(1 To 3) As Double
    Dim stage As String, message As String, number As Long
    On Error GoTo Failed
    stage = "材料・配列"
    Set sr.Ref = r: sr.FrameOrigin = r.Origin: sr.NF = r.elements.count * r.FieldCount: sr.ND = sr.NF * r.dimension
    sr.Mu = r.Young / (2 * (1 + r.Poisson)): sr.NuEff = r.Poisson
    If r.dimension = 2 And TextAt(WS("操作"), 6, 2) = "plane_stress" Then sr.NuEff = r.Poisson / (1 + r.Poisson)
    sr.IsPlaneStrain = r.dimension = 2 And TextAt(WS("操作"), 6, 2) = "plane_strain"
    sr.Lambda = 2 * sr.Mu * sr.NuEff / (1 - 2 * sr.NuEff)
    lower = r.vertices(1): upper = lower
    For Each point In r.vertices
        For j = 0 To r.dimension - 1: lower(j) = MinDouble(lower(j), point(j)): upper(j) = MaxDouble(upper(j), point(j)): Next j
    Next point
    For j = 0 To r.dimension - 1: sr.MeshScale = MaxDouble(sr.MeshScale, upper(j) - lower(j)): Next j
    sr.AllocateFields
    Set sr.Geometries = New Collection: Set rule = RegularRule(r.dimension, BoundaryOrder)
    For e = 0 To r.elements.count - 1
        stage = "要素幾何 " & e
        Set geom = New CElementData: geom.Initialize r, e, sr.FrameOrigin
        stage = "積分点準備 " & e
        Set geom.Regular = New CElementRule: geom.Regular.Initialize geom, rule: sr.Geometries.Add geom
        stage = "場節点 " & e
        For i = 1 To r.FieldCount
            FieldParameters r.dimension, r.ElementOrder, i, a, b
            MapGeometry geom, a, b, localPoint, normal, jac, ds, dt
            For j = 1 To r.dimension
                sr.SetField e * r.FieldCount + i, j, localPoint(j), normal(j)
            Next j
        Next i
    Next e
    Exit Sub
Failed:
    number = Err.number: message = Err.description
    Err.Raise number, "ElasticBEM", r.RegionName & "/" & stage & ": " & message
End Sub

Private Sub MeasureInterfaces()
    Dim first As Object, name As Variant, sr As CSolveRegion, i As Long, j As Long, index As Long, column As Long, prior As Variant
    Set first = CreateObject("Scripting.Dictionary"): InterfaceJump = 0: InterfaceImbalance = 0
    For Each name In SolveData.keys
        Set sr = SolveData(name)
        For i = 1 To sr.NF
            If sr.paired(i) Then
                For j = 1 To sr.Ref.dimension
                    index = (i - 1) * sr.Ref.dimension + j: column = sr.umap(index)
                    If first.Exists(CStr(column)) Then
                        prior = first(CStr(column))
                        InterfaceJump = MaxDouble(InterfaceJump, Abs(sr.u(index) - prior(0)))
                        InterfaceImbalance = MaxDouble(InterfaceImbalance, Abs(sr.t(index) + prior(1)))
                    Else
                        first.Add CStr(column), Array(sr.u(index), sr.t(index))
                    End If
                Next j
            End If
        Next i
    Next name
End Sub

Private Sub SetUnknownMaps()
    Dim name As Variant, sr As CSolveRegion, a As CSolveRegion, b As CSolveRegion, boundary As Variant, rule As Variant, pair As Variant
    Dim e As Variant, f As Long, i As Long, j As Long, l As Long, h As Long, ia As Long, ib As Long, best As Long, field As Long, d As Long, index As Long, cursor As Long
    Dim ucol As Long, tcol As Long, scaleFactor As Double, distance As Double, nearest As Double, ea As Collection, eb As Collection
    d = CurrentDimension()
    For Each pair In InterfacePairs
        Set a = SolveData(pair(0)): Set b = SolveData(pair(2)): Set ea = a.Ref.groups(pair(1)): Set eb = pair(4)
        scaleFactor = Sqr(a.Ref.Young) * Sqr(b.Ref.Young) / MaxDouble(a.MeshScale, b.MeshScale)
        For h = 1 To ea.count
            e = ea(h): f = eb(h)
            For l = 1 To a.Ref.FieldCount
                ia = CLng(e) * a.Ref.FieldCount + l: nearest = 1E+250: best = 0
                For j = 1 To b.Ref.FieldCount
                    ib = f * b.Ref.FieldCount + j: distance = 0
                    For i = 1 To d: distance = distance + ((a.FrameOrigin(i - 1) - b.FrameOrigin(i - 1)) + a.x(ia, i) - b.x(ib, i)) ^ 2: Next i
                    If distance < nearest Then nearest = distance: best = ib
                Next j
                If Sqr(nearest) > MaxDouble(a.MeshScale, b.MeshScale) * 0.000000001 Then Fail "接合面の場節点が一致しません。"
                If a.paired(ia) Or b.paired(best) Then Fail "接合面の場節点が重複しています。"
                a.SetInterface ia, cursor, cursor + d, scaleFactor
                b.SetInterface best, cursor, cursor + d, -scaleFactor
                cursor = cursor + 2 * d
            Next l
        Next h
    Next pair
    For Each name In SolveData.keys
        Set sr = SolveData(name): scaleFactor = sr.Ref.Young / sr.MeshScale
        For Each boundary In sr.Ref.groups.keys
            If Not sr.Ref.InterfaceGroups.Exists(boundary) Then
                rule = sr.Ref.RuleFor(CStr(boundary))
                For Each e In sr.Ref.groups(boundary)
                    For l = 1 To sr.Ref.FieldCount
                        field = CLng(e) * sr.Ref.FieldCount + l
                        For j = 1 To d
                            cursor = cursor + 1: index = (field - 1) * d + j
                            sr.SetExternal index, rule(0)(j - 1) = "u", rule(1)(j - 1), cursor, scaleFactor
                        Next j
                    Next l
                Next e
            End If
        Next boundary
    Next name
    If cursor <> SolveDOF Then Fail "境界未知数と方程式の数が一致しません。"
End Sub

Private Sub IntegrateBlock(ByVal sr As CSolveRegion, ByVal target As Long, ByVal e As Long, ByVal prepared As CElementRule, ByVal own As Boolean, ByRef gm() As Double, ByRef hm() As Double)
    Dim d As Long, k As Long, q As Long, a As Long, b As Long, l As Long, ownLocal As Long, row As Long, col As Long
    Dim x(1 To 3) As Double, y(1 To 3) As Double, normal(1 To 3) As Double
    Dim uk(1 To 3, 1 To 3) As Double, tk(1 To 3, 1 To 3) As Double, du(1 To 3, 1 To 3, 1 To 3) As Double, dt(1 To 3, 1 To 3, 1 To 3) As Double
    Dim ng As Double, nt As Double, wg(1 To 6) As Double, wt(1 To 6) As Double
    Dim positions() As Double, directions() As Double, weights() As Double, shapes() As Double
    positions = prepared.xyz: directions = prepared.Normals: weights = prepared.weight: shapes = prepared.n
    d = sr.Ref.dimension: k = sr.Ref.FieldCount: ownLocal = (target - 1) Mod k + 1
    For a = 1 To d: x(a) = sr.x(target, a): Next a
    For q = 1 To prepared.count
        For a = 1 To d: y(a) = positions(q, a): normal(a) = directions(q, a): Next a
        Kelvin sr, x, y, normal, uk, tk, du, dt
        For l = 1 To k
            wg(l) = weights(q) * shapes(q, l): wt(l) = wg(l)
            If own And l = ownLocal Then wt(l) = wt(l) - weights(q)
        Next l
        For a = 1 To d
            row = (target - 1) * d + a
            For b = 1 To d
                For l = 1 To k
                    col = (e * k + l - 1) * d + b
                    gm(row, col) = gm(row, col) + wg(l) * uk(a, b)
                    hm(row, col) = hm(row, col) + wt(l) * tk(a, b)
                Next l
            Next b
        Next a
    Next q
End Sub
Public Sub AssembleRegion(ByVal sr As CSolveRegion)
    Dim i As Long, e As Long, a As Long, b As Long, j As Long, d As Long, k As Long, row As Long, col As Long
    Dim geom As CElementData, rule As CIntegrationRule, prepared As CElementRule, own As Boolean, useRegular As Boolean
    Dim x(1 To 3) As Double, s As Double, t As Double, distance As Double, correction As Double
    Dim gm() As Double, hm() As Double
    d = sr.Ref.dimension: k = sr.Ref.FieldCount
    ReDim gm(1 To sr.ND, 1 To sr.ND): ReDim hm(1 To sr.ND, 1 To sr.ND)
    For i = 1 To sr.NF
        For a = 1 To d: x(a) = sr.x(i, a): Next a
        For e = 0 To sr.Ref.elements.count - 1
            Set geom = sr.Geometries(e + 1): own = (i - 1) \ k = e: useRegular = False
            If own Then
                Set rule = SelfRule(d, sr.Ref.ElementOrder, (i - 1) Mod k + 1)
            Else
                ClosestPoint geom, x, s, t, distance
                If distance / geom.length > 0.2 Then useRegular = True Else Set rule = NearRule(geom, s, t, distance, BoundaryDE)
            End If
            If useRegular Then
                Set prepared = geom.Regular
            Else
                Set prepared = New CElementRule: prepared.Initialize geom, rule
            End If
            IntegrateBlock sr, i, e, prepared, own, gm, hm
        Next e
        'Row-sum correction reproduces the source's rigid-translation identity H*1=0.
        For a = 1 To d
            row = (i - 1) * d + a
            For b = 1 To d
                correction = 0
                For j = 1 To sr.NF: correction = correction + hm(row, (j - 1) * d + b): Next j
                col = (i - 1) * d + b: hm(row, col) = hm(row, col) - correction
            Next b
        Next a
        If i Mod 4 = 0 Or i = sr.NF Then Application.StatusBar = "BEM係数積分 " & sr.Ref.RegionName & " " & i & "/" & sr.NF: DoEvents
    Next i
    sr.g = gm: sr.h = hm
End Sub
Public Sub ApplySystem(ByRef unknowns() As Double, ByRef answer() As Double)
    Dim name As Variant, sr As CSolveRegion, u() As Double, t() As Double, result() As Double
    Dim umap() As Long, tmap() As Long, uf() As Double, tf() As Double, j As Long
    ReDim answer(1 To SolveDOF)
    For Each name In SolveData.keys
        Set sr = SolveData(name): ReDim u(1 To sr.ND): ReDim t(1 To sr.ND)
        umap = sr.umap: tmap = sr.tmap: uf = sr.UFactor: tf = sr.TFactor
        For j = 1 To sr.ND
            If umap(j) > 0 Then u(j) = uf(j) * unknowns(umap(j))
            If tmap(j) > 0 Then t(j) = -tf(j) * unknowns(tmap(j))
        Next j
        ApplyRegion sr, t, u, result
        For j = 1 To sr.ND: answer(sr.Offset + j) = result(j): Next j
    Next name
End Sub
Private Sub BuildRHS()
    Dim name As Variant, sr As CSolveRegion, u() As Double, t() As Double, result() As Double, j As Long
    ReDim RHS(1 To SolveDOF)
    For Each name In SolveData.keys
        Set sr = SolveData(name): u = sr.u0: t = sr.t0
        For j = 1 To sr.ND: t(j) = -t(j): Next j
        ApplyRegion sr, t, u, result
        For j = 1 To sr.ND: RHS(sr.Offset + j) = -result(j): Next j
    Next name
End Sub
Private Sub AssembleSystem()
    Dim name As Variant, sr As CSolveRegion, i As Long, j As Long, row As Long, value As Double
    Dim gm() As Double, hm() As Double, u0() As Double, t0() As Double, umap() As Long, tmap() As Long, uf() As Double, tf() As Double
    ReDim SystemLU(1 To SolveDOF, 1 To SolveDOF): ReDim RHS(1 To SolveDOF)
    For Each name In SolveData.keys
        Set sr = SolveData(name)
        gm = sr.g: hm = sr.h: u0 = sr.u0: t0 = sr.t0: umap = sr.umap: tmap = sr.tmap: uf = sr.UFactor: tf = sr.TFactor
        For i = 1 To sr.ND
            row = sr.Offset + i
            For j = 1 To sr.ND
                RHS(row) = RHS(row) + gm(i, j) * t0(j) - hm(i, j) * u0(j)
                If umap(j) > 0 Then SystemLU(row, umap(j)) = SystemLU(row, umap(j)) + hm(i, j) * uf(j)
                If tmap(j) > 0 Then SystemLU(row, tmap(j)) = SystemLU(row, tmap(j)) - gm(i, j) * tf(j)
            Next j
        Next i
    Next name
End Sub
Private Sub FactorLU()
    Dim n As Long, i As Long, j As Long, k As Long, pivot As Long, prior As Long
    Dim largest As Double, value As Double, factor As Double
    n = SolveDOF: ReDim Permutation(1 To n): ReDim RowScale(1 To n)
    For i = 1 To n
        Permutation(i) = i
        For j = 1 To n: RowScale(i) = MaxDouble(RowScale(i), Abs(SystemLU(i, j))): Next j
        If RowScale(i) <= 1E-250 Then Fail "係数行列にゼロ行があります。境界条件を確認してください。"
        For j = 1 To n: SystemLU(i, j) = SystemLU(i, j) / RowScale(i): Next j
    Next i
    For k = 1 To n
        pivot = k: largest = Abs(SystemLU(k, k))
        For i = k + 1 To n
            If Abs(SystemLU(i, k)) > largest Then largest = Abs(SystemLU(i, k)): pivot = i
        Next i
        If largest < 0.0000000000001 Then Fail "係数行列が特異または悪条件です。拘束・材料・メッシュを確認してください。"
        If pivot <> k Then
            For j = 1 To n: value = SystemLU(k, j): SystemLU(k, j) = SystemLU(pivot, j): SystemLU(pivot, j) = value: Next j
            prior = Permutation(k): Permutation(k) = Permutation(pivot): Permutation(pivot) = prior
        End If
        For i = k + 1 To n
            factor = SystemLU(i, k) / SystemLU(k, k): SystemLU(i, k) = factor
            If factor <> 0 Then
                For j = k + 1 To n: SystemLU(i, j) = SystemLU(i, j) - factor * SystemLU(k, j): Next j
            End If
        Next i
        If k Mod 16 = 0 Then Application.StatusBar = "BEM連立方程式 LU " & k & "/" & n: DoEvents
    Next k
End Sub
Private Sub LUSolve(ByRef values() As Double, ByRef answer() As Double)
    Dim i As Long, j As Long, value As Double
    ReDim answer(1 To SolveDOF)
    For i = 1 To SolveDOF
        value = values(Permutation(i)) / RowScale(Permutation(i))
        For j = 1 To i - 1: value = value - SystemLU(i, j) * answer(j): Next j
        answer(i) = value
    Next i
    For i = SolveDOF To 1 Step -1
        value = answer(i)
        For j = i + 1 To SolveDOF: value = value - SystemLU(i, j) * answer(j): Next j
        answer(i) = value / SystemLU(i, i)
    Next i
End Sub
Private Sub DecodeSolution()
    Dim name As Variant, sr As CSolveRegion, j As Long
    Dim u() As Double, t() As Double, umap() As Long, tmap() As Long, uf() As Double, tf() As Double
    For Each name In SolveData.keys
        Set sr = SolveData(name)
        u = sr.u0: t = sr.t0: umap = sr.umap: tmap = sr.tmap: uf = sr.UFactor: tf = sr.TFactor
        For j = 1 To sr.ND
            If umap(j) > 0 Then u(j) = u(j) + uf(j) * Solution(umap(j))
            If tmap(j) > 0 Then t(j) = t(j) + tf(j) * Solution(tmap(j))
        Next j
        sr.u = u: sr.t = t
    Next name
End Sub
Private Function TrueResidual(ByRef residual() As Double) As Double
    Dim name As Variant, sr As CSolveRegion, i As Long, j As Long, row As Long, numerator As Double, denominator As Double, value As Double, magnitude As Double, termScale As Double
    Dim gm() As Double, hm() As Double, u() As Double, t() As Double
    Dim result() As Double, dg() As Double, dh() As Double, umax As Double, tmax As Double, a As Long, b As Long, d As Long, field As Long
    ReDim residual(1 To SolveDOF)
    For Each name In SolveData.keys
        Set sr = SolveData(name)
        u = sr.u: t = sr.t: umax = 0: tmax = 0
        For j = 1 To sr.ND: umax = MaxDouble(umax, Abs(u(j))): tmax = MaxDouble(tmax, Abs(t(j))): t(j) = -t(j): Next j
        ApplyRegion sr, t, u, result
        If sr.backend = "dense" Then gm = sr.g: hm = sr.h Else dg = sr.DiagG: dh = sr.DiagH
        For i = 1 To sr.ND
            value = result(i): magnitude = 0
            If sr.backend = "dense" Then
                For j = 1 To sr.ND: magnitude = magnitude + Abs(hm(i, j) * u(j)) + Abs(gm(i, j) * t(j)): Next j
            Else
                d = sr.Ref.dimension: field = (i - 1) \ d + 1: a = (i - 1) Mod d + 1
                For b = 1 To d: magnitude = magnitude + sr.NF * (Abs(dh(field, a, b)) * umax + Abs(dg(field, a, b)) * tmax): Next b
            End If
            termScale = termScale + magnitude * magnitude
            row = sr.Offset + i: residual(row) = -value: numerator = numerator + value * value
        Next i
    Next name
    For i = 1 To SolveDOF: denominator = denominator + RHS(i) * RHS(i): Next i
    'A rigid translation has zero physical RHS. Avoid dividing roundoff by roundoff.
    If Sqr(denominator) <= 128 * 2.22044604925031E-16 * Sqr(termScale) Then denominator = termScale
    TrueResidual = Sqr(numerator) / MaxDouble(Sqr(denominator), 1E-250)
End Function

Public Sub SolveBoundary()
    Dim name As Variant, r As CBemRegion, sr As CSolveRegion, total As Long, correction() As Double, residual() As Double, iteration As Long, i As Long
    Dim remainingMB As Double, workspaceMB As Double
    Call ReadSolveOptions
    Call ReadFmmOptions
    total = 0
    For Each name In Models.keys
        Set r = Models(name): total = total + r.elements.count * r.FieldCount * r.dimension
    Next name
    If total > MaxFmmUnknowns Then Fail "未知数 " & total & " がモデル上限 " & MaxFmmUnknowns & " を超えます。分割を減らすか上限を変更してください。"
    If LinearMethod = "lu" And total > MaxUnknowns Then Fail "LUの未知数 " & total & " が密行列上限 " & MaxUnknowns & " を超えます。gmres / fmmを選択してください。"
    If LinearMethod = "lu" Then workspaceMB = 8# * total * total / 1048576 Else workspaceMB = (8# * total * (MinLong(total, GMRESRestart) + 10) + 8# * (GMRESRestart + 1) ^ 2) / 1048576
    remainingMB = MemoryBudgetMB - workspaceMB
    If remainingMB <= 0 Then Fail "GMRES/LUの作業配列がメモリ上限を超えます。restartまたは未知数を減らしてください。"
    SolveDOF = total: Set SolveData = CreateObject("Scripting.Dictionary"): total = 0
    For Each name In Models.keys
        Set r = Models(name): Set sr = New CSolveRegion: InitializeRegion sr, r: sr.Offset = total: total = total + sr.ND
        SolveData.Add CStr(name), sr
    Next name
    SetUnknownMaps
    SolveBackend = "": SolveIterations = 0: SolveCycles = 0: SolveLinearInfo = 0
    For Each name In SolveData.keys
        Set sr = SolveData(name): SelectOperator sr, remainingMB
        If SolveBackend = "" Then SolveBackend = sr.backend Else If SolveBackend <> sr.backend Then SolveBackend = "mixed"
    Next name
    If LinearMethod = "lu" Then
        If SolveBackend <> "dense" Then Fail "LUには全領域denseが必要です。gmresを選択してください。"
        Call AssembleSystem
        Call FactorLU
        LUSolve RHS, Solution
    Else
        Call BuildRHS
        Call BuildPreconditioner
        GMRESSolve RHS, Solution
    End If
    Call DecodeSolution
    SolveResidual = TrueResidual(residual)
    If LinearMethod = "lu" Then
        For iteration = 1 To 3
            If SolveResidual <= LinearTolerance Then Exit For
            LUSolve residual, correction
            For i = 1 To SolveDOF: Solution(i) = Solution(i) + correction(i): Next i
            Call DecodeSolution
            SolveResidual = TrueResidual(residual)
        Next iteration
    End If
    SolveConverged = SolveResidual <= MaxDouble(2 * LinearTolerance, 0.000000000001) And SolveLinearInfo = 0
    Call MeasureInterfaces
End Sub

Public Function SolveCore() As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldStatus As Variant, oldCancel As Long, oldBusy As Boolean, start As Double, signature As String
    Dim boundary As Collection, interior As Collection, allGood As Boolean, description As String, errnum As Long
    If Busy Then Fail "ほかの処理を実行中です。完了を待ってください。"
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating: oldStatus = Application.StatusBar: oldCancel = Application.EnableCancelKey: oldBusy = Busy
    On Error GoTo Failed
    Call EnsureMesh
    LastSolvedFingerprint = "": SolveAllChecks = False
    signature = InputFingerprint(): start = Timer
    Busy = True: Application.EnableEvents = False: Application.ScreenUpdating = False: Application.EnableCancelKey = xlErrorHandler
    WS("操作").Range("B22").Value2 = "VBA解析中"
    Call SolveBoundary
    Set boundary = BoundaryResultRows()
    Set interior = New Collection: allGood = SolveConverged
    Call EvaluateAllFields(interior, allGood)
    If InputFingerprint() <> signature Then Fail "解析中に入力が変更されました。結果は保存せず、再解析してください。"
    ClearOutput "境界結果", 13: ClearOutput "内点結果", 17
    WS("境界結果").Range("A6").Resize(boundary.count, 13).Value2 = RowsMatrix(boundary, 13)
    WS("境界結果").Range("E6:M" & boundary.count + 5).NumberFormat = "0.000000E+00"
    If interior.count > 0 Then
        WS("内点結果").Range("A6").Resize(interior.count, 17).Value2 = RowsMatrix(interior, 17)
        WS("内点結果").Range("C6:N" & interior.count + 5).NumberFormat = "0.000000E+00"
    End If
    SolveSeconds = Timer - start: If SolveSeconds < 0 Then SolveSeconds = SolveSeconds + 86400
    LastImportFingerprint = MeshFingerprint: LastSolvedFingerprint = signature: SolveAllChecks = allGood
    WS("操作").Range("B22").Value2 = IIf(allGood, "VBA解析済み（収束OK）", "VBA解析済み（精度未達あり）")
    WS("操作").Range("B23").Value2 = "境界・内部場を計算済み"
    WS("操作").Range("D34").Value2 = "VBA " & SolveBackend & " / " & UCase$(LinearMethod)
    WS("操作").Range("D35").Value2 = "真の相対残差: " & JNum(SolveResidual)
    WS("操作").Range("D36").Value2 = "計算時間: " & Format$(SolveSeconds, "0.00") & " 秒"
    WS("操作").Range("D49").Value2 = "GMRES反復数: " & SolveIterations & " / 再始動: " & SolveCycles
    WS("操作").Range("D50").Value2 = "線形判定: " & IIf(SolveConverged, "OK", "LINEAR_SOLVE_TOLERANCE_UNMET")
    WS("境界結果").Range("A3").Value2 = "現在の入力をVBAで解析。変位[m]、表面力[kPa]。場節点と幾何節点は異なります。"
    WS("内点結果").Range("A3").Value2 = "現在の入力をVBAで解析。積分収束と値取得可否を確認してください。"
    LogAction "VBA解析", IIf(allGood, "OK", "精度未達"), "未知数 " & SolveDOF & " / 相対残差 " & JNum(SolveResidual) & " / " & Format$(SolveSeconds, "0.00") & "秒"
    SolveCore = IIf(allGood, "OK", "UNMET") & ": " & SolveDOF & " DOF / residual=" & JNum(SolveResidual)
    GoTo Finished
Failed:
    errnum = Err.number: description = Err.description
    If errnum = 18 Then description = "解析を中断しました。"
    WS("操作").Range("B22").Value2 = "VBA解析未完了・再解析が必要"
    LogAction "VBA解析", "ERROR", description
Finished:
    Application.EnableCancelKey = oldCancel: Application.StatusBar = oldStatus: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: Busy = oldBusy
    Erase SystemLU: Erase RHS: Erase Solution: Erase Permutation: Erase RowScale
    Call ClearPreconditioner
    If errnum <> 0 Then Err.Raise errnum, "ElasticBEM", description
End Function
Public Function BoundaryResultRows() As Collection
    Dim rows As Collection, name As Variant, sr As CSolveRegion, i As Long, j As Long, d As Long, k As Long, row As Variant
    Set rows = New Collection
    For Each name In SolveData.keys
        Set sr = SolveData(name): d = sr.Ref.dimension: k = sr.Ref.FieldCount
        For i = 1 To sr.NF
            row = Array(CStr(name), i, (i - 1) \ k + 1, (i - 1) Mod k + 1, sr.WorldCoordinate(i, 1), sr.WorldCoordinate(i, 2), Empty, Empty, Empty, Empty, Empty, Empty, Empty)
            If d = 3 Then row(6) = sr.WorldCoordinate(i, 3)
            For j = 1 To d: row(6 + j) = sr.u((i - 1) * d + j): row(9 + j) = sr.t((i - 1) * d + j): Next j
            rows.Add row
        Next i
    Next name
    Set BoundaryResultRows = rows
End Function
Public Sub BEM_Solve()
    Dim result As String
    On Error GoTo Failed
    result = SolveCore(): WS("境界結果").Activate: Exit Sub
Failed:
    MsgBox Err.description, vbExclamation, "BEM VBA解析"
End Sub

