Attribute VB_Name = "BEM_AutoSelect"
Option Explicit
Private Function Elapsed(ByVal start As Double) As Double
    Elapsed = Timer - start: If Elapsed < 0 Then Elapsed = Elapsed + 86400
    Elapsed = MaxDouble(Elapsed, 0.000001)
End Function
Private Function ProbeAssembly(ByVal sr As CSolveRegion, ByRef targets() As Long, ByRef sources() As Long) As Double
    Dim elements As Object, i As Long, element As Variant, e As Long, start As Double, g() As Double, h() As Double
    Set elements = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(sources)
        e = (sources(i) - 1) \ sr.Ref.FieldCount
        If Not elements.Exists(CStr(e)) Then elements.Add CStr(e), e
    Next i
    start = Timer
    For i = 1 To UBound(targets)
        For Each element In elements.items: ElementInfluence sr, targets(i), CLng(element), g, h: Next element
    Next i
    ProbeAssembly = Elapsed(start)
End Function
Private Function MultiplyProbe(ByVal rows As Long, ByVal cols As Long) As Double
    Dim matrix() As Double, vector() As Double, answer() As Double, i As Long, j As Long, rep As Long, passes As Long, start As Double, value As Double
    ReDim matrix(1 To rows, 1 To cols): ReDim vector(1 To cols): ReDim answer(1 To rows)
    For j = 1 To cols: vector(j) = Sin(j * 0.31): Next j
    For i = 1 To rows: For j = 1 To cols: matrix(i, j) = Cos(i * 0.11 + j * 0.37): Next j: Next i
    passes = MaxLong(3, CLng(300000# / (CDbl(rows) * cols)))
    start = Timer
    For rep = 1 To passes
        For i = 1 To rows
            value = 0
            For j = 1 To cols: value = value + matrix(i, j) * vector(j): Next j
            answer(i) = value
        Next i
    Next rep
    MultiplyProbe = Elapsed(start) / passes
End Function
Public Sub SelectOperator(ByVal sr As CSolveRegion, ByRef remainingMB As Double)
    Dim plan As CFmmOperator, d As Long, n As Long, q As Long, nc As Long, sample As Long, ntarget As Long, i As Long, groupIndex As Long, groups As Long
    Dim sources() As Long, targets() As Long, start As Double, selected As String, denseMB As Double, fmmMB As Double, cacheMB As Double
    Dim denseAssembly As Double, nearTime As Double, nearSampleEntries As Double, denseMV As Double, farMV As Double, transferMV As Double, nearMV As Double, generation As Double
    Dim denseSeconds As Double, fmmSeconds As Double, calibration As Double, denseAllowed As Boolean, fmmAllowed As Boolean
    d = sr.Ref.dimension: n = sr.NF: q = FmmOrder ^ d: nc = d + d * d: start = Timer
    Set plan = New CFmmOperator: plan.InitializePlan sr
    denseMB = plan.DenseEstimateMB: cacheMB = MinDouble(FmmCacheMB, 0.2 * remainingMB): fmmMB = plan.FmmEstimateMB(cacheMB)
    denseAllowed = denseMB <= remainingMB And sr.ND <= MaxUnknowns
    fmmAllowed = fmmMB <= remainingMB And sr.ND <= MaxFmmUnknowns And LinearMethod = "gmres"
    If RequestedBackend = "dense" And Not denseAllowed Then Fail sr.Ref.RegionName & ": denseの未知数またはメモリ上限を超えます。上限を変更するかfmmを選択してください。"
    If RequestedBackend = "fmm" And Not fmmAllowed Then Fail sr.Ref.RegionName & ": FMMの未知数またはメモリ上限を超えます。分割・p・キャッシュを減らしてください。"
    If Not denseAllowed And Not fmmAllowed Then Fail sr.Ref.RegionName & ": dense / FMMとも設定された未知数・メモリ上限を超えます。"
    selected = RequestedBackend
    If selected = "auto" Then
        If Not denseAllowed Then
            selected = "fmm"
        ElseIf Not fmmAllowed Then
            selected = "dense"
        Else
            Application.StatusBar = "演算方式の実測選択 " & sr.Ref.RegionName: DoEvents
            sample = MinLong(n, IIf(d = 2, 1600, 256)): ntarget = MinLong(n, IIf(d = 2, 6, 2))
            ReDim sources(1 To sample): ReDim targets(1 To ntarget)
            For i = 1 To sample: sources(i) = 1 + CLng(Fix((i - 1) * (n - 1) / MaxLong(sample - 1, 1))): Next i
            For i = 1 To ntarget: targets(i) = 1 + CLng(Fix((i - 1) * (n - 1) / MaxLong(ntarget - 1, 1))): Next i
            denseAssembly = ProbeAssembly(sr, targets, sources) * n / ntarget * n / sample
            groups = plan.NearGroupCount: nearTime = 0: nearSampleEntries = 0
            For i = 1 To MinLong(groups, 3)
                groupIndex = CLng(Fix((i - 1) * (groups - 1) / MaxLong(MinLong(groups, 3) - 1, 1)))
                targets = plan.CalibrationTargets(groupIndex): sources = plan.CalibrationSources(groupIndex)
                nearTime = nearTime + ProbeAssembly(sr, targets, sources): nearSampleEntries = nearSampleEntries + CDbl(UBound(targets)) * UBound(sources)
            Next i
            nearTime = nearTime / MaxDouble(nearSampleEntries, 1) * plan.EntryCount
            sample = MinLong(n, 256): denseMV = 2 * MultiplyProbe(d * sample, d * sample) * (CDbl(n) / sample) ^ 2
            transferMV = MultiplyProbe(q, q) * (nc + d) * MaxLong(plan.NodeCount - 1, 0)
            nearMV = MultiplyProbe(d * MinLong(n, 128), d * MinLong(n, 128)) * 2 * plan.EntryCount / CDbl(MinLong(n, 128)) ^ 2
            farMV = 0: generation = 0
            If plan.FarCount > 0 Then
                farMV = MultiplyProbe(d * q, nc * q) * plan.FarCount
                generation = 0.000001 * q * q * plan.TranslationCount * (1 + IIf(8# * d * nc * q * q * plan.TranslationCount > cacheMB * 1048576, ExpectedIterations, 0))
            End If
            denseSeconds = denseAssembly + ExpectedIterations * denseMV
            fmmSeconds = nearTime + ExpectedIterations * (farMV + nearMV + transferMV) + generation + 0.00002 * plan.NodeCount * q
            selected = IIf(denseSeconds <= fmmSeconds, "dense", "fmm")
        End If
    End If
    calibration = Elapsed(start): sr.Backend = selected
    sr.EstimatedMB = IIf(selected = "dense", denseMB, fmmMB): remainingMB = remainingMB - sr.EstimatedMB
    sr.SelectionJSON = "{""selected"":" & JQuote(selected) & ",""dense_estimated_seconds"":" & JNum(denseSeconds) & ",""fmm_estimated_seconds"":" & JNum(fmmSeconds)
    sr.SelectionJSON = sr.SelectionJSON & ",""dense_estimated_MB"":" & JNum(denseMB) & ",""fmm_estimated_MB"":" & JNum(fmmMB) & ",""selection_seconds"":" & JNum(calibration) & ",""expected_iterations"":" & ExpectedIterations
    sr.SelectionJSON = sr.SelectionJSON & ",""remaining_model_memory_budget_MB"":" & JNum(remainingMB) & ",""prototype_mode"":""VBA_measured_integration_and_matrix_products"",""estimate_is_guarantee"":false}"
    If selected = "fmm" Then
        Set sr.Fmm = plan: plan.Prepare sr, cacheMB
    Else
        AssembleRegion sr: SetDenseDiagonal sr
    End If
    LogAction "演算方式選択", "OK", sr.Ref.RegionName & "/" & selected & "/推定MB " & JNum(sr.EstimatedMB)
End Sub

