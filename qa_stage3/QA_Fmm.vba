Option Explicit
Public Function QA_Apply(ByVal name As String, ByVal seed As Double, Optional ByVal translation As Long = 0) As Variant
    Dim sr As CSolveRegion, t() As Double, u() As Double, out() As Double, i As Long
    Set sr = SolveData(name): ReDim t(1 To sr.ND): ReDim u(1 To sr.ND)
    For i = 1 To sr.ND
        If translation = 0 Then
            t(i) = Cos((i - 1) * 0.29 + seed): u(i) = 0.001 * Sin((i - 1) * 0.37 + seed)
        Else
            If (i - 1) Mod sr.Ref.Dimension + 1 = translation Then u(i) = 1
        End If
    Next i
    ApplyRegion sr, t, u, out: QA_Apply = out
End Function
Public Function QA_State() As String
    Dim sr As CSolveRegion, name As Variant, rows As Collection, details As String
    Set rows = New Collection
    For Each name In SolveData.Keys
        Set sr = SolveData(name)
        If sr.Backend = "fmm" Then details = sr.Fmm.StatsJSON Else details = "{""backend"":""dense""}"
        rows.Add "{""name"":" & JQuote(CStr(name)) & ",""dense_arrays_empty"":" & LCase$(CStr(IsEmpty(sr.G) And IsEmpty(sr.H))) & ",""operator"":" & details & ",""selection"":" & sr.SelectionJSON & "}"
    Next name
    QA_State = "{""residual"":" & JNum(SolveResidual) & ",""iterations"":" & SolveIterations & ",""cycles"":" & SolveCycles & ",""linear_info"":" & SolveLinearInfo & ",""regions"": [" & JoinCollection(rows, ",") & "]}"
End Function
Public Function QA_Plan(ByVal name As String, ByVal which As String) As Variant
    Dim sr As CSolveRegion: Set sr = SolveData(name)
    If which = "nodes" Then QA_Plan = sr.Fmm.PlanNodes Else QA_Plan = sr.Fmm.PlanPairs(which = "far")
End Function
Public Function QA_Diagonal(ByVal name As String, ByVal which As String) As Variant
    Dim sr As CSolveRegion: Set sr = SolveData(name)
    If which = "G" Then QA_Diagonal = sr.DiagG Else QA_Diagonal = sr.DiagH
End Function
Public Function QA_Cheb(ByVal order As Long, ByVal d As Long) As Variant
    Dim cheb As CChebyshev, x(1 To 3) As Double, center(1 To 3) As Double, result() As Double
    Set cheb = New CChebyshev: cheb.Initialize order, d
    x(1) = 0.12: x(2) = -0.17: x(3) = 0.61: center(1) = -0.04: center(2) = 0.02: center(3) = 0.09
    cheb.Basis x, center, 0.8, result: QA_Cheb = result
End Function
Public Function QA_Pinv(ByVal n As Long, ByVal singular As Boolean) As Variant
    Dim a() As Double, inverse() As Double, i As Long, j As Long
    ReDim a(1 To n, 1 To n)
    For i = 1 To n: For j = 1 To n: a(i, j) = Sin(i * 0.73 + j * 0.17) + IIf(i = j, 2#, 0#): Next j: Next i
    If singular Then For i = 1 To n: a(i, n) = a(i, 1): Next i
    SmallPseudoInverse a, inverse: QA_Pinv = inverse
End Function
