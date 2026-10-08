Attribute VB_Name = "BEM_Linear"
Option Explicit
Private Blocks As Collection
Public Function VectorNorm(ByRef values() As Double) As Double
    Dim scaleValue As Double, sumSquares As Double, i As Long, value As Double
    sumSquares = 1
    For i = LBound(values) To UBound(values)
        value = Abs(values(i))
        If value > 0 Then
            If scaleValue < value Then sumSquares = 1 + sumSquares * (scaleValue / value) ^ 2: scaleValue = value Else sumSquares = sumSquares + (value / scaleValue) ^ 2
        End If
    Next i
    VectorNorm = scaleValue * Sqr(sumSquares)
End Function
Public Sub SmallPseudoInverse(ByRef a() As Double, ByRef inverse() As Double)
    'One-sided Jacobi SVD avoids squaring the condition number of A.
    Dim b() As Double, v() As Double, norms() As Double, n As Long, i As Long, j As Long, p As Long, q As Long, sweep As Long
    Dim alpha As Double, beta As Double, gamma As Double, tau As Double, rotation As Double, c As Double, s As Double, first As Double, second As Double, largest As Double, changed As Boolean
    n = UBound(a, 1): b = a: ReDim v(1 To n, 1 To n): ReDim inverse(1 To n, 1 To n): ReDim norms(1 To n)
    For i = 1 To n: v(i, i) = 1: Next i
    For sweep = 1 To 80
        changed = False
        For p = 1 To n - 1: For q = p + 1 To n
            alpha = 0: beta = 0: gamma = 0
            For i = 1 To n: alpha = alpha + b(i, p) ^ 2: beta = beta + b(i, q) ^ 2: gamma = gamma + b(i, p) * b(i, q): Next i
            If Abs(gamma) > 4 * 2.22044604925031E-16 * Sqr(alpha) * Sqr(beta) Then
                changed = True: tau = (beta - alpha) / (2 * gamma)
                If tau = 0 Then rotation = 1 Else rotation = Sgn(tau) / (Abs(tau) + Sqr(1 + tau * tau))
                c = 1 / Sqr(1 + rotation * rotation): s = c * rotation
                For i = 1 To n
                    first = b(i, p): second = b(i, q): b(i, p) = c * first - s * second: b(i, q) = s * first + c * second
                    first = v(i, p): second = v(i, q): v(i, p) = c * first - s * second: v(i, q) = s * first + c * second
                Next i
            End If
        Next q: Next p
        If Not changed Then Exit For
    Next sweep
    For p = 1 To n
        For i = 1 To n: norms(p) = norms(p) + b(i, p) ^ 2: Next i
        largest = MaxDouble(largest, norms(p))
    Next p
    For p = 1 To n
        If norms(p) > largest * 1E-24 And norms(p) > 1E-300 Then
            For i = 1 To n: For j = 1 To n: inverse(i, j) = inverse(i, j) + v(i, p) * b(j, p) / norms(p): Next j: Next i
        End If
    Next p
End Sub
Private Sub AddBlock(ByRef rows() As Long, ByRef columns() As Long, ByRef matrix() As Double)
    Dim block As CUnknownBlock, inverse() As Double
    SmallPseudoInverse matrix, inverse
    Set block = New CUnknownBlock: block.count = UBound(rows): block.rows = rows: block.columns = columns: block.inverse = inverse: Blocks.Add block
End Sub
Public Sub BuildPreconditioner()
    Dim name As Variant, sr As CSolveRegion, a As CSolveRegion, first As Object, prior As Variant, field As Long, otherField As Long
    Dim i As Long, j As Long, index As Long, d As Long, key As String, matrix() As Double, rows() As Long, cols() As Long
    Dim dg() As Double, dh() As Double, ag() As Double, ah() As Double, umap() As Long, tmap() As Long, uf() As Double, tf() As Double
    Dim au() As Long, at() As Long, auf() As Double, atf() As Double, paired() As Boolean
    Set Blocks = New Collection: Set first = CreateObject("Scripting.Dictionary")
    For Each name In SolveData.keys
        Set sr = SolveData(name): d = sr.Ref.dimension
        dg = sr.DiagG: dh = sr.DiagH: umap = sr.umap: tmap = sr.tmap: uf = sr.UFactor: tf = sr.TFactor: paired = sr.paired
        For field = 1 To sr.NF
            If paired(field) Then
                key = CStr(umap((field - 1) * d + 1))
                If first.Exists(key) Then
                    prior = first(key): Set a = prior(0): otherField = prior(1)
                    ag = a.DiagG: ah = a.DiagH: au = a.umap: at = a.tmap: auf = a.UFactor: atf = a.TFactor
                    ReDim matrix(1 To 2 * d, 1 To 2 * d): ReDim rows(1 To 2 * d): ReDim cols(1 To 2 * d)
                    For i = 1 To d
                        rows(i) = a.Offset + (otherField - 1) * d + i: rows(d + i) = sr.Offset + (field - 1) * d + i
                        cols(i) = au((otherField - 1) * d + i): cols(d + i) = at((otherField - 1) * d + i)
                        For j = 1 To d
                            index = (otherField - 1) * d + j
                            matrix(i, j) = ah(otherField, i, j) * auf(index): matrix(i, d + j) = -ag(otherField, i, j) * atf(index)
                            index = (field - 1) * d + j
                            matrix(d + i, j) = dh(field, i, j) * uf(index): matrix(d + i, d + j) = -dg(field, i, j) * tf(index)
                        Next j
                    Next i
                    AddBlock rows, cols, matrix
                Else
                    first.Add key, Array(sr, field)
                End If
            Else
                ReDim matrix(1 To d, 1 To d): ReDim rows(1 To d): ReDim cols(1 To d)
                For i = 1 To d
                    rows(i) = sr.Offset + (field - 1) * d + i
                    index = (field - 1) * d + i: If umap(index) > 0 Then cols(i) = umap(index) Else cols(i) = tmap(index)
                    For j = 1 To d
                        index = (field - 1) * d + j
                        If umap(index) > 0 Then matrix(i, j) = dh(field, i, j) * uf(index) Else matrix(i, j) = -dg(field, i, j) * tf(index)
                    Next j
                Next i
                AddBlock rows, cols, matrix
            End If
        Next field
    Next name
End Sub
Public Sub Precondition(ByRef value() As Double, ByRef answer() As Double)
    Dim block As CUnknownBlock, rows() As Long, cols() As Long, inverse() As Double, i As Long, j As Long, total As Double
    ReDim answer(1 To SolveDOF)
    For Each block In Blocks
        rows = block.rows: cols = block.columns: inverse = block.inverse
        For i = 1 To block.count
            total = 0
            For j = 1 To block.count: total = total + inverse(i, j) * value(rows(j)): Next j
            answer(cols(i)) = total
        Next i
    Next block
End Sub
Private Sub Candidate(ByRef base() As Double, ByRef basis() As Double, ByRef h() As Double, ByRef gamma() As Double, ByVal count As Long, ByRef result() As Double)
    Dim coefficients() As Double, i As Long, j As Long, value As Double
    ReDim coefficients(1 To count): result = base
    For i = count To 1 Step -1
        value = gamma(i)
        For j = i + 1 To count: value = value - h(i, j) * coefficients(j): Next j
        If Abs(h(i, i)) > 1E-300 Then coefficients(i) = value / h(i, i)
    Next i
    For i = 1 To SolveDOF: For j = 1 To count: result(i) = result(i) + basis(i, j) * coefficients(j): Next j: Next i
End Sub
Public Sub GMRESSolve(ByRef rhs() As Double, ByRef solution() As Double)
    Dim basis() As Double, h() As Double, gamma() As Double, cosine() As Double, sine() As Double, residual() As Double, transformed() As Double
    Dim vec() As Double, av() As Double, w() As Double, base() As Double, candidateValue() As Double, trueResidual() As Double
    Dim n As Long, restart As Long, cycle As Long, inner As Long, i As Long, j As Long, pass As Long
    Dim rhsNorm As Double, preNorm As Double, beta As Double, value As Double, total As Double, length As Double, estimate As Double, actual As Double, happy As Boolean
    n = SolveDOF: restart = MinLong(GMRESRestart, n): ReDim solution(1 To n): ReDim residual(1 To n)
    rhsNorm = VectorNorm(rhs): Precondition rhs, transformed: preNorm = VectorNorm(transformed)
    SolveIterations = 0: SolveCycles = 0: SolveLinearInfo = 0
    If rhsNorm = 0 Then Exit Sub
    For cycle = 1 To GMRESMaxCycles
        SolveCycles = cycle: ApplySystem solution, av
        For i = 1 To n: residual(i) = rhs(i) - av(i): Next i
        actual = VectorNorm(residual) / rhsNorm
        If actual <= LinearTolerance Then Exit Sub
        Precondition residual, transformed: beta = VectorNorm(transformed)
        If beta <= 1E-300 Then Exit For
        base = solution
        ReDim basis(1 To n, 1 To restart + 1): ReDim h(1 To restart + 1, 1 To restart)
        ReDim gamma(1 To restart + 1): ReDim cosine(1 To restart): ReDim sine(1 To restart): ReDim vec(1 To n)
        For i = 1 To n: basis(i, 1) = transformed(i) / beta: Next i
        gamma(1) = beta
        For inner = 1 To restart
            For i = 1 To n: vec(i) = basis(i, inner): Next i
            ApplySystem vec, av: Precondition av, w
            length = VectorNorm(w)
            'A second MGS pass protects long restart spaces from loss of orthogonality.
            For pass = 1 To 2
                For j = 1 To inner
                    total = 0
                    For i = 1 To n: total = total + basis(i, j) * w(i): Next i
                    h(j, inner) = h(j, inner) + total
                    For i = 1 To n: w(i) = w(i) - total * basis(i, j): Next i
                Next j
            Next pass
            h(inner + 1, inner) = VectorNorm(w): happy = h(inner + 1, inner) <= 32 * 2.22044604925031E-16 * length
            If Not happy Then For i = 1 To n: basis(i, inner + 1) = w(i) / h(inner + 1, inner): Next i
            For j = 1 To inner - 1
                value = cosine(j) * h(j, inner) + sine(j) * h(j + 1, inner)
                h(j + 1, inner) = -sine(j) * h(j, inner) + cosine(j) * h(j + 1, inner): h(j, inner) = value
            Next j
            value = Sqr(h(inner, inner) ^ 2 + h(inner + 1, inner) ^ 2)
            If value = 0 Then cosine(inner) = 1 Else cosine(inner) = h(inner, inner) / value: sine(inner) = h(inner + 1, inner) / value
            h(inner, inner) = value: h(inner + 1, inner) = 0
            gamma(inner + 1) = -sine(inner) * gamma(inner): gamma(inner) = cosine(inner) * gamma(inner)
            SolveIterations = SolveIterations + 1: estimate = Abs(gamma(inner + 1))
            If happy Or estimate <= LinearTolerance * MaxDouble(preNorm, 1E-300) Or inner = restart Then
                Candidate base, basis, h, gamma, inner, candidateValue
                ApplySystem candidateValue, av
                For i = 1 To n: residual(i) = rhs(i) - av(i): Next i
                actual = VectorNorm(residual) / rhsNorm
                If actual <= LinearTolerance Then solution = candidateValue: Exit Sub
                If happy Or inner = restart Then solution = candidateValue: Exit For
            End If
            Application.StatusBar = "GMRES " & SolveIterations & " / ÄŽn“® " & cycle & " / „’èŽc· " & Format$(estimate / MaxDouble(preNorm, 1E-300), "0.00E+00"): DoEvents
        Next inner
    Next cycle
    SolveLinearInfo = SolveCycles
End Sub
Public Sub ClearPreconditioner()
    Set Blocks = Nothing
End Sub

