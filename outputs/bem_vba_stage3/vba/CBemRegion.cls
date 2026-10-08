VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBemRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RegionName As String
Public ShapeName As String
Public MaterialId As String
Public Young As Double
Public Poisson As Double
Public Dimension As Long
Public ElementOrder As Long
Public Origin As Variant
Public Sizes As Variant
Public Radius As Double
Public ModelScale As Double
Public Divisions As Variant
Public Vertices As Collection
Public Elements As Collection
Public Groups As Object
Public NodeLookup As Object
Public Rules As Object
Public InterfaceGroups As Object
Public Points As Collection
Public PointIds As Collection

Private Sub Class_Initialize()
    Set Vertices = New Collection
    Set Elements = New Collection
    Set Groups = CreateObject("Scripting.Dictionary")
    Set NodeLookup = CreateObject("Scripting.Dictionary")
    Set Rules = CreateObject("Scripting.Dictionary")
    Set InterfaceGroups = CreateObject("Scripting.Dictionary")
    Set Points = New Collection
    Set PointIds = New Collection
End Sub

Public Function NodeId(ByVal p As Variant) As Long
    Dim key As String, i As Long, tol As Double
    tol = ModelScale * 0.00000000001
    For i = 0 To 2
        key = key & "/" & Format$((p(i) - Origin(i)) / tol, "0")
    Next i
    If NodeLookup.Exists(key) Then
        NodeId = NodeLookup(key)
    Else
        NodeId = Vertices.count - 0
        Vertices.Add p
        NodeLookup.Add key, NodeId
    End If
End Function

Public Sub PushElement(ByVal ids As Variant, ByVal boundary As String)
    Dim g As Collection
    If Not Groups.Exists(boundary) Then
        Set g = New Collection
        Groups.Add boundary, g
    Else
        Set g = Groups(boundary)
    End If
    g.Add Elements.count
    Elements.Add ids
End Sub

Public Function FieldCount() As Long
    If ElementOrder = 0 Then
        FieldCount = 1
    ElseIf Dimension = 2 Then
        FieldCount = 3
    Else
        FieldCount = 6
    End If
End Function

Public Function Position(ByVal e As Long, ByVal a As Double, Optional ByVal b As Double = 0) As Variant
    Dim ids As Variant, n As Variant, p As Variant, q As Variant
    Dim c As Double, i As Long, j As Long
    ids = Elements(e + 1): p = Array(0#, 0#, 0#)
    If Dimension = 2 Then
        If ElementOrder = 0 Then
            n = Array((1 - a) / 2, (1 + a) / 2)
        Else
            n = Array(a * (a - 1) / 2, 1 - a * a, a * (a + 1) / 2)
        End If
    Else
        c = 1 - a - b
        If ElementOrder = 0 Then
            n = Array(c, a, b)
        Else
            n = Array(c * (2 * c - 1), a * (2 * a - 1), b * (2 * b - 1), 4 * c * a, 4 * a * b, 4 * b * c)
        End If
    End If
    For i = 0 To UBound(ids)
        q = Vertices(ids(i) + 1)
        For j = 0 To 2: p(j) = p(j) + n(i) * q(j): Next j
    Next i
    Position = p
End Function

Public Function CollocationPoint(ByVal e As Long, ByVal fieldIndex As Long) As Variant
    Dim s As Variant, a As Double, b As Double
    If Dimension = 2 Then
        If ElementOrder = 0 Then a = 0 Else a = (fieldIndex - 1) * 2 / 3
    Else
        a = 1 / 3: b = 1 / 3
        If ElementOrder = 2 Then
            s = Array(Array(0#, 0#), Array(1#, 0#), Array(0#, 1#), Array(0.5, 0#), Array(0.5, 0.5), Array(0#, 0.5))
            a = 1 / 3 + 0.65 * (s(fieldIndex)(0) - 1 / 3)
            b = 1 / 3 + 0.65 * (s(fieldIndex)(1) - 1 / 3)
        End If
    End If
    CollocationPoint = Position(e, a, b)
End Function

Public Function Differential(ByVal e As Long, ByVal a As Double, ByVal b As Double, ByRef jac As Double) As Variant
    Dim ids As Variant, da As Variant, db As Variant, u As Variant, v As Variant
    Dim p As Variant, w As Variant, c As Double, i As Long, j As Long
    ids = Elements(e + 1): u = Array(0#, 0#, 0#): v = Array(0#, 0#, 0#)
    If Dimension = 2 Then
        If ElementOrder = 0 Then da = Array(-0.5, 0.5) Else da = Array(a - 0.5, -2 * a, a + 0.5)
    Else
        c = 1 - a - b
        If ElementOrder = 0 Then
            da = Array(-1#, 1#, 0#): db = Array(-1#, 0#, 1#)
        Else
            da = Array(1 - 4 * c, 4 * a - 1, 0#, 4 * (c - a), 4 * b, -4 * b)
            db = Array(1 - 4 * c, 0#, 4 * b - 1, -4 * a, 4 * a, 4 * (c - b))
        End If
    End If
    For i = 0 To UBound(ids)
        p = Vertices(ids(i) + 1)
        For j = 0 To 2
            u(j) = u(j) + da(i) * p(j)
            If Dimension = 3 Then v(j) = v(j) + db(i) * p(j)
        Next j
    Next i
    If Dimension = 2 Then w = Array(u(1), -u(0), 0#) Else w = Cross3(u, v)
    jac = Norm3(w)
    If jac <= ModelScale ^ (Dimension - 1) * 0.00000000000001 Then Fail RegionName & ": —v‘f‚ª‘Þ‰»‚µ‚Ä‚¢‚Ü‚·B"
    Differential = Mul3(w, 1 / jac)
End Function

Public Function RuleFor(ByVal boundary As String) As Variant
    If Rules.Exists(boundary) Then
        RuleFor = Rules(boundary)
    ElseIf Rules.Exists("ALL") Then
        RuleFor = Rules("ALL")
    Else
        Fail RegionName & "/" & boundary & ": ‹«ŠEðŒ‚ª‚ ‚è‚Ü‚¹‚ñBALL‚Ü‚½‚Í‹«ŠE–¼‚ÅŽw’è‚µ‚Ä‚­‚¾‚³‚¢B"
    End If
End Function

